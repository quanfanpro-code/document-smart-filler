"""只读提取 Word 内容、页面图像和来源位置，并按请求预算切段。"""

import base64
import hashlib
import io
import shutil
import tempfile
import uuid
from pathlib import Path

import fitz
from docx import Document
from docx.oxml.ns import qn
from PIL import Image, ImageOps


def _warning(position, reason):
    return {"kind": "warning", "position": position, "text": reason}


def _image_blocks(picture, position):
    """大图按原像素分区域，邻接区域重叠，避免缩小后丢失小字。"""
    picture = ImageOps.exif_transpose(picture)
    if picture.mode not in ("RGB", "RGBA"):
        picture = picture.convert("RGB")
    edge, overlap = 2400, 80
    divided = picture.width > edge or picture.height > edge
    blocks = []
    for top in range(0, max(1, picture.height - overlap), edge - overlap):
        for left in range(0, max(1, picture.width - overlap), edge - overlap):
            right = min(left + edge, picture.width)
            bottom = min(top + edge, picture.height)
            crop = picture.crop((left, top, right, bottom))
            output = io.BytesIO()
            crop.save(output, format="PNG")
            location = position
            if divided:
                location += f"（区域{len(blocks) + 1}，像素范围{left + 1},{top + 1}至{right},{bottom}）"
            blocks.append({"kind": "image", "position": location, "text": "",
                           "image": "data:image/png;base64," + base64.b64encode(output.getvalue()).decode("ascii")})
    return blocks


def _content_children(node):
    # 同一区域的新旧兼容表示只能采用一份，否则正文和图片会重复。
    if node.tag.rsplit("}", 1)[-1] == "AlternateContent":
        choices = [child for child in node if child.tag.rsplit("}", 1)[-1] == "Choice"]
        return choices[:1] or [child for child in node if child.tag.rsplit("}", 1)[-1] == "Fallback"][:1]
    return node


def _read_docx(path):
    document = Document(str(path))
    blocks = []
    image_sources = {}

    def read_picture(node, part, position):
        relationship_id = node.get(qn("r:embed")) or node.get(qn("r:id"))
        if not relationship_id:
            blocks.append(_warning(position, "无法读取外部链接图片，原文该区域需人工核对"))
            return
        try:
            blob = part.related_parts[relationship_id].blob
            fingerprint = hashlib.sha256(blob).hexdigest()
            if fingerprint in image_sources:
                blocks.append({"kind": "text", "position": position,
                               "text": f"重复图片，内容见{image_sources[fingerprint]}"})
                return
            with Image.open(io.BytesIO(blob)) as picture:
                image_blocks = _image_blocks(picture, position)
            blocks.extend(image_blocks)
            image_sources[fingerprint] = position
        except Exception as exc:
            blocks.append(_warning(position, f"无法读取内嵌图片：{exc}"))

    def read_paragraph(paragraph, part, position):
        text_parts = []
        image_number = 0

        def flush():
            if text_parts:
                blocks.append({"kind": "text", "position": position, "text": "".join(text_parts)})
                text_parts.clear()

        def visit(node):
            nonlocal image_number
            tag = node.tag.rsplit("}", 1)[-1]
            if tag == "t":
                text_parts.append(node.text or "")
                return
            if tag in ("tab", "br", "cr", "noBreakHyphen"):
                text_parts.append({"tab": "\t", "br": "\n", "cr": "\n", "noBreakHyphen": "‑"}[tag])
                return
            if tag in ("blip", "imagedata"):
                flush()
                image_number += 1
                read_picture(node, part, f"{position}图片{image_number}")
                return
            unsupported = {"object": "嵌入对象", "altChunk": "外部嵌入内容", "chart": "图表",
                           "relIds": "智能图形", "footnoteReference": "脚注", "endnoteReference": "尾注",
                           "commentReference": "批注", "sym": "特殊字体符号"}
            if tag in unsupported:
                flush()
                blocks.append(_warning(position, f"无法完整读取{unsupported[tag]}，该区域需人工核对"))
                return
            if tag in ("ins", "del", "moveFrom", "moveTo"):
                flush()
                blocks.append(_warning(position, "存在修订内容，需在原文确认最终采用的版本"))
                if tag in ("del", "moveFrom"):
                    return
            if tag == "drawing" and not any(child.tag.rsplit("}", 1)[-1] in ("blip", "chart", "relIds") for child in node.iter()):
                flush()
                blocks.append(_warning(position, "无法完整还原绘图区域，已保留其中可读取的文字"))
            for child in _content_children(node):
                visit(child)

        visit(paragraph)
        flush()

    def read_container(container, part, prefix=""):
        paragraph_number = 0
        table_number = 0

        def visit(node):
            nonlocal paragraph_number, table_number
            tag = node.tag.rsplit("}", 1)[-1]
            if tag == "p":
                paragraph_number += 1
                read_paragraph(node, part, f"{prefix}段落{paragraph_number}")
            elif tag == "tbl":
                table_number += 1
                table_position = f"{prefix}表格{table_number}"
                rows = (row for row in node.iter(qn("w:tr"))
                        if next(row.iterancestors(qn("w:tbl")), None) is node)
                for row_number, row in enumerate(rows, 1):
                    column = 1
                    cells = (cell for cell in row.iter(qn("w:tc"))
                             if next(cell.iterancestors(qn("w:tr")), None) is row)
                    for cell in cells:
                        read_container(cell, part, f"{table_position}第{row_number}行第{column}列")
                        span = cell.find("./" + qn("w:tcPr") + "/" + qn("w:gridSpan"))
                        column += int(span.get(qn("w:val"), "1")) if span is not None else 1
            elif tag == "altChunk":
                blocks.append(_warning(prefix or "正文", "无法读取外部嵌入内容，该区域需人工核对"))
            elif tag not in ("sectPr", "tcPr", "tblPr", "tblGrid", "trPr"):
                if tag in ("ins", "del", "moveFrom", "moveTo"):
                    blocks.append(_warning(prefix or "正文", "存在修订区域，需确认最终采用的版本"))
                    if tag in ("del", "moveFrom"):
                        return
                for child in _content_children(node):
                    visit(child)

        for child in _content_children(container):
            visit(child)

    read_container(document.element.body, document.part)
    seen_parts = set()
    for section_number, section in enumerate(document.sections, 1):
        for reference in section._sectPr:
            name = reference.tag.rsplit("}", 1)[-1]
            if name not in ("headerReference", "footerReference"):
                continue
            label = "页眉" if name == "headerReference" else "页脚"
            page_type = {"default": "", "first": "首页", "even": "偶数页"}.get(reference.get(qn("w:type")), "")
            position = f"第{section_number}节{page_type}{label}"
            try:
                part = document.part.related_parts[reference.get(qn("r:id"))]
                if part.partname not in seen_parts:
                    read_container(part.element, part, position)
                    seen_parts.add(part.partname)
            except Exception as exc:
                blocks.append(_warning(position, f"无法读取{label}：{exc}"))
    return blocks or [_warning("正文", "未读取到正文或图像，请检查文件是否为空或包含不支持的对象")]


def _convert_doc(path, work_dir):
    """Word 仅打开工作副本，转换结果始终放入独立的新目录。"""
    import pythoncom
    import win32com.client

    work_dir.mkdir(parents=True, exist_ok=True)
    conversion_dir = Path(tempfile.mkdtemp(prefix="Word转换_", dir=str(work_dir)))
    source_copy = conversion_dir / path.name
    shutil.copy2(path, source_copy)
    converted = conversion_dir / (path.stem + ".docx")
    word = None
    opened = None
    pythoncom.CoInitialize()
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        word.AutomationSecurity = 3
        # 随机口令让加密文档直接返回错误，不弹出隐藏的密码窗口。
        opened = word.Documents.Open(str(source_copy), ReadOnly=True, AddToRecentFiles=False,
                                     ConfirmConversions=False, Visible=False,
                                     PasswordDocument=uuid.uuid4().hex,
                                     PasswordTemplate=uuid.uuid4().hex, NoEncodingDialog=True)
        if converted.exists():
            raise ValueError("转换目标已存在，已停止以避免覆盖")
        opened.SaveAs2(str(converted), FileFormat=16, AddToRecentFiles=False)
    finally:
        try:
            if opened is not None:
                opened.Close(SaveChanges=False)
        finally:
            try:
                if word is not None:
                    word.Quit(SaveChanges=False)
            finally:
                pythoncom.CoUninitialize()
    if not converted.is_file():
        raise ValueError("Word 未生成转换文件")
    return converted


def read_document(path: Path, work_dir: Path) -> list[dict]:
    """读取 DOCX、DOC、PDF、PNG、JPEG、TIFF；错误带中文原因。"""
    path = Path(path).resolve()
    work_dir = Path(work_dir).resolve()
    try:
        if not path.is_file():
            raise ValueError("文件不存在或不是普通文件")
        suffix = path.suffix.lower()
        if suffix == ".doc":
            return _read_docx(_convert_doc(path, work_dir))
        if suffix == ".docx":
            return _read_docx(path)
        blocks = []
        if suffix == ".pdf":
            with fitz.open(path) as document:
                if document.needs_pass:
                    raise ValueError("PDF 已加密，需要先提供可读取的版本")
                if not document.page_count:
                    raise ValueError("PDF 没有可读取页面")
                for index in range(document.page_count):
                    position = f"第{index + 1}页"
                    try:
                        page = document.load_page(index)
                        pixels = page.get_pixmap(dpi=160, colorspace=fitz.csRGB, alpha=False)
                        picture = Image.frombytes("RGB", (pixels.width, pixels.height), pixels.samples)
                        blocks.extend(_image_blocks(picture, position))
                    except Exception as exc:
                        blocks.append(_warning(position, f"无法渲染该页：{exc}"))
        elif suffix in (".png", ".jpg", ".jpeg", ".tif", ".tiff"):
            with Image.open(path) as picture:
                for index in range(getattr(picture, "n_frames", 1)):
                    position = f"第{index + 1}页"
                    try:
                        picture.seek(index)
                        blocks.extend(_image_blocks(picture.copy(), position))
                    except Exception as exc:
                        blocks.append(_warning(position, f"无法读取该图像页：{exc}"))
        else:
            raise ValueError("不支持此格式，请选择 DOCX、DOC、PDF、PNG、JPEG 或 TIFF")
        if not any(block["kind"] == "image" for block in blocks):
            raise ValueError("没有成功读取任何页面；" + "；".join(block["text"] for block in blocks))
        return blocks
    except Exception as exc:
        raise ValueError(f"无法读取“{path.name}”：{exc}") from exc


def chunk_blocks(blocks: list[dict], max_chars: int = 18000, max_images: int = 4) -> list[list[dict]]:
    """按文字量和图片数切段；长单块保留重叠，相邻页尽量同段。"""
    if not isinstance(max_chars, int) or not isinstance(max_images, int) or max_chars < 1 or max_images < 1:
        raise ValueError("文字和图片预算必须是正整数")
    overlap = min(500, max_chars // 10)
    pieces = []
    for block in blocks:
        if block.get("kind") not in ("text", "image", "warning") or not isinstance(block.get("text"), str):
            raise ValueError("读取块必须包含有效类型和文字")
        text = block["text"]
        if len(text) <= max_chars:
            pieces.append(dict(block))
            continue
        start = 0
        while start < len(text):
            piece = dict(block)
            piece["text"] = text[start:start + max_chars]
            if start and piece["kind"] == "image":
                piece["kind"] = "text"
                piece.pop("image", None)
            pieces.append(piece)
            if start + max_chars >= len(text):
                break
            start += max_chars - overlap

    chunks = []
    current = []

    def fits(candidate, extra):
        together = candidate + [extra]
        return (sum(len(block["text"]) for block in together) <= max_chars
                and sum(block["kind"] == "image" for block in together) <= max_images)

    for piece in pieces:
        if current and not fits(current, piece):
            chunks.append(current)
            previous = dict(current[-1])
            current = []
            if previous["kind"] == "image":
                if max_images > 1 and fits([previous], piece):
                    current.append(previous)
            elif overlap and previous["text"]:
                previous["text"] = previous["text"][-overlap:]
                if fits([previous], piece):
                    current.append(previous)
        current.append(piece)
    if current:
        chunks.append(current)
    return chunks