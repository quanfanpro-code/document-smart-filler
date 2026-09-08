"""严格解析及必要内容检查，保留未知、冲突和来源。"""
import copy
import json
import re
from collections import defaultdict
from datetime import date
from decimal import Decimal, InvalidOperation
from .字段配置 import FIELDS, MONEY_FIELDS, IDENTIFIER_FIELDS


class ResponseError(ValueError):
    """回答无法按约定读取，可在统一次数内重问。"""


class RangeTooLarge(ResponseError):
    """输入或输出过长，需要缩小处理范围。"""


def _unique_pairs(pairs):
    result = {}
    for key, value in pairs:
        if key in result:
            raise ResponseError("回答含重复字段，无法可靠取数")
        result[key] = value
    return result


def parse_response(text, finish_reason="stop"):
    if finish_reason in {"length", "max_tokens", "max_output_tokens"}:
        raise RangeTooLarge("模型输出被截断，需要缩小处理范围")
    if finish_reason != "stop":
        raise ResponseError("模型回答未正常结束")
    if not isinstance(text, str) or not text.strip():
        raise ResponseError("模型未返回可读取的内容")
    decoder = json.JSONDecoder(object_pairs_hook=_unique_pairs,
                               parse_float=Decimal, parse_constant=lambda _: (_ for _ in ()).throw(
                                   ResponseError("回答含非有限数字")))
    candidates = []
    # 仅从外围说明中取完整对象；不补括号、不替换字段内容。
    cursor = 0
    while cursor < len(text):
        start = text.find("{", cursor)
        if start < 0:
            break
        try:
            obj, consumed = decoder.raw_decode(text[start:])
        except ResponseError:
            raise
        except (ValueError, InvalidOperation):
            cursor = start + 1
            continue
        if isinstance(obj, dict) and "records" in obj:
            candidates.append(obj)
        cursor = start + consumed
    if len(candidates) != 1:
        raise ResponseError("回答未包含唯一完整结果对象")
    result = candidates[0]
    if (not isinstance(result.get("records"), list) or type(result.get("complete")) is not bool
            or not _strings(result.get("unprocessed"))):
        raise ResponseError("回答缺少结果、完整性或未处理范围")
    if result["complete"] and result["unprocessed"]:
        raise ResponseError("完整标记与未处理范围矛盾")
    if not result["complete"] and not result["unprocessed"]:
        raise ResponseError("未完整回答必须列出未处理范围")
    for row in result["records"]:
        if (not isinstance(row, dict) or not isinstance(row.get("document_id"), str)
                or not row["document_id"].strip() or not isinstance(row.get("fields"), dict)
                or not _strings(row.get("positions")) or not _strings(row.get("issues"))
                or not isinstance(row.get("missing"), dict)):
            raise ResponseError("结果行结构不完整")
        if any(not isinstance(k, str) or not isinstance(v, str) for k, v in row["missing"].items()):
            raise ResponseError("字段缺失原因格式不正确")
        for key, value in row["fields"].items():
            if not isinstance(key, str) or isinstance(value, bool) or (
                    value is not None and not isinstance(value, (str, int, float, Decimal))):
                raise ResponseError("字段必须是文字、数值或空白")
    return result


def _strings(value):
    return isinstance(value, list) and all(isinstance(item, str) for item in value)


def _empty(value):
    return value is None or isinstance(value, str) and not value.strip()


def _money(value):
    if _empty(value):
        return None
    if isinstance(value, bool):
        raise ValueError
    text = str(value).strip()
    if not re.fullmatch(r"[+-]?(?:\d+|\d{1,3}(?:,\d{3})+)(?:\.\d+)?", text):
        raise ValueError
    number = Decimal(text.replace(",", ""))
    if not number.is_finite():
        raise ValueError
    return number


def _precision(row):
    value = row.get("precision")
    if value is not None:
        try:
            precision = Decimal(str(value))
            if precision.is_finite() and 0 < precision <= 1:
                return precision
        except InvalidOperation:
            pass
    # 原文字符串有小数位才能确认最低精度；整数JSON本身不能证明票面精度。
    decimals = [len(str(v).split(".")[1]) for k, v in row["fields"].items()
                if k in MONEY_FIELDS and isinstance(v, str) and re.fullmatch(r"[+-]?\d+\.\d+", v)]
    return Decimal(10) ** -max(decimals) if decimals else None


def _identity(row, kind):
    fields = row.get("fields", {})
    extra = ()
    if kind == "会计凭证":
        extra = (fields.get("分录序号"),)
    elif kind == "会议纪要":
        extra = (fields.get("议题"), fields.get("待办事项"))
    return (row.get("source_file", ""), row.get("document_id", "")) + extra


def _adjacent_positions(left, right):
    for first in left:
        for second in right:
            pages = [re.match(r"^第(\d+)页", value) for value in (first, second)]
            if all(pages) and abs(int(pages[0][1]) - int(pages[1][1])) <= 1:
                return True
            paragraphs = [re.fullmatch(r"(.*?段落)(\d+)(.*)", value) for value in (first, second)]
            if (all(paragraphs) and paragraphs[0][1] == paragraphs[1][1]
                    and paragraphs[0][3] == paragraphs[1][3]
                    and abs(int(paragraphs[0][2]) - int(paragraphs[1][2])) <= 1):
                return True
    return False


def _merge_evidence(prior, row, kind, complementary):
    left = [value for value in prior.get("positions", []) if value.strip()]
    right = [value for value in row.get("positions", []) if value.strip()]
    if not left or not right or not row.get("document_id"):
        return False
    if kind == "会计凭证" and _empty(row["fields"].get("分录序号")):
        return False
    if set(left) & set(right):
        return True
    # 没有共同来源且字段毫无互补时，邻页也可能是独立副本。
    if not complementary:
        return False
    a, b = prior["fields"], row["fields"]
    identifier = {"发票": "发票号码", "会计凭证": "凭证字号", "合同": "编号"}.get(kind)
    if identifier:
        numbers = [value for value in (a.get(identifier), b.get(identifier)) if not _empty(value)]
        shared_number = len(numbers) == 2 and numbers[0] == numbers[1]
        anchored_number = len(numbers) == 1 and numbers[0] == row.get("document_id")
        if shared_number or anchored_number:
            # 原文编号支持补齐不同页的字段；完全相同的远处副本仍保留核对。
            return True
    if not _adjacent_positions(left, right):
        return False
    if kind == "合同":
        return all(not _empty(a.get(key)) and a.get(key) == b.get(key) for key in ("合同名称", "甲方", "乙方"))
    if kind == "会议纪要":
        return not _empty(a.get("议题")) and a.get("议题") == b.get("议题")
    return False


def _payment_terms(left, right):
    # ponytail: 只拼接有明确期次的完整条款，其他语义差异保留候选供核对。
    if not isinstance(left, str) or not isinstance(right, str):
        return None
    stages = {}
    for line in left.splitlines() + right.splitlines():
        if not line.strip():
            continue
        match = re.fullmatch(r"第(\d+|[一二三四五六七八九十])期[：:](.+)", line.strip())
        if not match:
            return None
        label = match[1]
        number = int(label) if label.isdigit() else "一二三四五六七八九十".index(label) + 1
        if number in stages and stages[number] != line:
            return None
        stages[number] = line
    return "\n".join(stages.values()) if stages else None


def merge_records(records, kind):
    if kind not in FIELDS:
        raise ValueError("请选择四类文档之一")
    merged = []
    for original in records:
        row = copy.deepcopy(original)
        consumed = False
        for prior in merged:
            if _identity(prior, kind) != _identity(row, kind):
                continue
            fields = dict(prior["fields"])
            conflicts = []
            complementary = False
            for key in fields.keys() | row["fields"].keys():
                left, right = fields.get(key), row["fields"].get(key)
                if _empty(left) != _empty(right):
                    complementary = True
                if _empty(left) and not _empty(right):
                    fields[key] = right
                elif not _empty(left) and not _empty(right) and left != right:
                    joined = _payment_terms(left, right) if kind == "合同" and key == "付款方式" else None
                    if joined is None:
                        conflicts.append(key)
                    else:
                        fields[key] = joined
                        complementary = True
            if conflicts:
                reason = "同一单据或事项的来源冲突：" + "、".join(sorted(conflicts))
                for item in (prior, row):
                    item.setdefault("issues", []).append(reason)
                continue
            if not _merge_evidence(prior, row, kind, complementary):
                reason = "相同标识的多条结果缺少足够对应依据，需核对单据或分录边界"
                for item in (prior, row):
                    item.setdefault("issues", []).append(reason)
                continue
            missing = prior.setdefault("missing", {})
            for key, value in fields.items():
                if _empty(prior["fields"].get(key)) and not _empty(value):
                    missing.pop(key, None)
            for key, value in row.get("missing", {}).items():
                if _empty(fields.get(key)) or not _empty(row["fields"].get(key)):
                    # 自身既有字段值又有缺失说明的矛盾不能被合并掩盖。
                    missing[key] = value
            prior["fields"] = fields
            prior["positions"] = list(dict.fromkeys(prior.get("positions", []) + row.get("positions", [])))
            prior["issues"] = list(dict.fromkeys(prior.get("issues", []) + row.get("issues", [])))
            consumed = True
            break
        if not consumed:
            merged.append(row)
    return merged


def validate_records(records, kind, complete=True):
    if kind not in FIELDS:
        raise ValueError("请选择四类文档之一")
    checked = []
    precisions = {}
    mandatory = {"发票": {"票种", "价税合计"}, "会计凭证": {"凭证字号", "日期", "分录序号", "会计科目"},
                 "合同": {"甲方", "乙方"}, "会议纪要": {"议题"}}[kind]
    for original in records:
        row = copy.deepcopy(original)
        fields = row.setdefault("fields", {})
        missing = row.setdefault("missing", {})
        reasons = list(row.get("issues", []))
        precision = _precision(row)
        precisions[id(row)] = precision
        if not row.get("positions") or not all(isinstance(p, str) and p.strip() for p in row["positions"]):
            reasons.append("缺少可回到原文的来源位置")
        if not row.get("document_id"):
            reasons.append("缺少单据或事项序号")
        if not complete or row.get("complete") is False:
            reasons.append("该来源尚未完整处理")
        for key in FIELDS[kind]:
            if key not in fields:
                reasons.append("字段未返回：" + key)
                continue
            value = fields[key]
            if _empty(value):
                fields[key] = None
                reason = missing.get(key, "")
                if key in mandatory or not reason:
                    reasons.append("关键字段缺失：" + key if key in mandatory else "空白未说明原因：" + key)
            if key in MONEY_FIELDS:
                try:
                    fields[key] = _money(value)
                except (ValueError, InvalidOperation):
                    reasons.append("金额不能按原文可靠读取：" + key)
            if key in IDENTIFIER_FIELDS and not _empty(value) and not isinstance(value, str):
                fields[key] = str(value)
                reasons.append("编号返回为数值，无法确认前导零：" + key)
            if key in {"日期", "开票日期", "签订日期", "会议时间", "截止日期"} and isinstance(value, str):
                match = re.fullmatch(r"(\d{4})[-/年](\d{1,2})[-/月](\d{1,2})日?", value.strip())
                if match:
                    try:
                        fields[key] = date(*map(int, match.groups())).isoformat()
                    except ValueError:
                        reasons.append("日期无效：" + key)
        for key, reason in missing.items():
            if not _empty(fields.get(key)):
                reasons.append("字段值与缺失说明矛盾：" + key)
            if not reason.startswith(("原文未记载", "不适用")):
                reasons.append(key + "：" + reason)
        if kind == "发票":
            amounts = [fields.get(k) for k in ("不含税金额", "税额", "价税合计")]
            if all(isinstance(v, Decimal) for v in amounts):
                if precision is None:
                    reasons.append("金额原文精度不明")
                elif abs(amounts[0] + amounts[1] - amounts[2]) > min(precision, Decimal("0.01")):
                    reasons.append("不含税金额加税额与价税合计不符")
            if isinstance(fields.get("价税合计"), Decimal) and (
                    _empty(fields.get("币种")) or _empty(fields.get("金额单位"))):
                reasons.append("票据币种或金额单位未确认")
        if kind == "会议纪要":
            lines = {key: sum(bool(line.strip()) for line in str(fields.get(key) or "").splitlines())
                     for key in ("待办事项", "责任人", "截止日期")}
            if lines["截止日期"] > 1 or lines["责任人"] > 1 and lines["待办事项"] > 1:
                reasons.append("多项待办被合在一行，需分别对应责任人和截止日期")
        if kind == "合同" and not _empty(fields.get("金额")):
            if any(_empty(fields.get(k)) for k in ("币种", "金额单位", "含税口径")):
                reasons.append("合同金额的币种、单位或含税口径未确认")
        row["reasons"] = reasons
        checked.append(row)
    if kind == "会计凭证":
        groups = defaultdict(list)
        for row in checked:
            groups[(row.get("source_file", ""), row.get("document_id"))].append(row)
        for group in groups.values():
            reasons = [reason for row in group for reason in row["reasons"]]
            for key in ("核算单位", "凭证字号", "日期", "币种", "金额单位"):
                values = {str(row["fields"].get(key)) for row in group}
                if len(values) != 1:
                    reasons.append("凭证公共信息冲突：" + key)
                if key in {"币种", "金额单位"} and any(_empty(row["fields"].get(key)) for row in group):
                    reasons.append("凭证币种或金额单位未确认")
            sequences = [str(row["fields"].get("分录序号", "")) for row in group]
            if (not all(s.isdigit() for s in sequences) or len(set(sequences)) != len(sequences)
                    or sorted(map(int, sequences)) != list(range(1, len(group) + 1))):
                reasons.append("凭证分录序号缺失、重复或不连续")
            for row in group:
                debit, credit = row["fields"].get("借方金额"), row["fields"].get("贷方金额")
                if debit is None and credit is None or debit is not None and credit is not None:
                    reasons.append("分录借贷方向不明确")
                if not all(v is None or isinstance(v, Decimal) for v in (debit, credit)):
                    reasons.append("分录金额无法可靠计算")
            units = [precisions[id(row)] for row in group]
            if not all(units):
                reasons.append("凭证金额原文精度不明")
            if not reasons:
                debit = sum((row["fields"].get("借方金额") or Decimal(0) for row in group), Decimal(0))
                credit = sum((row["fields"].get("贷方金额") or Decimal(0) for row in group), Decimal(0))
                if abs(debit - credit) > min(min(units), Decimal("0.01")):
                    reasons.append("凭证借贷合计不平衡")
            for row in group:
                row["reasons"] = list(reasons)
    if kind == "发票":
        numbers = defaultdict(list)
        for row in checked:
            number = row["fields"].get("发票号码")
            if not _empty(number):
                numbers[str(number)].append(row)
        for group in numbers.values():
            if len(group) > 1:
                for row in group:
                    row["reasons"].append("票号相同的多条结果，需核对是否重复票据")
    for row in checked:
        row["reasons"] = list(dict.fromkeys(row["reasons"]))
        row["status"] = "review" if row["reasons"] else "normal"
    return checked
