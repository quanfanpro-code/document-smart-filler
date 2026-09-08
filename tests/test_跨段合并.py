"""用明确编号和人工确定的分期条款检查跨段合并，不推断单据身份。"""
import copy

from src.字段配置 import FIELDS
from src.结果校验 import merge_records, validate_records


def row(kind, identity, position, values):
    fields = {name: None for name in FIELDS[kind]}
    fields.update(values)
    return {"document_id": identity, "source_file": "合成材料.pdf", "positions": [position],
            "fields": fields, "missing": {key: "原文未记载" for key, value in fields.items() if value is None},
            "issues": [], "precision": "0.01"}


def test_invoice_complementary_pages_become_one_complete_record():
    first = row("发票", "票号000123", "第1页", {"票种": "增值税普通发票", "发票号码": "000123",
                "开票日期": "2026-09-08", "销售方": "甲公司", "购买方": "乙公司"})
    second = row("发票", "票号000123", "第2页", {"发票号码": "000123", "不含税金额": "100.00",
                 "税额": "13.00", "价税合计": "113.00", "币种": "人民币", "金额单位": "元"})
    original = copy.deepcopy([first, second])
    merged = merge_records([first, second], "发票")
    assert len(merged) == 1
    assert merged[0]["fields"]["销售方"] == "甲公司"
    assert merged[0]["fields"]["价税合计"] == "113.00"
    assert merged[0]["positions"] == ["第1页", "第2页"]
    assert "价税合计" not in merged[0]["missing"] and "销售方" not in merged[0]["missing"]
    assert validate_records(merged, "发票")[0]["status"] == "normal"
    assert [first, second] == original


def test_voucher_same_entry_in_overlapping_page_is_not_a_new_entry():
    common = {"核算单位": "甲公司", "凭证字号": "记0012", "日期": "2026-09-08",
              "币种": "人民币", "金额单位": "元", "摘要": "采购", "附件张数": "1"}
    debit = row("会计凭证", "记0012", "第1页", {**common, "分录序号": "1", "会计科目": "原材料", "借方金额": "100.00"})
    debit["positions"] = ["第1页", "第2页"]
    repeated = copy.deepcopy(debit)
    repeated["positions"] = ["第2页"]
    credit = row("会计凭证", "记0012", "第2页", {**common, "分录序号": "2", "会计科目": "应付账款", "贷方金额": "100.00"})
    for item in (debit, repeated):
        item["missing"]["贷方金额"] = "不适用"
    credit["missing"]["借方金额"] = "不适用"
    merged = merge_records([debit, repeated, credit], "会计凭证")
    assert len(merged) == 2
    assert [item["fields"]["分录序号"] for item in merged] == ["1", "2"]
    assert [item["status"] for item in validate_records(merged, "会计凭证")] == ["normal", "normal"]


def test_voucher_complementary_entry_keeps_its_own_amount():
    common = {"凭证字号": "记0012", "分录序号": "1", "核算单位": "甲公司", "日期": "2026-09-08"}
    first = row("会计凭证", "记0012", "第1页", {**common, "会计科目": "原材料"})
    second = row("会计凭证", "记0012", "第2页", {**common, "借方金额": "100.00"})
    other = row("会计凭证", "记0012", "第2页", {**common, "分录序号": "2", "借方金额": "13.00"})
    merged = merge_records([first, second, other], "会计凭证")
    assert len(merged) == 2
    assert merged[0]["fields"]["会计科目"] == "原材料"
    assert merged[0]["fields"]["借方金额"] == "100.00"
    assert merged[1]["fields"]["借方金额"] == "13.00"


def test_meeting_todo_complements_owner_and_deadline_without_reassigning_another_todo():
    common = {"会议时间": "2026-09-08", "议题": "预算调整", "待办事项": "提交预算清单"}
    first = row("会议纪要", "第1页事项1", "第1页", {**common, "责任人": "张某"})
    second = row("会议纪要", "第1页事项1", "第2页", {**common, "截止日期": "2026-09-15"})
    another = row("会议纪要", "第1页事项1", "第2页", {**common, "待办事项": "核对余额", "责任人": "李某"})
    merged = merge_records([first, second, another], "会议纪要")
    assert len(merged) == 2
    assert merged[0]["fields"]["责任人"] == "张某" and merged[0]["fields"]["截止日期"] == "2026-09-15"
    assert merged[1]["fields"]["责任人"] == "李某" and merged[1]["fields"]["截止日期"] is None
    assert validate_records(merged, "会议纪要")[0]["status"] == "normal"


def test_contract_distinct_payment_stages_keep_complete_lines_in_order():
    common = {"编号": "合同008", "甲方": "甲公司", "乙方": "乙公司", "金额": "1000.00",
              "币种": "人民币", "金额单位": "元", "含税口径": "含税"}
    first_clause = "第1期：合同签订后5日支付30%，金额300.00元。"
    last_clause = "第2期：验收通过后10日支付70%，金额700.00元。"
    first = row("合同", "合同008", "第1页", {**common, "付款方式": first_clause})
    second = row("合同", "合同008", "第4页", {**common, "付款方式": last_clause})
    merged = merge_records([first, second], "合同")
    assert len(merged) == 1
    assert merged[0]["fields"]["付款方式"] == first_clause + "\n" + last_clause
    assert merged[0]["fields"]["金额"] == "1000.00"
    assert validate_records(merged, "合同")[0]["status"] == "normal"


def test_same_stage_payment_conflict_keeps_both_original_candidates():
    common = {"编号": "合同008", "甲方": "甲公司", "乙方": "乙公司"}
    first = row("合同", "合同008", "第1页", {**common, "付款方式": "第一期：签订后支付30%。"})
    second = row("合同", "合同008", "第2页", {**common, "付款方式": "第1期：签订后支付40%。"})
    merged = merge_records([first, second], "合同")
    assert len(merged) == 2
    assert [item["fields"]["付款方式"] for item in merged] == ["第一期：签订后支付30%。", "第1期：签订后支付40%。"]
    assert all(any("付款方式" in issue for issue in item["issues"]) for item in merged)


def test_same_page_identical_text_with_different_document_ids_is_two_documents():
    fields = {"票种": "收据", "发票号码": "000123", "价税合计": "50.00"}
    first = row("发票", "第1页左票", "第1页", fields)
    second = row("发票", "第1页右票", "第1页", fields)
    assert len(merge_records([first, second], "发票")) == 2
    second["document_id"] = first["document_id"]
    second["source_file"] = "另一份合成材料.pdf"
    assert len(merge_records([first, second], "发票")) == 2


def test_no_invoice_number_does_not_turn_generic_repeated_label_into_identity_evidence():
    first = row("发票", "票1", "第1页", {"票种": "收据", "销售方": "甲公司"})
    second = row("发票", "票1", "第2页", {"票种": "收据", "价税合计": "50.00"})
    merged = merge_records([first, second], "发票")
    assert len(merged) == 2
    assert all(item["issues"] for item in merged)


def test_missing_voucher_sequence_does_not_swallow_two_equal_entries():
    first = row("会计凭证", "记0012", "第1页", {"凭证字号": "记0012", "会计科目": "原材料", "借方金额": "100.00"})
    second = copy.deepcopy(first)
    assert len(merge_records([first, second], "会计凭证")) == 2


def test_invoice_amount_conflicts_remain_separate_for_review():
    first = row("发票", "票号000123", "第1页", {"发票号码": "000123", "价税合计": "113.00"})
    second = row("发票", "票号000123", "第2页", {"发票号码": "000123", "价税合计": "114.00"})
    merged = merge_records([first, second], "发票")
    assert len(merged) == 2
    assert [item["fields"]["价税合计"] for item in merged] == ["113.00", "114.00"]
    assert all(item["issues"] for item in merged)

def test_meeting_multiple_todos_in_one_row_need_review_without_guessing_pairs():
    cases = [
        {"议题": "采购安排", "待办事项": "李四完成询价\n王五完成采购", "责任人": "李四\n王五",
         "截止日期": "2026年9月10日\n2026年9月12日"},
        {"议题": "采购安排", "待办事项": "完成询价\n完成采购", "责任人": "李四\n王五", "截止日期": None},
        {"议题": "采购安排", "待办事项": "完成采购", "责任人": "李四", "截止日期": "2026年9月10日\n2026年9月12日"},
    ]
    for values in cases:
        source = row("会议纪要", "第1页事项1", "第1页", values)
        checked = validate_records([source], "会议纪要")
        assert len(checked) == 1 and checked[0]["status"] == "review"
        assert "多项待办被合在一行，需分别对应责任人和截止日期" in checked[0]["reasons"]
        assert checked[0]["fields"]["待办事项"] == values["待办事项"]
        assert checked[0]["fields"]["责任人"] == values["责任人"]
        assert checked[0]["fields"]["截止日期"] == values["截止日期"]

def test_original_number_as_document_id_links_a_continuation_without_repeated_number_field():
    cases = [
        ("发票", "000123", {"发票号码": "000123", "销售方": "甲公司"}, {"价税合计": "113.00"}),
        ("会计凭证", "记0012", {"凭证字号": "记0012", "分录序号": "1", "会计科目": "原材料"}, {"分录序号": "1", "借方金额": "100.00"}),
        ("合同", "合同008", {"编号": "合同008", "甲方": "甲公司"}, {"乙方": "乙公司"}),
    ]
    for kind, identity, first_fields, later_fields in cases:
        first = row(kind, identity, "第1页", first_fields)
        later = row(kind, identity, "第2页", later_fields)
        merged = merge_records([first, later], kind)
        assert len(merged) == 1
        for key, value in {**first_fields, **later_fields}.items():
            assert merged[0]["fields"][key] == value

def test_相邻页同号相同字段不能仅凭页码吞掉独立副本():
    first=row('发票','000123','第1页',{'票种':'收据','发票号码':'000123','价税合计':'50.00','币种':'人民币','金额单位':'元'})
    second=copy.deepcopy(first);second['positions']=['第2页']
    merged=merge_records([first,second],'发票')
    assert len(merged)==2
    assert all(item['status']=='review' for item in validate_records(merged,'发票'))
