"""真实接口请求和业务结果校验。"""
import copy
import json
import threading
import time
from concurrent.futures import ThreadPoolExecutor
from contextlib import contextmanager
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
import pytest

def record(fields, doc="票1", positions=None, missing=None):
    return {"document_id": doc, "fields": fields, "positions": positions or ["第1页"],
            "issues": [], "missing": missing or {}, "precision": "0.01"}

def invoice(total="113.00", doc="票1", number="000123"):
    return record({"票种": "增值税普通发票", "发票号码": number, "开票日期": "2026年9月8日",
        "销售方": "销售公司", "销售方纳税人识别号": "001234567890123456",
        "购买方": "购买公司", "购买方纳税人识别号": "009876543210987654",
        "不含税金额": "100.00", "税额": "13.00", "价税合计": total,
        "币种": "人民币", "金额单位": "元", "商品/服务明细": "办公用品"}, doc)

def voucher(seq, debit, credit):
    return record({"核算单位": "测试公司", "凭证字号": "记0012", "日期": "2026-09-08",
        "分录序号": str(seq), "摘要": "采购", "会计科目": ["原材料", "应交税费", "应付账款"][seq-1],
        "借方金额": debit, "贷方金额": credit, "币种": "人民币", "金额单位": "元",
        "附件张数": "1"}, "凭证1", missing={"贷方金额" if credit is None else "借方金额": "不适用"})

def envelope(records=None, complete=True):
    return {"records": records if records is not None else [invoice()], "complete": complete,
            "unprocessed": [] if complete else ["第2页"]}

@contextmanager
def endpoint(responses=None, delay=0, streamed=False):
    state = {"requests": [], "active": 0, "peak": 0}
    lock = threading.Lock()
    responses = list(responses or [(200, envelope())])
    class Handler(BaseHTTPRequestHandler):
        def log_message(self, *args):
            pass
        def do_POST(self):
            body = json.loads(self.rfile.read(int(self.headers["Content-Length"])))
            with lock:
                index = len(state["requests"])
                state["requests"].append((self.path, dict(self.headers), body))
                state["active"] += 1
                state["peak"] = max(state["peak"], state["active"])
            try:
                time.sleep(delay)
                status, result = responses[min(index, len(responses)-1)]
                if status == 200:
                    reply = {"id": "controlled", "object": "chat.completion", "created": 1,
                        "model": body["model"], "choices": [{"index": 0,
                        "message": {"role": "assistant", "content": result if isinstance(result, str)
                                    else json.dumps(result, ensure_ascii=False)},
                        "finish_reason": "length" if result == "TRUNCATED" else "stop"}],
                        "usage": {"prompt_tokens": 15, "completion_tokens": 30, "total_tokens": 45}}
                else:
                    reply = {"error": {"message": result if isinstance(result, str) else "controlled error", "type": "invalid_request_error"}}
                data = json.dumps(reply, ensure_ascii=False).encode()
                if streamed and status == 200:
                    answer = reply["choices"][0]["message"]["content"]
                    frames = [{"choices": [{"delta": {"content": answer[:30]}, "finish_reason": None}]},
                              {"choices": [{"delta": {"content": answer[30:]}, "finish_reason": "stop"}]},
                              {"choices": [], "usage": reply["usage"]}]
                    data = ("".join("data: " + json.dumps(frame, ensure_ascii=False) + "\n\n"
                                    for frame in frames) + "data: [DONE]\n\n").encode()
                self.send_response(status)
                self.send_header("Content-Type", "text/event-stream" if streamed else "application/json")
                self.send_header("Content-Length", str(len(data)))
                self.end_headers()
                self.wfile.write(data)
            finally:
                with lock:
                    state["active"] -= 1
    server = ThreadingHTTPServer(("127.0.0.1", 0), Handler)
    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()
    try:
        yield {"base_url": f"http://127.0.0.1:{server.server_port}/v1",
               "api_key": "controlled-key", "model": "controlled-model",
               "retry_delay": 0, "read_timeout": 2, "request_timeout": 5}, state
    finally:
        server.shutdown()
        server.server_close()
        thread.join(timeout=2)

def test_strict_json_keeps_complete_data_and_rejects_truncation():
    from src.结果校验 import parse_response, ResponseError, RangeTooLarge
    text = json.dumps(envelope(), ensure_ascii=False)
    assert parse_response("说明\n```json\n" + text + "\n```")["records"][0]["fields"]["发票号码"] == "000123"
    with pytest.raises(RangeTooLarge):
        parse_response(text, "length")
    with pytest.raises(ResponseError):
        parse_response(text[:-2])
    with pytest.raises(ResponseError):
        parse_response('{"records": [], "complete": "true", "unprocessed": []}')
    bad = envelope()
    bad["records"][0]["fields"]["价税合计"] = {"amount": 113}
    with pytest.raises(ResponseError):
        parse_response(json.dumps(bad))
    with pytest.raises(ResponseError):
        parse_response('{"records": [], "records": [], "complete": true, "unprocessed": []}')

def test_money_uses_decimal_and_preserves_numbers_and_blanks():
    from src.结果校验 import validate_records
    from decimal import Decimal
    result = validate_records([invoice()], "发票")[0]
    assert result["status"] == "normal"
    assert result["fields"]["价税合计"] == Decimal("113.00")
    assert result["fields"]["发票号码"] == "000123"
    assert result["fields"]["开票日期"] == "2026-09-08"
    assert validate_records([invoice("113.02")], "发票")[0]["status"] == "review"
    negative = invoice("-113.00")
    negative["fields"].update({"不含税金额": "-100.00", "税额": "-13.00"})
    assert validate_records([negative], "发票")[0]["status"] == "normal"
    zero = invoice("0.00")
    zero["fields"].update({"不含税金额": "0.00", "税额": "0.00"})
    assert validate_records([zero], "发票")[0]["status"] == "normal"
    blank = invoice(None)
    blank["missing"]["价税合计"] = "无法辨认"
    checked = validate_records([blank], "发票")[0]
    assert checked["fields"]["价税合计"] is None and checked["status"] == "review"

def test_voucher_groups_all_entries_and_never_balances_incomplete_input():
    from src.结果校验 import validate_records
    rows = [voucher(1, "100.00", None), voucher(2, "13.00", None), voucher(3, None, "113.00")]
    assert [r["status"] for r in validate_records(rows, "会计凭证")] == ["normal"] * 3
    rows[2]["fields"]["贷方金额"] = "112.00"
    assert [r["status"] for r in validate_records(rows, "会计凭证")] == ["review"] * 3
    rows[2]["fields"]["贷方金额"] = "113.00"
    assert [r["status"] for r in validate_records(rows, "会计凭证", complete=False)] == ["review"] * 3
    rows[2]["fields"]["币种"] = "美元"
    assert [r["status"] for r in validate_records(rows, "会计凭证")] == ["review"] * 3

def test_missing_reason_and_no_todo_agenda_are_retained():
    from src.结果校验 import validate_records
    fields = {"会议时间": "2026-09-08", "地点": None, "主持人": None, "参会人员": None,
              "议题": "预算讨论", "决议事项": "下次继续讨论", "待办事项": None, "责任人": None, "截止日期": None}
    row = record(fields, "议题1", missing={k: "原文未记载" for k, v in fields.items() if v is None})
    assert validate_records([row], "会议纪要")[0]["status"] == "normal"
    row["missing"]["责任人"] = "无法辨认"
    assert validate_records([row], "会议纪要")[0]["status"] == "review"

def test_merge_only_overlap_duplicates_and_preserve_conflicting_candidates():
    from src.结果校验 import merge_records, validate_records
    first = invoice()
    overlap = copy.deepcopy(first)
    overlap["positions"].append("第2页")
    assert len(merge_records([first, overlap, invoice(doc="票2", number="000124")], "发票")) == 2
    separated = copy.deepcopy(first)
    separated["positions"] = ["第5页"]
    assert len(merge_records([first, separated], "发票")) == 2
    merged = merge_records([first, invoice("114.00")], "发票")
    assert len(merged) == 2
    assert all(r["status"] == "review" for r in validate_records(merged, "发票"))
    assert len(merge_records([voucher(1, "100.00", None), voucher(2, "100.00", None)], "会计凭证")) == 2

def test_real_http_uses_settings_images_thinking_and_keeps_raw_attempts():
    from src.模型调用 import Client
    with endpoint() as (settings, state):
        client = Client(settings)
        picture = "data:image/png;base64,YQ=="
        result = client.extract([{"kind": "image", "position": "第2页", "image": picture}],
                                "发票", fields=["合同价"], instructions="只提取原文")
        path, headers, sent = state["requests"][0]
        assert path == "/v1/chat/completions"
        assert headers["Authorization"] == "Bearer controlled-key"
        assert sent["model"] == "controlled-model"
        assert sent["chat_template_kwargs"]["enable_thinking"] is False
        assert picture in json.dumps(sent)
        assert "发票号码" in sent["messages"][0]["content"] and "合同价" in sent["messages"][0]["content"]
        assert result["usage"]["total_tokens"] == 45
        assert len(client.attempts) == 1 and "000123" in client.attempts[0]["raw"]
        client.extract([{"kind": "text", "position": "段落1", "text": "测试"}], "合同")
        sent = state["requests"][1][2]
        assert sent["chat_template_kwargs"]["enable_thinking"] is True
        assert sent["reasoning_effort"] == "medium"

def test_configuration_error_stops_and_retry_budget_is_shared():
    from src.模型调用 import Client, ConfigurationError, ModelError
    with endpoint([(401, {})]) as (settings, state):
        with pytest.raises(ConfigurationError) as caught:
            Client(settings).extract([], "发票")
        assert len(state["requests"]) == 1
        assert settings["api_key"] not in str(caught.value) and settings["base_url"] not in str(caught.value)
    with endpoint([(429, {}), (200, "{broken"), (200, envelope())]) as (settings, state):
        client = Client(settings)
        assert client.extract([], "发票")["complete"] is True
        assert len(state["requests"]) == 3 and len(client.attempts) == 3
    with endpoint([(429, {})]) as (settings, state):
        with pytest.raises(ModelError):
            Client(settings).extract([], "发票")
        assert len(state["requests"]) == 3

def test_output_truncation_does_not_repeat_same_request():
    from src.模型调用 import Client, RangeTooLarge
    with endpoint([(200, "TRUNCATED")]) as (settings, state):
        with pytest.raises(RangeTooLarge):
            Client(settings).extract([], "发票")
        assert len(state["requests"]) == 1

def test_two_slots_are_global_and_cancel_does_not_release_inflight_request():
    from src.模型调用 import Client, Cancelled
    with endpoint(delay=0.35) as (settings, state):
        cancel = threading.Event()
        def run_cancelled():
            with pytest.raises(Cancelled):
                Client(settings).extract([], "发票", cancel_event=cancel)
        with ThreadPoolExecutor(max_workers=5) as pool:
            first = pool.submit(run_cancelled)
            deadline = time.monotonic() + 2
            while len(state["requests"]) < 1 and time.monotonic() < deadline:
                time.sleep(0.01)
            cancel.set()
            first.result(timeout=0.25)
            later = [pool.submit(Client(settings).extract, [], "发票") for _ in range(4)]
            for future in later:
                assert future.result(timeout=4)["complete"] is True
        assert state["peak"] <= 2 and len(state["requests"]) == 5


def test_streaming_response_and_cancelled_late_callback_are_preserved():
    from src.模型调用 import Client, Cancelled
    saved = []
    with endpoint(streamed=True) as (settings, state):
        client = Client(settings)
        result = client.extract([], "发票", on_attempt=saved.append)
        assert result["records"][0]["fields"]["发票号码"] == "000123"
        assert result["usage"]["total_tokens"] == 45
        assert len(saved) == 1 and "data: [DONE]" in saved[0]["raw"]
    with endpoint(delay=0.25) as (settings, state):
        client, cancel = Client(settings), threading.Event()
        saved.clear()
        with ThreadPoolExecutor(max_workers=1) as pool:
            pending = pool.submit(client.extract, [], "发票", cancel_event=cancel, on_attempt=saved.append)
            deadline = time.monotonic() + 2
            while len(state["requests"]) < 1 and time.monotonic() < deadline:
                time.sleep(0.01)
            cancel.set()
            with pytest.raises(Cancelled):
                pending.result(timeout=0.2)
        deadline = time.monotonic() + 2
        while client.pending_count and time.monotonic() < deadline:
            time.sleep(0.01)
        assert len(saved) == 1 and "000123" in saved[0]["raw"]
        assert "controlled-key" not in saved[0]["raw"]

def test_unknown_precision_invalid_date_and_voucher_conflicts_require_review():
    from src.结果校验 import validate_records
    row = invoice()
    row.pop("precision")
    row["fields"].update({"不含税金额": 100, "税额": 13, "价税合计": 113})
    assert validate_records([row], "发票")[0]["status"] == "review"
    row = invoice()
    row["fields"]["开票日期"] = "2026-02-31"
    assert validate_records([row], "发票")[0]["status"] == "review"
    rows = [voucher(1, "100.00", None), voucher(2, "13.00", None), voucher(3, None, "113.00")]
    rows[2]["fields"]["日期"] = "2026-09-07"
    assert all(row["status"] == "review" for row in validate_records(rows, "会计凭证"))
    rows[2]["fields"]["日期"] = "2026-09-08"
    rows[1]["fields"]["分录序号"] = "1"
    assert all(row["status"] == "review" for row in validate_records(rows, "会计凭证"))

def test_empty_source_position_and_contradictory_missing_reasons_cannot_pass():
    from src.结果校验 import validate_records
    row = invoice()
    row["positions"] = [""]
    assert validate_records([row], "发票")[0]["status"] == "review"
    row = invoice()
    row["missing"]["销售方"] = "原文未记载"
    assert validate_records([row], "发票")[0]["status"] == "review"

def test_contract_complementary_parts_merge_but_versions_with_conflicting_amount_stay_separate():
    from src.结果校验 import merge_records
    one = record({"合同名称": "采购", "甲方": "甲", "乙方": "乙", "金额": None}, doc="合同1",
                 missing={"金额": "原文未记载"})
    two = record({"合同名称": "采购", "甲方": "甲", "乙方": "乙", "金额": "100.00"}, doc="合同1",
                 positions=["第2页"])
    rows = merge_records([one, two], "合同")
    assert len(rows) == 1 and rows[0]["fields"]["金额"] == "100.00"
    assert "金额" not in rows[0]["missing"]
    conflict = copy.deepcopy(two)
    conflict["fields"]["金额"] = "200.00"
    assert len(merge_records([two, conflict], "合同")) == 2


def test_real_http_error_archives_redact_standalone_host_and_key(monkeypatch):
    import socket
    from src.模型调用 import Client, ConfigurationError
    actual_resolver = socket.getaddrinfo
    example_host = "203.0.113.24"
    def local_only_resolver(host, *args, **kwargs):
        return actual_resolver("127.0.0.1" if host == example_host else host, *args, **kwargs)
    monkeypatch.setattr(socket, "getaddrinfo", local_only_resolver)
    monkeypatch.setenv("NO_PROXY", example_host)
    monkeypatch.setenv("no_proxy", example_host)
    message = "host=203.0.113.24; endpoint=http://203.0.113.24:1234/v1; token=fakekey"
    with endpoint([(401, message)]) as (settings, state):
        settings["base_url"] = settings["base_url"].replace("127.0.0.1", example_host)
        settings["api_key"] = "fakekey"
        saved = []
        client = Client(settings)
        with pytest.raises(ConfigurationError) as caught:
            client.extract([], "发票", on_attempt=saved.append)
        assert len(state["requests"]) == 1 and len(saved) == 1
        exposed = json.dumps(client.attempts + saved, ensure_ascii=False) + str(caught.value)
        assert example_host not in exposed
        assert settings["base_url"] not in exposed
        assert "fakekey" not in exposed
        assert "401" in str(caught.value)

def test_credentials_embedded_in_url_are_rejected_before_request():
    from src.模型调用 import Client, ConfigurationError
    with endpoint() as (settings, state):
        settings["base_url"] = settings["base_url"].replace("http://", "http://fakeuser:fakepassword@")
        with pytest.raises(ConfigurationError) as caught:
            Client(settings).extract([], "发票")
        assert not state["requests"]
        assert "fakeuser" not in str(caught.value) and "fakepassword" not in str(caught.value)
