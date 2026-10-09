from email.header import Header

from modules.mail_reader import _decode_str


def test_mislabeled_subject_uses_lossless_fallback():
    original = "示例科技有限公司-法国电池法注册"
    raw = original.encode("gb18030")
    header = "=?utf-8?B?" + __import__("base64").b64encode(raw).decode("ascii") + "?="
    assert _decode_str(header) == original


def test_valid_utf8_subject_still_decodes():
    original = "示例贸易有限公司"
    assert _decode_str(Header(original, "utf-8").encode()) == original


def test_utf8_bytes_mislabeled_as_gbk_prefer_utf8():
    original = "示例科技有限公司"
    header = "=?gbk?B?" + __import__("base64").b64encode(original.encode("utf-8")).decode("ascii") + "?="
    assert _decode_str(header) == original
