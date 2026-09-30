from html.parser import HTMLParser
from pathlib import Path

WORKBENCH_HTML = Path(__file__).resolve().parents[1] / "workbench.html"

class ModalControls(HTMLParser):
    def __init__(self):
        super().__init__()
        self.inputs = set()
        self.fill_targets = set()
    def handle_starttag(self, tag, attrs):
        attrs = dict(attrs)
        if tag == "input" and attrs.get("id", "").startswith("f-"):
            self.inputs.add(attrs["id"])
        if tag == "button" and attrs.get("data-fill-evidence"):
            self.fill_targets.add(attrs["data-fill-evidence"])

def test_compact_editor_uses_native_copy_paste_without_fill_buttons():
    html = WORKBENCH_HTML.read_text(encoding="utf-8")
    parser = ModalControls()
    parser.feed(html)
    assert {"f-company", "f-agent", "f-country", "f-program", "f-request"} <= parser.inputs
    assert not parser.fill_targets
    assert 'id="f-reason"' not in html
    assert "#modal-overlay.show{pointer-events:none}" in html
    assert "#modal-overlay.show .modal{pointer-events:auto}" in html
    assert 'aria-modal="false"' in html
    assert "selectedPageEvidence" not in html

def test_editor_save_is_bound_to_original_detail_and_failure_keeps_draft():
    html = WORKBENCH_HTML.read_text(encoding="utf-8")
    assert "const target={mailId:mail.id,detailId:d?.id}" in html
    assert "if(saved)$('#modal-overlay').classList.remove('show')" in html
    assert "targetDetail" in html
    for source in ("subject", "body", "attachment-name"):
        assert f'data-evidence-source="{source}"' in html
    assert "eprApplication?'epr-application':'attachment-content'" in html
    assert "user-select:text" in html

def test_approved_layout_preserves_scoped_actions_and_resizers():
    html = WORKBENCH_HTML.read_text(encoding="utf-8")
    for action in ('edit', 'confirm', 'return', 'delete'):
        assert f'data-business-action="{action}"' in html
    assert "mail_ids:[mail.id]" in html
    assert "ecopv-workbench-splitters-v3" in html
    assert "splitter.onpointercancel=end" in html
    assert "splitter.ondblclick" in html
    assert 'id="approved-ui03"' in html


def test_original_evidence_tabs_keep_all_workbooks_and_source_priorities():
    html = WORKBENCH_HTML.read_text(encoding="utf-8")
    assert 'id="source-workspace-ui"' in html
    assert "for(const detail of (mail.details||[]))" in html
    assert "member.member_token||member.path" in html
    assert "book.resolvedMember" in html
    assert "book.isEpr&&bookEvidenceHits(book,needles)" in html
    assert "else if(active&&countEvidenceHits(body,needles))" in html
    assert "else if(active&&countEvidenceHits(title,needles))" in html
    assert "data-source-more" in html
    assert "data-source-retry" in html
    assert "sheet.next_offset" in html


def test_table_tab_includes_non_epr_workbooks_without_duplicate_accordion():
    html = WORKBENCH_HTML.read_text(encoding="utf-8")
    assert "['epr','EPR申请表及其他注册表',tableCount]" in html
    assert "const tableCount=state.books.length" in html
    assert "其他附件文本与表格" not in html
    assert 'aria-label="切换附件表格"' in html
    # Display unification must not change the original EPR evidence priority.
    assert "book.isEpr&&bookEvidenceHits(book,needles)" in html
    assert "state.books.map((item,index)" in html
