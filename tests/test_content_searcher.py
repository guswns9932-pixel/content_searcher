import re
import sys
import zipfile
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import content_searcher as cs


def test_normalize_text_collapses_whitespace_and_none():
    assert cs.normalize_text(None) == ""
    assert cs.normalize_text("  a\n\tb   c ") == "a b c"
    assert cs.normalize_text(123) == "123"


def test_make_snippet_marks_truncation():
    text = "0123456789" * 20
    snippet = cs.make_snippet(text, 100, 105, radius=5)
    assert snippet.startswith("…")
    assert snippet.endswith("…")
    assert "…" not in snippet[1:-1]


def test_make_snippet_no_truncation_marker_at_edges():
    text = "hello world"
    snippet = cs.make_snippet(text, 0, 5, radius=10)
    assert snippet == "hello world"


def test_find_matches_returns_start_end_snippet():
    pattern = re.compile("cat")
    results = cs.find_matches("a cat sat on a cat mat", pattern)
    assert len(results) == 2
    assert results[0][0] == 2 and results[0][1] == 5


def test_find_matches_caps_at_max_hits_per_file():
    text = "x" * (cs.MAX_HITS_PER_FILE + 10)
    pattern = re.compile("x")
    results = cs.find_matches(text, pattern)
    assert len(results) == cs.MAX_HITS_PER_FILE


def test_build_pattern_literal_is_case_insensitive_by_default():
    pattern = cs.build_pattern("Cat", use_regex=False, case_sensitive=False)
    assert pattern.search("a cat sat")


def test_build_pattern_literal_case_sensitive():
    pattern = cs.build_pattern("Cat", use_regex=False, case_sensitive=True)
    assert pattern.search("a Cat sat")
    assert not pattern.search("a cat sat")


def test_build_pattern_literal_escapes_special_chars():
    pattern = cs.build_pattern("a.b", use_regex=False, case_sensitive=True)
    assert pattern.search("a.b")
    assert not pattern.search("aXb")


def test_build_pattern_invalid_regex_raises():
    with pytest.raises(re.error):
        cs.build_pattern("(", use_regex=True, case_sensitive=True)


def test_read_text_file_utf8(tmp_path):
    path = tmp_path / "sample.txt"
    path.write_text("안녕하세요 keyword", encoding="utf-8")
    assert cs.read_text_file(path) == "안녕하세요 keyword"


def test_read_text_file_cp949(tmp_path):
    path = tmp_path / "sample_cp949.txt"
    path.write_bytes("안녕 keyword".encode("cp949"))
    assert cs.read_text_file(path) == "안녕 keyword"


def test_search_text_file_finds_snippet(tmp_path):
    path = tmp_path / "a.txt"
    path.write_text("hello keyword world", encoding="utf-8")
    pattern = re.compile("keyword")
    rows = cs.search_text_file(path, pattern)
    assert rows == [("본문", "hello keyword world")]


def test_search_pdf(tmp_path):
    fitz = pytest.importorskip("fitz")
    path = tmp_path / "sample.pdf"
    doc = fitz.open()
    page = doc.new_page()
    page.insert_text((72, 72), "hello keyword world")
    doc.save(path)
    doc.close()

    pattern = re.compile("keyword")
    rows = cs.search_pdf(path, pattern)
    assert len(rows) == 1
    assert rows[0][0] == "1페이지"
    assert "keyword" in rows[0][1]


def test_search_xlsx(tmp_path):
    openpyxl = pytest.importorskip("openpyxl")
    path = tmp_path / "sample.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"] = "hello keyword"
    wb.save(path)

    pattern = re.compile("keyword")
    rows = cs.search_xlsx(path, pattern)
    assert rows == [("Sheet1!A1", "hello keyword")]


def test_search_xls(tmp_path):
    xlwt = pytest.importorskip("xlwt")
    pytest.importorskip("xlrd")
    path = tmp_path / "sample.xls"
    wb = xlwt.Workbook()
    ws = wb.add_sheet("Sheet1")
    ws.write(0, 0, "hello keyword")
    wb.save(str(path))

    pattern = re.compile("keyword")
    rows = cs.search_xls(path, pattern)
    assert len(rows) == 1
    assert rows[0][0] == "Sheet1!A1"
    assert "keyword" in rows[0][1]


def test_search_docx(tmp_path):
    docx = pytest.importorskip("docx")
    path = tmp_path / "sample.docx"
    document = docx.Document()
    document.add_paragraph("hello keyword world")
    table = document.add_table(rows=1, cols=1)
    table.rows[0].cells[0].text = "table keyword cell"
    document.save(path)

    pattern = re.compile("keyword")
    rows = cs.search_docx(path, pattern)
    locations = [loc for loc, _ in rows]
    assert "문단 1" in locations
    assert any(loc.startswith("표 1") for loc in locations)


def test_search_pptx(tmp_path):
    pptx_mod = pytest.importorskip("pptx")
    path = tmp_path / "sample.pptx"
    presentation = pptx_mod.Presentation()
    slide = presentation.slides.add_slide(presentation.slide_layouts[6])
    textbox = slide.shapes.add_textbox(0, 0, pptx_mod.util.Inches(4), pptx_mod.util.Inches(1))
    textbox.text_frame.text = "hello keyword world"
    presentation.save(path)

    pattern = re.compile("keyword")
    rows = cs.search_pptx(path, pattern)
    assert len(rows) == 1
    assert rows[0][0] == "1슬라이드"


def test_extract_shape_text_handles_group_shapes():
    class FakeShape:
        def __init__(self, text=None, has_table=False, table=None, shape_type=None, shapes=None):
            self.text = text
            self.has_table = has_table
            self.table = table
            self.shape_type = shape_type
            self.shapes = shapes or []

    child = FakeShape(text="child text")
    group = FakeShape(shape_type=6, shapes=[child])

    assert cs.extract_shape_text(group) == ["child text"]


def test_search_hwpx(tmp_path):
    path = tmp_path / "sample.hwpx"
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr(
            "Contents/section0.xml",
            "<root><p>hello keyword world</p></root>"
        )

    pattern = re.compile("keyword")
    rows = cs.search_hwpx(path, pattern)
    assert len(rows) == 1
    assert rows[0][0] == "HWPX/section0.xml"


@pytest.mark.parametrize("ext,expected_dispatch", [
    (".txt", "search_text_file"),
    (".pdf", "search_pdf"),
    (".xlsx", "search_xlsx"),
    (".xls", "search_xls"),
    (".docx", "search_docx"),
    (".pptx", "search_pptx"),
    (".hwpx", "search_hwpx"),
    (".unknown", None),
])
def test_search_file_contents_dispatches_by_extension(monkeypatch, ext, expected_dispatch):
    calls = []

    def make_stub(name):
        def stub(path, pattern):
            calls.append(name)
            return []
        return stub

    for fn_name in (
        "search_text_file", "search_pdf", "search_xlsx",
        "search_xls", "search_docx", "search_pptx", "search_hwpx",
    ):
        monkeypatch.setattr(cs, fn_name, make_stub(fn_name))

    path = Path(f"dummy{ext}")
    result = cs.search_file_contents(path, re.compile("x"))

    if expected_dispatch is None:
        assert result == []
        assert calls == []
    else:
        assert calls == [expected_dispatch]
