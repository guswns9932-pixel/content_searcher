import re
import zipfile
from pathlib import Path

import pytest

from content_search import parsers
from content_search.text_utils import build_pattern


def test_read_text_file_utf8(tmp_path):
    path = tmp_path / "sample.txt"
    path.write_text("안녕하세요 keyword", encoding="utf-8")
    assert parsers.read_text_file(path) == "안녕하세요 keyword"


def test_read_text_file_cp949(tmp_path):
    path = tmp_path / "sample_cp949.txt"
    path.write_bytes("안녕 keyword".encode("cp949"))
    assert parsers.read_text_file(path) == "안녕 keyword"


def test_search_text_file_finds_snippet(tmp_path):
    path = tmp_path / "a.txt"
    path.write_text("hello keyword world", encoding="utf-8")
    pattern = build_pattern("keyword", use_wildcard=False, case_sensitive=False)
    rows = parsers.search_text_file(path, pattern)
    assert rows == [("본문", "keyword", "hello keyword world")]


def test_search_pdf(tmp_path):
    fitz = pytest.importorskip("fitz")
    path = tmp_path / "sample.pdf"
    doc = fitz.open()
    page = doc.new_page()
    page.insert_text((72, 72), "hello keyword world")
    doc.save(path)
    doc.close()

    pattern = build_pattern("keyword", use_wildcard=False, case_sensitive=False)
    rows = parsers.search_pdf(path, pattern)
    assert len(rows) == 1
    assert rows[0][0] == "1페이지"
    assert rows[0][1] == "keyword"
    assert "keyword" in rows[0][2]


def test_search_xlsx(tmp_path):
    openpyxl = pytest.importorskip("openpyxl")
    path = tmp_path / "sample.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"] = "hello keyword"
    wb.save(path)

    pattern = build_pattern("keyword", use_wildcard=False, case_sensitive=False)
    rows = parsers.search_xlsx(path, pattern)
    assert rows == [("Sheet1!A1", "keyword", "hello keyword")]


def test_search_xls(tmp_path):
    xlwt = pytest.importorskip("xlwt")
    pytest.importorskip("xlrd")
    path = tmp_path / "sample.xls"
    wb = xlwt.Workbook()
    ws = wb.add_sheet("Sheet1")
    ws.write(0, 0, "hello keyword")
    wb.save(str(path))

    pattern = build_pattern("keyword", use_wildcard=False, case_sensitive=False)
    rows = parsers.search_xls(path, pattern)
    assert len(rows) == 1
    assert rows[0][0] == "Sheet1!A1"
    assert rows[0][1] == "keyword"
    assert "keyword" in rows[0][2]


def test_search_docx(tmp_path):
    docx = pytest.importorskip("docx")
    path = tmp_path / "sample.docx"
    document = docx.Document()
    document.add_paragraph("hello keyword world")
    table = document.add_table(rows=1, cols=1)
    table.rows[0].cells[0].text = "table keyword cell"
    document.save(path)

    pattern = build_pattern("keyword", use_wildcard=False, case_sensitive=False)
    rows = parsers.search_docx(path, pattern)
    locations = [loc for loc, _, _ in rows]
    assert "문단 1" in locations
    assert any(loc.startswith("표 1") for loc in locations)
    assert all(keyword == "keyword" for _, keyword, _ in rows)


def test_search_pptx(tmp_path):
    pptx_mod = pytest.importorskip("pptx")
    path = tmp_path / "sample.pptx"
    presentation = pptx_mod.Presentation()
    slide = presentation.slides.add_slide(presentation.slide_layouts[6])
    textbox = slide.shapes.add_textbox(0, 0, pptx_mod.util.Inches(4), pptx_mod.util.Inches(1))
    textbox.text_frame.text = "hello keyword world"
    presentation.save(path)

    pattern = build_pattern("keyword", use_wildcard=False, case_sensitive=False)
    rows = parsers.search_pptx(path, pattern)
    assert len(rows) == 1
    assert rows[0][0] == "1슬라이드"
    assert rows[0][1] == "keyword"


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

    assert parsers.extract_shape_text(group) == ["child text"]


def test_search_hwpx(tmp_path):
    path = tmp_path / "sample.hwpx"
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr(
            "Contents/section0.xml",
            "<root><p>hello keyword world</p></root>"
        )

    pattern = build_pattern("keyword", use_wildcard=False, case_sensitive=False)
    rows = parsers.search_hwpx(path, pattern)
    assert len(rows) == 1
    assert rows[0][0] == "HWPX/section0.xml"
    assert rows[0][1] == "keyword"


def test_search_hwpx_attributes_correct_keyword_among_several(tmp_path):
    path = tmp_path / "sample.hwpx"
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr(
            "Contents/section0.xml",
            "<root><p>invoice paid, receipt attached</p></root>"
        )

    pattern = build_pattern(["invoice", "receipt"], use_wildcard=False, case_sensitive=False)
    rows = parsers.search_hwpx(path, pattern)
    matched_keywords = {keyword for _, keyword, _ in rows}
    assert matched_keywords == {"invoice", "receipt"}


def test_file_contains_excluded_text_checks_across_different_units(tmp_path):
    # 검색어("keyword")는 xlsx의 A1 셀에, 제외어("draft")는 다른 셀인 B1에 있다.
    # 예전 방식(같은 단위 안에서만 제외 검사)이면 놓쳤을 케이스를 검증한다.
    openpyxl = pytest.importorskip("openpyxl")
    path = tmp_path / "sample.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = "hello keyword"
    ws["B1"] = "draft version"
    wb.save(path)

    exclude_pattern = build_pattern("draft", use_wildcard=False, case_sensitive=False)
    assert parsers.file_contains_excluded_text(path, exclude_pattern) is True

    missing_pattern = build_pattern("nonexistent", use_wildcard=False, case_sensitive=False)
    assert parsers.file_contains_excluded_text(path, missing_pattern) is False


def test_file_contains_excluded_text_unsupported_extension_is_false(tmp_path):
    path = tmp_path / "sample.unknownext"
    path.write_text("draft", encoding="utf-8")
    exclude_pattern = build_pattern("draft", use_wildcard=False, case_sensitive=False)
    assert parsers.file_contains_excluded_text(path, exclude_pattern) is False


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
        monkeypatch.setattr(parsers, fn_name, make_stub(fn_name))

    path = Path(f"dummy{ext}")
    result = parsers.search_file_contents(path, re.compile("x"))

    if expected_dispatch is None:
        assert result == []
        assert calls == []
    else:
        assert calls == [expected_dispatch]
