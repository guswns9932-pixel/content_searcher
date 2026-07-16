"""파일 형식별 텍스트 추출 및 검색. 선택적 라이브러리가 없으면 해당 형식은 건너뛴다."""

from pathlib import Path
from xml.etree import ElementTree as ET
import zipfile

from .text_utils import MAX_HITS_PER_FILE, find_matches

try:
    import fitz  # PyMuPDF
except ImportError:
    fitz = None

try:
    import openpyxl
except ImportError:
    openpyxl = None

try:
    from docx import Document
except ImportError:
    Document = None

try:
    from pptx import Presentation
except ImportError:
    Presentation = None

try:
    import xlrd
except ImportError:
    xlrd = None


TEXT_EXTENSIONS = {
    ".txt", ".csv", ".tsv", ".log", ".ini", ".cfg", ".conf",
    ".json", ".xml", ".yaml", ".yml", ".md", ".py", ".java",
    ".c", ".cpp", ".h", ".hpp", ".js", ".ts", ".css", ".html",
    ".htm", ".sql", ".bat", ".cmd", ".ps1"
}

SUPPORTED_EXTENSIONS = (
    TEXT_EXTENSIONS
    | {".pdf", ".xlsx", ".xlsm", ".xltx", ".xltm", ".xls",
       ".docx", ".pptx", ".hwpx"}
)


def read_text_file(path):
    """텍스트 파일을 여러 인코딩으로 시도해 읽는다."""
    encodings = ("utf-8-sig", "utf-8", "cp949", "euc-kr", "utf-16")
    last_error = None

    for encoding in encodings:
        try:
            with open(path, "r", encoding=encoding, errors="strict") as f:
                return f.read()
        except UnicodeError as exc:
            last_error = exc

    # 마지막 수단: 읽을 수 없는 문자를 대체
    try:
        with open(path, "r", encoding="utf-8", errors="replace") as f:
            return f.read()
    except Exception:
        if last_error:
            raise last_error
        raise


def search_text_file(path, pattern, exclude_pattern=None):
    text = read_text_file(path)
    rows = []
    for _, _, keyword, snippet in find_matches(text, pattern, exclude_pattern):
        rows.append(("본문", keyword, snippet))
    return rows


def search_pdf(path, pattern, exclude_pattern=None):
    if fitz is None:
        raise RuntimeError("PyMuPDF가 설치되지 않았습니다.")

    rows = []
    with fitz.open(path) as doc:
        for page_index, page in enumerate(doc):
            text = page.get_text("text")
            for _, _, keyword, snippet in find_matches(text, pattern, exclude_pattern):
                rows.append((f"{page_index + 1}페이지", keyword, snippet))
                if len(rows) >= MAX_HITS_PER_FILE:
                    return rows
    return rows


def search_xlsx(path, pattern, exclude_pattern=None):
    if openpyxl is None:
        raise RuntimeError("openpyxl이 설치되지 않았습니다.")

    rows = []
    workbook = openpyxl.load_workbook(
        path,
        read_only=True,
        data_only=False
    )

    try:
        for worksheet in workbook.worksheets:
            for row in worksheet.iter_rows():
                for cell in row:
                    value = cell.value
                    if value is None:
                        continue
                    text = str(value)
                    matches = find_matches(text, pattern, exclude_pattern)
                    for _, _, keyword, snippet in matches:
                        rows.append((f"{worksheet.title}!{cell.coordinate}", keyword, snippet))
                        if len(rows) >= MAX_HITS_PER_FILE:
                            return rows
    finally:
        workbook.close()

    return rows


def search_xls(path, pattern, exclude_pattern=None):
    if xlrd is None:
        raise RuntimeError("xlrd가 설치되지 않았습니다.")

    rows = []
    workbook = xlrd.open_workbook(path, on_demand=True)
    try:
        for sheet in workbook.sheets():
            for row_index in range(sheet.nrows):
                for col_index in range(sheet.ncols):
                    value = sheet.cell_value(row_index, col_index)
                    if value in (None, ""):
                        continue
                    text = str(value)
                    for _, _, keyword, snippet in find_matches(text, pattern, exclude_pattern):
                        cell_name = f"{xlrd.formula.colname(col_index)}{row_index + 1}"
                        rows.append((f"{sheet.name}!{cell_name}", keyword, snippet))
                        if len(rows) >= MAX_HITS_PER_FILE:
                            return rows
    finally:
        workbook.release_resources()

    return rows


def search_docx(path, pattern, exclude_pattern=None):
    if Document is None:
        raise RuntimeError("python-docx가 설치되지 않았습니다.")

    rows = []
    doc = Document(path)

    for index, paragraph in enumerate(doc.paragraphs, start=1):
        text = paragraph.text
        for _, _, keyword, snippet in find_matches(text, pattern, exclude_pattern):
            rows.append((f"문단 {index}", keyword, snippet))
            if len(rows) >= MAX_HITS_PER_FILE:
                return rows

    for table_index, table in enumerate(doc.tables, start=1):
        for row_index, row in enumerate(table.rows, start=1):
            for col_index, cell in enumerate(row.cells, start=1):
                text = cell.text
                for _, _, keyword, snippet in find_matches(text, pattern, exclude_pattern):
                    rows.append((f"표 {table_index} / {row_index}행 {col_index}열", keyword, snippet))
                    if len(rows) >= MAX_HITS_PER_FILE:
                        return rows

    return rows


def extract_shape_text(shape):
    """PowerPoint 도형에서 텍스트와 표 내용을 추출한다."""
    texts = []

    if hasattr(shape, "text") and shape.text:
        texts.append(shape.text)

    if getattr(shape, "has_table", False):
        for row in shape.table.rows:
            for cell in row.cells:
                if cell.text:
                    texts.append(cell.text)

    if getattr(shape, "shape_type", None) == 6 and hasattr(shape, "shapes"):  # 그룹 도형
        for child in shape.shapes:
            texts.extend(extract_shape_text(child))

    return texts


def search_pptx(path, pattern, exclude_pattern=None):
    if Presentation is None:
        raise RuntimeError("python-pptx가 설치되지 않았습니다.")

    rows = []
    presentation = Presentation(path)

    for slide_index, slide in enumerate(presentation.slides, start=1):
        for shape in slide.shapes:
            for text in extract_shape_text(shape):
                for _, _, keyword, snippet in find_matches(text, pattern, exclude_pattern):
                    rows.append((f"{slide_index}슬라이드", keyword, snippet))
                    if len(rows) >= MAX_HITS_PER_FILE:
                        return rows

    return rows


def search_hwpx(path, pattern, exclude_pattern=None):
    """
    HWPX는 ZIP 기반 XML 형식이다.
    XML에 포함된 텍스트를 추출해 검색한다.
    """
    rows = []

    with zipfile.ZipFile(path, "r") as archive:
        xml_names = [
            name for name in archive.namelist()
            if name.lower().endswith(".xml")
            and (
                name.startswith("Contents/")
                or name.startswith("Preview/")
            )
        ]

        for name in xml_names:
            raw = archive.read(name)
            try:
                root = ET.fromstring(raw)
                text = " ".join(part for part in root.itertext() if part)
            except ET.ParseError:
                text = raw.decode("utf-8", errors="ignore")

            for _, _, keyword, snippet in find_matches(text, pattern, exclude_pattern):
                location = Path(name).name
                rows.append((f"HWPX/{location}", keyword, snippet))
                if len(rows) >= MAX_HITS_PER_FILE:
                    return rows

    return rows


def search_file_contents(path, pattern, exclude_pattern=None):
    ext = path.suffix.lower()

    if ext in TEXT_EXTENSIONS:
        return search_text_file(path, pattern, exclude_pattern)
    if ext == ".pdf":
        return search_pdf(path, pattern, exclude_pattern)
    if ext in {".xlsx", ".xlsm", ".xltx", ".xltm"}:
        return search_xlsx(path, pattern, exclude_pattern)
    if ext == ".xls":
        return search_xls(path, pattern, exclude_pattern)
    if ext == ".docx":
        return search_docx(path, pattern, exclude_pattern)
    if ext == ".pptx":
        return search_pptx(path, pattern, exclude_pattern)
    if ext == ".hwpx":
        return search_hwpx(path, pattern, exclude_pattern)

    return []
