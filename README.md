# content_searcher

지정한 폴더(하위 폴더 포함) 안의 파일명과 파일 내용을 키워드 또는 와일드카드로 검색하는
Tkinter 기반 데스크톱 GUI 도구입니다.

## 주요 기능

- 파일명 검색 / 파일 내부 검색 (동시 선택 가능)
- 하위 폴더 포함 여부, 대소문자 구분 옵션
- 와일드카드 사용(`*`는 임의 길이 문자열, `?`는 문자 1개)으로 `202?년*보고서`처럼
  파일 탐색기에서 익숙한 방식으로 패턴 검색 가능
- 확장자 필터: `txt, pdf`처럼 입력하면 해당 확장자 파일만 검색 대상에 포함 (비워두면 전체)
- 폴더 탐색과 파일 검색을 스레드 풀로 병렬 처리해 대용량 폴더에서도 빠르게 검색하며,
  발견 즉시 스트리밍으로 처리하므로 시작 전 대기 시간이 없고 언제든 중지 가능
- 결과를 표 형태(줄무늬 스타일)로 확인하고, 더블클릭으로 파일/폴더 바로 열기
- 검색 결과 CSV 저장

### 지원 파일 형식

| 종류 | 확장자 | 필요 라이브러리 |
|---|---|---|
| 일반 텍스트 | `.txt .csv .tsv .log .ini .cfg .conf .json .xml .yaml .yml .md .py .java .c .cpp .h .hpp .js .ts .css .html .htm .sql .bat .cmd .ps1` | 없음 |
| PDF | `.pdf` | PyMuPDF |
| Excel | `.xlsx .xlsm .xltx .xltm` | openpyxl |
| Excel (구버전) | `.xls` | xlrd |
| Word | `.docx` | python-docx |
| PowerPoint | `.pptx` | python-pptx |
| 한글 | `.hwpx` | 없음 (ZIP+XML 직접 파싱) |

선택적 라이브러리가 설치되어 있지 않으면 해당 형식은 검색에서 자동으로 제외됩니다.

## 설치

Python 3.9 이상이 필요합니다. (Tkinter는 대부분의 공식 Python 배포판에 기본 포함되어 있습니다.
Linux에서 누락된 경우 `sudo apt install python3-tk` 등으로 설치하세요.)

```bash
pip install -r requirements.txt
```

모든 파일 형식이 필요하지 않다면 `requirements.txt`에서 필요한 항목만 골라 설치해도 됩니다.

## 실행

```bash
python content_searcher.py
```

1. "폴더 선택"으로 검색할 폴더를 지정합니다.
2. 검색 키워드를 입력하고 필요한 옵션(하위 폴더, 대소문자, 와일드카드, 파일명/내용)을 선택합니다.
3. 특정 형식만 검색하려면 "확장자 필터"에 `txt`, `.pdf`, `*.docx`처럼 쉼표/공백으로 구분해 입력합니다.
4. "검색 시작"을 누르면 결과가 실시간으로 표에 채워집니다.
5. 결과를 더블클릭하면 파일이 열리고, "CSV 저장"으로 결과를 내보낼 수 있습니다.

## 프로젝트 구조

```
content_searcher.py        실행 진입점 (python content_searcher.py)
content_search/
  text_utils.py             정규화, 스니펫, 패턴, 확장자 필터 파싱 (형식 비의존)
  parsers.py                형식별 텍스트 추출 (txt/pdf/xlsx/xls/docx/pptx/hwpx)
  scanner.py                폴더 순회 + 스레드 풀 기반 병렬 검색 엔진 (GUI 비의존)
  gui.py                    Tkinter 화면 구성과 이벤트 처리
tests/                      pytest 단위 테스트
```

로직을 형식/스캔/화면 레이어로 분리해 각각 독립적으로 테스트하고 재사용할 수 있게 했습니다.

## 테스트

핵심 로직·파서·스캔 엔진에 대한 단위 테스트가 `tests/`에 있습니다.

```bash
pip install -r requirements-dev.txt
pytest
```

선택적 라이브러리(PyMuPDF, openpyxl, python-docx, python-pptx, xlrd/xlwt)가 설치되어 있지 않으면
해당 형식 테스트는 자동으로 건너뜁니다(skip). GUI 자체의 시각적 동작은 자동 테스트 대상이 아니며,
Xvfb 등 가상 디스플레이 환경에서 수동으로 확인했습니다.
