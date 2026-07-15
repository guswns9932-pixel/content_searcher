# content_searcher

지정한 폴더(하위 폴더 포함) 안의 파일명과 파일 내용을 키워드 또는 정규식으로 검색하는
Tkinter 기반 데스크톱 GUI 도구입니다.

## 주요 기능

- 파일명 검색 / 파일 내부 검색 (동시 선택 가능)
- 하위 폴더 포함 여부, 대소문자 구분, 정규식 검색 옵션
- 검색은 백그라운드 스레드에서 실행되며 언제든 중지 가능
- 결과를 표 형태로 확인하고, 더블클릭으로 파일/폴더 바로 열기
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
2. 검색 키워드를 입력하고 필요한 옵션(하위 폴더, 대소문자, 정규식, 파일명/내용)을 선택합니다.
3. "검색 시작"을 누르면 결과가 실시간으로 표에 채워집니다.
4. 결과를 더블클릭하면 파일이 열리고, "CSV 저장"으로 결과를 내보낼 수 있습니다.

## 테스트

각 파일 형식 파서와 검색/스니펫 로직에 대한 단위 테스트가 `tests/`에 있습니다.

```bash
pip install -r requirements-dev.txt
pytest
```

선택적 라이브러리(PyMuPDF, openpyxl, python-docx, python-pptx, xlrd/xlwt)가 설치되어 있지 않으면
해당 형식 테스트는 자동으로 건너뜁니다(skip). GUI 동작 자체는 자동 테스트 대상이 아니며,
Xvfb 등 가상 디스플레이 환경에서 수동으로 확인했습니다.
