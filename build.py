"""PyInstaller onefile 빌드 스크립트.

실행: python build.py
결과물: Windows에서는 dist/quote_Auto.exe, 그 외 OS에서는 dist/quote_Auto

주의: PyInstaller는 크로스 컴파일을 지원하지 않는다. 즉 실제 .exe가 필요하면
반드시 Windows 환경에서 이 스크립트를 실행해야 한다(Linux/Mac에서 실행하면
그 OS용 실행 파일이 만들어진다).

파일명을 다른 이름으로 바꾸고 싶다면 아래 APP_NAME만 바꾸면 되고,
바뀐 이름 그대로 빌드된 실행 파일명 = 작업관리자에 표시되는 프로세스명이 된다
(일반적인 PyInstaller onefile 빌드는 항상 이렇게 파일명과 프로세스명이 같다).
"""

import shutil
import subprocess
import sys
from pathlib import Path

APP_NAME = "quote_Auto"
ENTRY_POINT = "content_searcher.py"

ROOT = Path(__file__).resolve().parent

# PyInstaller가 자동으로 못 찾을 수 있는 선택적 라이브러리들.
# 설치되어 있는 것만 --hidden-import로 넘긴다.
OPTIONAL_MODULES = ["fitz", "openpyxl", "docx", "pptx", "xlrd"]


def ensure_pyinstaller():
    try:
        import PyInstaller  # noqa: F401
    except ImportError:
        print("PyInstaller가 설치되어 있지 않아 설치를 진행합니다...")
        subprocess.run([sys.executable, "-m", "pip", "install", "pyinstaller"], check=True)


def installed_optional_modules():
    available = []
    for name in OPTIONAL_MODULES:
        try:
            __import__(name)
            available.append(name)
        except ImportError:
            pass
    return available


def clean_previous_build():
    for stale in ("build", "dist", f"{APP_NAME}.spec"):
        path = ROOT / stale
        if path.is_dir():
            shutil.rmtree(path)
        elif path.is_file():
            path.unlink()


def build():
    ensure_pyinstaller()
    clean_previous_build()

    command = [
        sys.executable, "-m", "PyInstaller",
        "--onefile",
        "--windowed",
        "--name", APP_NAME,
        "--noconfirm",
        "--clean",
    ]

    for module in installed_optional_modules():
        command += ["--hidden-import", module]

    command.append(str(ROOT / ENTRY_POINT))

    print("실행 명령:", " ".join(command))
    subprocess.run(command, check=True, cwd=ROOT)

    exe_name = f"{APP_NAME}.exe" if sys.platform.startswith("win") else APP_NAME
    result_path = ROOT / "dist" / exe_name

    if result_path.exists():
        print(f"\n빌드 완료: {result_path}")
    else:
        print("\n빌드는 끝났지만 결과 파일을 찾지 못했습니다. dist/ 폴더를 확인하세요.")


if __name__ == "__main__":
    build()
