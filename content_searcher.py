"""실행 진입점. 실제 구현은 content_search 패키지에 있다."""

from content_search.gui import ContentSearchApp

if __name__ == "__main__":
    app = ContentSearchApp()
    app.mainloop()
