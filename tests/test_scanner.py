import re
import threading

import pytest

from content_search.scanner import iter_paths, scan
from content_search.text_utils import build_pattern


def never_cancelled():
    return False


def test_iter_paths_recurses_into_subfolders(tmp_path):
    (tmp_path / "sub").mkdir()
    (tmp_path / "a.txt").write_text("a")
    (tmp_path / "sub" / "b.txt").write_text("b")

    paths = set(iter_paths(tmp_path, include_subfolders=True, extensions=None, is_cancelled=never_cancelled))
    names = {p.name for p in paths}
    assert names == {"a.txt", "b.txt"}


def test_iter_paths_without_subfolders_stays_flat(tmp_path):
    (tmp_path / "sub").mkdir()
    (tmp_path / "a.txt").write_text("a")
    (tmp_path / "sub" / "b.txt").write_text("b")

    paths = set(iter_paths(tmp_path, include_subfolders=False, extensions=None, is_cancelled=never_cancelled))
    names = {p.name for p in paths}
    assert names == {"a.txt"}


def test_iter_paths_filters_by_extension(tmp_path):
    (tmp_path / "a.txt").write_text("a")
    (tmp_path / "b.log").write_text("b")

    paths = set(iter_paths(tmp_path, include_subfolders=True, extensions={".txt"}, is_cancelled=never_cancelled))
    names = {p.name for p in paths}
    assert names == {"a.txt"}


def test_scan_finds_filename_and_content_matches(tmp_path):
    (tmp_path / "a.txt").write_text("hello keyword world", encoding="utf-8")
    (tmp_path / "keyword_file.txt").write_text("nothing relevant here", encoding="utf-8")
    (tmp_path / "plain.txt").write_text("no match here", encoding="utf-8")

    pattern = re.compile("keyword")
    results = []
    errors = []

    processed, total_hits, file_count = scan(
        folder=tmp_path,
        pattern=pattern,
        include_subfolders=True,
        search_filename=True,
        search_contents=True,
        extensions=set(),
        is_cancelled=never_cancelled,
        on_file_start=lambda index, path: None,
        on_file_result=lambda path, location, keyword, snippet: results.append((path.name, location, snippet)),
        on_error=errors.append,
        max_workers=4,
    )

    assert processed == 3
    assert file_count == 2
    assert total_hits == 2
    assert not errors

    locations = {(name, location) for name, location, _ in results}
    assert ("a.txt", "본문") in locations
    assert ("keyword_file.txt", "파일명") in locations


def test_scan_attributes_matched_keyword_via_build_pattern(tmp_path):
    (tmp_path / "a.txt").write_text("this invoice is paid", encoding="utf-8")
    (tmp_path / "b.txt").write_text("here is your receipt", encoding="utf-8")
    (tmp_path / "c.txt").write_text("nothing relevant here", encoding="utf-8")

    pattern = build_pattern(["invoice", "receipt"], use_wildcard=False, case_sensitive=True)
    results = []

    processed, total_hits, file_count = scan(
        folder=tmp_path,
        pattern=pattern,
        include_subfolders=True,
        search_filename=False,
        search_contents=True,
        extensions=set(),
        is_cancelled=never_cancelled,
        on_file_start=lambda index, path: None,
        on_file_result=lambda path, location, keyword, snippet: results.append((path.name, keyword)),
        on_error=lambda message: None,
        max_workers=4,
    )

    assert file_count == 2
    assert set(results) == {("a.txt", "invoice"), ("b.txt", "receipt")}


def test_scan_exclude_pattern_drops_units_containing_it(tmp_path):
    (tmp_path / "a.txt").write_text("this is a draft invoice", encoding="utf-8")
    (tmp_path / "b.txt").write_text("this is a final invoice", encoding="utf-8")

    pattern = build_pattern("invoice", use_wildcard=False, case_sensitive=True)
    exclude_pattern = build_pattern("draft", use_wildcard=False, case_sensitive=True)
    results = []

    processed, total_hits, file_count = scan(
        folder=tmp_path,
        pattern=pattern,
        exclude_pattern=exclude_pattern,
        include_subfolders=True,
        search_filename=False,
        search_contents=True,
        extensions=set(),
        is_cancelled=never_cancelled,
        on_file_start=lambda index, path: None,
        on_file_result=lambda path, location, keyword, snippet: results.append(path.name),
        on_error=lambda message: None,
        max_workers=4,
    )

    assert file_count == 1
    assert results == ["b.txt"]


def test_scan_exclude_pattern_excludes_whole_file_across_units(tmp_path):
    # 검색어("keyword")는 A1 셀에, 제외어("draft")는 다른 셀 B1에 있다.
    # 예전 방식(같은 단위 안에서만 제외 검사)이면 이 파일이 결과에 남았을 것이다.
    openpyxl = pytest.importorskip("openpyxl")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = "hello keyword"
    ws["B1"] = "draft version"
    wb.save(tmp_path / "mixed.xlsx")

    pattern = build_pattern("keyword", use_wildcard=False, case_sensitive=False)
    exclude_pattern = build_pattern("draft", use_wildcard=False, case_sensitive=False)
    results = []

    processed, total_hits, file_count = scan(
        folder=tmp_path,
        pattern=pattern,
        exclude_pattern=exclude_pattern,
        include_subfolders=True,
        search_filename=False,
        search_contents=True,
        extensions=set(),
        is_cancelled=never_cancelled,
        on_file_start=lambda index, path: None,
        on_file_result=lambda path, location, keyword, snippet: results.append(path.name),
        on_error=lambda message: None,
        max_workers=4,
    )

    assert file_count == 0
    assert results == []


def test_scan_exclude_pattern_checks_filename_even_though_match_is_in_content(tmp_path):
    # 파일명 자체에 제외어("소개")가 있고, 검색어("keyword")는 파일 내용에 있다.
    (tmp_path / "회사소개.txt").write_text("hello keyword world", encoding="utf-8")
    (tmp_path / "plain.txt").write_text("hello keyword again", encoding="utf-8")

    pattern = build_pattern("keyword", use_wildcard=False, case_sensitive=False)
    exclude_pattern = build_pattern("소개", use_wildcard=False, case_sensitive=False)
    results = []

    processed, total_hits, file_count = scan(
        folder=tmp_path,
        pattern=pattern,
        exclude_pattern=exclude_pattern,
        include_subfolders=True,
        search_filename=True,
        search_contents=True,
        extensions=set(),
        is_cancelled=never_cancelled,
        on_file_start=lambda index, path: None,
        on_file_result=lambda path, location, keyword, snippet: results.append(path.name),
        on_error=lambda message: None,
        max_workers=4,
    )

    assert file_count == 1
    assert results == ["plain.txt"]


def test_scan_respects_extension_filter(tmp_path):
    (tmp_path / "a.txt").write_text("keyword here", encoding="utf-8")
    (tmp_path / "b.log").write_text("keyword here too", encoding="utf-8")

    pattern = re.compile("keyword")
    results = []

    processed, total_hits, file_count = scan(
        folder=tmp_path,
        pattern=pattern,
        include_subfolders=True,
        search_filename=False,
        search_contents=True,
        extensions={".txt"},
        is_cancelled=never_cancelled,
        on_file_start=lambda index, path: None,
        on_file_result=lambda path, location, keyword, snippet: results.append(path.name),
        on_error=lambda message: None,
    )

    assert processed == 1
    assert file_count == 1
    assert results == ["a.txt"]


def test_scan_reports_errors_without_raising(tmp_path):
    (tmp_path / "broken.xlsx").write_text("this is not a real xlsx file", encoding="utf-8")

    pattern = re.compile("anything")
    errors = []

    processed, total_hits, file_count = scan(
        folder=tmp_path,
        pattern=pattern,
        include_subfolders=True,
        search_filename=False,
        search_contents=True,
        extensions=set(),
        is_cancelled=never_cancelled,
        on_file_start=lambda index, path: None,
        on_file_result=lambda path, location, keyword, snippet: None,
        on_error=errors.append,
    )

    assert processed == 1
    assert total_hits == 0
    assert len(errors) == 1


def test_scan_stops_early_when_cancelled_mid_scan(tmp_path):
    for i in range(500):
        (tmp_path / f"file_{i}.txt").write_text("keyword", encoding="utf-8")

    pattern = re.compile("keyword")
    cancel_event = threading.Event()
    seen = []

    def on_file_start(index, path):
        seen.append(index)
        if index >= 5:
            cancel_event.set()

    processed, total_hits, file_count = scan(
        folder=tmp_path,
        pattern=pattern,
        include_subfolders=True,
        search_filename=True,
        search_contents=False,
        extensions=set(),
        is_cancelled=cancel_event.is_set,
        on_file_start=on_file_start,
        on_file_result=lambda path, location, keyword, snippet: None,
        on_error=lambda message: None,
        max_workers=4,
    )

    assert processed < 500
    assert processed >= 5
