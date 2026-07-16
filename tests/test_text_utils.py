import re

import pytest

from content_search.text_utils import (
    MAX_HITS_PER_FILE,
    build_pattern,
    find_matches,
    make_snippet,
    normalize_text,
    parse_extension_filter,
)


def test_normalize_text_collapses_whitespace_and_none():
    assert normalize_text(None) == ""
    assert normalize_text("  a\n\tb   c ") == "a b c"
    assert normalize_text(123) == "123"


def test_make_snippet_marks_truncation():
    text = "0123456789" * 20
    snippet = make_snippet(text, 100, 105, radius=5)
    assert snippet.startswith("…")
    assert snippet.endswith("…")
    assert "…" not in snippet[1:-1]


def test_make_snippet_no_truncation_marker_at_edges():
    text = "hello world"
    snippet = make_snippet(text, 0, 5, radius=10)
    assert snippet == "hello world"


def test_find_matches_returns_start_end_snippet():
    pattern = re.compile("cat")
    results = find_matches("a cat sat on a cat mat", pattern)
    assert len(results) == 2
    assert results[0][0] == 2 and results[0][1] == 5


def test_find_matches_caps_at_max_hits_per_file():
    text = "x" * (MAX_HITS_PER_FILE + 10)
    pattern = re.compile("x")
    results = find_matches(text, pattern)
    assert len(results) == MAX_HITS_PER_FILE


def test_build_pattern_literal_is_case_insensitive_by_default():
    pattern = build_pattern("Cat", use_regex=False, case_sensitive=False)
    assert pattern.search("a cat sat")


def test_build_pattern_literal_case_sensitive():
    pattern = build_pattern("Cat", use_regex=False, case_sensitive=True)
    assert pattern.search("a Cat sat")
    assert not pattern.search("a cat sat")


def test_build_pattern_literal_escapes_special_chars():
    pattern = build_pattern("a.b", use_regex=False, case_sensitive=True)
    assert pattern.search("a.b")
    assert not pattern.search("aXb")


def test_build_pattern_invalid_regex_raises():
    with pytest.raises(re.error):
        build_pattern("(", use_regex=True, case_sensitive=True)


def test_parse_extension_filter_empty_means_no_filter():
    assert parse_extension_filter("") == set()
    assert parse_extension_filter("   ") == set()


def test_parse_extension_filter_normalizes_various_forms():
    assert parse_extension_filter("txt") == {".txt"}
    assert parse_extension_filter(".txt") == {".txt"}
    assert parse_extension_filter("*.txt") == {".txt"}
    assert parse_extension_filter("TXT") == {".txt"}


def test_parse_extension_filter_multiple_tokens():
    assert parse_extension_filter("txt, pdf .docx  xlsx") == {
        ".txt", ".pdf", ".docx", ".xlsx"
    }
