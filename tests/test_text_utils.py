import re

from content_search.text_utils import (
    MAX_HITS_PER_FILE,
    build_pattern,
    find_matches,
    make_snippet,
    normalize_text,
    parse_extension_filter,
    parse_keywords,
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


def test_find_matches_keyword_is_none_for_plain_regex():
    # build_pattern()을 거치지 않은 일반 re.Pattern에는 keywords 정보가 없다.
    pattern = re.compile("cat")
    results = find_matches("a cat sat", pattern)
    assert results[0][2] is None


def test_find_matches_attributes_correct_keyword_from_build_pattern():
    pattern = build_pattern(["invoice", "receipt"], use_wildcard=False, case_sensitive=True)
    results = find_matches("paid the invoice, got a receipt", pattern)
    matched_keywords = [keyword for _, _, keyword, _ in results]
    assert matched_keywords == ["invoice", "receipt"]


def test_find_matches_attributes_keyword_for_wildcard_match():
    pattern = build_pattern(["invoice*2024"], use_wildcard=True, case_sensitive=True)
    results = find_matches("the invoice number 2024 was paid", pattern)
    assert len(results) == 1
    assert results[0][2] == "invoice*2024"


def test_build_pattern_literal_is_case_insensitive_by_default():
    pattern = build_pattern("Cat", use_wildcard=False, case_sensitive=False)
    assert pattern.search("a cat sat")


def test_build_pattern_literal_case_sensitive():
    pattern = build_pattern("Cat", use_wildcard=False, case_sensitive=True)
    assert pattern.search("a Cat sat")
    assert not pattern.search("a cat sat")


def test_build_pattern_literal_escapes_special_chars():
    pattern = build_pattern("a.b", use_wildcard=False, case_sensitive=True)
    assert pattern.search("a.b")
    assert not pattern.search("aXb")


def test_build_pattern_wildcard_star_matches_any_length():
    pattern = build_pattern("invoice*2024", use_wildcard=True, case_sensitive=True)
    assert pattern.search("this is invoice number 2024 total")
    assert not pattern.search("invoice 2023")


def test_build_pattern_wildcard_question_mark_matches_exactly_one_char():
    pattern = build_pattern("b?d", use_wildcard=True, case_sensitive=True)
    assert pattern.search("the bad dog")
    assert not pattern.search("the bd dog")  # '?' requires exactly one character
    assert not pattern.search("the byyd dog")  # two characters, not one


def test_build_pattern_wildcard_still_escapes_literal_regex_chars():
    pattern = build_pattern("a.b*c", use_wildcard=True, case_sensitive=True)
    assert pattern.search("a.b---c")
    assert not pattern.search("aXb---c")  # literal '.' must not act as regex wildcard


def test_build_pattern_wildcard_never_raises_on_arbitrary_input():
    # 순수 문자 치환이라 정규식 특수문자가 섞여도 컴파일 오류가 나지 않는다.
    pattern = build_pattern("(weird [input", use_wildcard=True, case_sensitive=True)
    assert pattern.search("(weird [input")


def test_build_pattern_multiple_keywords_matches_any_of_them():
    pattern = build_pattern(["invoice", "receipt"], use_wildcard=False, case_sensitive=True)
    assert pattern.search("paid the invoice today")
    assert pattern.search("here is your receipt")
    assert not pattern.search("nothing relevant here")


def test_build_pattern_multiple_keywords_supports_wildcard_per_keyword():
    pattern = build_pattern(["invoice*2024", "b?d"], use_wildcard=True, case_sensitive=True)
    assert pattern.search("this invoice number 2024 is paid")
    assert pattern.search("the bad dog")
    assert not pattern.search("nothing relevant here")


def test_parse_keywords_empty_returns_empty_list():
    assert parse_keywords("") == []
    assert parse_keywords("   ") == []
    assert parse_keywords(", ,;\n") == []


def test_parse_keywords_splits_on_comma_semicolon_and_newline():
    assert parse_keywords("invoice, receipt;세금계산서\n영수증") == [
        "invoice", "receipt", "세금계산서", "영수증"
    ]


def test_parse_keywords_preserves_spaces_within_a_keyword():
    assert parse_keywords("2024 report, final draft") == ["2024 report", "final draft"]


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
