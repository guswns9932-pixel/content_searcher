"""검색 패턴, 스니펫, 확장자 필터 등 파일 형식과 무관한 순수 텍스트 로직."""

import re

MAX_HITS_PER_FILE = 30


def normalize_text(value):
    """검색 및 결과 표시용으로 문자열을 정리한다."""
    if value is None:
        return ""
    text = str(value)
    return re.sub(r"\s+", " ", text).strip()


def make_snippet(text, start, end, radius=70):
    """일치 위치 전후 문맥을 짧게 잘라 반환한다."""
    flat = normalize_text(text)
    if not flat:
        return ""
    start = max(0, min(start, len(flat)))
    end = max(start, min(end, len(flat)))
    left = max(0, start - radius)
    right = min(len(flat), end + radius)
    prefix = "…" if left > 0 else ""
    suffix = "…" if right < len(flat) else ""
    return prefix + flat[left:right] + suffix


def find_matches(text, pattern):
    """
    text 안에서 패턴을 찾아 (시작, 끝, 스니펫) 목록을 반환한다.
    pattern은 re.Pattern 객체다.
    """
    if text is None:
        return []
    text = str(text)
    results = []
    for match in pattern.finditer(text):
        results.append((match.start(), match.end(), make_snippet(text, match.start(), match.end())))
        if len(results) >= MAX_HITS_PER_FILE:
            break
    return results


def _wildcard_to_regex(keyword):
    """
    '*'(임의 길이 문자열)와 '?'(문자 1개)만 지원하는 와일드카드 표현을
    부분 문자열 검색용 정규식 조각으로 바꾼다(전체 일치가 아닌 finditer용).
    """
    parts = []
    for char in keyword:
        if char == "*":
            parts.append(".*")
        elif char == "?":
            parts.append(".")
        else:
            parts.append(re.escape(char))
    return "".join(parts)


def build_pattern(keyword, use_wildcard, case_sensitive):
    flags = 0 if case_sensitive else re.IGNORECASE
    expression = _wildcard_to_regex(keyword) if use_wildcard else re.escape(keyword)
    return re.compile(expression, flags)


def parse_extension_filter(text):
    """
    "txt, .pdf *.docx" 같은 입력을 {'.txt', '.pdf', '.docx'} 형태로 정규화한다.
    비어 있으면 빈 집합(필터 없음, 전체 검색)을 반환한다.
    """
    if not text:
        return set()

    tokens = re.split(r"[,;\s]+", text.strip())
    extensions = set()
    for token in tokens:
        token = token.strip().lstrip("*")
        if not token:
            continue
        if not token.startswith("."):
            token = "." + token
        extensions.add(token.lower())
    return extensions
