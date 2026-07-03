"""점수 기반 표 선별기.

여러 기준(헤더 일치율, 제목 키워드, 열 개수, 셀 패턴)에 가중치를 부여해
합산 점수가 threshold 이상인 표만 추출 대상으로 선정한다.
"""
from __future__ import annotations

import re

from .models import MatchResult, Table

DEFAULT_WEIGHTS = {"header": 50, "title": 30, "cols": 10, "pattern": 10}


def normalize(s: str) -> str:
    """공백/줄바꿈 제거 후 소문자화 ("품 명" == "품명")."""
    return re.sub(r"\s+", "", s).lower()


def score_table(table: Table, rule: dict) -> tuple[float, list[str]]:
    """표 하나를 규칙 하나와 대조하여 (점수, 매칭 근거) 반환."""
    weights = {**DEFAULT_WEIGHTS, **(rule.get("weights") or {})}
    score = 0.0
    reasons: list[str] = []

    # 1) 헤더 일치율
    expected = {normalize(h) for h in (rule.get("expected_headers") or []) if h}
    if expected:
        actual = {normalize(h) for h in table.header_row() if h}
        matched = expected & actual
        ratio = len(matched) / len(expected)
        if ratio > 0:
            score += ratio * weights["header"]
            reasons.append(f"헤더 {len(matched)}/{len(expected)} 일치 (+{ratio * weights['header']:.0f})")

    # 2) 캡션/직전 문단 키워드
    context = normalize(" ".join([table.caption, *table.preceding_texts]))
    for kw in (rule.get("title_keywords") or []):
        if kw and normalize(kw) in context:
            score += weights["title"]
            reasons.append(f"제목 키워드 '{kw}' (+{weights['title']})")
            break

    # 3) 열 개수 일치
    expected_cols = rule.get("expected_cols")
    if expected_cols and table.n_cols == int(expected_cols):
        score += weights["cols"]
        reasons.append(f"열 개수 {table.n_cols} 일치 (+{weights['cols']})")

    # 4) 셀 내 정규식 패턴
    pattern = rule.get("cell_pattern")
    if pattern:
        if any(re.search(pattern, c.text) for c in table.cells):
            score += weights["pattern"]
            reasons.append(f"셀 패턴 '{pattern}' 발견 (+{weights['pattern']})")

    return score, reasons


def select_tables(tables: list[Table], rules: list[dict], logger=None) -> list[MatchResult]:
    """모든 표를 모든 규칙과 대조하여 threshold 를 넘긴 (표, 규칙) 쌍을 반환.

    한 표가 여러 규칙에 매칭되면 점수가 가장 높은 규칙 하나만 채택한다.
    """
    log = logger.info if logger else print
    results: list[MatchResult] = []

    for table in tables:
        best: MatchResult | None = None
        for rule in rules:
            threshold = float(rule.get("threshold", 60))
            score, reasons = score_table(table, rule)
            log(f"  [표 #{table.index}] 규칙 '{rule.get('name')}' → {score:.0f}점"
                f" (기준 {threshold:.0f}) {'; '.join(reasons) if reasons else '매칭 없음'}")
            if score >= threshold and (best is None or score > best.score):
                best = MatchResult(table=table, rule_name=rule.get("name", "rule"),
                                   score=score, reasons=reasons)
        if best:
            results.append(best)
            log(f"  => 표 #{best.table.index} 선정 (규칙 '{best.rule_name}', {best.score:.0f}점)")

    return results
