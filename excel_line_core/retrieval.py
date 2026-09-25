"""Deterministic local retrieval helpers for the Excel memory tree.

The store remains the audit source of truth.  This module provides a small,
optional ranking layer that is fast enough for ``prefetch`` and deliberately has
no vector/embedding dependency.
"""
from __future__ import annotations

import math
import re
import unicodedata
from collections import Counter
from datetime import datetime
from typing import Any, Dict, Iterable, List, Mapping, Optional, Sequence

_TOKEN_RE = re.compile(r"[^\W_]+", re.UNICODE)

# Function words only: domain words (model, token, accounting, …) stay searchable.
STOP_WORDS = {
    "a", "an", "and", "are", "as", "at", "be", "by", "can", "do", "for", "from",
    "how", "i", "in", "is", "it", "of", "on", "or", "that", "the", "this", "to",
    "was", "what", "when", "where", "which", "who", "why", "with", "you", "your",
    "có", "và", "hay", "cũng", "còn", "để", "trong", "với", "về", "của", "cho",
    "đến", "bằng", "qua", "khi", "nếu", "thì", "mà", "nhưng", "hoặc", "vì", "vậy",
    "do", "đó", "như", "nên", "lại", "theo", "tôi", "là", "từ", "gì", "được",
    "nào", "bao", "nhiêu", "ra", "vào", "lên", "xuống", "một", "các", "những",
}


def normalize(value: Any) -> str:
    """Case-fold and normalize Unicode without destroying non-Latin text."""
    text = unicodedata.normalize("NFKC", str(value or "")).casefold()
    return " ".join(text.split())


def tokenize(value: Any) -> List[str]:
    text = normalize(value)
    return [token for token in _TOKEN_RE.findall(text) if len(token) > 1 and token not in STOP_WORDS]


def extract_keywords(query: Any, limit: int = 8) -> List[str]:
    """Return de-duplicated query terms while preserving meaningful words."""
    terms: List[str] = []
    seen = set()
    for token in tokenize(query):
        if token not in seen:
            seen.add(token)
            terms.append(token)
    return terms[:limit]


def row_text(row: Mapping[str, Any]) -> str:
    metadata = row.get("metadata") if isinstance(row.get("metadata"), Mapping) else {}
    return normalize(" ".join(str(row.get(key) or metadata.get(key) or "") for key in (
        "title", "content", "brief", "tags", "branch", "source", "profile_id",
        "session_id", "entity_links", "entities",
    )))


def _term_scores(query_terms: Sequence[str], text: str) -> Dict[str, float]:
    text_tokens = tokenize(text)
    counts = Counter(text_tokens)
    length = max(1, len(text_tokens))
    total = max(1, sum(counts.values()))
    scores: Dict[str, float] = {}
    for term in query_terms:
        term_norm = normalize(term)
        if not term_norm:
            continue
        if term_norm in counts:
            # Presence is more valuable than repetition; cap repetition so a
            # noisy long row cannot dominate ranking.
            scores[term_norm] = min(3.0, 1.0 + math.log1p(counts[term_norm] - 1)) * (1.0 + 1.0 / math.sqrt(length))
        else:
            # Prefix matching helps Vietnamese compounds and English stems.
            prefix_hits = sum(1 for token in counts if token.startswith(term_norm) or term_norm.startswith(token))
            if prefix_hits:
                scores[term_norm] = 0.55 * (1.0 + math.log1p(prefix_hits))
    return scores


def score_row(query: Any, row: Mapping[str, Any], *, metadata: Optional[Mapping[str, Any]] = None) -> float:
    """Return a bounded relevance score (0..1) for one row.

    Title/tags/entity matches receive modest boosts; confidence and recency are
    only tie-breakers, never fabricated when absent.
    """
    terms = extract_keywords(query, limit=12)
    if not terms:
        return 0.0
    text = row_text(row)
    scores = _term_scores(terms, text)
    if not scores:
        return 0.0
    raw = sum(scores.values())
    # Normalize against a generous upper bound to keep output deterministic and
    # readable while preserving ordering.
    score = min(1.0, raw / max(3.0, len(terms) * 2.5))
    title_terms = set(tokenize(row.get("title")))
    tag_terms = set(tokenize(row.get("tags")))
    entity_terms = set(tokenize(row.get("entity_links") or row.get("entities") or
                                (row.get("metadata") or {}).get("entity_links")
                                if isinstance(row.get("metadata"), Mapping) else ""))
    for term in scores:
        if term in title_terms:
            score += 0.08
        if term in tag_terms:
            score += 0.04
        if term in entity_terms:
            score += 0.06
    try:
        confidence = float((metadata or row).get("confidence", 1.0))
    except (TypeError, ValueError):
        confidence = 1.0
    score *= max(0.5, min(1.25, confidence))
    return min(1.0, round(score, 6))


def rank_rows(query: Any, rows: Iterable[Mapping[str, Any]], *, limit: int = 10,
              filters: Optional[Mapping[str, Any]] = None) -> List[Dict[str, Any]]:
    """Filter and rank rows, returning copies with ``relevance`` and ``match_terms``."""
    terms = set(extract_keywords(query, limit=12))
    filters = filters or {}
    ranked: List[Dict[str, Any]] = []
    for row in rows:
        if not _matches_filters(row, filters):
            continue
        relevance = score_row(query, row)
        if relevance <= 0:
            continue
        copy = dict(row)
        copy["relevance"] = relevance
        row_terms = set(tokenize(row_text(row)))
        copy["match_terms"] = sorted(terms & row_terms)
        ranked.append(copy)
    ranked.sort(key=lambda item: (-item["relevance"], -_updated_key(item), int(item.get("id") or 0)))
    return ranked[: max(0, int(limit or 0))]


def _updated_key(row: Mapping[str, Any]) -> float:
    text = str(row.get("updated") or row.get("created") or "")
    if not text:
        return 0.0
    try:
        return datetime.fromisoformat(text.replace("Z", "+00:00")).timestamp()
    except (TypeError, ValueError, OverflowError):
        # Legacy/custom timestamps retain deterministic lexical ordering.
        return float(sum((index + 1) * ord(char) for index, char in enumerate(text[:64])))


def _matches_filters(row: Mapping[str, Any], filters: Mapping[str, Any]) -> bool:
    for key, wanted in filters.items():
        if wanted in (None, "", [], {}):
            continue
        metadata = row.get("metadata") if isinstance(row.get("metadata"), Mapping) else {}
        value = row.get(key, metadata.get(key))
        if isinstance(wanted, (list, tuple, set)):
            values = {normalize(item) for item in wanted}
            actual = normalize(value)
            if actual not in values:
                # Entity lists may be comma/semicolon separated.
                parts = {normalize(item) for item in str(value or "").replace(";", ",").split(",")}
                if not values.intersection(parts):
                    return False
        elif normalize(value) != normalize(wanted):
            return False
    return True


__all__ = [
    "STOP_WORDS", "extract_keywords", "normalize", "rank_rows", "row_text", "score_row", "tokenize",
]
