"""Suggest groups of names that probably refer to the same entity.

Two passes:

1. **Exact on a comparison key.** Names are upper-cased, stripped of
   punctuation and legal suffixes (INC, LLC, N.A. ...), and common
   abbreviations are expanded (BK -> BANK, INTL -> INTERNATIONAL).
   "Harbor Point Logistics, LLC" and "HARBOR POINT LOGISTICS" share a key.
2. **Fuzzy.** Remaining keys are compared with a token-sort similarity
   score, which also catches word-order flips ("SMITH, JOHN" vs
   "JOHN SMITH") and small typos.

Suggestions are never applied automatically; the investigator approves them.
Uses rapidfuzz when installed, otherwise Python's built-in difflib.
"""

from __future__ import annotations

import re
from collections import defaultdict
from difflib import SequenceMatcher

try:  # Fast path
    from rapidfuzz import fuzz as _rf_fuzz

    def _token_sort(a: str, b: str) -> float:
        return _rf_fuzz.token_sort_ratio(a, b)

    def _plain(a: str, b: str) -> float:
        return _rf_fuzz.ratio(a, b)

    BACKEND = "rapidfuzz"
except ImportError:  # pragma: no cover - exercised when rapidfuzz is absent

    def _plain(a: str, b: str) -> float:
        return SequenceMatcher(None, a, b).ratio() * 100

    def _token_sort(a: str, b: str) -> float:
        return _plain(" ".join(sorted(a.split())), " ".join(sorted(b.split())))

    BACKEND = "difflib"


def _similarity(a: str, b: str) -> float:
    """Best of word-order-insensitive and spacing-insensitive similarity.

    Token sort catches "SMITH JOHN" vs "JOHN SMITH"; comparing with spaces
    removed catches "BLUE WATER" vs "BLUEWATER".
    """
    return max(_token_sort(a, b), _plain(a.replace(" ", ""), b.replace(" ", "")))


LEGAL_SUFFIXES = {
    "INC", "INCORPORATED", "LLC", "LLP", "LP", "LTD", "LIMITED", "CORP",
    "CORPORATION", "CO", "COMPANY", "NA", "PLC", "GMBH", "AG", "SA", "PC",
    "PLLC", "THE",
}

ABBREVIATIONS = {
    "BK": "BANK",
    "BNK": "BANK",
    "NATL": "NATIONAL",
    "INTL": "INTERNATIONAL",
    "INTERNATL": "INTERNATIONAL",
    "SVCS": "SERVICES",
    "SVC": "SERVICES",
    "SERVICE": "SERVICES",
    "MGMT": "MANAGEMENT",
    "MFG": "MANUFACTURING",
    "ASSOC": "ASSOCIATES",
    "ASSN": "ASSOCIATION",
    "GRP": "GROUP",
    "HLDGS": "HOLDINGS",
    "HOLDING": "HOLDINGS",
    "TR": "TRUST",
    "INV": "INVESTMENTS",
    "INVESTMENT": "INVESTMENTS",
    "&": "AND",
}

_PUNCT = re.compile(r"[^\w\s&]")


def comparison_key(name: str) -> str:
    """Aggressively simplified form of a name, used only for matching."""
    if not isinstance(name, str):
        return ""
    text = name.upper().replace("&", " & ")
    text = re.sub(r"^\s*1/\s*", "", text)
    text = _PUNCT.sub(" ", text)
    tokens = [ABBREVIATIONS.get(t, t) for t in text.split()]

    # Join runs of single letters: "J P MORGAN" -> "JP MORGAN", "N A" -> "NA"
    joined: list[str] = []
    in_initials = False
    for t in tokens:
        is_initial = len(t) == 1 and t.isalpha()
        if is_initial and in_initials:
            joined[-1] += t
        else:
            joined.append(t)
        in_initials = is_initial

    core = [t for t in joined if t not in LEGAL_SUFFIXES]
    return " ".join(core or joined)


class _UnionFind:
    def __init__(self, items):
        self.parent = {i: i for i in items}

    def find(self, x):
        while self.parent[x] != x:
            self.parent[x] = self.parent[self.parent[x]]
            x = self.parent[x]
        return x

    def union(self, a, b):
        ra, rb = self.find(a), self.find(b)
        if ra != rb:
            self.parent[rb] = ra


def suggest_clusters(names, counts=None, threshold: float = 88.0, max_names: int = 5000):
    """Group likely-duplicate names.

    Parameters
    ----------
    names : sequence of str
        Distinct names (e.g. ``session.distinct_names()["Name"]``).
    counts : sequence of int, optional
        Occurrences of each name, used to rank clusters and pick the
        suggested canonical spelling.
    threshold : float
        Minimum similarity (0-100) for the fuzzy pass.

    Returns
    -------
    list of dict, largest total count first, each with keys
    ``names``, ``counts``, ``total``, ``suggested``, ``score``.
    """
    names = list(names)
    counts = list(counts) if counts is not None else [1] * len(names)
    if len(names) > max_names:
        raise ValueError(f"Too many names for pairwise matching ({len(names)} > {max_names}).")

    keys = [comparison_key(n) for n in names]
    uf = _UnionFind(range(len(names)))
    best_score: dict[tuple[int, int], float] = {}

    # Pass 1: identical comparison keys
    by_key: dict[str, list[int]] = defaultdict(list)
    for i, k in enumerate(keys):
        if k:
            by_key[k].append(i)
    for members in by_key.values():
        for other in members[1:]:
            uf.union(members[0], other)

    # Pass 2: fuzzy, one representative per distinct key, blocked by first
    # letter of any token so person names in either order still meet.
    reps = {k: members[0] for k, members in by_key.items()}
    blocks: dict[str, set[str]] = defaultdict(set)
    for k in reps:
        for token in set(k.split()):
            if len(token) >= 3:
                blocks[token[:3]].add(k)

    compared: set[tuple[str, str]] = set()
    for block in blocks.values():
        block = sorted(block)
        for i, a in enumerate(block):
            for b in block[i + 1:]:
                if (a, b) in compared:
                    continue
                compared.add((a, b))
                score = _similarity(a, b)
                if score >= threshold:
                    ia, ib = reps[a], reps[b]
                    uf.union(ia, ib)
                    best_score[(ia, ib)] = score

    clusters: dict[int, list[int]] = defaultdict(list)
    for i in range(len(names)):
        clusters[uf.find(i)].append(i)

    results = []
    for members in clusters.values():
        if len(members) < 2:
            continue
        # Suggested spelling = the most "central" name (highest average
        # similarity to the rest of the cluster), so a frequent typo is not
        # proposed as canonical. Ties go to the more frequent spelling.
        centrality = {
            i: sum(_similarity(keys[i], keys[j]) for j in members if j != i) for i in members
        }
        members.sort(key=lambda i: (-round(centrality[i], 1), -counts[i], names[i]))
        root_scores = [s for (a, b), s in best_score.items() if a in members and b in members]
        results.append(
            {
                "names": [names[i] for i in members],
                "counts": [counts[i] for i in members],
                "total": sum(counts[i] for i in members),
                "suggested": names[members[0]],
                "score": round(min(root_scores), 1) if root_scores else 100.0,
            }
        )
    results.sort(key=lambda c: (-c["total"], c["suggested"]))
    return results
