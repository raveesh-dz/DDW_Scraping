"""
Keyword loading and matching for the DDW scraper.

Loads the four keyword lists (Indication, Company, Product Name, MOA) from
``DDW Keywords.xlsx`` and matches them against session text fields to populate
the corresponding columns in the output Excel.

Typical usage
-------------
::

    from ddw_scraper import keywords as kw

    patterns = kw.build_patterns(kw.load_keywords())
    matches  = kw.match_keywords(["Crohn's Disease", "Liver"], patterns)
    # → {"Indication": "Crohnsdisease", "Company": "", "Product Name": "", "MOA": ""}
"""

import re
from pathlib import Path

import openpyxl

from ddw_scraper import config

# Column names in the keywords file (and in the output Excel), in order.
COLUMNS = ["Indication", "Company", "Product Name", "MOA"]

# ---------------------------------------------------------------------------
# Unicode normalisation
# ---------------------------------------------------------------------------

# Map typographic / smart quotes to their plain ASCII equivalents so that
# keywords and scraped text are compared on a level playing field.
_UNICODE_NORMALIZE = str.maketrans({
    "\u2019": "'",   # RIGHT SINGLE QUOTATION MARK  →  apostrophe
    "\u2018": "'",   # LEFT SINGLE QUOTATION MARK   →  apostrophe
    "\u201c": '"',   # LEFT DOUBLE QUOTATION MARK   →  quote
    "\u201d": '"',   # RIGHT DOUBLE QUOTATION MARK  →  quote
    "\u2013": "-",   # EN DASH                       →  hyphen
    "\u2014": "-",   # EM DASH                       →  hyphen
})

# ---------------------------------------------------------------------------
# Per-keyword matching overrides
# ---------------------------------------------------------------------------

# Keywords matched case-sensitively (re.IGNORECASE is NOT applied).
# Use this for names whose exact capitalisation distinguishes them from
# common English words (e.g. "selectION" vs "selection").
CASE_SENSITIVE_KEYWORDS: frozenset[str] = frozenset({"selectION", "Tract"})

# Extra fixed-width negative-lookbehind fragments prepended to the pattern
# for specific keywords, to block known false-positive contexts that the
# generic word-boundary alone cannot prevent.
# "Tract" (company) must not fire inside "Biliary Tract" or "GI Tract".
_EXTRA_LOOKBEHIND: dict[str, str] = {
    "Tract": r"(?<!Biliary )(?<!biliary )(?<!GI )(?<!gi )",
}


# ---------------------------------------------------------------------------
# Loading
# ---------------------------------------------------------------------------

def load_keywords(path=None) -> dict[str, list[str]]:
    """Read keywords from *path* (defaults to ``config.KEYWORDS_FILE``).

    Returns a dict ``{column_name: [keyword, ...]}``.  The first header row
    is skipped; ``None`` cells and blank strings are ignored.

    Typographic characters (curly apostrophes, en/em dashes) in keyword
    cells are normalised to plain ASCII equivalents so that patterns compiled
    from them can match the normalised scraped text.
    """
    path = Path(path or config.KEYWORDS_FILE)
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active

    keyword_map: dict[str, list[str]] = {col: [] for col in COLUMNS}

    for row in ws.iter_rows(min_row=2, values_only=True):
        for col_idx, col_name in enumerate(COLUMNS):
            raw = row[col_idx] if col_idx < len(row) else None
            if raw is not None:
                keyword = str(raw).strip().translate(_UNICODE_NORMALIZE)
                if keyword:
                    keyword_map[col_name].append(keyword)

    wb.close()
    return keyword_map


# ---------------------------------------------------------------------------
# Pattern compilation
# ---------------------------------------------------------------------------

def _compile(keyword: str) -> re.Pattern:
    """Return a word-boundary-aware pattern for *keyword*.

    Uses lookahead / lookbehind on ``[a-zA-Z0-9]`` rather than ``\\b`` so
    that keywords containing non-ASCII characters (Greek letters, hyphens,
    special chars like ``&``) are handled correctly.

    Keywords listed in ``CASE_SENSITIVE_KEYWORDS`` are compiled without
    ``re.IGNORECASE`` so their exact capitalisation is required — preventing
    ambiguous short names (e.g. ``selectION``) from matching common words.

    Keywords listed in ``_EXTRA_LOOKBEHIND`` have additional fixed-width
    negative-lookbehind assertions prepended to suppress known false-positive
    contexts (e.g. ``Tract`` must not fire inside ``Biliary Tract``).
    """
    escaped = re.escape(keyword)
    extra_lb = _EXTRA_LOOKBEHIND.get(keyword, "")
    pattern = extra_lb + r"(?<![a-zA-Z0-9])" + escaped + r"(?![a-zA-Z0-9])"
    flags = 0 if keyword in CASE_SENSITIVE_KEYWORDS else re.IGNORECASE
    return re.compile(pattern, flags)


def build_patterns(
    keyword_map: dict[str, list[str]],
) -> dict[str, list[tuple[str, re.Pattern]]]:
    """Pre-compile regexes for every keyword.

    Returns ``{column_name: [(keyword, compiled_pattern), ...]}``.
    Call this once per run; reuse the result for every session.
    """
    return {
        col: [(kw, _compile(kw)) for kw in keywords]
        for col, keywords in keyword_map.items()
    }


# ---------------------------------------------------------------------------
# Matching
# ---------------------------------------------------------------------------

def _deduplicate_substrings(matched: list[str]) -> list[str]:
    """Remove any keyword that is a case-insensitive substring of another
    matched keyword in the same column.

    Example: if both ``"Inflammatory Bowel Disease"`` and
    ``"Inflammatory Bowel Diseases"`` matched, the shorter one is suppressed
    because it is contained within the longer match.
    """
    if len(matched) <= 1:
        return matched
    lowered = [k.lower() for k in matched]
    return [
        kw for i, kw in enumerate(matched)
        if not any(
            lowered[i] in lowered[j]
            for j in range(len(matched)) if j != i
        )
    ]


def match_keywords(
    text_fields: list[str],
    patterns: dict[str, list[tuple[str, re.Pattern]]],
) -> dict[str, str]:
    """Search *text_fields* for every keyword and return matched values.

    *text_fields* is a list of strings (session name, track, format, …).
    They are joined into one combined string before matching so a keyword
    only needs to appear in any field to be considered a match.

    Typographic characters in the text are normalised to plain ASCII before
    matching (e.g. curly apostrophe U+2019 → straight apostrophe) so that
    scraped content from the website is compared fairly against keywords.

    Matched keywords within the same column are deduplicated: if keyword A
    is a substring of keyword B and both matched, only B is kept (e.g.
    ``"Inflammatory Bowel Disease"`` is suppressed when
    ``"Inflammatory Bowel Diseases"`` also matched the same text).

    Returns ``{column_name: comma_separated_matched_keywords}``.
    Columns with no matches contain an empty string.
    """
    combined = " ".join(
        f.translate(_UNICODE_NORMALIZE) for f in text_fields if f
    )
    result: dict[str, str] = {}
    for col, kw_patterns in patterns.items():
        matched = [kw for kw, pat in kw_patterns if pat.search(combined)]
        matched = _deduplicate_substrings(matched)
        result[col] = ", ".join(matched)
    return result
