"""Core engine for the Entity Normalizer.

The workbook is never edited cell by cell. Instead the tool keeps:

* a read-only snapshot of every original name column (``<column> Raw``,
  appended at the end of the sheet), and
* an ordered list of rename steps (``original names -> normalized name``).

Current values are always derived by replaying the rename steps over the
cleaned raw values. That design gives unlimited undo, a complete audit log,
and an alias table that can be saved and reapplied to the next case.
"""

from __future__ import annotations

import io
import re
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from datetime import datetime
from typing import Iterable

import pandas as pd

RAW_SUFFIX = " Raw"
LOG_SHEET = "Mapping Log"
ALIAS_SHEET = "Alias Table"

# Which columns hold counterparty names, per sheet. A sheet can optionally be
# limited to rows where ``filter_column`` equals ``filter_value``.
DEFAULT_CONFIG: dict[str, dict] = {
    "Wire": {"columns": ["Orig Acct Name", "Bene Acct Name"]},
    "Combined": {
        "columns": ["Description", "Account Name"],
        "filter_column": "Tab Name",
        "filter_value": "Wire",
    },
}

# Structured SWIFT name/address fields prefix the name line with "1/".
# Only a prefix at the very start is removed, so a name such as
# "UNIT 21/B HOLDINGS" is left intact.
_SWIFT_NAME_PREFIX = re.compile(r"^\s*1/\s*")
_WHITESPACE = re.compile(r"\s+")


def strip_swift_prefix(value):
    """Remove a leading SWIFT "1/" name-line prefix and tidy whitespace."""
    if not isinstance(value, str):
        return value
    cleaned = _SWIFT_NAME_PREFIX.sub("", value)
    return _WHITESPACE.sub(" ", cleaned).strip()


def match_key(value) -> str:
    """Key used to decide whether two cells hold the same name.

    Case-insensitive and whitespace-insensitive, so "Acme Corp" and
    "ACME  CORP " are treated as one entity everywhere: in the name list
    *and* when renaming.
    """
    if not isinstance(value, str):
        return ""
    return _WHITESPACE.sub(" ", value).strip().casefold()


@dataclass
class RenameStep:
    source_keys: frozenset[str]
    source_names: tuple[str, ...]
    target: str
    timestamp: datetime = field(default_factory=datetime.now)
    cells_changed: int = 0


@dataclass
class _Field:
    sheet: str
    column: str

    @property
    def raw_column(self) -> str:
        return f"{self.column}{RAW_SUFFIX}"

    @property
    def label(self) -> str:
        return f"{self.sheet} · {self.column}"


class NormalizerSession:
    """Holds one workbook and every rename applied to it."""

    def __init__(
        self,
        sheets: dict[str, pd.DataFrame],
        config: dict[str, dict] | None = None,
        rename_all_rows: bool = False,
    ):
        self.config = config or DEFAULT_CONFIG
        self.rename_all_rows = rename_all_rows
        self.sheets: dict[str, pd.DataFrame] = {}
        self.fields: list[_Field] = []
        self.steps: list[RenameStep] = []
        self.warnings: list[str] = []
        self.restored_steps = 0

        for name, df in sheets.items():
            self.sheets[name] = df.copy()

        for sheet, spec in self.config.items():
            df = self.sheets.get(sheet)
            if df is None:
                self.warnings.append(f"Sheet '{sheet}' not found; skipped.")
                continue
            filter_col = spec.get("filter_column")
            if filter_col and filter_col not in df.columns:
                self.warnings.append(
                    f"Column '{filter_col}' not found on '{sheet}'; "
                    "all rows on that sheet will be used."
                )
            for column in spec["columns"]:
                if column not in df.columns:
                    self.warnings.append(f"Column '{column}' not found on '{sheet}'.")
                    continue
                f = _Field(sheet, column)
                # Snapshot the original values once, at the END of the sheet.
                # (Appending is far faster and more stable than inserting
                # columns mid-sheet.) Re-running never overwrites a snapshot.
                if f.raw_column not in df.columns:
                    df[f.raw_column] = df[column]
                self.fields.append(f)

        self._base: dict[str, pd.Series] = {
            f.label: self.sheets[f.sheet][f.raw_column].map(strip_swift_prefix)
            for f in self.fields
        }

    # ------------------------------------------------------------------
    # Loading
    # ------------------------------------------------------------------
    @classmethod
    def from_excel(cls, source, restore: bool = True, **kwargs) -> "NormalizerSession":
        """Open a workbook.

        If it was exported by this tool, its "Alias Table" sheet is reapplied
        (``restore=True``) so work can resume exactly where it stopped.
        """
        sheets = pd.read_excel(source, sheet_name=None)
        aliases = sheets.pop(ALIAS_SHEET, None)
        sheets.pop(LOG_SHEET, None)
        session = cls(sheets, **kwargs)
        if restore and aliases is not None and not aliases.empty:
            session.restored_steps = session.apply_alias_table(aliases)
        return session

    # ------------------------------------------------------------------
    # Scope
    # ------------------------------------------------------------------
    def _in_scope(self, f: _Field) -> pd.Series:
        df = self.sheets[f.sheet]
        spec = self.config.get(f.sheet, {})
        col, val = spec.get("filter_column"), spec.get("filter_value")
        if self.rename_all_rows or not col or col not in df.columns:
            return pd.Series(True, index=df.index)
        return df[col].astype(str).str.strip().str.casefold() == str(val).casefold()

    # ------------------------------------------------------------------
    # Resolving current values
    # ------------------------------------------------------------------
    def _resolve(self, value, steps: Iterable[RenameStep] | None = None):
        if not isinstance(value, str) or not value:
            return value
        current = value
        for step in self.steps if steps is None else steps:
            if match_key(current) in step.source_keys:
                current = step.target
        return current

    def _current_series(self, f: _Field, steps=None) -> pd.Series:
        base = self._base[f.label]
        scope = self._in_scope(f)
        distinct = base[scope].dropna().unique()
        lookup = {v: self._resolve(v, steps) for v in distinct}
        out = base.copy()
        out[scope] = base[scope].map(lambda v: lookup.get(v, v))
        return out

    def current_sheets(self) -> dict[str, pd.DataFrame]:
        """Every sheet with normalized name columns (raw columns untouched)."""
        result = {name: df.copy() for name, df in self.sheets.items()}
        for f in self.fields:
            result[f.sheet][f.column] = self._current_series(f)
        return result

    # ------------------------------------------------------------------
    # The entity list (left-hand panel)
    # ------------------------------------------------------------------
    def distinct_names(self) -> pd.DataFrame:
        spellings: dict[str, Counter] = defaultdict(Counter)
        locations: dict[str, Counter] = defaultdict(Counter)
        for f in self.fields:
            values = self._current_series(f)[self._in_scope(f)].dropna()
            for v in values:
                if not isinstance(v, str) or not v.strip():
                    continue
                k = match_key(v)
                spellings[k][v.strip()] += 1
                locations[k][f.label] += 1

        rows = []
        for k, spelled in spellings.items():
            rows.append(
                {
                    "key": k,
                    "Name": spelled.most_common(1)[0][0],
                    "Count": sum(spelled.values()),
                    "Spellings": len(spelled),
                    "Appears in": "; ".join(
                        f"{loc} ({n})" for loc, n in locations[k].most_common()
                    ),
                }
            )
        df = pd.DataFrame(rows, columns=["key", "Name", "Count", "Spellings", "Appears in"])
        return df.sort_values("Name", key=lambda s: s.str.casefold()).reset_index(drop=True)

    def suggested_name(self, keys: Iterable[str]) -> str:
        """Most frequent spelling among the selected names (prefill text)."""
        names = self.distinct_names().set_index("key")
        keys = [k for k in keys if k in names.index]
        if not keys:
            return ""
        best = names.loc[keys].sort_values("Count", ascending=False)
        return str(best.iloc[0]["Name"])

    # ------------------------------------------------------------------
    # Rename / undo
    # ------------------------------------------------------------------
    def _changed_cells(self, steps_before, steps_after) -> int:
        total = 0
        for f in self.fields:
            before = self._current_series(f, steps_before)
            after = self._current_series(f, steps_after)
            total += int((before.fillna("") != after.fillna("")).sum())
        return total

    def rename(self, names: Iterable[str], target: str) -> RenameStep:
        """Rename every cell matching any of ``names`` (any case) to ``target``."""
        target = _WHITESPACE.sub(" ", str(target)).strip()
        names = [n for n in names if isinstance(n, str) and n.strip()]
        if not target:
            raise ValueError("Enter a new name first.")
        if not names:
            raise ValueError("Select at least one name to rename.")
        step = RenameStep(
            source_keys=frozenset(match_key(n) for n in names),
            source_names=tuple(names),
            target=target,
        )
        step.cells_changed = self._changed_cells(self.steps, self.steps + [step])
        self.steps.append(step)
        return step

    def undo(self) -> RenameStep | None:
        """Reverse the most recent rename. Can be repeated back to the start."""
        return self.steps.pop() if self.steps else None

    # ------------------------------------------------------------------
    # Audit trail and reusable aliases
    # ------------------------------------------------------------------
    def mapping_log(self) -> pd.DataFrame:
        return pd.DataFrame(
            [
                {
                    "Step": i,
                    "Timestamp": s.timestamp.strftime("%Y-%m-%d %H:%M:%S"),
                    "Original names": " | ".join(s.source_names),
                    "Normalized to": s.target,
                    "Cells changed": s.cells_changed,
                }
                for i, s in enumerate(self.steps, start=1)
            ],
            columns=["Step", "Timestamp", "Original names", "Normalized to", "Cells changed"],
        )

    def alias_table(self) -> pd.DataFrame:
        """Original spelling -> final normalized name, for every changed name."""
        pairs = {}
        for f in self.fields:
            base = self._base[f.label][self._in_scope(f)].dropna()
            for v in base.unique():
                final = self._resolve(v)
                if isinstance(v, str) and final != v:
                    pairs[v] = final
        return (
            pd.DataFrame(sorted(pairs.items()), columns=["Original", "Normalized"])
            if pairs
            else pd.DataFrame(columns=["Original", "Normalized"])
        )

    def apply_alias_table(self, aliases: pd.DataFrame) -> int:
        """Reapply a saved alias table (columns Original, Normalized).

        Returns the number of rename steps added.
        """
        missing = {"Original", "Normalized"} - set(aliases.columns)
        if missing:
            raise ValueError(f"Alias table is missing column(s): {', '.join(sorted(missing))}")
        added = 0
        clean = aliases.dropna(subset=["Original", "Normalized"])
        for target, group in clean.groupby("Normalized", sort=False):
            step = self.rename(group["Original"].astype(str).tolist(), str(target))
            if step.cells_changed == 0:
                self.steps.pop()  # nothing in this workbook matched
            else:
                added += 1
        return added

    # ------------------------------------------------------------------
    # Export
    # ------------------------------------------------------------------
    def to_excel_bytes(self) -> bytes:
        buffer = io.BytesIO()
        self.export(buffer)
        return buffer.getvalue()

    def export(self, target) -> None:
        """Write the normalized workbook with formatting and audit sheets."""
        from .export import write_workbook

        write_workbook(self, target)
