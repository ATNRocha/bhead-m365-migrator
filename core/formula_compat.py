from __future__ import annotations

import math
import re
import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterable

from .profile import normalize_text
from .xlsx_ooxml import M, _workbook_sheet_map

ERROR_TOKENS = ("#REF!", "#N/A", "#DIV/0!", "#VALUE!", "#NAME?", "#NOME?", "#NUM!", "#NULL!")
IFS_RE = re.compile(r"(?i)(?<![A-Z0-9_])(?:_xlfn\.)?IFS\s*\(")
XLOOKUP_RE = re.compile(r"(?i)(?<![A-Z0-9_])(?:_xlfn\.)?XLOOKUP\s*\(")
GOOGLE_WRAPPER_RE = re.compile(r"(?i)\b(?:ARRAYFORMULA|ARRAY_CONSTRAIN)\s*\(")


@dataclass
class FormulaIssue:
    sheet: str
    cell: str
    kind: str
    formula: str = ""
    detail: str = ""


@dataclass
class FormulaSheetSummary:
    name: str
    formula_cells: int = 0
    explicit_formulas: int = 0
    shared_formula_followers: int = 0
    array_formulas: int = 0
    single_cell_array_formulas: int = 0
    multi_cell_array_formulas: int = 0
    ifs_formulas: int = 0
    xlookup_formulas: int = 0
    google_wrapper_formulas: int = 0
    broken_reference_formulas: int = 0
    cached_error_cells: int = 0
    formula_text_by_cell: dict[str, str] = field(default_factory=dict)
    cached_value_by_cell: dict[str, str] = field(default_factory=dict)
    issues: list[FormulaIssue] = field(default_factory=list)


@dataclass
class FormulaCompatibilityAnalysis:
    sheets: list[FormulaSheetSummary] = field(default_factory=list)
    html_reference_values: int = 0
    html_error_cells: list[FormulaIssue] = field(default_factory=list)
    source_value_mismatch_count: int = 0
    source_value_mismatches: list[FormulaIssue] = field(default_factory=list)

    @property
    def formula_cells(self) -> int:
        return sum(s.formula_cells for s in self.sheets)

    @property
    def explicit_formulas(self) -> int:
        return sum(s.explicit_formulas for s in self.sheets)

    @property
    def array_formulas(self) -> int:
        return sum(s.array_formulas for s in self.sheets)

    @property
    def single_cell_array_formulas(self) -> int:
        return sum(s.single_cell_array_formulas for s in self.sheets)

    @property
    def multi_cell_array_formulas(self) -> int:
        return sum(s.multi_cell_array_formulas for s in self.sheets)

    @property
    def ifs_formulas(self) -> int:
        return sum(s.ifs_formulas for s in self.sheets)

    @property
    def xlookup_formulas(self) -> int:
        return sum(s.xlookup_formulas for s in self.sheets)

    @property
    def google_wrapper_formulas(self) -> int:
        return sum(s.google_wrapper_formulas for s in self.sheets)

    @property
    def broken_reference_formulas(self) -> int:
        return sum(s.broken_reference_formulas for s in self.sheets)

    @property
    def cached_error_cells(self) -> int:
        return sum(s.cached_error_cells for s in self.sheets)

    @property
    def has_source_errors(self) -> bool:
        return bool(self.broken_reference_formulas or self.cached_error_cells or self.html_error_cells)

    @property
    def wanted_cells_by_sheet(self) -> dict[str, set[str]]:
        return {normalize_text(s.name): set(s.formula_text_by_cell) for s in self.sheets if s.formula_text_by_cell}


@dataclass
class FormulaConversionStats:
    ifs_converted: int = 0
    single_cell_arrays_normalized: int = 0
    xlookup_arrays_normalized: int = 0
    formulas_left_for_review: int = 0
    workbook_recalc_enabled: bool = False
    warnings: list[str] = field(default_factory=list)


def _cell_cached_text(cell: ET.Element) -> str:
    v = cell.find(f"{M}v")
    if v is not None and v.text is not None:
        return v.text
    inline = cell.find(f"{M}is")
    if inline is not None:
        return "".join(t.text or "" for t in inline.iter(f"{M}t"))
    return ""


def _is_single_ref(ref: str, cell_ref: str) -> bool:
    cleaned = (ref or "").replace("$", "").strip().upper()
    return bool(cleaned and ":" not in cleaned and cleaned == (cell_ref or "").upper())


def analyze_formula_compatibility(path: str | Path) -> FormulaCompatibilityAnalysis:
    source = Path(path)
    result = FormulaCompatibilityAnalysis()
    with zipfile.ZipFile(source, "r") as zf:
        for sheet_name, xml_path in _workbook_sheet_map(zf):
            root = ET.fromstring(zf.read(xml_path))
            summary = FormulaSheetSummary(sheet_name)
            for cell in root.iter(f"{M}c"):
                cell_ref = cell.attrib.get("r", "")
                f = cell.find(f"{M}f")
                if f is None:
                    continue
                summary.formula_cells += 1
                formula = f.text or ""
                if formula:
                    summary.explicit_formulas += 1
                    summary.formula_text_by_cell[cell_ref] = formula
                    summary.cached_value_by_cell[cell_ref] = _cell_cached_text(cell)
                elif f.attrib.get("t") == "shared":
                    summary.shared_formula_followers += 1

                if f.attrib.get("t") == "array":
                    summary.array_formulas += 1
                    ref = f.attrib.get("ref", "")
                    if _is_single_ref(ref, cell_ref):
                        summary.single_cell_array_formulas += 1
                    else:
                        summary.multi_cell_array_formulas += 1
                        summary.issues.append(FormulaIssue(
                            sheet_name, cell_ref, "ARRAY_MULTI", formula,
                            f"Fórmula matricial ocupa o intervalo {ref or '(não informado)'} e não será scalarizada automaticamente."
                        ))

                upper = formula.upper()
                if formula and IFS_RE.search(formula):
                    summary.ifs_formulas += 1
                if formula and XLOOKUP_RE.search(formula):
                    summary.xlookup_formulas += 1
                if formula and GOOGLE_WRAPPER_RE.search(formula):
                    summary.google_wrapper_formulas += 1
                if formula and "#REF!" in upper:
                    summary.broken_reference_formulas += 1
                    summary.issues.append(FormulaIssue(
                        sheet_name, cell_ref, "BROKEN_REF", formula,
                        "A fórmula já contém #REF! no XLSX de origem."
                    ))

                cached = _cell_cached_text(cell).strip()
                ctype = cell.attrib.get("t", "")
                if cached and (ctype == "e" or cached.upper() in ERROR_TOKENS):
                    summary.cached_error_cells += 1
                    summary.issues.append(FormulaIssue(
                        sheet_name, cell_ref, "CACHED_ERROR", formula,
                        f"Resultado em cache do XLSX: {cached}"
                    ))
            result.sheets.append(summary)
    return result


def _find_matching_paren(text: str, open_idx: int) -> int | None:
    depth = 0
    in_string = False
    i = open_idx
    while i < len(text):
        ch = text[i]
        if ch == '"':
            if in_string and i + 1 < len(text) and text[i + 1] == '"':
                i += 2
                continue
            in_string = not in_string
        elif not in_string:
            if ch == "(":
                depth += 1
            elif ch == ")":
                depth -= 1
                if depth == 0:
                    return i
        i += 1
    return None


def _split_args(text: str) -> tuple[list[str], str] | tuple[None, None]:
    # OOXML normally stores commas even when Excel is localized. We still accept
    # semicolons for resilience with hand-modified workbooks.
    depth = 0
    in_string = False
    comma_positions: list[int] = []
    semicolon_positions: list[int] = []
    i = 0
    while i < len(text):
        ch = text[i]
        if ch == '"':
            if in_string and i + 1 < len(text) and text[i + 1] == '"':
                i += 2
                continue
            in_string = not in_string
        elif not in_string:
            if ch in "([{":
                depth += 1
            elif ch in ")]}" and depth > 0:
                depth -= 1
            elif depth == 0:
                if ch == ",":
                    comma_positions.append(i)
                elif ch == ";":
                    semicolon_positions.append(i)
        i += 1
    positions = comma_positions if comma_positions else semicolon_positions
    sep = "," if comma_positions else ";"
    if not positions:
        return [text.strip()], sep
    out: list[str] = []
    start = 0
    for pos in positions:
        out.append(text[start:pos].strip())
        start = pos + 1
    out.append(text[start:].strip())
    return out, sep


def _convert_one_ifs(function_text: str) -> str | None:
    # function_text includes IFS(...)
    m = IFS_RE.match(function_text)
    if not m:
        return None
    open_idx = function_text.find("(", m.start())
    close_idx = _find_matching_paren(function_text, open_idx)
    if close_idx is None or close_idx != len(function_text) - 1:
        return None
    args, sep = _split_args(function_text[open_idx + 1:close_idx])
    if not args or len(args) < 2 or len(args) % 2:
        return None
    # Preserve IFS semantics: if no condition is TRUE, IFS returns #N/A.
    fallback = "NA()"
    for i in range(len(args) - 2, -1, -2):
        test = args[i]
        value = args[i + 1]
        fallback = f"IF({test}{sep}{value}{sep}{fallback})"
    return fallback


def convert_ifs_to_nested_if(formula: str) -> tuple[str, int]:
    """Convert every parseable IFS() occurrence to nested IF(), innermost first."""
    text = formula
    converted = 0
    while True:
        matches = list(IFS_RE.finditer(text))
        if not matches:
            break
        changed = False
        for m in reversed(matches):
            open_idx = text.find("(", m.start())
            close_idx = _find_matching_paren(text, open_idx)
            if close_idx is None:
                continue
            candidate = text[m.start():close_idx + 1]
            replacement = _convert_one_ifs(candidate)
            if replacement is None:
                continue
            text = text[:m.start()] + replacement + text[close_idx + 1:]
            converted += 1
            changed = True
        if not changed:
            break
    return text, converted


def normalize_formula_xml(root: ET.Element) -> FormulaConversionStats:
    stats = FormulaConversionStats()
    for cell in root.iter(f"{M}c"):
        cell_ref = cell.attrib.get("r", "")
        f = cell.find(f"{M}f")
        if f is None or not f.text:
            continue
        original = f.text
        converted_formula, count = convert_ifs_to_nested_if(original)
        if count:
            f.text = converted_formula
            stats.ifs_converted += count

        is_single_array = f.attrib.get("t") == "array" and _is_single_ref(f.attrib.get("ref", ""), cell_ref)
        if is_single_array and (count or XLOOKUP_RE.search(original)):
            # Google frequently exports scalar ARRAYFORMULA/ARRAY_CONSTRAIN results
            # as a one-cell legacy array formula. IFS and scalar XLOOKUP do not need
            # that wrapper in Microsoft 365. Removing only the one-cell array marker
            # avoids changing genuine multi-cell array formulas.
            if XLOOKUP_RE.search(original):
                stats.xlookup_arrays_normalized += 1
            stats.single_cell_arrays_normalized += 1
            for attr in ("t", "ref", "aca", "ca"):
                f.attrib.pop(attr, None)
        elif f.attrib.get("t") == "array" and not is_single_array:
            stats.formulas_left_for_review += 1
    return stats


def enable_full_recalculation(workbook_xml: bytes) -> bytes:
    root = ET.fromstring(workbook_xml)
    calc = root.find(f"{M}calcPr")
    if calc is None:
        calc = ET.SubElement(root, f"{M}calcPr")
    calc.attrib.update({
        "calcMode": "auto",
        "fullCalcOnLoad": "1",
        "forceFullCalc": "1",
    })
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _normalized_value(value: str) -> str:
    s = (value or "").strip()
    if not s:
        return ""
    # Tolerate numeric formatting differences such as 1 vs 1.0.
    try:
        n = float(s.replace(",", "."))
        if math.isfinite(n):
            if n.is_integer():
                return str(int(n))
            return f"{n:.12g}"
    except ValueError:
        pass
    return " ".join(s.split()).casefold()


def attach_html_reference_data(
    analysis: FormulaCompatibilityAnalysis,
    html_sheets: Iterable,
    mismatch_limit: int = 100,
) -> None:
    by_name = {normalize_text(s.name): s for s in analysis.sheets}
    for hs in html_sheets:
        fs = by_name.get(normalize_text(hs.name))
        if fs is None:
            continue
        ref_values = getattr(hs, "reference_values", {}) or {}
        error_cells = getattr(hs, "error_cells", {}) or {}
        analysis.html_reference_values += len(ref_values)
        for cell, error in error_cells.items():
            analysis.html_error_cells.append(FormulaIssue(
                hs.name, cell, "HTML_ERROR", fs.formula_text_by_cell.get(cell, ""),
                f"O Google HTML exibe {error} nesta célula."
            ))
        for cell, html_value in ref_values.items():
            cached = fs.cached_value_by_cell.get(cell, "")
            if not cached or not html_value:
                continue
            if _normalized_value(cached) != _normalized_value(html_value):
                analysis.source_value_mismatch_count += 1
                if len(analysis.source_value_mismatches) < mismatch_limit:
                    analysis.source_value_mismatches.append(FormulaIssue(
                        hs.name, cell, "XLSX_HTML_VALUE", fs.formula_text_by_cell.get(cell, ""),
                        f"Cache XLSX='{cached}' | valor Google HTML='{html_value}'"
                    ))
