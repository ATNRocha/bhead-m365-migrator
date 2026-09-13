"""Branch coverage for HybridAnalysis.status (core/hybrid.py).

Each test isolates exactly one decision path. All collaborators that
status() doesn't need for that branch are held in a known-clean state
so only the intended condition can cause the result.
"""
from pathlib import Path
from types import SimpleNamespace

from core.hybrid import HybridAnalysis, HtmlChipSheet
from core.formula_compat import (
    FormulaCompatibilityAnalysis,
    FormulaSheetSummary,
    FormulaIssue,
)

DUMMY_XLSX_PATH = Path("dummy.xlsx")
DUMMY_HTML_PATH = Path("dummy.html")


def clean_xlsx_analysis():
    """Minimal stub: status() only ever reads .google_functions off xlsx_analysis."""
    return SimpleNamespace(google_functions=set())


def clean_html_sheets():
    return [HtmlChipSheet(name="Sheet1")]


def clean_formula_compat():
    """A real, fully-default FormulaCompatibilityAnalysis: has_source_errors,
    multi_cell_array_formulas and source_value_mismatch_count all evaluate falsy."""
    return FormulaCompatibilityAnalysis()


def make_analysis(
    *,
    xlsx_analysis=None,
    html_sheets=None,
    matched_sheets=1,
    formula_compat=None,
    missing_validation_columns=None,
):
    return HybridAnalysis(
        xlsx_path=DUMMY_XLSX_PATH,
        html_path=DUMMY_HTML_PATH,
        xlsx_analysis=xlsx_analysis if xlsx_analysis is not None else clean_xlsx_analysis(),
        html_sheets=html_sheets if html_sheets is not None else clean_html_sheets(),
        matched_sheets=matched_sheets,
        formula_compat=formula_compat,
        missing_validation_columns=missing_validation_columns or [],
    )


def test_status_revisar_when_google_functions_present():
    """B1: any Google-only function detected in the XLSX forces REVISAR,
    even though formula_compat and every other field are clean."""
    analysis = make_analysis(
        xlsx_analysis=SimpleNamespace(google_functions={"ARRAYFORMULA"}),
        formula_compat=clean_formula_compat(),
    )
    assert analysis.status == "REVISAR"


def test_status_revisar_when_formula_compat_has_source_errors():
    """B2a: has_source_errors becomes True via a real html_error_cells entry;
    no Google functions, no other trigger present."""
    formula_compat = FormulaCompatibilityAnalysis(
        html_error_cells=[FormulaIssue(sheet="Sheet1", cell="A1", kind="html_error")]
    )
    assert formula_compat.has_source_errors is True  # sanity check on the real object

    analysis = make_analysis(formula_compat=formula_compat)
    assert analysis.status == "REVISAR"


def test_status_revisar_when_multi_cell_array_formulas_without_source_errors():
    """B2b: has_source_errors is False, but multi_cell_array_formulas > 0
    still trips the second operand of the `or`."""
    formula_compat = FormulaCompatibilityAnalysis(
        sheets=[FormulaSheetSummary(name="Sheet1", multi_cell_array_formulas=1)]
    )
    assert formula_compat.has_source_errors is False  # isolates B2b from B2a
    assert formula_compat.multi_cell_array_formulas == 1

    analysis = make_analysis(formula_compat=formula_compat)
    assert analysis.status == "REVISAR"


def test_status_revisar_when_source_value_mismatch_count_positive():
    """B3: has_source_errors and multi_cell_array_formulas both falsy,
    but a cache/HTML mismatch count still trips REVISAR."""
    formula_compat = FormulaCompatibilityAnalysis(source_value_mismatch_count=1)
    assert formula_compat.has_source_errors is False
    assert formula_compat.multi_cell_array_formulas == 0

    analysis = make_analysis(formula_compat=formula_compat)
    assert analysis.status == "REVISAR"


def test_status_revisar_when_formula_compat_none_and_validation_columns_missing():
    """guard + B4: formula_compat=None must be skipped without raising,
    and missing_validation_columns alone must still force REVISAR."""
    analysis = make_analysis(
        formula_compat=None,
        missing_validation_columns=["Data Nascimento"],
    )
    assert analysis.status == "REVISAR"


def test_status_revisar_when_html_sheets_empty():
    """B5: no HTML sheets at all forces REVISAR, isolated from B6
    by keeping matched_sheets at a non-zero value."""
    analysis = make_analysis(
        html_sheets=[],
        matched_sheets=1,
        formula_compat=clean_formula_compat(),
    )
    assert analysis.status == "REVISAR"


def test_status_revisar_when_no_sheets_matched():
    """B6: html_sheets is non-empty but matched_sheets == 0,
    isolated from B5 by keeping html_sheets populated."""
    analysis = make_analysis(
        html_sheets=clean_html_sheets(),
        matched_sheets=0,
        formula_compat=clean_formula_compat(),
    )
    assert analysis.status == "REVISAR"


def test_status_compativel_when_all_clean():
    """B7: every condition is clean -> fall through to the approved status."""
    analysis = make_analysis(
        xlsx_analysis=clean_xlsx_analysis(),
        html_sheets=clean_html_sheets(),
        matched_sheets=1,
        formula_compat=clean_formula_compat(),
        missing_validation_columns=[],
    )
    assert analysis.status == "COMPATÍVEL PARA CONVERSÃO"
