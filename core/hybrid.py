from __future__ import annotations

import csv
import html as html_lib
import re
import time
import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from html.parser import HTMLParser
from pathlib import Path
from typing import Callable

from .profile import MigrationProfile, normalize_text
from .xlsx_ooxml import (
    M, _read_shared_strings, _workbook_sheet_map, _dimension, _headers,
    _google_functions, _read_dxf_style, _make_dxf, _existing_cf_signatures,
    _max_priority, _col_to_num, _num_to_col, _split_cell_ref, _detect_header_row, analyze_xlsx,
)
from .formula_compat import (
    FormulaCompatibilityAnalysis, analyze_formula_compatibility, attach_html_reference_data,
    normalize_formula_xml, enable_full_recalculation,
)


@dataclass
class HtmlChipSheet:
    name: str
    headers: dict[int, str] = field(default_factory=dict)
    header_row: int = 1
    chip_styles: dict[int, dict[str, dict[str, str]]] = field(default_factory=dict)
    chip_cells: int = 0
    reference_values: dict[str, str] = field(default_factory=dict)
    error_cells: dict[str, str] = field(default_factory=dict)

    @property
    def by_header(self) -> dict[str, dict[str, dict[str, str]]]:
        out: dict[str, dict[str, dict[str, str]]] = {}
        for col, values in self.chip_styles.items():
            h = normalize_text(self.headers.get(col, ""))
            if h:
                out[h] = values
        return out


@dataclass
class HybridAnalysis:
    xlsx_path: Path
    html_path: Path
    xlsx_analysis: object
    html_sheets: list[HtmlChipSheet]
    matched_sheets: int = 0
    html_chip_columns: int = 0
    formula_compat: FormulaCompatibilityAnalysis | None = None
    missing_validation_columns: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)

    @property
    def status(self) -> str:
        if self.xlsx_analysis.google_functions:
            return "REVISAR"
        if self.formula_compat is not None:
            if self.formula_compat.has_source_errors or self.formula_compat.multi_cell_array_formulas:
                return "REVISAR"
            if self.formula_compat.source_value_mismatch_count:
                return "REVISAR"
        if self.missing_validation_columns:
            return "REVISAR"
        if not self.html_sheets or self.matched_sheets == 0:
            return "REVISAR"
        return "COMPATÍVEL PARA CONVERSÃO"


@dataclass
class HybridConversionResult:
    output: Path
    blocks_added: int = 0
    rules_added: int = 0
    profile_rules_added: int = 0
    html_rules_added: int = 0
    dynamic_styles_added: int = 0
    date_validations_added: int = 0
    sheets_processed: int = 0
    verified_nonempty_sheets: int = 0
    verified_content_cells: int = 0
    ifs_formulas_converted: int = 0
    array_formulas_normalized: int = 0
    xlookup_arrays_normalized: int = 0
    formulas_left_for_review: int = 0
    workbook_recalc_enabled: bool = False
    warnings: list[str] = field(default_factory=list)


class _ChipParser(HTMLParser):
    """Extrai apenas cabeçalhos e estilos dos chips do HTML exportado pelo Google."""

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.in_tbody = False
        self.row = -1
        self.col = 0
        self._cell = None
        self._occupied: dict[int, set[int]] = {}
        self.skip_row = False
        self.headers: dict[int, str] = {}
        self.styles: dict[int, dict[str, dict[str, str]]] = {}
        self.chip_cells = 0

    @staticmethod
    def _attrs(attrs):
        return {k.lower(): (v or "") for k, v in attrs}

    @staticmethod
    def _props(style_text: str) -> dict[str, str]:
        out = {}
        for part in (style_text or "").split(";"):
            if ":" in part:
                k, v = part.split(":", 1)
                out[k.strip().lower()] = v.strip()
        return out

    @staticmethod
    def _color(value: str | None, fallback: str) -> str:
        s = (value or "").strip()
        if re.fullmatch(r"#[0-9a-fA-F]{6}", s):
            return s[1:].upper()
        m = re.fullmatch(r"rgb\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)", s, re.I)
        if m:
            return "".join(f"{max(0, min(255, int(x))):02X}" for x in m.groups())
        return fallback

    def handle_starttag(self, tag, attrs):
        tag = tag.lower()
        a = self._attrs(attrs)
        if tag == "tbody":
            self.in_tbody = True
        elif tag == "tr" and self.in_tbody:
            self.row += 1
            self.col = 0
            self.skip_row = False
        elif tag == "th" and self.in_tbody:
            if "freezebar-cell" in a.get("class", ""):
                self.skip_row = True
        elif tag == "td" and self.in_tbody and not self.skip_row:
            occupied = self._occupied.get(self.row, set())
            while self.col in occupied:
                self.col += 1
            colspan = int(a.get("colspan") or "1")
            rowspan = int(a.get("rowspan") or "1")
            self._cell = {
                "row": self.row,
                "col": self.col,
                "colspan": colspan,
                "rowspan": rowspan,
                "text": [],
                "chip": {},
            }
            if rowspan > 1:
                for rr in range(self.row + 1, self.row + rowspan):
                    self._occupied.setdefault(rr, set()).update(range(self.col, self.col + colspan))
        elif tag == "span" and self._cell is not None:
            p = self._props(a.get("style", ""))
            if "background-color" in p:
                self._cell["chip"].update(p)
        elif tag == "br" and self._cell is not None:
            self._cell["text"].append("\n")

    def handle_data(self, data):
        if self._cell is not None:
            self._cell["text"].append(data)

    def handle_endtag(self, tag):
        tag = tag.lower()
        if tag == "tbody":
            self.in_tbody = False
        elif tag == "tr" and self.in_tbody and self.skip_row:
            self.row -= 1
            self.skip_row = False
        elif tag == "td" and self._cell is not None:
            c = self._cell
            text = re.sub(r"[ \t\r\f\v]+", " ", "".join(c["text"])).strip()
            if c["row"] == 0:
                self.headers[c["col"]] = text
            chip = c["chip"]
            if chip.get("background-color") and text and c["row"] > 0:
                style = {
                    "fill": self._color(chip.get("background-color"), "FFFFFF"),
                    "font": self._color(chip.get("color"), "000000"),
                }
                self.styles.setdefault(c["col"], {})[text] = style
                self.chip_cells += 1
            self.col += c["colspan"]
            self._cell = None



# Fast path for Google Sheets HTML exports. The files can be tens of MB after
# decompression; HTMLParser is accurate but unnecessarily expensive when we only
# need the header row plus inline styles/values of dropdown chips.
_TR_RE = re.compile(r"<tr\b[^>]*>(.*?)</tr>", re.I | re.S)
_TD_RE = re.compile(r"<td\b([^>]*)>(.*?)</td>", re.I | re.S)
_SPAN_STYLE_RE = re.compile(r'<span\b[^>]*style="([^"]*background-color:[^"]*)"[^>]*>(.*?)</span>', re.I | re.S)
_TAG_RE = re.compile(r"<[^>]+>")
_COLSPAN_RE = re.compile(r"colspan=[\"']?(\d+)", re.I)
_BG_RE = re.compile(r"background-color:\s*(#[0-9a-fA-F]{6}|rgb\([^)]*\))", re.I)
_FG_RE = re.compile(r"(?:^|;)\s*color:\s*(#[0-9a-fA-F]{6}|rgb\([^)]*\))", re.I)
_ERROR_TOKEN_RE = re.compile(r"#(?:REF!|N/A|DIV/0!|VALUE!|NAME\?|NOME\?|NUM!|NULL!)", re.I)


def _clean_html_text(value: str) -> str:
    return re.sub(r"[ \t\r\f\v]+", " ", html_lib.unescape(_TAG_RE.sub("", value))).strip()


def _css_color(value: str | None, fallback: str) -> str:
    s = (value or "").strip()
    if re.fullmatch(r"#[0-9a-fA-F]{6}", s):
        return s[1:].upper()
    m = re.fullmatch(r"rgb\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)", s, re.I)
    if m:
        return "".join(f"{max(0, min(255, int(x))):02X}" for x in m.groups())
    return fallback


def _fast_chip_parse(raw: str, progress_cb: Callable[[float, str], None] | None = None,
                     progress_start: float = 0.0, progress_end: float = 100.0,
                     stage: str = "Lendo HTML", header_row: int = 1,
                     wanted_cells: set[str] | None = None) -> HtmlChipSheet:
    lo = raw.find("<tbody")
    if lo < 0:
        p = _ChipParser(); p.feed(raw)
        return HtmlChipSheet(name="", headers=p.headers, header_row=header_row, chip_styles=p.styles, chip_cells=p.chip_cells)
    lo = raw.find(">", lo)
    hi = raw.find("</tbody>", lo)
    if lo < 0 or hi < 0:
        p = _ChipParser(); p.feed(raw)
        return HtmlChipSheet(name="", headers=p.headers, header_row=header_row, chip_styles=p.styles, chip_cells=p.chip_cells)

    body = raw[lo + 1:hi]
    headers: dict[int, str] = {}
    styles: dict[int, dict[str, dict[str, str]]] = {}
    chip_cells = 0
    reference_values: dict[str, str] = {}
    error_cells: dict[str, str] = {}
    wanted_cells = wanted_cells or set()
    logical_row = -1
    span = max(0.0, progress_end - progress_start)
    body_len = max(1, len(body))

    for physical_row, tm in enumerate(_TR_RE.finditer(body)):
        row_html = tm.group(1)
        if "freezebar-cell" in row_html:
            continue
        logical_row += 1
        col = 0
        for cm in _TD_RE.finditer(row_html):
            attrs, inner = cm.groups()
            csm = _COLSPAN_RE.search(attrs)
            colspan = int(csm.group(1)) if csm else 1
            row_number = logical_row + 1
            cell_ref = f"{_num_to_col(col + 1)}{row_number}"
            if row_number == header_row:
                headers[col] = _clean_html_text(inner)
            else:
                needs_text = cell_ref in wanted_cells or bool(_ERROR_TOKEN_RE.search(inner))
                cell_text = _clean_html_text(inner) if needs_text else ""
                if cell_ref in wanted_cells and cell_text:
                    reference_values[cell_ref] = cell_text
                em = _ERROR_TOKEN_RE.search(cell_text) if cell_text else None
                if em:
                    error_cells[cell_ref] = em.group(0).upper()

            if row_number > header_row and "background-color:" in inner:
                sm = _SPAN_STYLE_RE.search(inner)
                if sm:
                    style_text, chip_inner = sm.groups()
                    value = _clean_html_text(chip_inner)
                    bg = _BG_RE.search(style_text)
                    fg = _FG_RE.search(style_text)
                    if bg and value:
                        style = {
                            "fill": _css_color(bg.group(1), "FFFFFF"),
                            "font": _css_color(fg.group(1) if fg else None, "000000"),
                        }
                        styles.setdefault(col, {})[value] = style
                        chip_cells += 1
            col += colspan

        if progress_cb and physical_row % 750 == 0:
            pct = progress_start + span * (tm.start() / body_len)
            progress_cb(min(progress_end, pct), stage)
            # Give Tk's main thread a scheduling opportunity during CPU-heavy parsing.
            time.sleep(0)

    if progress_cb:
        progress_cb(progress_end, stage)
    return HtmlChipSheet("", headers, header_row, styles, chip_cells, reference_values, error_cells)

def _html_root_files(zf: zipfile.ZipFile) -> list[str]:
    return [n for n in zf.namelist() if n.lower().endswith((".html", ".htm")) and "/" not in n.rstrip("/")]


def extract_html_chip_styles(zip_path: str | Path,
                             progress_cb: Callable[[float, str], None] | None = None,
                             progress_start: float = 0.0,
                             progress_end: float = 100.0,
                             header_rows_by_sheet: dict[str, int] | None = None,
                             wanted_cells_by_sheet: dict[str, set[str]] | None = None) -> list[HtmlChipSheet]:
    path = Path(zip_path)
    if not path.exists():
        raise FileNotFoundError("Pacote HTML/ZIP não encontrado.")
    sheets: list[HtmlChipSheet] = []
    with zipfile.ZipFile(path) as zf:
        names = _html_root_files(zf)
        sizes = [max(1, zf.getinfo(n).file_size) for n in names]
        total = max(1, sum(sizes))
        done = 0
        overall_span = max(0.0, progress_end - progress_start)
        for name, size in zip(names, sizes):
            file_start = progress_start + overall_span * (done / total)
            file_end = progress_start + overall_span * ((done + size) / total)
            if progress_cb:
                progress_cb(file_start, f"Lendo HTML: {Path(name).stem}")
            raw = zf.read(name).decode("utf-8", "ignore")
            stem = Path(name).stem
            norm = normalize_text(stem)
            header_row = (header_rows_by_sheet or {}).get(norm, 1)
            wanted = (wanted_cells_by_sheet or {}).get(norm, set())
            parsed = _fast_chip_parse(
                raw, progress_cb, file_start, file_end, f"Lendo HTML: {stem}",
                header_row=header_row, wanted_cells=wanted,
            )
            parsed.name = stem
            sheets.append(parsed)
            done += size
    return sheets

def _list_validation_columns(sheet_root: ET.Element) -> set[str]:
    cols: set[str] = set()
    dvs = sheet_root.find(f"{M}dataValidations")
    if dvs is None:
        return cols
    for dv in dvs.findall(f"{M}dataValidation"):
        if dv.attrib.get("type") != "list":
            continue
        for ref in (dv.attrib.get("sqref") or "").split():
            # Pode ser C2:C13020 ou C2.
            start = ref.split(":", 1)[0]
            col, _ = _split_cell_ref(start.replace("$", ""))
            cols.add(col)
    return cols


def analyze_hybrid(xlsx_path: str | Path, html_path: str | Path, profile: MigrationProfile,
                   progress_cb: Callable[[float, str], None] | None = None) -> HybridAnalysis:
    if progress_cb:
        progress_cb(2, "Analisando estrutura XLSX")
    xa = analyze_xlsx(xlsx_path, profile, html_path)
    formula_compat = analyze_formula_compatibility(xlsx_path)
    header_rows = {normalize_text(s.name): s.header_row for s in xa.sheets}
    if progress_cb:
        progress_cb(18, "Estrutura e fórmulas XLSX analisadas")
    hs = extract_html_chip_styles(
        html_path, progress_cb, 18, 78,
        header_rows_by_sheet=header_rows,
        wanted_cells_by_sheet=formula_compat.wanted_cells_by_sheet,
    )
    attach_html_reference_data(formula_compat, hs)
    html_by_name = {normalize_text(s.name): s for s in hs}
    matched = 0
    chip_cols = 0
    missing: list[str] = []
    warnings: list[str] = []

    with zipfile.ZipFile(Path(xlsx_path)) as z:
        shared = _read_shared_strings(z)  # read once, not once per worksheet
        sheet_map = _workbook_sheet_map(z)
        total_sheets = max(1, len(sheet_map))
        for idx, (sheet_name, xml_path) in enumerate(sheet_map):
            h = html_by_name.get(normalize_text(sheet_name))
            if h:
                matched += 1
                chip_cols += len(h.chip_styles)
                root = ET.fromstring(z.read(xml_path))
                _header_row, xheaders = _detect_header_row(root, shared, profile)
                by_norm = {normalize_text(v): col for col, v in xheaders.items()}
                dv_cols = _list_validation_columns(root)
                for html_header_norm, values in h.by_header.items():
                    col = by_norm.get(html_header_norm)
                    if not col or not values:
                        continue
                    if col not in dv_cols:
                        missing.append(f"{sheet_name} — {xheaders.get(col, html_header_norm)}")
            if progress_cb:
                progress_cb(78 + 20 * ((idx + 1) / total_sheets), f"Conferindo planilhas: {idx + 1}/{total_sheets}")

    if matched < len(hs):
        warnings.append(f"Somente {matched} de {len(hs)} planilhas HTML foram associadas por nome ao XLSX.")
    if missing:
        warnings.append(
            "Há colunas com chips no HTML sem lista de validação correspondente no XLSX. "
            "As cores podem ser reconstruídas, mas as opções completas do dropdown não podem ser garantidas."
        )
    if progress_cb:
        progress_cb(100, "Análise concluída")
    if formula_compat.broken_reference_formulas:
        warnings.append(
            f"Foram encontradas {formula_compat.broken_reference_formulas} fórmulas com #REF! já no XLSX de origem; "
            "essas referências não serão inventadas automaticamente."
        )
    if formula_compat.html_error_cells:
        warnings.append(
            f"O HTML do Google contém {len(formula_compat.html_error_cells)} células com erro visível na origem."
        )
    if formula_compat.source_value_mismatch_count:
        warnings.append(
            f"Foram observadas {formula_compat.source_value_mismatch_count} divergências entre o cache do XLSX e "
            "o valor exibido no HTML; revise o relatório técnico."
        )
    return HybridAnalysis(
        Path(xlsx_path), Path(html_path), xa, hs, matched, chip_cols, formula_compat, missing, warnings
    )

def _style_key(style: dict[str, str]) -> tuple[str, str]:
    return style.get("fill", "FFFFFF").upper().replace("#", ""), style.get("font", "000000").upper().replace("#", "")


def _ensure_all_styles(styles_xml: bytes, profile: MigrationProfile, html_sheets: list[HtmlChipSheet]):
    root = ET.fromstring(styles_xml)
    dxfs = root.find(f"{M}dxfs")
    if dxfs is None:
        dxfs = ET.Element(f"{M}dxfs", {"count": "0"})
        table_styles = root.find(f"{M}tableStyles")
        if table_styles is not None:
            root.insert(list(root).index(table_styles), dxfs)
        else:
            root.append(dxfs)

    existing: dict[tuple[str, str], int] = {}
    for idx, dxf in enumerate(dxfs.findall(f"{M}dxf")):
        key = _read_dxf_style(dxf)
        if key:
            existing[key] = idx

    profile_ids: dict[str, int] = {}
    dynamic_ids: dict[tuple[str, str], int] = {}
    dynamic_added = 0

    def ensure(style: dict[str, str]) -> int:
        nonlocal dynamic_added
        key = _style_key(style)
        if key in existing:
            return existing[key]
        idx = len(dxfs.findall(f"{M}dxf"))
        dxfs.append(_make_dxf(style))
        existing[key] = idx
        dynamic_added += 1
        return idx

    for name, style in profile.data.get("styles", {}).items():
        profile_ids[name] = ensure(style)

    for sh in html_sheets:
        for values in sh.chip_styles.values():
            for st in values.values():
                dynamic_ids[_style_key(st)] = ensure(st)

    dxfs.attrib["count"] = str(len(dxfs.findall(f"{M}dxf")))
    return ET.tostring(root, encoding="utf-8", xml_declaration=True), profile_ids, dynamic_ids, dynamic_added


def _cf_insert_index(root: ET.Element) -> int:
    insert_before_tags = {
        f"{M}dataValidations", f"{M}hyperlinks", f"{M}printOptions", f"{M}pageMargins",
        f"{M}pageSetup", f"{M}headerFooter", f"{M}rowBreaks", f"{M}colBreaks",
        f"{M}customProperties", f"{M}cellWatches", f"{M}ignoredErrors", f"{M}smartTags",
        f"{M}drawing", f"{M}legacyDrawing", f"{M}legacyDrawingHF", f"{M}picture",
        f"{M}oleObjects", f"{M}controls", f"{M}webPublishItems", f"{M}tableParts", f"{M}extLst"
    }
    for i, child in enumerate(list(root)):
        if child.tag in insert_before_tags:
            return i
    return len(list(root))


def _excel_literal(value: str) -> str:
    return value.replace('"', '""')


def _date_rule_for_header(profile: MigrationProfile, header: str):
    n = normalize_text(header)
    for rule in profile.data.get("date_columns", []):
        if n in {normalize_text(h) for h in rule.get("headers", [])}:
            return rule
    return None


def _dv_insert_index(root: ET.Element) -> int:
    """Posição OOXML válida para dataValidations quando a planilha ainda não possui o bloco."""
    insert_before_tags = {
        f"{M}hyperlinks", f"{M}printOptions", f"{M}pageMargins", f"{M}pageSetup",
        f"{M}headerFooter", f"{M}rowBreaks", f"{M}colBreaks", f"{M}customProperties",
        f"{M}cellWatches", f"{M}ignoredErrors", f"{M}smartTags", f"{M}drawing",
        f"{M}legacyDrawing", f"{M}legacyDrawingHF", f"{M}picture", f"{M}oleObjects",
        f"{M}controls", f"{M}webPublishItems", f"{M}tableParts", f"{M}extLst",
    }
    for i, child in enumerate(list(root)):
        if child.tag in insert_before_tags:
            return i
    return len(list(root))


def _insert_date_guidance(root: ET.Element, headers: dict[str, str], max_row: int, profile: MigrationProfile, header_row: int = 1) -> int:
    """Adiciona validação de data + mensagem de entrada nas colunas de data homologadas.

    No modo híbrido o XLSX continua sendo a fonte estrutural. Essa rotina não altera
    valores nem fórmulas; apenas acrescenta uma proteção nativa do Excel para novas
    edições e a orientação sobre anos anteriores a 2000 no seletor do Excel Online.
    """
    date_cols = []
    for col, header in sorted(headers.items(), key=lambda kv: _col_to_num(kv[0])):
        rule = _date_rule_for_header(profile, header)
        if rule:
            date_cols.append((col, rule))
    if not date_cols:
        return 0

    dvs = root.find(f"{M}dataValidations")
    if dvs is None:
        dvs = ET.Element(f"{M}dataValidations", {"count": "0"})
        root.insert(_dv_insert_index(root), dvs)

    # Evita duplicar a mesma validação caso o arquivo seja processado novamente.
    existing = {
        (dv.attrib.get("type", ""), dv.attrib.get("sqref", ""))
        for dv in dvs.findall(f"{M}dataValidation")
    }

    added = 0
    for col, rule in date_cols:
        start_row = max(int(rule.get("start_row", 2)), header_row + 1)
        end_row = max(max_row, int(rule.get("minimum_end_row", max_row or start_row)))
        sqref = f"{col}{start_row}:{col}{end_row}"
        key = ("date", sqref)
        prompt = rule.get(
            "input_message",
            "Digite a data no formato DD/MM/AAAA.\nPara anos anteriores a 2000, prefira digitar manualmente.",
        )
        if key in existing:
            # Atualiza mensagem de uma validação de data já existente.
            for dv in dvs.findall(f"{M}dataValidation"):
                if dv.attrib.get("type") == "date" and dv.attrib.get("sqref") == sqref:
                    dv.attrib.update({
                        "allowBlank": "1",
                        "showInputMessage": "1",
                        "showErrorMessage": "1",
                        "promptTitle": rule.get("input_title", "Data"),
                        "prompt": prompt,
                        "errorTitle": rule.get("error_title", "Data inválida"),
                        "error": rule.get("error_message", "Informe uma data válida no formato DD/MM/AAAA."),
                    })
            continue

        dv = ET.SubElement(dvs, f"{M}dataValidation", {
            "type": "date",
            "operator": "between",
            "allowBlank": "1",
            "showInputMessage": "1",
            "showErrorMessage": "1",
            "promptTitle": rule.get("input_title", "Data"),
            "prompt": prompt,
            "errorTitle": rule.get("error_title", "Data inválida"),
            "error": rule.get("error_message", "Informe uma data válida no formato DD/MM/AAAA."),
            "sqref": sqref,
        })
        f1 = ET.SubElement(dv, f"{M}formula1")
        f1.text = "DATE(1900,1,1)"
        f2 = ET.SubElement(dv, f"{M}formula2")
        f2.text = "DATE(2100,12,31)"
        existing.add(key)
        added += 1

    dvs.attrib["count"] = str(len(dvs.findall(f"{M}dataValidation")))
    return added


def _insert_hybrid_cf(
    root: ET.Element,
    headers: dict[str, str],
    max_row: int,
    profile: MigrationProfile,
    profile_ids: dict[str, int],
    dynamic_ids: dict[tuple[str, str], int],
    html_sheet: HtmlChipSheet | None,
    header_row: int = 1,
):
    existing = _existing_cf_signatures(root)
    priority = _max_priority(root)
    insertion_index = _cf_insert_index(root)
    offset = 0
    blocks = rules = profile_rules = html_rules = 0
    warnings: list[str] = []

    html_by_header = html_sheet.by_header if html_sheet else {}

    def add_rule(bucket, sqref, formula, dxf_id):
        nonlocal priority
        if (sqref, formula) in existing:
            return False
        priority += 1
        r = ET.Element(f"{M}cfRule", {
            "type": "expression", "dxfId": str(dxf_id), "priority": str(priority), "stopIfTrue": "1"
        })
        f = ET.SubElement(r, f"{M}formula")
        f.text = formula
        bucket.append(r)
        existing.add((sqref, formula))
        return True

    for col, header in sorted(headers.items(), key=lambda kv: _col_to_num(kv[0])):
        pr = profile.find_header_rule(header)
        start_row = max(int(pr.get("start_row", 2)), header_row + 1) if pr else header_row + 1
        end_row = max(max_row, int(pr.get("minimum_end_row", max_row))) if pr else max(max_row, 2)
        sqref = f"{col}{start_row}:{col}{end_row}"
        anchor = f"{col}{start_row}"
        new_rules: list[ET.Element] = []

        # Cabeçalho já homologado: prioriza a regra semântica do perfil.
        # Isso evita milhares de regras em dropdowns multi-seleção, onde o HTML
        # apresenta cada combinação de itens como um valor distinto.
        if pr:
            for rd in pr.get("rules", []):
                formula = rd["formula"].format(cell=anchor, col=col, row=start_row)
                if add_rule(new_rules, sqref, formula, profile_ids[rd["style"]]):
                    profile_rules += 1
        else:
            observed = {
                value: style for value, style in html_by_header.get(normalize_text(header), {}).items()
                if value.strip().replace("\u200b", "")
            }
            if observed:
                # Para coluna nova, inferimos a menor quantidade de regras possível.
                # 1 estilo para todos os valores => regra genérica não-vazio.
                groups: dict[tuple[str, str], list[str]] = {}
                style_by_key: dict[tuple[str, str], dict[str, str]] = {}
                for value, style in observed.items():
                    key = _style_key(style)
                    groups.setdefault(key, []).append(value)
                    style_by_key[key] = style

                if len(groups) == 1:
                    key = next(iter(groups))
                    if add_rule(new_rules, sqref, f'{anchor}<>""', dynamic_ids[key]):
                        html_rules += 1
                elif len(observed) <= 20:
                    for value, style in observed.items():
                        formula = f'{anchor}="{_excel_literal(value)}"'
                        if add_rule(new_rules, sqref, formula, dynamic_ids[_style_key(style)]):
                            html_rules += 1
                else:
                    # Usa o estilo dominante como fallback não-vazio e cria regras
                    # específicas apenas para os valores minoritários, desde que o
                    # total permaneça pequeno. Caso contrário sinaliza revisão.
                    ordered = sorted(groups.items(), key=lambda kv: len(kv[1]), reverse=True)
                    default_key, default_values = ordered[0]
                    minorities = [(k, v) for k, v in ordered[1:]]
                    minority_count = sum(len(v) for _, v in minorities)
                    if minority_count <= 15:
                        for key, values in minorities:
                            for value in values:
                                formula = f'{anchor}="{_excel_literal(value)}"'
                                if add_rule(new_rules, sqref, formula, dynamic_ids[key]):
                                    html_rules += 1
                        if add_rule(new_rules, sqref, f'{anchor}<>""', dynamic_ids[default_key]):
                            html_rules += 1
                    else:
                        warnings.append(
                            f'{header}: {len(observed)} valores coloridos / {len(groups)} estilos; '
                            'regras visuais não foram inferidas automaticamente para evitar excesso de formatação condicional.'
                        )

        if new_rules:
            cf = ET.Element(f"{M}conditionalFormatting", {"sqref": sqref})
            for r in new_rules:
                cf.append(r)
            root.insert(insertion_index + offset, cf)
            offset += 1
            blocks += 1
            rules += len(new_rules)

    return blocks, rules, profile_rules, html_rules, warnings

def convert_hybrid(xlsx_path: str | Path, html_path: str | Path, output_path: str | Path,
                   profile: MigrationProfile,
                   html_sheets: list[HtmlChipSheet] | None = None,
                   progress_cb: Callable[[float, str], None] | None = None) -> HybridConversionResult:
    src = Path(xlsx_path)
    html = Path(html_path)
    out = Path(output_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    if src.resolve() == out.resolve():
        raise RuntimeError("O arquivo de saída deve ser diferente do XLSX original.")

    if html_sheets is None:
        html_sheets = extract_html_chip_styles(html, progress_cb, 0, 38)
    elif progress_cb:
        progress_cb(5, "Reutilizando análise HTML já realizada")
    html_by_name = {normalize_text(s.name): s for s in html_sheets}
    result = HybridConversionResult(out)

    with zipfile.ZipFile(src, "r") as zin:
        if progress_cb:
            progress_cb(10, "Preparando estilos do Excel")
        shared = _read_shared_strings(zin)
        styles_bytes, profile_ids, dynamic_ids, dynamic_added = _ensure_all_styles(
            zin.read("xl/styles.xml"), profile, html_sheets
        )
        result.dynamic_styles_added = dynamic_added
        replacements: dict[str, bytes] = {
            "xl/styles.xml": styles_bytes,
            "xl/workbook.xml": enable_full_recalculation(zin.read("xl/workbook.xml")),
        }
        result.workbook_recalc_enabled = True
        sheet_map = _workbook_sheet_map(zin)
        total_sheets = max(1, len(sheet_map))

        for idx, (sheet_name, xml_path) in enumerate(sheet_map):
            if progress_cb:
                progress_cb(12 + 48 * (idx / total_sheets), f"Convertendo planilha: {sheet_name}")
            root = ET.fromstring(zin.read(xml_path))
            max_row, _ = _dimension(root)
            header_row, headers = _detect_header_row(root, shared, profile)
            hs = html_by_name.get(normalize_text(sheet_name))
            b, r, pr, hr, warns = _insert_hybrid_cf(
                root, headers, max_row, profile, profile_ids, dynamic_ids, hs, header_row=header_row
            )
            date_added = _insert_date_guidance(root, headers, max_row, profile, header_row=header_row)
            formula_stats = normalize_formula_xml(root)
            result.ifs_formulas_converted += formula_stats.ifs_converted
            result.array_formulas_normalized += formula_stats.single_cell_arrays_normalized
            result.xlookup_arrays_normalized += formula_stats.xlookup_arrays_normalized
            result.formulas_left_for_review += formula_stats.formulas_left_for_review
            formula_changed = bool(
                formula_stats.ifs_converted or formula_stats.single_cell_arrays_normalized
            )
            if b or r or date_added or formula_changed:
                replacements[xml_path] = ET.tostring(root, encoding="utf-8", xml_declaration=True)
            result.blocks_added += b
            result.rules_added += r
            result.profile_rules_added += pr
            result.html_rules_added += hr
            result.date_validations_added += date_added
            result.sheets_processed += 1
            result.warnings.extend(f"{sheet_name} — {w}" for w in warns)

        if progress_cb:
            progress_cb(62, "Gravando arquivo XLSX")
        infos = zin.infolist()
        with zipfile.ZipFile(out, "w") as zout:
            total_items = max(1, len(infos))
            for idx, item in enumerate(infos):
                if item.filename in replacements:
                    data = replacements[item.filename]
                else:
                    data = zin.read(item.filename)
                zout.writestr(item, data)
                if progress_cb and idx % 5 == 0:
                    progress_cb(62 + 20 * ((idx + 1) / total_items), "Gravando pacote XLSX")

    # Verificação pós-conversão: mantém a checagem forte, mas mostra progresso
    # por planilha para a interface continuar responsiva.
    if progress_cb:
        progress_cb(84, "Validando arquivo gerado")
    with zipfile.ZipFile(out, "r") as ztest:
        bad = ztest.testzip()
        if bad:
            raise RuntimeError(f"Falha de integridade no XLSX híbrido: {bad}")
        if "xl/workbook.xml" not in ztest.namelist():
            raise RuntimeError("Arquivo final não contém xl/workbook.xml.")

        content_cells = 0
        nonempty_sheets = 0
        check_map = _workbook_sheet_map(ztest)
        total_checks = max(1, len(check_map))
        for idx, (_sheet_name, xml_path) in enumerate(check_map):
            root_check = ET.fromstring(ztest.read(xml_path))
            sheet_data = root_check.find(f"{M}sheetData")
            sheet_cells = 0
            if sheet_data is not None:
                for cell in sheet_data.iter(f"{M}c"):
                    if (cell.find(f"{M}v") is not None or
                        cell.find(f"{M}f") is not None or
                        cell.find(f"{M}is") is not None):
                        sheet_cells += 1
            if sheet_cells:
                nonempty_sheets += 1
                content_cells += sheet_cells
            if progress_cb:
                progress_cb(85 + 14 * ((idx + 1) / total_checks), f"Validando planilha: {_sheet_name}")

        result.verified_nonempty_sheets = nonempty_sheets
        result.verified_content_cells = content_cells
        if content_cells == 0:
            try:
                out.unlink()
            except OSError:
                pass
            raise RuntimeError(
                "Falha de segurança: o XLSX gerado passou no teste ZIP, mas não contém células com conteúdo. "
                "A saída foi descartada para evitar uma migração vazia."
            )
    if progress_cb:
        progress_cb(100, "Conversão concluída")
    return result

def write_hybrid_report_csv(path: str | Path, analysis: HybridAnalysis, result: HybridConversionResult | None = None):
    p = Path(path)
    p.parent.mkdir(parents=True, exist_ok=True)
    with p.open("w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["BHead M365 Migrator", "Relatório híbrido V0.4.5 — Formula Compatibility Engine"])
        w.writerow(["XLSX estrutural", str(analysis.xlsx_path)])
        w.writerow(["HTML visual", str(analysis.html_path)])
        w.writerow(["Status análise", analysis.status])
        w.writerow(["Planilhas XLSX", len(analysis.xlsx_analysis.sheets)])
        w.writerow(["Planilhas HTML", len(analysis.html_sheets)])
        w.writerow(["Planilhas associadas", analysis.matched_sheets])
        w.writerow(["Colunas com chips no HTML", analysis.html_chip_columns])
        w.writerow(["Funções Google detectadas", ", ".join(sorted(analysis.xlsx_analysis.google_functions)) or "-"])
        if analysis.formula_compat:
            fc = analysis.formula_compat
            w.writerow(["Células com fórmula", fc.formula_cells])
            w.writerow(["Fórmulas explícitas", fc.explicit_formulas])
            w.writerow(["Fórmulas IFS detectadas", fc.ifs_formulas])
            w.writerow(["Fórmulas XLOOKUP detectadas", fc.xlookup_formulas])
            w.writerow(["Fórmulas matriciais", fc.array_formulas])
            w.writerow(["Matrizes de uma célula", fc.single_cell_array_formulas])
            w.writerow(["Matrizes multi-célula", fc.multi_cell_array_formulas])
            w.writerow(["#REF! em fórmulas de origem", fc.broken_reference_formulas])
            w.writerow(["Erros visíveis no HTML Google", len(fc.html_error_cells)])
            w.writerow(["Valores HTML disponíveis para referência", fc.html_reference_values])
            w.writerow(["Divergências cache XLSX x HTML", fc.source_value_mismatch_count])
        w.writerow(["Colunas chip sem validação XLSX", len(analysis.missing_validation_columns)])
        if result:
            w.writerow(["Arquivo convertido", str(result.output)])
            w.writerow(["Blocos CF adicionados", result.blocks_added])
            w.writerow(["Regras CF adicionadas", result.rules_added])
            w.writerow(["Regras vindas do HTML", result.html_rules_added])
            w.writerow(["Regras fallback do perfil", result.profile_rules_added])
            w.writerow(["Estilos diferenciais adicionados", result.dynamic_styles_added])
            w.writerow(["Validações/orientações de data adicionadas", result.date_validations_added])
            w.writerow(["Planilhas com conteúdo verificadas", result.verified_nonempty_sheets])
            w.writerow(["Células com conteúdo verificadas", result.verified_content_cells])
            w.writerow(["IFS convertidos para IF aninhado", result.ifs_formulas_converted])
            w.writerow(["Matrizes de uma célula normalizadas", result.array_formulas_normalized])
            w.writerow(["XLOOKUPs scalarizados", result.xlookup_arrays_normalized])
            w.writerow(["Fórmulas matriciais deixadas para revisão", result.formulas_left_for_review])
            w.writerow(["Recálculo completo no Excel habilitado", "SIM" if result.workbook_recalc_enabled else "NÃO"])
        w.writerow([])
        w.writerow(["Diagnóstico de fórmulas"])
        if analysis.formula_compat:
            for fs in analysis.formula_compat.sheets:
                for issue in fs.issues[:200]:
                    w.writerow([issue.kind, issue.sheet, issue.cell, issue.detail, issue.formula])
            for issue in analysis.formula_compat.html_error_cells[:200]:
                w.writerow([issue.kind, issue.sheet, issue.cell, issue.detail, issue.formula])
            for issue in analysis.formula_compat.source_value_mismatches[:200]:
                w.writerow([issue.kind, issue.sheet, issue.cell, issue.detail, issue.formula])
        w.writerow([])
        w.writerow(["Ressalvas"])
        for s in analysis.missing_validation_columns:
            w.writerow(["Dropdown não garantido", s])
        for s in analysis.warnings:
            w.writerow([s])
