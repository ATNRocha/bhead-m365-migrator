from __future__ import annotations

import csv
import html as html_lib
import calendar
import re
import zipfile
from dataclasses import dataclass, field
from datetime import datetime
from html.parser import HTMLParser
from pathlib import Path
from typing import Callable

try:
    import xlsxwriter
except ImportError as exc:  # pragma: no cover - handled by UI
    xlsxwriter = None
    _XLSXWRITER_IMPORT_ERROR = exc
else:
    _XLSXWRITER_IMPORT_ERROR = None

from .profile import MigrationProfile, normalize_text


STYLE_RE = re.compile(r"\.ritz\s+\.waffle\s+\.(s\d+)\{([^}]*)\}", re.I)
PROP_RE = re.compile(r"([\w-]+)\s*:\s*([^;]+)")
PX_RE = re.compile(r"([\d.]+)px", re.I)
PT_RE = re.compile(r"([\d.]+)pt", re.I)
DATE_RE = re.compile(r"^(\d{1,2})/(\d{1,2})/(\d{4})$")
INT_RE = re.compile(r"^-?(?:0|[1-9]\d*)$")


@dataclass
class HtmlSheetAnalysis:
    name: str
    rows: int = 0
    cols: int = 0
    chip_cells: int = 0
    chip_columns: dict[int, set[str]] = field(default_factory=dict)
    headers: dict[int, str] = field(default_factory=dict)
    profiled_headers: list[str] = field(default_factory=list)
    merged_cells: int = 0
    hyperlinks: int = 0


@dataclass
class HtmlWorkbookAnalysis:
    source: Path
    sheets: list[HtmlSheetAnalysis]
    html_files: int
    css_files: int

    @property
    def status(self) -> str:
        if not self.sheets:
            return "INCOMPATÍVEL"
        if any(s.profiled_headers for s in self.sheets):
            return "COMPATÍVEL PARA CONVERSÃO"
        return "REVISAR"


@dataclass
class HtmlConversionResult:
    output: Path
    sheets_created: int = 0
    cells_written: int = 0
    chip_cells: int = 0
    validations_added: int = 0
    conditional_rules_added: int = 0
    formulas_reconstructed: int = 0
    inferred_validations: int = 0
    date_columns_formatted: int = 0
    date_validations_added: int = 0
    warnings: list[str] = field(default_factory=list)


def ensure_xlsxwriter():
    if xlsxwriter is None:
        raise RuntimeError(
            "A conversão HTML direta precisa da biblioteca XlsxWriter. "
            "Execute setup_dev.bat uma vez e abra o aplicativo novamente."
        ) from _XLSXWRITER_IMPORT_ERROR


def _parse_style_props(style_text: str) -> dict[str, str]:
    return {m.group(1).lower(): m.group(2).strip() for m in PROP_RE.finditer(style_text or "")}


def _extract_class_styles(html_text: str) -> dict[str, dict[str, str]]:
    # O bloco de estilos das planilhas Google fica no início de cada HTML.
    head = html_text.split("<div class=\"ritz", 1)[0]
    out: dict[str, dict[str, str]] = {}
    for m in STYLE_RE.finditer(head):
        out[m.group(1)] = _parse_style_props(m.group(2))
    return out


def _px(style_text: str, prop: str) -> float | None:
    props = _parse_style_props(style_text)
    m = PX_RE.search(props.get(prop, ""))
    return float(m.group(1)) if m else None


def _html_sheet_names(zf: zipfile.ZipFile) -> list[str]:
    return [n for n in zf.namelist() if n.lower().endswith((".html", ".htm")) and "/" not in n.rstrip("/")]


def _safe_sheet_name(name: str, used: set[str]) -> str:
    base = re.sub(r"[\[\]:*?/\\]", "_", name).strip() or "Planilha"
    base = base[:31]
    candidate = base
    i = 2
    while candidate in used:
        suffix = f"_{i}"
        candidate = base[: 31 - len(suffix)] + suffix
        i += 1
    used.add(candidate)
    return candidate


def _convert_value(text: str):
    value = text.replace("\xa0", " ").strip()
    if not value:
        return "", "text"
    m = DATE_RE.match(value)
    if m:
        try:
            return datetime(int(m.group(3)), int(m.group(2)), int(m.group(1))), "date"
        except ValueError:
            pass
    if INT_RE.match(value) and (value == "0" or not (value.startswith("0") and len(value) > 1)):
        try:
            return int(value), "number"
        except ValueError:
            pass
    return value, "text"


class GoogleTableParser(HTMLParser):
    """Streaming parser para a tabela .waffle exportada pelo Google Sheets."""

    def __init__(self, on_cell: Callable, on_col_width: Callable, on_row_height: Callable):
        super().__init__(convert_charrefs=True)
        self.on_cell = on_cell
        self.on_col_width = on_col_width
        self.on_row_height = on_row_height
        self.in_tbody = False
        self.in_thead = False
        self.row = -1
        self.col = 0
        self.col_header_index = -1
        self._cell = None
        self._in_row_th = False
        self._occupied: dict[int, set[int]] = {}
        self.max_col = 0
        self.max_row = -1
        self.skip_row = False

    @staticmethod
    def _attrs(attrs):
        return {k.lower(): (v or "") for k, v in attrs}

    def handle_starttag(self, tag, attrs):
        a = self._attrs(attrs)
        tag = tag.lower()
        if tag == "thead":
            self.in_thead = True
        elif tag == "tbody":
            self.in_tbody = True
        elif tag == "tr" and self.in_tbody:
            self.row += 1
            self.col = 0
            self.skip_row = False
        elif tag == "th":
            classes = a.get("class", "")
            if self.in_tbody and "freezebar-cell" in classes:
                self.skip_row = True
            if self.in_thead and "column-headers-background" in classes:
                # O primeiro TH é o canto da linha e não tem id Cn.
                if re.search(r"C\d+$", a.get("id", "")):
                    self.col_header_index += 1
                    width = _px(a.get("style", ""), "width")
                    if width is not None:
                        self.on_col_width(self.col_header_index, width)
            elif self.in_tbody and "row-headers-background" in classes:
                self._in_row_th = True
                height = _px(a.get("style", ""), "height")
                if height is not None:
                    self.on_row_height(self.row, height)
        elif tag == "td" and self.in_tbody:
            if self.skip_row:
                return
            occupied = self._occupied.get(self.row, set())
            while self.col in occupied:
                self.col += 1
            colspan = int(a.get("colspan") or "1")
            rowspan = int(a.get("rowspan") or "1")
            self._cell = {
                "row": self.row,
                "col": self.col,
                "class": (a.get("class", "").split() or [""])[0],
                "style": a.get("style", ""),
                "colspan": colspan,
                "rowspan": rowspan,
                "text": [],
                "chip_style": {},
                "href": None,
            }
            if rowspan > 1:
                for rr in range(self.row + 1, self.row + rowspan):
                    self._occupied.setdefault(rr, set()).update(range(self.col, self.col + colspan))
        elif tag == "span" and self._cell is not None:
            props = _parse_style_props(a.get("style", ""))
            if "background-color" in props or "color" in props:
                self._cell["chip_style"].update(props)
        elif tag == "a" and self._cell is not None:
            self._cell["href"] = a.get("href") or None
        elif tag == "br" and self._cell is not None:
            self._cell["text"].append("\n")

    def handle_endtag(self, tag):
        tag = tag.lower()
        if tag == "thead":
            self.in_thead = False
        elif tag == "tbody":
            self.in_tbody = False
        elif tag == "tr" and self.in_tbody:
            if self.skip_row:
                self.row -= 1
                self.skip_row = False
        elif tag == "th":
            self._in_row_th = False
        elif tag == "td" and self._cell is not None:
            cell = self._cell
            text = "".join(cell["text"])
            text = re.sub(r"[ \t\r\f\v]+", " ", text).strip()
            cell["value"] = html_lib.unescape(text)
            self.on_cell(cell)
            self.max_row = max(self.max_row, self.row + cell["rowspan"] - 1)
            self.col += cell["colspan"]
            self.max_col = max(self.max_col, self.col)
            self._cell = None

    def handle_data(self, data):
        if self._cell is not None:
            self._cell["text"].append(data)


class _AnalyzerSink:
    def __init__(self):
        self.analysis = HtmlSheetAnalysis(name="")
        self._first_row: dict[int, str] = {}

    def on_cell(self, cell):
        r, c = cell["row"], cell["col"]
        self.analysis.rows = max(self.analysis.rows, r + cell["rowspan"])
        self.analysis.cols = max(self.analysis.cols, c + cell["colspan"])
        if r == 0:
            self._first_row[c] = cell["value"]
        if cell["chip_style"].get("background-color"):
            self.analysis.chip_cells += 1
            if cell["value"]:
                self.analysis.chip_columns.setdefault(c, set()).add(cell["value"])
        if cell["colspan"] > 1 or cell["rowspan"] > 1:
            self.analysis.merged_cells += 1
        if cell.get("href"):
            self.analysis.hyperlinks += 1

    def on_col_width(self, col, width):
        pass

    def on_row_height(self, row, height):
        pass

    def finish(self, name: str, profile: MigrationProfile):
        self.analysis.name = name
        self.analysis.headers = dict(self._first_row)
        normalized_profile_headers = {
            normalize_text(h)
            for rule in profile.data.get("header_rules", [])
            for h in rule.get("headers", [])
        }
        self.analysis.profiled_headers = [
            text for text in self._first_row.values() if normalize_text(text) in normalized_profile_headers
        ]
        return self.analysis


def analyze_html_zip(zip_path: str | Path, profile: MigrationProfile) -> HtmlWorkbookAnalysis:
    path = Path(zip_path)
    if not path.exists():
        raise FileNotFoundError("Selecione um pacote HTML/ZIP válido.")
    sheets: list[HtmlSheetAnalysis] = []
    with zipfile.ZipFile(path) as zf:
        html_names = _html_sheet_names(zf)
        css_count = sum(1 for n in zf.namelist() if n.lower().endswith(".css"))
        for html_name in html_names:
            text = zf.read(html_name).decode("utf-8", "ignore")
            sink = _AnalyzerSink()
            parser = GoogleTableParser(sink.on_cell, sink.on_col_width, sink.on_row_height)
            parser.feed(text)
            sheet_name = Path(html_name).stem
            sheets.append(sink.finish(sheet_name, profile))
    return HtmlWorkbookAnalysis(path, sheets, len(sheets), css_count)


def _fmt_from_props(workbook, props: dict[str, str], chip: dict[str, str] | None = None, date=False, cache=None):
    chip = chip or {}
    key = tuple(sorted(props.items())) + (("chip", tuple(sorted(chip.items()))), ("date", date))
    if cache is not None and key in cache:
        return cache[key]
    opts = {}
    fill = chip.get("background-color") or props.get("background-color")
    color = chip.get("color") or props.get("color")
    if fill and fill.lower() not in {"transparent", "none"}:
        opts["bg_color"] = fill
        opts["pattern"] = 1
    if color:
        opts["font_color"] = color
    if props.get("font-weight", "").lower() in {"bold", "700", "600"}:
        opts["bold"] = True
    if props.get("font-style", "").lower() == "italic":
        opts["italic"] = True
    m = PT_RE.search(props.get("font-size", ""))
    if m:
        opts["font_size"] = float(m.group(1))
    font_name = props.get("font-family", "").split(",", 1)[0].strip(' "')
    if font_name and not font_name.lower().startswith("docs-"):
        opts["font_name"] = font_name
    align = props.get("text-align", "").lower()
    if align in {"left", "center", "right", "justify"}:
        opts["align"] = align
    valign = props.get("vertical-align", "").lower()
    if valign in {"top", "middle", "bottom"}:
        opts["valign"] = "vcenter" if valign == "middle" else valign
    if props.get("white-space", "").lower() not in {"nowrap", ""}:
        opts["text_wrap"] = True
    border_css = " ".join(v for k, v in props.items() if k.startswith("border-")).lower()
    if "solid" in border_css:
        opts["border"] = 1
        color_match = re.search(r"#[0-9a-f]{6}", border_css)
        if color_match:
            opts["border_color"] = color_match.group(0)
    if date:
        opts["num_format"] = "dd/mm/yyyy"
    fmt = workbook.add_format(opts)
    if cache is not None:
        cache[key] = fmt
    return fmt


def _profile_rule_for_header(profile: MigrationProfile, header: str):
    n = normalize_text(header)
    for rule in profile.data.get("header_rules", []):
        if n in {normalize_text(h) for h in rule.get("headers", [])}:
            return rule
    return None


def _date_rule_for_header(profile: MigrationProfile, header: str):
    n = normalize_text(header)
    for rule in profile.data.get("date_columns", []):
        if n in {normalize_text(h) for h in rule.get("headers", [])}:
            return rule
    return None


def _col_letter(col_zero: int) -> str:
    n = col_zero + 1
    out = ""
    while n:
        n, rem = divmod(n - 1, 26)
        out = chr(65 + rem) + out
    return out


def _formula_values_from_rule(rule: dict) -> list[str]:
    if rule.get("allowed_values"):
        return [str(x) for x in rule["allowed_values"]]
    vals = []
    for r in rule.get("rules", []):
        formula = r.get("formula", "")
        # Captura strings entre aspas; serve como fallback, não como fonte autoritativa.
        for v in re.findall(r'="([^"]+)"', formula):
            if v and v not in vals:
                vals.append(v)
        for v in re.findall(r'\{cell\}="([^"]+)"', formula):
            if v and v not in vals:
                vals.append(v)
    return vals


def _formula_rule_apply(ws, headers: dict[int, str], max_row: int, profile: MigrationProfile, nonempty_by_col: dict[int, set[int]], fmt_by_cell: dict[tuple[int, int], object], value_by_cell: dict[tuple[int, int], object]) -> int:
    count = 0
    by_norm = {normalize_text(v): c for c, v in headers.items()}
    for fr in profile.data.get("formula_rules", []):
        target = None
        for h in fr.get("target_headers", []):
            if normalize_text(h) in by_norm:
                target = by_norm[normalize_text(h)]
                break
        if target is None:
            continue
        sources = {}
        missing = False
        for key, aliases in fr.get("source_headers", {}).items():
            found = None
            for h in aliases:
                if normalize_text(h) in by_norm:
                    found = by_norm[normalize_text(h)]
                    break
            if found is None:
                missing = True
                break
            sources[key] = found
        if missing:
            continue
        start = int(fr.get("start_row", 2)) - 1
        # HTML não informa onde existiam fórmulas em branco. Para não inflar o workbook,
        # reconstruímos apenas linhas cujo resultado renderizado estava preenchido.
        rows = sorted(r for r in nonempty_by_col.get(target, set()) if r >= start)
        guard = fr.get("reconstruct_if") or {}
        for rr in rows:
            # Preserva exceções digitadas manualmente: só substitui o valor por fórmula
            # quando o resultado renderizado é coerente com a regra de negócio conhecida.
            if guard:
                gtype = guard.get("type")
                source_key = guard.get("source")
                source_col = sources.get(source_key) if source_key else None
                source_val = value_by_cell.get((rr, source_col)) if source_col is not None else None
                target_val = value_by_cell.get((rr, target))
                ok = True
                if gtype == "age_years":
                    if isinstance(source_val, datetime) and isinstance(target_val, int):
                        today = datetime.today().date()
                        birth = source_val.date()
                        expected = today.year - birth.year - ((today.month, today.day) < (birth.month, birth.day))
                        ok = (target_val == expected)
                    else:
                        ok = False
                elif gtype == "edate_months":
                    months = int(guard.get("months", 0))
                    if isinstance(source_val, datetime) and isinstance(target_val, datetime):
                        d = source_val.date()
                        y = d.year + (d.month - 1 + months) // 12
                        m = (d.month - 1 + months) % 12 + 1
                        day = min(d.day, calendar.monthrange(y, m)[1])
                        expected = datetime(y, m, day).date()
                        ok = (target_val.date() == expected)
                    else:
                        ok = False
                if not ok:
                    continue
            row1 = rr + 1
            mapping = {"row": row1, "target": f"{_col_letter(target)}{row1}"}
            for key, cc in sources.items():
                mapping[key] = f"{_col_letter(cc)}{row1}"
            formula = fr["formula"].format(**mapping)
            target_val = value_by_cell.get((rr, target))
            cached = target_val
            if isinstance(target_val, datetime):
                cached = (target_val - datetime(1899, 12, 30)).total_seconds() / 86400.0
            ws.write_formula(rr, target, formula, fmt_by_cell.get((rr, target)), cached)
            count += 1
    return count

def _exact_formula_apply(ws, sheet_name: str, profile: MigrationProfile, fmt_by_cell: dict[tuple[int, int], object]) -> int:
    count = 0
    for group in profile.data.get("exact_formula_groups", []):
        if normalize_text(group.get("sheet", "")) != normalize_text(sheet_name):
            continue
        for item in group.get("cells", []):
            cell_ref = item["cell"]
            m = re.match(r"([A-Z]+)(\d+)$", cell_ref.upper())
            fmt = None
            if m:
                cc = 0
                for ch in m.group(1): cc = cc * 26 + ord(ch) - 64
                fmt = fmt_by_cell.get((int(m.group(2)) - 1, cc - 1))
            ws.write_formula(cell_ref, item["formula"], fmt)
            count += 1
        for item in group.get("series", []):
            start_col = item["start_col"].upper()
            end_col = item["end_col"].upper()
            row = int(item["row"])
            def col_num(s):
                n=0
                for ch in s: n=n*26+ord(ch)-64
                return n
            for n in range(col_num(start_col), col_num(end_col)+1):
                x=n; col=""
                while x:
                    x,rem=divmod(x-1,26); col=chr(65+rem)+col
                formula=item["formula"].format(cell=f"{col}{row}", col=col, row=row)
                ws.write_formula(f"{col}{row}", formula, fmt_by_cell.get((row - 1, n - 1)))
                count += 1
    return count


def convert_html_zip(zip_path: str | Path, output_path: str | Path, profile: MigrationProfile) -> HtmlConversionResult:
    ensure_xlsxwriter()
    src = Path(zip_path)
    out = Path(output_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    result = HtmlConversionResult(out)
    formula_header_norms = set()
    for fr in profile.data.get("formula_rules", []):
        formula_header_norms.update(normalize_text(h) for h in fr.get("target_headers", []))
        for aliases in fr.get("source_headers", {}).values():
            formula_header_norms.update(normalize_text(h) for h in aliases)

    workbook = xlsxwriter.Workbook(str(out), {"constant_memory": False})
    workbook.set_properties({"title": src.stem, "comments": "Convertido pelo BHead M365 Migrator - HTML direto"})
    format_cache = {}
    used_names: set[str] = set()
    list_sheet = workbook.add_worksheet("__LISTAS_M365")
    list_col = 0
    first_visible = None

    with zipfile.ZipFile(src) as zf:
        html_names = _html_sheet_names(zf)
        for html_name in html_names:
            raw = zf.read(html_name).decode("utf-8", "ignore")
            class_styles = _extract_class_styles(raw)
            name = _safe_sheet_name(Path(html_name).stem, used_names)
            ws = workbook.add_worksheet(name)
            if first_visible is None:
                first_visible = ws
            result.sheets_created += 1

            headers: dict[int, str] = {}
            chip_values: dict[int, set[str]] = {}
            max_row = 0
            max_col = 0
            nonempty_by_col: dict[int, set[int]] = {}
            fmt_by_cell: dict[tuple[int, int], object] = {}
            value_by_cell: dict[tuple[int, int], object] = {}
            col_widths: dict[int, float] = {}

            def on_col_width(col, width):
                col_widths[col] = width
                try:
                    ws.set_column_pixels(col, col, max(10, int(width)))
                except AttributeError:
                    ws.set_column(col, col, max(2, width / 7.0))

            def on_row_height(row, height):
                try:
                    ws.set_row_pixels(row, max(5, int(height)))
                except AttributeError:
                    ws.set_row(row, max(5, height * 0.75))

            def on_cell(cell):
                nonlocal max_row, max_col
                r, c = cell["row"], cell["col"]
                max_row = max(max_row, r + cell["rowspan"])
                max_col = max(max_col, c + cell["colspan"])
                if r == 0:
                    headers[c] = cell["value"]
                if cell["chip_style"].get("background-color") and cell["value"]:
                    chip_values.setdefault(c, set()).add(cell["value"])
                    result.chip_cells += 1
                value, value_type = _convert_value(cell["value"])
                props = class_styles.get(cell["class"], {})
                fmt = _fmt_from_props(workbook, props, cell["chip_style"], date=value_type == "date", cache=format_cache)
                if cell["value"].strip():
                    nonempty_by_col.setdefault(c, set()).add(r)
                if r > 0 and normalize_text(headers.get(c, "")) in formula_header_norms:
                    value_by_cell[(r, c)] = value
                # Guardamos apenas referências de formato (baratas) para reescrever fórmulas sem perder bordas/alinhamento.
                fmt_by_cell[(r, c)] = fmt
                if cell["colspan"] > 1 or cell["rowspan"] > 1:
                    r2 = r + cell["rowspan"] - 1
                    c2 = c + cell["colspan"] - 1
                    try:
                        ws.merge_range(r, c, r2, c2, value, fmt)
                    except Exception:
                        ws.write(r, c, value, fmt)
                elif cell.get("href") and isinstance(value, str) and value:
                    try:
                        ws.write_url(r, c, cell["href"], fmt, string=value)
                    except Exception:
                        ws.write(r, c, value, fmt)
                else:
                    ws.write(r, c, value, fmt)
                result.cells_written += 1

            parser = GoogleTableParser(on_cell, on_col_width, on_row_height)
            parser.feed(raw)

            # Datas precisam de formatação semântica persistente na coluna.
            # Na V0.2.1 a máscara era aplicada apenas à célula importada; ao editar
            # no Excel Online, o valor podia ser reinterpretado como número geral.
            # Aqui aplicamos a máscara também como estilo padrão da coluna e
            # uma validação de data para proteger novas edições.
            for col, header in headers.items():
                date_rule = _date_rule_for_header(profile, header)
                if not date_rule:
                    continue
                num_format = date_rule.get("number_format", "dd/mm/yyyy")
                date_col_fmt = workbook.add_format({"num_format": num_format})
                width_px = max(10, int(col_widths.get(col, 80)))
                try:
                    ws.set_column_pixels(col, col, width_px, date_col_fmt)
                except AttributeError:
                    ws.set_column(col, col, max(2, width_px / 7.0), date_col_fmt)
                result.date_columns_formatted += 1

                start = int(date_rule.get("start_row", 2)) - 1
                end = max(max_row - 1, int(date_rule.get("minimum_end_row", max_row)) - 1)
                if end >= start:
                    try:
                        ws.data_validation(start, col, end, col, {
                            "validate": "date",
                            "criteria": "between",
                            "minimum": datetime(1900, 1, 1),
                            "maximum": datetime(2100, 12, 31),
                            "input_title": date_rule.get("input_title", "Data"),
                            "input_message": date_rule.get(
                                "input_message",
                                "Digite a data no formato DD/MM/AAAA.\nPara anos anteriores a 2000, prefira digitar manualmente.",
                            ),
                            "error_title": date_rule.get("error_title", "Data inválida"),
                            "error_message": date_rule.get(
                                "error_message",
                                "Informe uma data válida no formato DD/MM/AAAA.",
                            ),
                        })
                        result.date_validations_added += 1
                    except Exception:
                        # A máscara da coluna é obrigatória; a validação é uma proteção adicional.
                        pass

            # Congela o cabeçalho quando a planilha tem estrutura tabular.
            if headers and profile.data.get("freeze_header", True):
                ws.freeze_panes(1, 0)
            if headers and profile.data.get("autofilter", True) and max_col > 0 and max_row > 1:
                try:
                    ws.autofilter(0, 0, max_row - 1, max_col - 1)
                except Exception:
                    pass

            # Fórmulas conhecidas do modelo (HTML não carrega a fórmula original).
            result.formulas_reconstructed += _formula_rule_apply(ws, headers, max_row, profile, nonempty_by_col, fmt_by_cell, value_by_cell)
            result.formulas_reconstructed += _exact_formula_apply(ws, name, profile, fmt_by_cell)

            # Validações e formatação condicional por perfil.
            for col, header in headers.items():
                rule = _profile_rule_for_header(profile, header)
                if not rule:
                    # Em colunas com chips desconhecidas, cria lista inferida apenas se houver 2+ valores.
                    observed = sorted(chip_values.get(col, set()))
                    if len(observed) >= 2:
                        start = 1
                        end = max(max_row - 1, 1)
                        for rr, val in enumerate(observed):
                            list_sheet.write(rr, list_col, val)
                        dv_name = f"_BHM_DV_{list_col + 1}"
                        workbook.define_name(dv_name, f"='__LISTAS_M365'!${_col_letter(list_col)}$1:${_col_letter(list_col)}${len(observed)}")
                        ws.data_validation(start, col, end, col, {"validate": "list", "source": f"={dv_name}"})
                        list_col += 1
                        result.validations_added += 1
                        result.inferred_validations += 1
                    continue

                start = int(rule.get("start_row", 2)) - 1
                end = max(max_row - 1, int(rule.get("minimum_end_row", max_row)) - 1)
                values = [str(x) for x in rule.get("allowed_values", [])]
                if not values:
                    values = sorted(chip_values.get(col, set()))
                if not values:
                    values = _formula_values_from_rule(rule)
                if values:
                    for rr, val in enumerate(values):
                        list_sheet.write(rr, list_col, val)
                    dv_name = f"_BHM_DV_{list_col + 1}"
                    workbook.define_name(dv_name, f"='__LISTAS_M365'!${_col_letter(list_col)}$1:${_col_letter(list_col)}${len(values)}")
                    ws.data_validation(start, col, end, col, {"validate": "list", "source": f"={dv_name}"})
                    list_col += 1
                    result.validations_added += 1

                first_cell = f"{_col_letter(col)}{start + 1}"
                for cr in rule.get("rules", []):
                    style = profile.data.get("styles", {}).get(cr.get("style"), {})
                    formula = cr.get("formula", "").replace("{cell}", first_cell)
                    if not formula:
                        continue
                    fmt = workbook.add_format({
                        "bg_color": "#" + style.get("fill", "FFFFFF").lstrip("#"),
                        "font_color": "#" + style.get("font", "000000").lstrip("#"),
                    })
                    ws.conditional_format(start, col, end, col, {
                        "type": "formula", "criteria": "=" + formula.lstrip("="), "format": fmt
                    })
                    result.conditional_rules_added += 1

    if result.sheets_created == 0:
        result.warnings.append("Nenhuma planilha HTML foi encontrada no ZIP.")
    result.warnings.append(
        "O HTML do Google não contém as fórmulas nem as definições completas das validações. "
        "O aplicativo reconstrói o que estiver mapeado no perfil e infere listas desconhecidas pelos valores observados."
    )
    if first_visible is not None:
        first_visible.activate()
        list_sheet.hide()
    workbook.close()
    return result


def write_html_report_csv(path: str | Path, analysis: HtmlWorkbookAnalysis, result: HtmlConversionResult):
    p = Path(path)
    p.parent.mkdir(parents=True, exist_ok=True)
    with p.open("w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["BHead M365 Migrator", "Relatório HTML direto"])
        w.writerow(["Origem", str(analysis.source)])
        w.writerow(["Status análise", analysis.status])
        w.writerow(["Planilhas HTML", analysis.html_files])
        w.writerow(["CSS", analysis.css_files])
        w.writerow(["Planilhas criadas", result.sheets_created])
        w.writerow(["Células gravadas", result.cells_written])
        w.writerow(["Chips detectados", result.chip_cells])
        w.writerow(["Validações adicionadas", result.validations_added])
        w.writerow(["Validações inferidas", result.inferred_validations])
        w.writerow(["Colunas de data formatadas", result.date_columns_formatted])
        w.writerow(["Validações de data adicionadas", result.date_validations_added])
        w.writerow(["Regras condicionais adicionadas", result.conditional_rules_added])
        w.writerow(["Fórmulas reconstruídas pelo perfil", result.formulas_reconstructed])
        w.writerow([])
        w.writerow(["Planilha", "Linhas", "Colunas", "Chips", "Cabeçalhos do perfil", "Mesclagens", "Links"])
        for s in analysis.sheets:
            w.writerow([s.name, s.rows, s.cols, s.chip_cells, len(s.profiled_headers), s.merged_cells, s.hyperlinks])
        w.writerow([])
        w.writerow(["Ressalvas"])
        for warning in result.warnings:
            w.writerow([warning])
