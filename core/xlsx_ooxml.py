from __future__ import annotations

import csv
import re
import shutil
import tempfile
import zipfile
from copy import deepcopy
from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterable
import xml.etree.ElementTree as ET

from .profile import MigrationProfile, normalize_text

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
ET.register_namespace("", NS_MAIN)
ET.register_namespace("r", NS_REL)

M = f"{{{NS_MAIN}}}"
R = f"{{{NS_REL}}}"
P = f"{{{NS_PKG_REL}}}"

GOOGLE_ONLY_FUNCTIONS = {
    "QUERY", "IMPORTRANGE", "GOOGLEFINANCE", "GOOGLETRANSLATE",
    "REGEXMATCH", "REGEXEXTRACT", "REGEXREPLACE", "ARRAYFORMULA"
}


@dataclass
class SheetAnalysis:
    name: str
    xml_path: str
    max_row: int = 1
    max_col: int = 1
    headers: dict[str, str] = field(default_factory=dict)  # col -> text
    header_row: int = 1
    validations: int = 0
    list_validations: int = 0
    conditional_blocks: int = 0
    conditional_rules: int = 0
    formulas: int = 0
    google_functions: set[str] = field(default_factory=set)
    profiled_headers: list[str] = field(default_factory=list)


@dataclass
class WorkbookAnalysis:
    source: Path
    sheets: list[SheetAnalysis]
    html_reference: Path | None = None

    @property
    def google_functions(self) -> set[str]:
        result: set[str] = set()
        for sh in self.sheets:
            result |= sh.google_functions
        return result

    @property
    def status(self) -> str:
        if self.google_functions:
            return "REVISAR"
        return "APROVADO"


class XlsxError(RuntimeError):
    pass


def _col_to_num(col: str) -> int:
    n = 0
    for ch in col:
        if ch.isalpha():
            n = n * 26 + ord(ch.upper()) - 64
    return n


def _num_to_col(n: int) -> str:
    out = ""
    while n:
        n, rem = divmod(n - 1, 26)
        out = chr(65 + rem) + out
    return out or "A"


def _split_cell_ref(ref: str) -> tuple[str, int]:
    m = re.match(r"([A-Z]+)(\d+)$", ref or "")
    if not m:
        return "A", 1
    return m.group(1), int(m.group(2))


def _read_shared_strings(z: zipfile.ZipFile) -> list[str]:
    try:
        root = ET.fromstring(z.read("xl/sharedStrings.xml"))
    except KeyError:
        return []
    strings: list[str] = []
    for si in root.findall(f"{M}si"):
        strings.append("".join(t.text or "" for t in si.iter(f"{M}t")))
    return strings


def _cell_text(cell: ET.Element, shared: list[str]) -> str:
    ctype = cell.attrib.get("t")
    if ctype == "inlineStr":
        return "".join(t.text or "" for t in cell.iter(f"{M}t"))
    v = cell.find(f"{M}v")
    if v is None or v.text is None:
        return ""
    if ctype == "s":
        try:
            return shared[int(v.text)]
        except (ValueError, IndexError):
            return ""
    return v.text


def _workbook_sheet_map(z: zipfile.ZipFile) -> list[tuple[str, str]]:
    wb = ET.fromstring(z.read("xl/workbook.xml"))
    rels_root = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    rels = {r.attrib["Id"]: r.attrib["Target"] for r in rels_root.findall(f"{P}Relationship")}
    result = []
    sheets = wb.find(f"{M}sheets")
    if sheets is None:
        return result
    for s in sheets:
        rid = s.attrib.get(f"{R}id")
        target = rels.get(rid or "")
        if target:
            target = target.lstrip("/")
            if not target.startswith("xl/"):
                target = "xl/" + target
            result.append((s.attrib.get("name", "Planilha"), target))
    return result


def _dimension(root: ET.Element) -> tuple[int, int]:
    dim = root.find(f"{M}dimension")
    if dim is not None:
        ref = dim.attrib.get("ref", "A1")
        end = ref.split(":")[-1]
        col, row = _split_cell_ref(end)
        return max(1, row), max(1, _col_to_num(col))
    max_row = 1
    max_col = 1
    for c in root.iter(f"{M}c"):
        col, row = _split_cell_ref(c.attrib.get("r", "A1"))
        max_row = max(max_row, row)
        max_col = max(max_col, _col_to_num(col))
    return max_row, max_col


def _headers(root: ET.Element, shared: list[str], row_number: int = 1) -> dict[str, str]:
    result: dict[str, str] = {}
    row = root.find(f".//{M}row[@r='{row_number}']")
    if row is None:
        return result
    for c in row.findall(f"{M}c"):
        col, _ = _split_cell_ref(c.attrib.get("r", "A1"))
        result[col] = _cell_text(c, shared)
    return result


def _detect_header_row(root: ET.Element, shared: list[str], profile: MigrationProfile | None = None, max_scan: int = 20) -> tuple[int, dict[str, str]]:
    """Detecta o cabeçalho real em planilhas com notas/linhas vazias antes da tabela.

    O Google Sheets pode exportar folhas cujo cabeçalho começa na linha 2, 3 ou
    posterior. A versão anterior assumia sempre a linha 1, o que impedia o modo
    híbrido de associar corretamente cores/regras em modelos com múltiplas regras de cores.
    """
    best_row = 1
    best_headers: dict[str, str] = _headers(root, shared, 1)
    best_score = -1.0
    for row_number in range(1, max_scan + 1):
        headers = _headers(root, shared, row_number)
        values = [str(v).strip() for v in headers.values() if str(v).strip()]
        if not values:
            continue
        profile_hits = 0
        if profile is not None:
            profile_hits = sum(1 for v in values if profile.find_header_rule(v))
        text_like = sum(1 for v in values if not re.fullmatch(r"[-+]?\d+(?:[.,]\d+)?", v))
        # Perfil é o sinal mais forte; densidade textual desempata folhas novas.
        score = profile_hits * 1000 + text_like * 5 + len(values) - (row_number * 0.01)
        if score > best_score:
            best_score = score
            best_row = row_number
            best_headers = headers
    return best_row, best_headers


def _google_functions(root: ET.Element) -> tuple[int, set[str]]:
    count = 0
    found: set[str] = set()
    for f in root.iter(f"{M}f"):
        if not f.text:
            continue
        count += 1
        upper = f.text.upper()
        for fn in GOOGLE_ONLY_FUNCTIONS:
            if re.search(rf"\b{re.escape(fn)}\s*\(", upper):
                found.add(fn)
    return count, found


def analyze_xlsx(path: str | Path, profile: MigrationProfile, html_reference: str | Path | None = None) -> WorkbookAnalysis:
    source = Path(path)
    if not source.exists():
        raise XlsxError(f"Arquivo não encontrado: {source}")
    if source.suffix.lower() != ".xlsx":
        raise XlsxError("A V0.1 trabalha com arquivos .xlsx exportados do Google Sheets.")

    analyses: list[SheetAnalysis] = []
    with zipfile.ZipFile(source, "r") as z:
        shared = _read_shared_strings(z)
        for name, xml_path in _workbook_sheet_map(z):
            try:
                root = ET.fromstring(z.read(xml_path))
            except KeyError:
                continue
            max_row, max_col = _dimension(root)
            header_row, headers = _detect_header_row(root, shared, profile)
            dvs = root.find(f"{M}dataValidations")
            dv_list = [] if dvs is None else list(dvs.findall(f"{M}dataValidation"))
            list_count = sum(1 for dv in dv_list if dv.attrib.get("type") == "list")
            cf_blocks = root.findall(f"{M}conditionalFormatting")
            cf_rules = sum(len(cf.findall(f"{M}cfRule")) for cf in cf_blocks)
            formula_count, google = _google_functions(root)
            profiled = [h for h in headers.values() if profile.find_header_rule(h)]
            analyses.append(SheetAnalysis(
                name=name,
                xml_path=xml_path,
                max_row=max_row,
                max_col=max_col,
                headers=headers,
                header_row=header_row,
                validations=len(dv_list),
                list_validations=list_count,
                conditional_blocks=len(cf_blocks),
                conditional_rules=cf_rules,
                formulas=formula_count,
                google_functions=google,
                profiled_headers=profiled,
            ))
    return WorkbookAnalysis(source=source, sheets=analyses, html_reference=Path(html_reference) if html_reference else None)


def _style_key(style: dict[str, str]) -> tuple[str, str]:
    return style["fill"].upper().replace("#", ""), style["font"].upper().replace("#", "")


def _read_dxf_style(dxf: ET.Element) -> tuple[str, str] | None:
    fill = dxf.find(f"{M}fill/{M}patternFill/{M}fgColor")
    font = dxf.find(f"{M}font/{M}color")
    if fill is None or font is None:
        return None
    fg = fill.attrib.get("rgb", "").upper()
    fc = font.attrib.get("rgb", "").upper()
    if len(fg) == 8:
        fg = fg[2:]
    if len(fc) == 8:
        fc = fc[2:]
    return fg, fc


def _make_dxf(style: dict[str, str]) -> ET.Element:
    fill_hex, font_hex = _style_key(style)
    dxf = ET.Element(f"{M}dxf")
    font = ET.SubElement(dxf, f"{M}font")
    ET.SubElement(font, f"{M}color", {"rgb": "FF" + font_hex})
    fill = ET.SubElement(dxf, f"{M}fill")
    pattern = ET.SubElement(fill, f"{M}patternFill", {"patternType": "solid"})
    ET.SubElement(pattern, f"{M}fgColor", {"rgb": "FF" + fill_hex})
    ET.SubElement(pattern, f"{M}bgColor", {"rgb": "FF" + fill_hex})
    return dxf


def _ensure_styles(styles_xml: bytes, profile: MigrationProfile) -> tuple[bytes, dict[str, int]]:
    root = ET.fromstring(styles_xml)
    dxfs = root.find(f"{M}dxfs")
    if dxfs is None:
        dxfs = ET.Element(f"{M}dxfs", {"count": "0"})
        table_styles = root.find(f"{M}tableStyles")
        if table_styles is not None:
            idx = list(root).index(table_styles)
            root.insert(idx, dxfs)
        else:
            root.append(dxfs)

    existing: dict[tuple[str, str], int] = {}
    for idx, dxf in enumerate(dxfs.findall(f"{M}dxf")):
        key = _read_dxf_style(dxf)
        if key:
            existing[key] = idx

    ids: dict[str, int] = {}
    for name, style in profile.data.get("styles", {}).items():
        key = _style_key(style)
        if key in existing:
            ids[name] = existing[key]
            continue
        idx = len(dxfs.findall(f"{M}dxf"))
        dxfs.append(_make_dxf(style))
        existing[key] = idx
        ids[name] = idx

    dxfs.attrib["count"] = str(len(dxfs.findall(f"{M}dxf")))
    return ET.tostring(root, encoding="utf-8", xml_declaration=True), ids


def _existing_cf_signatures(root: ET.Element) -> set[tuple[str, str]]:
    sigs: set[tuple[str, str]] = set()
    for cf in root.findall(f"{M}conditionalFormatting"):
        sqref = cf.attrib.get("sqref", "")
        for rule in cf.findall(f"{M}cfRule"):
            formula = rule.find(f"{M}formula")
            sigs.add((sqref, formula.text if formula is not None and formula.text else ""))
    return sigs


def _max_priority(root: ET.Element) -> int:
    out = 0
    for rule in root.iter(f"{M}cfRule"):
        try:
            out = max(out, int(rule.attrib.get("priority", "0")))
        except ValueError:
            pass
    return out


def _insert_conditional_formatting(root: ET.Element, profile: MigrationProfile, dxf_ids: dict[str, int], headers: dict[str, str], max_row: int) -> tuple[int, int]:
    existing = _existing_cf_signatures(root)
    priority = _max_priority(root)
    blocks_added = 0
    rules_added = 0

    # Localiza ponto correto da sequência OOXML: antes de dataValidations/hyperlinks/printOptions/...
    children = list(root)
    insert_before_tags = {
        f"{M}dataValidations", f"{M}hyperlinks", f"{M}printOptions", f"{M}pageMargins",
        f"{M}pageSetup", f"{M}headerFooter", f"{M}rowBreaks", f"{M}colBreaks",
        f"{M}customProperties", f"{M}cellWatches", f"{M}ignoredErrors", f"{M}smartTags",
        f"{M}drawing", f"{M}legacyDrawing", f"{M}legacyDrawingHF", f"{M}picture",
        f"{M}oleObjects", f"{M}controls", f"{M}webPublishItems", f"{M}tableParts", f"{M}extLst"
    }
    insertion_index = len(children)
    for i, child in enumerate(children):
        if child.tag in insert_before_tags:
            insertion_index = i
            break

    offset = 0
    for col, header in sorted(headers.items(), key=lambda kv: _col_to_num(kv[0])):
        rule_def = profile.find_header_rule(header)
        if not rule_def:
            continue
        start_row = int(rule_def.get("start_row", 2))
        end_row = max(max_row, int(rule_def.get("minimum_end_row", max_row)))
        sqref = f"{col}{start_row}:{col}{end_row}"
        anchor = f"{col}{start_row}"
        new_rules: list[ET.Element] = []
        for rule in rule_def.get("rules", []):
            formula = rule["formula"].format(cell=anchor, col=col, row=start_row)
            if (sqref, formula) in existing:
                continue
            style_name = rule["style"]
            priority += 1
            cf_rule = ET.Element(f"{M}cfRule", {
                "type": "expression",
                "dxfId": str(dxf_ids[style_name]),
                "priority": str(priority),
                "stopIfTrue": "1",
            })
            f = ET.SubElement(cf_rule, f"{M}formula")
            f.text = formula
            new_rules.append(cf_rule)
            existing.add((sqref, formula))
        if new_rules:
            cf = ET.Element(f"{M}conditionalFormatting", {"sqref": sqref})
            for r in new_rules:
                cf.append(r)
            root.insert(insertion_index + offset, cf)
            offset += 1
            blocks_added += 1
            rules_added += len(new_rules)
    return blocks_added, rules_added


def _should_process_sheet(name: str, profile: MigrationProfile) -> bool:
    wanted = profile.sheet_names
    return not wanted or name in wanted


def convert_xlsx(source: str | Path, output: str | Path, profile: MigrationProfile) -> dict[str, object]:
    source = Path(source)
    output = Path(output)
    output.parent.mkdir(parents=True, exist_ok=True)
    if source.resolve() == output.resolve():
        raise XlsxError("O arquivo de saída deve ser diferente do original.")

    with zipfile.ZipFile(source, "r") as zin:
        shared = _read_shared_strings(zin)
        sheet_map = _workbook_sheet_map(zin)
        styles_bytes, dxf_ids = _ensure_styles(zin.read("xl/styles.xml"), profile)

        replacements: dict[str, bytes] = {"xl/styles.xml": styles_bytes}
        details = []
        total_blocks = 0
        total_rules = 0

        for sheet_name, xml_path in sheet_map:
            if not _should_process_sheet(sheet_name, profile):
                continue
            root = ET.fromstring(zin.read(xml_path))
            max_row, _ = _dimension(root)
            header_row, headers = _detect_header_row(root, shared, profile)
            blocks, rules = _insert_conditional_formatting(root, profile, dxf_ids, headers, max_row)
            if blocks or rules:
                replacements[xml_path] = ET.tostring(root, encoding="utf-8", xml_declaration=True)
            details.append({"sheet": sheet_name, "blocks_added": blocks, "rules_added": rules})
            total_blocks += blocks
            total_rules += rules

        with zipfile.ZipFile(output, "w") as zout:
            for item in zin.infolist():
                data = replacements.get(item.filename, zin.read(item.filename))
                zout.writestr(item, data)

    # Integridade básica: o pacote final deve ser um ZIP/XLSX legível e conter workbook.xml.
    with zipfile.ZipFile(output, "r") as ztest:
        bad = ztest.testzip()
        if bad:
            raise XlsxError(f"Falha de integridade no arquivo gerado: {bad}")
        if "xl/workbook.xml" not in ztest.namelist():
            raise XlsxError("Arquivo gerado não contém xl/workbook.xml")

    return {
        "source": str(source),
        "output": str(output),
        "blocks_added": total_blocks,
        "rules_added": total_rules,
        "sheets": details,
    }


def html_zip_summary(path: str | Path | None) -> dict[str, object]:
    if not path:
        return {"provided": False, "html_files": 0, "css_files": 0}
    p = Path(path)
    if not p.exists():
        return {"provided": True, "error": "Arquivo HTML/ZIP não encontrado"}
    if p.suffix.lower() != ".zip":
        return {"provided": True, "error": "Na V0.1 a referência deve ser um .zip exportado do Google"}
    with zipfile.ZipFile(p, "r") as z:
        names = z.namelist()
        return {
            "provided": True,
            "html_files": sum(1 for n in names if n.lower().endswith(".html")),
            "css_files": sum(1 for n in names if n.lower().endswith(".css")),
            "files": names[:30],
        }


def write_report_csv(report_path: str | Path, before: WorkbookAnalysis, after: WorkbookAnalysis | None, conversion: dict[str, object] | None, html_info: dict[str, object]) -> None:
    report_path = Path(report_path)
    report_path.parent.mkdir(parents=True, exist_ok=True)
    rows: list[list[object]] = []
    rows.append(["BHead M365 Migrator", "Relatório de Migração V0.1"])
    rows.append(["Arquivo origem", str(before.source)])
    rows.append(["Status pré-análise", before.status])
    rows.append(["Referência HTML/ZIP", "Sim" if html_info.get("provided") else "Não"])
    if html_info.get("provided"):
        rows.append(["HTMLs no ZIP", html_info.get("html_files", 0)])
        rows.append(["CSS no ZIP", html_info.get("css_files", 0)])
    if conversion:
        rows.append(["Arquivo convertido", conversion.get("output", "")])
        rows.append(["Blocos de formatação adicionados", conversion.get("blocks_added", 0)])
        rows.append(["Regras de formatação adicionadas", conversion.get("rules_added", 0)])
    if after:
        rows.append(["Status pós-conversão", after.status])
    rows.append([])
    rows.append(["Planilha", "Linhas", "Colunas", "Validações", "Dropdowns", "Blocos CF", "Regras CF", "Fórmulas", "Funções Google", "Cabeçalhos com perfil"])
    after_by_name = {s.name: s for s in after.sheets} if after else {}
    for sh in before.sheets:
        target = after_by_name.get(sh.name, sh)
        rows.append([
            sh.name, sh.max_row, sh.max_col, sh.validations, sh.list_validations,
            target.conditional_blocks, target.conditional_rules, sh.formulas,
            ", ".join(sorted(sh.google_functions)) or "-",
            ", ".join(sh.profiled_headers) or "-",
        ])
    with report_path.open("w", newline="", encoding="utf-8-sig") as f:
        csv.writer(f, delimiter=";").writerows(rows)
