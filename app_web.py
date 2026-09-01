from __future__ import annotations

import hashlib
import os
import shutil
import tempfile
import time
import uuid
from pathlib import Path

import streamlit as st

from core.profile import load_profile
from core.hybrid import analyze_hybrid, convert_hybrid, write_hybrid_report_csv
from core.html_to_xlsx import analyze_html_zip, convert_html_zip, write_html_report_csv

APP_VERSION = "0.4.6"
APP_TITLE = "BHead M365 Migrator Web"
BASE_DIR = Path(__file__).resolve().parent
PROFILE_PATH = Path(os.environ.get("BHEAD_M365_PROFILE", str(BASE_DIR / "profiles" / "example_profile.json")))
TEMP_PREFIX = "bhead_m365_web_"
TEMP_MAX_AGE_SECONDS = 6 * 60 * 60

st.set_page_config(
    page_title=f"{APP_TITLE} V{APP_VERSION}",
    page_icon="📊",
    layout="wide",
)


@st.cache_resource
def get_profile():
    return load_profile(PROFILE_PATH)


def purge_old_temp_dirs() -> None:
    root = Path(tempfile.gettempdir())
    now = time.time()
    for path in root.glob(f"{TEMP_PREFIX}*"):
        try:
            if path.is_dir() and now - path.stat().st_mtime > TEMP_MAX_AGE_SECONDS:
                shutil.rmtree(path, ignore_errors=True)
        except OSError:
            pass


def ensure_session_dir() -> Path:
    if "session_dir" not in st.session_state:
        path = Path(tempfile.gettempdir()) / f"{TEMP_PREFIX}{uuid.uuid4().hex}"
        path.mkdir(parents=True, exist_ok=True)
        st.session_state.session_dir = str(path)
    path = Path(st.session_state.session_dir)
    path.mkdir(parents=True, exist_ok=True)
    try:
        os.utime(path, None)
    except OSError:
        pass
    return path


def clear_session_workdir() -> None:
    """Remove temporários da migração atual sem expor essa manutenção ao operador."""
    old = st.session_state.get("session_dir")
    if old:
        shutil.rmtree(old, ignore_errors=True)
    st.session_state.pop("session_dir", None)


def upload_digest(uploaded) -> str:
    if uploaded is None:
        return "-"
    raw = uploaded.getvalue()
    return hashlib.sha256(raw).hexdigest()


def save_upload(uploaded, target: Path) -> Path:
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_bytes(uploaded.getvalue())
    return target


def original_xlsx_name(uploaded_name: str) -> str:
    """Preserva o nome original do XLSX enviado pelo operador."""
    name = Path(uploaded_name).name
    if not name.lower().endswith(".xlsx"):
        name = f"{Path(name).stem}.xlsx"
    return name


def html_output_name(uploaded_name: str) -> str:
    """Deriva um nome de planilha limpo quando a origem é somente HTML/ZIP."""
    stem = Path(uploaded_name).stem.strip()
    for suffix in (" (HTML)", "(HTML)", " - HTML", "_HTML"):
        if stem.upper().endswith(suffix.upper()):
            stem = stem[: -len(suffix)].rstrip()
            break
    return f"{stem or 'Planilha'}.xlsx"


def input_signature(mode: str, xlsx_upload, html_upload) -> str:
    return "|".join([mode, upload_digest(xlsx_upload), upload_digest(html_upload)])


def clear_result_if_inputs_changed(signature: str) -> None:
    previous = st.session_state.get("input_signature")
    if previous != signature:
        # Ao trocar XLSX, HTML/ZIP ou modo, a migração anterior deixa de ser válida.
        # O WebApp limpa automaticamente análise, downloads e temporários associados.
        clear_session_workdir()
        for key in [
            "analysis", "analysis_mode", "analysis_signature", "output_name", "output_bytes",
            "report_name", "report_bytes", "last_status", "last_warnings",
        ]:
            st.session_state.pop(key, None)
        st.session_state.input_signature = signature


def hybrid_progress_callback(bar, label_slot):
    def cb(pct: float, label: str):
        pct = max(0.0, min(100.0, float(pct)))
        bar.progress(int(pct))
        label_slot.caption(f"{label} — {pct:.0f}%")
    return cb


def format_hybrid_analysis(a):
    rows = []
    formula_by_sheet = {}
    if getattr(a, "formula_compat", None):
        formula_by_sheet = {s.name: s for s in a.formula_compat.sheets}
    for sh in a.xlsx_analysis.sheets:
        fc = formula_by_sheet.get(sh.name)
        rows.append({
            "Planilha": sh.name,
            "Cabeçalho": sh.header_row,
            "Linhas": sh.max_row,
            "Colunas": sh.max_col,
            "Validações": sh.validations,
            "Listas": sh.list_validations,
            "Fórmulas": fc.formula_cells if fc else sh.formulas,
            "IFS": fc.ifs_formulas if fc else 0,
            "Matrizes": fc.array_formulas if fc else 0,
            "#REF! origem": fc.broken_reference_formulas if fc else 0,
            "Regras CF": sh.conditional_rules,
            "Cabeçalhos do perfil": len(sh.profiled_headers),
        })
    return rows


def format_html_analysis(a):
    rows = []
    for sh in a.sheets:
        rows.append({
            "Planilha": sh.name,
            "Linhas": sh.rows,
            "Colunas": sh.cols,
            "Chips": sh.chip_cells,
            "Colunas com chips": len(sh.chip_columns),
            "Cabeçalhos do perfil": len(sh.profiled_headers),
            "Mesclagens": sh.merged_cells,
            "Links": sh.hyperlinks,
        })
    return rows


def render_analysis():
    a = st.session_state.get("analysis")
    mode = st.session_state.get("analysis_mode")
    if a is None:
        return

    st.subheader("Resultado da análise")
    status = a.status
    if status == "COMPATÍVEL PARA CONVERSÃO":
        st.success(status)
    elif status == "REVISAR":
        st.warning(status)
    else:
        st.error(status)

    if mode == "hybrid":
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Planilhas XLSX", len(a.xlsx_analysis.sheets))
        c2.metric("Planilhas HTML", len(a.html_sheets))
        c3.metric("Associadas", a.matched_sheets)
        c4.metric("Colunas com chips", a.html_chip_columns)
        st.dataframe(format_hybrid_analysis(a), use_container_width=True, hide_index=True)
        fc = getattr(a, "formula_compat", None)
        if fc:
            st.markdown("#### Compatibilidade de fórmulas")
            f1, f2, f3, f4, f5 = st.columns(5)
            f1.metric("Células com fórmula", fc.formula_cells)
            f2.metric("IFS → IF", fc.ifs_formulas)
            f3.metric("Matrizes 1 célula", fc.single_cell_array_formulas)
            f4.metric("#REF! na origem", fc.broken_reference_formulas)
            f5.metric("Erros no HTML", len(fc.html_error_cells))
            st.caption(
                "Na conversão, IFS é normalizado para IF aninhado, fórmulas matriciais de uma célula "
                "reconhecidas como escalares são normalizadas e o arquivo é marcado para recálculo completo no Microsoft 365."
            )
            if fc.html_reference_values:
                st.info(
                    f"Referência Google disponível para {fc.html_reference_values:,} células de fórmula "
                    "a partir do HTML exportado.".replace(",", ".")
                )
            source_issues = []
            for fs in fc.sheets:
                source_issues.extend(fs.issues)
            source_issues.extend(fc.html_error_cells)
            if source_issues:
                with st.expander("Erros/referências que já existem na origem", expanded=True):
                    for issue in source_issues[:30]:
                        st.warning(f"{issue.sheet} · {issue.cell} · {issue.detail}")
                    if len(source_issues) > 30:
                        st.caption(f"+ {len(source_issues) - 30} ocorrências adicionais no relatório técnico.")
            if fc.source_value_mismatch_count:
                with st.expander(f"Divergências XLSX × HTML ({fc.source_value_mismatch_count})", expanded=False):
                    for issue in fc.source_value_mismatches[:20]:
                        st.warning(f"{issue.sheet} · {issue.cell} · {issue.detail}")
        if a.xlsx_analysis.google_functions:
            st.warning("Funções Google detectadas: " + ", ".join(sorted(a.xlsx_analysis.google_functions)))
        if a.missing_validation_columns:
            st.warning(
                "Colunas coloridas no HTML sem validação correspondente no XLSX:\n\n- "
                + "\n- ".join(a.missing_validation_columns[:30])
            )
        for warning in a.warnings:
            st.warning(warning)
    else:
        c1, c2, c3 = st.columns(3)
        c1.metric("HTML", a.html_files)
        c2.metric("CSS", a.css_files)
        c3.metric("Planilhas", len(a.sheets))
        st.dataframe(format_html_analysis(a), use_container_width=True, hide_index=True)
        st.info(
            "O HTML não contém as fórmulas originais nem a definição completa das validações. "
            "O modo homologado reconstrói o que estiver mapeado no perfil."
        )


def render_downloads():
    output_bytes = st.session_state.get("output_bytes")
    if not output_bytes:
        return

    st.subheader("Conversão concluída")
    st.success("Arquivo pronto para homologação no Excel Online.")
    col1, col2 = st.columns([1, 1])
    with col1:
        st.download_button(
            "⬇️ Baixar XLSX convertido",
            data=output_bytes,
            file_name=st.session_state.output_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )
    report_bytes = st.session_state.get("report_bytes")
    if report_bytes:
        with col2:
            st.download_button(
                "⬇️ Baixar relatório técnico",
                data=report_bytes,
                file_name=st.session_state.report_name,
                mime="text/csv",
                use_container_width=True,
            )

    warnings = st.session_state.get("last_warnings", [])
    if warnings:
        with st.expander("Ressalvas técnicas", expanded=True):
            for w in warnings:
                st.warning(w)


purge_old_temp_dirs()
profile = get_profile()

st.title("BHEAD M365 MIGRATOR WEB")
st.caption(f"Google Workspace → Microsoft 365 Online · V{APP_VERSION}")
st.info(
    "Os arquivos são processados temporariamente no computador que hospeda este WebApp. "
    "As pastas temporárias antigas são removidas automaticamente após 6 horas."
)

mode_label = st.radio(
    "Modo de migração",
    [
        "NOVO MODELO — XLSX + HTML/ZIP",
        "MODELO HOMOLOGADO — somente HTML/ZIP",
    ],
    index=0,
    horizontal=True,
)
mode = "hybrid" if mode_label.startswith("NOVO") else "html"

if mode == "hybrid":
    st.caption(
        "XLSX = fonte estrutural (fórmulas, validações e referências) · "
        "HTML = fonte visual e valor de referência do Google (chips, cores e resultados calculados)."
    )
else:
    st.caption(
        "Use somente em modelos já conhecidos pelo perfil. Fórmulas/listas ausentes no HTML "
        "são reconstruídas pelo perfil homologado."
    )

left, right = st.columns(2)
with left:
    xlsx_upload = None
    if mode == "hybrid":
        xlsx_upload = st.file_uploader(
            "Arquivo XLSX exportado do Google",
            type=["xlsx"],
            accept_multiple_files=False,
            key="xlsx_upload",
        )
with right:
    html_upload = st.file_uploader(
        "Pacote HTML/ZIP exportado do Google",
        type=["zip"],
        accept_multiple_files=False,
        key="html_upload",
    )

report_requested = st.checkbox("Gerar relatório técnico da migração", value=False)

signature = input_signature(mode, xlsx_upload, html_upload)
clear_result_if_inputs_changed(signature)

analysis = st.session_state.get("analysis")
analysis_mode = st.session_state.get("analysis_mode")
analysis_signature = st.session_state.get("analysis_signature")
analysis_ready = (
    analysis is not None
    and analysis_mode == mode
    and analysis_signature == signature
    and getattr(analysis, "status", None) != "INCOMPATÍVEL"
)

buttons = st.columns([1, 1, 5])
analyze_clicked = buttons[0].button("🔎 ANALISAR", type="primary", use_container_width=True)
convert_clicked = buttons[1].button(
    "⚙️ CONVERTER PARA M365",
    use_container_width=True,
    disabled=not analysis_ready,
    help=(
        "Disponível após analisar com sucesso os arquivos selecionados."
        if not analysis_ready
        else "Converter os arquivos já analisados para Microsoft 365."
    ),
)

if not analysis_ready:
    if analysis is not None and getattr(analysis, "status", None) == "INCOMPATÍVEL":
        st.caption("⛔ Conversão bloqueada: a análise classificou o arquivo como INCOMPATÍVEL.")
    else:
        st.caption("🔒 Analise os arquivos selecionados para habilitar a conversão.")


def validate_inputs() -> bool:
    if html_upload is None:
        st.error("Selecione o pacote HTML/ZIP exportado do Google.")
        return False
    if mode == "hybrid" and xlsx_upload is None:
        st.error("No modo NOVO MODELO, selecione também o arquivo XLSX.")
        return False
    return True


def prepare_input_paths():
    workdir = ensure_session_dir()
    html_path = save_upload(html_upload, workdir / "entrada_google_html.zip")
    xlsx_path = None
    if xlsx_upload is not None:
        xlsx_path = save_upload(xlsx_upload, workdir / "entrada_google.xlsx")
    return workdir, xlsx_path, html_path


if analyze_clicked and validate_inputs():
    workdir, xlsx_path, html_path = prepare_input_paths()
    bar = st.progress(0)
    label = st.empty()
    try:
        if mode == "hybrid":
            analysis = analyze_hybrid(
                xlsx_path,
                html_path,
                profile,
                progress_cb=hybrid_progress_callback(bar, label),
            )
        else:
            label.caption("Analisando pacote HTML/ZIP...")
            bar.progress(20)
            analysis = analyze_html_zip(html_path, profile)
            bar.progress(100)
            label.caption("Análise concluída — 100%")
        st.session_state.analysis = analysis
        st.session_state.analysis_mode = mode
        st.session_state.analysis_signature = signature
        # Recarrega a interface para que o botão CONVERTER seja habilitado
        # somente após a análise concluída dos arquivos atuais.
        st.rerun()
    except Exception as exc:
        st.session_state.pop("analysis_signature", None)
        st.exception(exc)


if convert_clicked and validate_inputs():
    workdir, xlsx_path, html_path = prepare_input_paths()
    analysis = st.session_state.get("analysis")
    analysis_mode = st.session_state.get("analysis_mode")
    analysis_signature = st.session_state.get("analysis_signature")
    bar = st.progress(0)
    label = st.empty()

    try:
        if (
            analysis is None
            or analysis_mode != mode
            or analysis_signature != signature
        ):
            st.error("Faça a análise dos arquivos atuais antes de converter.")
            st.stop()
        if getattr(analysis, "status", None) == "INCOMPATÍVEL":
            st.error("Conversão bloqueada: a análise classificou o arquivo como INCOMPATÍVEL.")
            st.stop()

        if mode == "hybrid":
            base_name = Path(xlsx_upload.name).stem
            download_name = original_xlsx_name(xlsx_upload.name)
            output = workdir / download_name
            label.caption("Convertendo XLSX híbrido...")
            result = convert_hybrid(
                xlsx_path,
                html_path,
                output,
                profile,
                html_sheets=analysis.html_sheets,
                progress_cb=hybrid_progress_callback(bar, label),
            )
            warnings = list(analysis.warnings) + list(result.warnings)
            auto_report = analysis.status == "REVISAR" or bool(warnings)
            report = workdir / f"{base_name}_Relatorio_Hibrido.csv"
            if report_requested or auto_report:
                write_hybrid_report_csv(report, analysis, result)
        else:
            base_name = Path(html_upload.name).stem
            download_name = html_output_name(html_upload.name)
            output = workdir / download_name
            label.caption("Convertendo HTML para Excel Online...")
            bar.progress(25)
            result = convert_html_zip(html_path, output, profile)
            bar.progress(90)
            warnings = list(result.warnings)
            auto_report = analysis.status == "REVISAR" or bool(warnings)
            report = workdir / f"{base_name}_Relatorio_Migracao.csv"
            if report_requested or auto_report:
                write_html_report_csv(report, analysis, result)
            bar.progress(100)
            label.caption("Conversão concluída — 100%")

        if mode == "hybrid" and (result.ifs_formulas_converted or result.array_formulas_normalized):
            st.info(
                f"Formula Compatibility Engine: {result.ifs_formulas_converted} IFS convertido(s) para IF aninhado; "
                f"{result.array_formulas_normalized} fórmula(s) matricial(is) de uma célula normalizada(s)."
            )
        st.session_state.output_name = output.name
        st.session_state.output_bytes = output.read_bytes()
        st.session_state.last_status = analysis.status
        st.session_state.last_warnings = warnings
        if report.exists():
            st.session_state.report_name = report.name
            st.session_state.report_bytes = report.read_bytes()
        else:
            st.session_state.pop("report_name", None)
            st.session_state.pop("report_bytes", None)
        st.success("Conversão concluída. Faça a homologação do arquivo no Excel Online.")
    except Exception as exc:
        st.exception(exc)

render_analysis()
render_downloads()

with st.expander("Sobre o perfil e o destino"):
    st.write(f"**Perfil:** {profile.name}")
    st.write(f"**Destino:** {profile.target}")
    st.write(
        "O modo híbrido é recomendado para planilhas novas. Na V0.4.5, o Formula Compatibility Engine "
        "também analisa fórmulas, erros de origem e valores de referência do HTML. O modo HTML direto deve "
        "ser usado somente para modelos já homologados pelo perfil."
    )
