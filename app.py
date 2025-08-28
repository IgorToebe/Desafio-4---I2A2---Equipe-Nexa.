import os
import sys
import io
import time
from datetime import date
from typing import List, Tuple, Optional
import base64

import streamlit as st


# Paths
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
XLSX_DIR = os.path.join(BASE_DIR, "xlsx")
RUN_LOG_PATH = os.path.join(BASE_DIR, "run.log")
def _inject_css():
        st.markdown(
                """
                <style>
                :root {
                    --bg: #ffffff;
                    --panel: #ffffff;
                    --text: #000000;
                    --muted: #6b7280;
                    --accent: #6C63FF;
                    --accent-2: #22c55e;
                    --warn: #f59e0b;
                    --border: #e5e7eb;
                    /* Button palette (soft backgrounds + opposite text tone) */
                    --btn-bg: #eef2ff;       /* soft indigo */
                    --btn-text: #1f2937;     /* dark slate */
                    --btn-bg-hover: #e0e7ff;
                    --btn2-bg: #ecfdf5;      /* soft green */
                    --btn2-text: #065f46;    /* dark teal */
                    --btn2-bg-hover: #d1fae5;
                }
                [data-testid="stAppViewContainer"] {
                    background: var(--bg);
                }
                body, [data-testid="stAppViewContainer"] { color: var(--text); }
                p, span, label, li, div, code, pre { color: var(--text); }
                [data-testid="stHeader"] {
                    background: linear-gradient(90deg, rgba(108,99,255,0.10), rgba(34,197,94,0.10));
                }
                h1, h2, h3 { color: var(--text); }
                .hero {
                    padding: 24px 20px;
                    margin: -16px -16px 8px -16px;
                    background: linear-gradient(135deg, #eef2ff 0%, #ecfdf5 100%);
                    border-bottom: 1px solid var(--border);
                    text-align: center;
                }
                .hero .title {
                    font-size: 40px; font-weight: 800; letter-spacing: 0.5px;
                    color: #1f2937;
                }
                .hero .badge { font-size: 14px; padding: 8px 12px; }
                .hero .subtitle { font-size: 18px; color:#4b5563; margin-top: 4px; }
                .hero-center { display:flex; flex-direction:column; align-items:center; justify-content:center; gap:12px; }
                .badge {
                    display: inline-block;
                    font-size: 12px; font-weight: 600; color: #374151;
                    background: #e0e7ff; padding: 6px 10px; border-radius: 999px;
                }
                .accent-divider { height: 3px; border: 0; background: linear-gradient(90deg, var(--accent), var(--accent-2)); border-radius: 3px; }
                .panel {
                    padding: 18px 18px 8px 18px; border-radius: 12px; background: var(--panel);
                    border: 1px solid var(--border); box-shadow: 0 6px 16px rgba(31,41,55,0.05);
                    margin-bottom: 16px;
                }
                .panel h2 { margin-top: 0; }
                .footer {
                    margin-top: 18px; padding: 12px 8px; text-align: center; color: var(--muted);
                    border-top: 1px solid var(--border);
                }
                /* Buttons (soft + opposite text tone) */
                .stButton>button {
                    background: var(--btn-bg);
                    color: var(--btn-text);
                    border: 1px solid var(--border);
                    border-radius: 10px;
                    padding: 10px 16px;
                    box-shadow: 0 2px 6px rgba(31,41,55,0.08);
                    transition: background 120ms ease, box-shadow 120ms ease, transform 60ms ease;
                }
                .stButton>button:hover { background: var(--btn-bg-hover); box-shadow: 0 3px 8px rgba(31,41,55,0.10); }
                .stButton>button:active { transform: translateY(1px); }
                .stButton>button:disabled { filter: grayscale(0.1); opacity: 0.65; }

                .stDownloadButton>button {
                    background: var(--btn2-bg);
                    color: var(--btn2-text);
                    border: 1px solid var(--border);
                    border-radius: 10px;
                    padding: 10px 16px;
                    box-shadow: 0 2px 6px rgba(31,41,55,0.08);
                    transition: background 120ms ease, box-shadow 120ms ease, transform 60ms ease;
                }
                .stDownloadButton>button:hover { background: var(--btn2-bg-hover); box-shadow: 0 3px 8px rgba(31,41,55,0.10); }
                .stDownloadButton>button:active { transform: translateY(1px); }
                /* Progress */
                .stProgress > div > div { background: linear-gradient(90deg, var(--accent), var(--accent-2)) !important; }
                /* Text areas */
                textarea {
                    background: #fbfbff !important; border-radius: 10px !important; border: 1px solid var(--border) !important;
                    color: #111827 !important; /* dark gray text */
                }
                textarea:disabled {
                    color: #111827 !important;
                    -webkit-text-fill-color: #111827 !important; /* ensure in webkit */
                    opacity: 1 !important; /* avoid washed out look */
                }
                .stTextArea textarea { color: #111827 !important; }
                /* File uploader */
                [data-testid="stFileUploader"]>div {
                    background: var(--panel); border: 1px dashed var(--border); padding: 12px; border-radius: 12px;
                }
                /* Make the 'Browse files' button text white */
                [data-testid="stFileUploader"] button {
                    color: #ffffff !important;
                }
                </style>
                """,
                unsafe_allow_html=True,
        )


def _hero_header():
    """Render a hero header with optional logo at BASE_DIR/logo.png."""
    logo_path = os.path.join(BASE_DIR, "logo.png")
    logo_img_tag = ""
    if os.path.exists(logo_path):
        try:
            with open(logo_path, "rb") as f:
                b64 = base64.b64encode(f.read()).decode("ascii")
            # 5x enlargement from 64px -> 320px height
            logo_img_tag = (
                f'<img src="data:image/png;base64,{b64}" '
                f'style="height:320px; width:auto; max-width:100%; '
                f'border-radius:8px; box-shadow:0 2px 8px rgba(0,0,0,0.08);" />'
            )
        except Exception:
            logo_img_tag = ""
    # Build hero with optional logo
    html = f"""
            <div class=\"hero\">
                <div class=\"hero-center\">
                    {logo_img_tag}
                    <div>
                        <div class=\"badge\">Automação de VR</div>
                        <div class=\"subtitle\">Front-end para cálculo e geração de planilhas</div>
                    </div>
                </div>
            </div>
            <hr class=\"accent-divider\" />
    """
    st.markdown(html, unsafe_allow_html=True)


def get_expected_filenames() -> List[str]:
    """Read expected .xlsx filenames from consolidar_bases.ARQUIVOS plus template file."""
    try:
        sys.path.insert(0, BASE_DIR)
        import consolidar_bases as cb  # type: ignore

        expected = list({os.path.basename(p) for p in cb.ARQUIVOS.values()})
        # Include template used by vr_agent
        expected.append("VR MENSAL 05.2025.xlsx")
        # Ensure unique and stable order
        expected = sorted(set(expected), key=lambda s: s.lower())
        return expected
    except Exception:
        # Fallback to a known set if import fails
        return sorted({
            "ATIVOS.xlsx",
            "ADMISSÃO ABRIL.xlsx",
            "FÉRIAS.xlsx",
            "DESLIGADOS.xlsx",
            "Base sindicato x valor.xlsx",
            "Base dias uteis.xlsx",
            "AFASTAMENTOS.xlsx",
            "EXTERIOR.xlsx",
            "VR MENSAL 05.2025.xlsx",
        })


def save_uploaded_files(files: List[Tuple[str, bytes]]):
    os.makedirs(XLSX_DIR, exist_ok=True)
    for fname, data in files:
        out_path = os.path.join(XLSX_DIR, fname)
        with open(out_path, "wb") as f:
            f.write(data)


def _expected_vr_filename_for_date(fim: Optional[date]) -> Optional[str]:
    if not fim:
        return None
    return f"VR_MENSAL_{fim.month:02d}.{fim.year}_FINAL.xlsx"


def _resolve_vr_output_path(fim: Optional[date]) -> Optional[str]:
    """Return the expected VR output path for the given end date, or fallback to latest found."""
    try:
        os.makedirs(XLSX_DIR, exist_ok=True)
        # First, try by expected name from fim
        exp_name = _expected_vr_filename_for_date(fim)
        if exp_name:
            p = os.path.join(XLSX_DIR, exp_name)
            if os.path.exists(p):
                return p
        # Fallback: find latest VR_MENSAL_*_FINAL.xlsx
        candidates = []
        for fn in os.listdir(XLSX_DIR):
            if fn.startswith("VR_MENSAL_") and fn.endswith("_FINAL.xlsx"):
                full = os.path.join(XLSX_DIR, fn)
                try:
                    mtime = os.path.getmtime(full)
                except Exception:
                    mtime = 0
                candidates.append((mtime, full))
        if candidates:
            candidates.sort(key=lambda t: t[0], reverse=True)
            return candidates[0][1]
    except Exception:
        pass
    return None


def stream_process_and_logs(inicio: date, fim: date):
    """Run consolidar_bases.py with dates and stream stdout + run.log into UI."""
    import subprocess

    # Prepare command
    py = sys.executable
    script = os.path.join(BASE_DIR, "consolidar_bases.py")
    args = [py, script, "--inicio", inicio.isoformat(), "--fim", fim.isoformat()]

    # Ensure GEMINI_API_KEY is available to the child process (Streamlit Cloud secrets don't
    # automatically propagate to subprocesses). If present, inject into env.
    env = os.environ.copy()
    try:
        api_key = None
        try:
            # st is already imported at module top
            api_key = st.secrets.get("GEMINI_API_KEY") or st.secrets.get("gemini_api_key")
            if not api_key:
                nested = st.secrets.get("secrets")
                if nested and isinstance(nested, dict):
                    api_key = nested.get("GEMINI_API_KEY") or nested.get("gemini_api_key")
        except Exception:
            api_key = None
        if api_key:
            env["GEMINI_API_KEY"] = str(api_key)
    except Exception:
        pass

    # Prepare run.log tailing (binary safe)
    log_pos = 0
    if os.path.exists(RUN_LOG_PATH):
        try:
            with open(RUN_LOG_PATH, "rb") as f:
                f.seek(0, os.SEEK_END)
                log_pos = f.tell()
        except Exception:
            log_pos = 0

    # Start process
    proc = subprocess.Popen(
        args,
        cwd=BASE_DIR,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        bufsize=1,
        universal_newlines=True,
        encoding="utf-8",
        errors="replace",
    env=env,
    )
    # Progress UI
    progress = st.progress(0.0)
    status = st.empty()
    log_area = st.empty()
    collected: List[str] = []

    # Persist session state for this run
    st.session_state["run_in_progress"] = True
    st.session_state["last_inicio"] = inicio
    st.session_state["last_fim"] = fim
    st.session_state["logs_buffer"] = ""

    # Markers to update progress
    markers = [
        ("Iniciando consolidação de bases", 0.05),
        ("Executando merges", 0.20),
        ("Aplicando filtros de exclusão", 0.35),
        ("Validando e corrigindo dados", 0.50),
        ("Salvando resultado", 0.65),
        ("Iniciando geração do arquivo de VR", 0.75),
        ("Output file written successfully", 1.00),
        ("[CONCLUIDO] VR gerado com sucesso", 1.00),
    ]
    seen = set()
    cur_prog = 0.0

    def append(text: str):
        collected.append(text)
        # Limit memory usage
        if len(collected) > 2000:
            del collected[: len(collected) - 2000]
        buffer = "".join(collected)
        # Use a disabled text area for scrollable logs
        log_area.text_area("Logs", value=buffer, height=300, disabled=True)
        st.session_state["logs_buffer"] = buffer

        # Update progress markers
        nonlocal cur_prog
        for token, val in markers:
            if token in text and token not in seen:
                seen.add(token)
                if val is not None:
                    cur_prog = max(cur_prog, float(val))
                    progress.progress(cur_prog)
                    status.write(f"Etapa: {token}")

    # Stream loop with spinner
    with st.spinner("Executando e carregando logs…"):
        last_log_read_time = 0.0
        while True:
            line = proc.stdout.readline() if proc.stdout else ""
            if line:
                append(line)
            # Periodically read run.log for agent entries
            now = time.time()
            if now - last_log_read_time > 0.4:
                last_log_read_time = now
                try:
                    if os.path.exists(RUN_LOG_PATH):
                        with open(RUN_LOG_PATH, "rb") as f:
                            f.seek(log_pos)
                            chunk = f.read()
                            if chunk:
                                try:
                                    text = chunk.decode("utf-8", errors="ignore")
                                except Exception:
                                    text = ""
                                if text:
                                    append(text)
                                log_pos += len(chunk)
                except Exception:
                    pass
            if proc.poll() is not None:
                # Drain remaining stdout
                rest = proc.stdout.read() if proc.stdout else ""
                if rest:
                    append(rest)
                # Final read from run.log
                try:
                    if os.path.exists(RUN_LOG_PATH):
                        with open(RUN_LOG_PATH, "rb") as f:
                            f.seek(log_pos)
                            chunk = f.read()
                            if chunk:
                                try:
                                    text = chunk.decode("utf-8", errors="ignore")
                                except Exception:
                                    text = ""
                                if text:
                                    append(text)
                                log_pos += len(chunk)
                except Exception:
                    pass
                break
            # Yield control to UI
            time.sleep(0.05)
    st.session_state["run_in_progress"] = False
    st.session_state["last_rc"] = proc.returncode or 0
    # Mark files ready state
    base_path = os.path.join(XLSX_DIR, "BaseConsolidada.xlsx")
    vr_path_dyn = _resolve_vr_output_path(fim)
    if vr_path_dyn:
        st.session_state["last_vr_path"] = vr_path_dyn
    st.session_state["files_ready"] = os.path.exists(base_path) or (vr_path_dyn and os.path.exists(vr_path_dyn))
    return proc.returncode or 0


def main():
    st.set_page_config(page_title="VR App", page_icon="🍽️", layout="wide")
    _inject_css()
    _hero_header()

    # Section: Upload
    st.markdown("<div class='panel'>", unsafe_allow_html=True)
    st.header("📁 Carregar planilhas (.xlsx)")
    expected = get_expected_filenames()
    st.caption("Arquivos esperados:")
    st.write("\n".join(f"- {name}" for name in expected))

    uploaded = st.file_uploader(
        "Selecione as planilhas (múltiplos arquivos)",
        type=["xlsx"],
        accept_multiple_files=True,
    )

    # Validate names
    invalid_names: List[str] = []
    collisions: List[str] = []
    staged: List[Tuple[str, bytes]] = []
    if uploaded:
        provided_names = [u.name for u in uploaded]
        for u in uploaded:
            if u.name not in expected:
                invalid_names.append(u.name)
            out_path = os.path.join(XLSX_DIR, u.name)
            if os.path.exists(out_path):
                collisions.append(u.name)
            # buffer data
            staged.append((u.name, bytes(u.getbuffer())))

    if invalid_names:
        st.error(
            "Foram encontrados arquivos com nomes não esperados: "
            + ", ".join(invalid_names)
            + ". Renomeie conforme a lista esperada e tente novamente."
        )

    # Overwrite confirmation flow
    if uploaded and not invalid_names:
        if collisions:
            st.warning(
                "Os seguintes arquivos já existem em xlsx/ e serão substituídos: "
                + ", ".join(sorted(set(collisions)))
            )
            if st.button("Confirmar substituição e salvar"):
                save_uploaded_files(staged)
                st.success("Arquivos salvos em xlsx/ com substituição.")
                st.session_state["uploaded_saved_last"] = [name for name, _ in staged]
        else:
            if st.button("Salvar arquivos"):
                save_uploaded_files(staged)
                st.success("Arquivos salvos em xlsx/.")
                st.session_state["uploaded_saved_last"] = [name for name, _ in staged]

    st.markdown("</div>", unsafe_allow_html=True)
    st.markdown("<div class='panel'>", unsafe_allow_html=True)

    # Section: Dates
    st.header("📅 Seleção de competência")
    c1, c2 = st.columns(2)
    # Persist date selections
    if "last_inicio" not in st.session_state:
        st.session_state["last_inicio"] = date(2025, 4, 15)
    if "last_fim" not in st.session_state:
        st.session_state["last_fim"] = date(2025, 5, 15)
    with c1:
        d_inicio = st.date_input("Data de Início", value=st.session_state["last_inicio"], format="YYYY-MM-DD")
    with c2:
        d_fim = st.date_input("Data de Fim", value=st.session_state["last_fim"], format="YYYY-MM-DD")
    st.session_state["last_inicio"] = d_inicio
    st.session_state["last_fim"] = d_fim

    st.markdown("</div>", unsafe_allow_html=True)
    st.markdown("<div class='panel'>", unsafe_allow_html=True)

    # Section: Execute
    st.header("⚙️ Execução do processo")
    run_clicked = st.button("Executar Cálculo", disabled=st.session_state.get("run_in_progress", False))

    if run_clicked:
        # Pre-check: ensure required files exist
        missing = [name for name in expected if not os.path.exists(os.path.join(XLSX_DIR, name))]
        if missing:
            st.error("Arquivos ausentes em xlsx/: " + ", ".join(missing))
            st.stop()

        st.subheader("📝 Log de execução")
        rc = stream_process_and_logs(d_inicio, d_fim)
        if rc == 0:
            st.success("Processo concluído.")
        else:
            st.error(f"Processo finalizado com código {rc}. Verifique os logs.")

        # Downloads (fresh run)
        st.subheader("🧾 Downloads")
        base_path = os.path.join(XLSX_DIR, "BaseConsolidada.xlsx")
        vr_path = _resolve_vr_output_path(d_fim)
        if os.path.exists(base_path):
            with open(base_path, "rb") as f:
                st.download_button(
                    label="Baixar BaseConsolidada.xlsx",
                    data=f.read(),
                    file_name="BaseConsolidada.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
        if vr_path and os.path.exists(vr_path):
            with open(vr_path, "rb") as f:
                st.download_button(
                    label=f"Baixar {os.path.basename(vr_path)}",
                    data=f.read(),
                    file_name=os.path.basename(vr_path),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
    else:
        # If not running, show last logs and downloads if available (persisted state)
        last_buf = st.session_state.get("logs_buffer")
        if last_buf:
            st.subheader("📝 Últimos logs")
            st.text_area("Logs", value=last_buf, height=300, disabled=True)

        # Downloads (persist)
        base_path = os.path.join(XLSX_DIR, "BaseConsolidada.xlsx")
        vr_path = st.session_state.get("last_vr_path") or _resolve_vr_output_path(st.session_state.get("last_fim"))
        if os.path.exists(base_path) or (vr_path and os.path.exists(vr_path)):
            st.subheader("🧾 Downloads")
        if os.path.exists(base_path):
            with open(base_path, "rb") as f:
                st.download_button(
                    label="Baixar BaseConsolidada.xlsx",
                    data=f.read(),
                    file_name="BaseConsolidada.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
        if vr_path and os.path.exists(vr_path):
            with open(vr_path, "rb") as f:
                st.download_button(
                    label=f"Baixar {os.path.basename(vr_path)}",
                    data=f.read(),
                    file_name=os.path.basename(vr_path),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

    st.markdown("</div>", unsafe_allow_html=True)
    st.markdown("<div class='footer'>Automação de VR</div>", unsafe_allow_html=True)


if __name__ == "__main__":
    main()
