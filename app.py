# -*- coding: utf-8 -*-
import streamlit as st
import tempfile, os, shutil, time, io, zipfile, re, base64, pytz
from pathlib import Path
from datetime import datetime
from typing import Tuple, List, Dict
from concurrent.futures import ThreadPoolExecutor, as_completed
from openpyxl import load_workbook

# Usa a tua função existente que cria os .xlsx por PDF
from xylella_processor import process_pdf

# ───────────────────────────────────────────────
# Configuração base
# ───────────────────────────────────────────────
st.set_page_config(page_title="Xylella Processor", page_icon="🧪", layout="centered")
st.title("🧪 Xylella Processor")
st.caption("Processa PDFs de requisições Xylella e gera automaticamente 1 ficheiro Excel por requisição.")

# ───────────────────────────────────────────────
# CSS — estilo azul + animações suaves
# ───────────────────────────────────────────────
st.markdown("""
<style>
.stButton > button[kind="primary"]{
  background:#CA4300!important;border:1px solid #CA4300!important;color:#fff!important;
  font-weight:600!important;border-radius:6px!important;transition:background-color .2s ease-in-out!important;
}
.stButton > button[kind="primary"]:hover{background:#A13700!important;border-color:#A13700!important;}
[data-testid="stFileUploader"]>div:first-child{border:2px dashed #CA4300!important;border-radius:10px!important;padding:1rem!important}

.file-box{border-radius:8px;padding:.6rem 1rem;margin-bottom:.5rem;opacity:0;animation:fadeIn .3s ease forwards}
@keyframes fadeIn{from{opacity:0;transform:translateY(-4px)}to{opacity:1;transform:translateY(0)}}
.file-box.active{background:#E8F1FB;border-left:4px solid #2B6CB0}
.file-box.success{background:#e6f9ee;border-left:4px solid #1a7f37}
.file-box.warning{background:#fff8e5;border-left:4px solid #e6a100}
.file-box.error{background:#fdeaea;border-left:4px solid #cc0000}
.file-title{font-size:.9rem;font-weight:600;color:#1A365D}
.file-sub{font-size:.8rem;color:#2A4365}
.dots::after{content:'...';display:inline-block;animation:dots 1.5s steps(4,end) infinite}
@keyframes dots{
  0%,20%{color:rgba(42,67,101,0);text-shadow:.25em 0 0 rgba(42,67,101,0),.5em 0 0 rgba(42,67,101,0)}
  40%{color:#2A4365;text-shadow:.25em 0 0 rgba(42,67,101,0),.5em 0 0 rgba(42,67,101,0)}
  60%{text-shadow:.25em 0 0 #2A4365,.5em 0 0 rgba(42,67,101,0)}
  80%,100%{text-shadow:.25em 0 0 #2A4365,.5em 0 0 #2A4365}
}
.clean-btn{background:#fff!important;border:1px solid #ccc!important;color:#333!important;font-weight:600!important;
border-radius:8px!important;padding:.5rem 1.2rem!important;transition:all .2s ease}
.clean-btn:hover{border-color:#999!important;color:#000!important}
</style>
""", unsafe_allow_html=True)

# ───────────────────────────────────────────────
# Estado e reset
# ───────────────────────────────────────────────
if "stage" not in st.session_state:
    st.session_state.stage = "idle"
if "upload_paths" not in st.session_state:
    st.session_state.upload_paths = []

def reset_app():
    st.session_state.stage = "idle"
    st.session_state.upload_paths = []

# ───────────────────────────────────────────────
# Auxiliares
# ───────────────────────────────────────────────
def read_e1_counts(xlsx_path: str) -> Tuple[int | None, int | None]:
    """
    Lê E1 (ou célula equivalente) com o formato 'Nº Amostras: X / Y'
    e devolve (esperado=declaradas, processado=efetivas),
    assumindo que o primeiro número é o declarado.
    """
    try:
        wb = load_workbook(xlsx_path, data_only=True)
        ws = wb.active
        candidates = [ws["E1"].value, ws.cell(1, 5).value, ws.cell(1, 6).value]
        val = next((v for v in candidates if isinstance(v, str) and "/" in v), "")
        m = re.search(r"(\d+)\s*/\s*(\d+)", val)
        if m:
            declared, processed = int(m.group(1)), int(m.group(2))
            return declared, processed
    except Exception:
        pass
    return None, None

def collect_debug_files(output_dirs: List[Path]) -> List[str]:
    debug_files = []
    for pattern in ["*_ocr_debug.txt", "process_log.csv", "process_summary_*.txt"]:
        for d in output_dirs:
            debug_files += [str(f) for f in d.glob(pattern)]
    return debug_files

def build_zip_with_summary(excel_files: List[str], debug_files: List[str], summary_text: str) -> bytes:
    mem = io.BytesIO()
    with zipfile.ZipFile(mem, "w", zipfile.ZIP_DEFLATED) as z:
        for p in excel_files:
            if os.path.exists(p):
                z.write(p, arcname=os.path.basename(p))
        for d in debug_files:
            if os.path.exists(d):
                z.write(d, arcname=f"debug/{os.path.basename(d)}")
        z.writestr("summary.txt", summary_text)
    mem.seek(0)
    return mem.read()

def render_box(placeholder, kind: str, title: str, subtitle_html: str):
    html = f"<div class='file-box {kind}'><div class='file-title'>{title}</div><div class='file-sub'>{subtitle_html}</div></div>"
    placeholder.markdown(html, unsafe_allow_html=True)

def process_one_pdf(pdf_path: str, final_dir: Path) -> Dict:
    created = process_pdf(pdf_path) or []
    req_count = len(created)
    samples_total = 0
    discrepancies = []

    for fp in created:
        dest = final_dir / Path(fp).name
        try:
            shutil.copy(fp, dest)
        except Exception:
            pass
        exp, proc = read_e1_counts(str(dest))
        if proc:
            samples_total += proc
        if exp is not None and proc is not None and exp != proc:
            discrepancies.append({"xlsx": dest.name, "proc": proc, "exp": exp})

    return {
        "pdf_name": os.path.basename(pdf_path),
        "created": [str(final_dir / Path(fp).name) for fp in created],
        "req_count": req_count,
        "samples_total": samples_total,
        "discrepancies": discrepancies,
    }

# ───────────────────────────────────────────────
# Interface principal
# ───────────────────────────────────────────────
if st.session_state.stage == "idle":
    uploads = st.file_uploader("📂 Carrega um ou vários PDFs", type=["pdf"], accept_multiple_files=True, key="file_uploader")
    start = st.button("📄 Processar ficheiros de Input", type="primary", disabled=not uploads)

    if start and uploads:
        session_dir = Path(tempfile.mkdtemp(prefix="xylella_session_"))
        saved = []
        for up in uploads:
            tmp_pdf = session_dir / up.name
            with open(tmp_pdf, "wb") as f:
                f.write(up.getbuffer())
            saved.append({"name": up.name, "path": str(tmp_pdf)})
        st.session_state.upload_paths = saved
        st.session_state.stage = "processing"
        st.rerun()
    else:
        if not uploads:
            st.info("💡 Carrega um ficheiro PDF para ativar o botão de processamento.")

elif st.session_state.stage == "processing":
    st.info("⏳ A processar ficheiros... aguarde até o processo terminar.")
    items = st.session_state.upload_paths or []
    if not items:
        st.warning("⚠️ Nenhum ficheiro encontrado. Volte atrás e carregue novamente os PDFs.")
        st.session_state.stage = "idle"
        st.rerun()

    final_dir = Path.cwd() / "output_final"
    final_dir.mkdir(exist_ok=True)

    placeholders = [st.empty() for _ in items]
    progress = st.progress(0.0)
    start_ts = time.time()
    results: List[Dict] = []
    futures = {}
    max_workers = min(4, len(items)) or 1

    with ThreadPoolExecutor(max_workers=max_workers) as ex:
        for idx, it in enumerate(items):
            os.environ["OUTPUT_DIR"] = str(Path(tempfile.mkdtemp(prefix="xylella_out_")))
            render_box(
                placeholders[idx],
                "active",
                f"📄 {it['name']}",
                f"Ficheiro {idx+1} de {len(items)} — a processar<span class='dots'></span>"
            )
            futures[ex.submit(process_one_pdf, it["path"], final_dir)] = idx

        done_count = 0
        for fut in as_completed(futures):
            idx = futures[fut]
            try:
                res = fut.result()
            except Exception as e:
                res = {"pdf_name": os.path.basename(items[idx]["path"]), "created": [], "req_count": 0, "samples_total": 0, "discrepancies": [], "error": str(e)}
            results.append(res)

            ph = placeholders[idx]
            title = f"📄 {res['pdf_name']}"
            if res.get("error"):
                render_box(ph, "error", title, "❌ Erro: nenhum ficheiro gerado.")
            else:
                warn = bool(res["discrepancies"])
                box = "warning" if warn else "success"
                sub = f"<b>{res['req_count']}</b> requisição(ões), <b>{res['samples_total']}</b> amostras."
                if warn:
                    sub += f"<br>⚠️ <b>{len(res['discrepancies'])}</b> discrepância(s)."
                render_box(ph, box, title, sub)

            done_count += 1
            progress.progress(done_count / len(items))
            time.sleep(0.05)

    # ──────────────── SECÇÃO FINAL ────────────────
    total_time = time.time() - start_ts
    lisbon_tz = pytz.timezone("Europe/Lisbon")
    now_local = datetime.now(lisbon_tz)

    all_excel = []
    summary_lines = []
    results_sorted = [None] * len(items)
    for res in results:
        try:
            pos = next(i for i, it in enumerate(items) if os.path.basename(it["path"]) == res["pdf_name"])
        except StopIteration:
            pos = 0
        results_sorted[pos] = res
    results = results_sorted

    total_samples_overall = 0
    warning_count = 0
    error_count = 0

    for res in results:
        if not res:
            continue
        pdf_name = res["pdf_name"]
        req_count = res["req_count"]
        samples_total = res["samples_total"]
        total_samples_overall += samples_total

        if res.get("error") or req_count == 0:
            error_count += 1
            summary_lines.append(f"📄 {pdf_name}: erro - nenhum ficheiro gerado.")
            continue

        line = f"📄 {pdf_name}: {req_count} requisição(ões), {samples_total} amostras"
        if res["discrepancies"]:
            line += f" ⚠️ {len(res['discrepancies'])} discrepância(s)"
            warning_count += 1
        summary_lines.append(line)

        disc_map = {d["xlsx"]: (d["proc"], d["exp"]) for d in res["discrepancies"]}
        for p in res["created"]:
            all_excel.append(p)
            name = Path(p).name
            if name in disc_map:
                proc, exp = disc_map[name]
                summary_lines.append(f"   ↳ ⚠️ {name} (processadas: {proc} / declaradas: {exp})")
            else:
                summary_lines.append(f"   ↳ {name}")

    summary_lines.append(f"📊 Total: {len(all_excel)} ficheiro(s) Excel")
    summary_lines.append(f"🧪 Total de amostras: {total_samples_overall}")
    summary_lines.append(f"⏱️ Tempo total: {total_time:.1f} segundos")
    summary_lines.append(f"📅 Executado em: {now_local:%d/%m/%Y às %H:%M:%S}")
    if warning_count:
        summary_lines.append(f"⚠️ {warning_count} ficheiro(s) com discrepâncias")
    if error_count:
        summary_lines.append(f"❌ {error_count} ficheiro(s) com erro (sem ficheiros Excel gerados)")

    summary_text = "\n".join(summary_lines)
    worker_dirs = [Path(os.environ.get("OUTPUT_DIR", tempfile.gettempdir()))]
    debug_files = collect_debug_files(worker_dirs)
    zip_bytes = build_zip_with_summary(all_excel, debug_files, summary_text)
    zip_name = f"xylella_output_{now_local:%Y%m%d_%H%M%S}.zip"

    st.markdown(f"""
    <div style='text-align:center;margin-top:1.2rem;'>
      <h3>🏁 Processamento concluído!</h3>
      <p>Foram gerados <b>{len(all_excel)}</b> ficheiro(s) Excel,
      com um total de <b>{total_samples_overall}</b> amostras processadas.<br>
      Tempo total de execução: <b>{total_time:.1f} segundos</b>.<br>
      Executado em: <b>{now_local:%d/%m/%Y às %H:%M:%S}</b>.</p>
    </div>""", unsafe_allow_html=True)

    zip_b64 = base64.b64encode(zip_bytes).decode()
    col1, col2 = st.columns(2)
    with col1:
        st.markdown(
            f"<a href='data:application/zip;base64,{zip_b64}' download='{zip_name}'>"
            f"<button class='clean-btn' style='width:100%;'>⬇️ Descarregar resultados (ZIP)</button>"
            f"</a>",
            unsafe_allow_html=True
        )
    with col2:
        st.button("🔁 Novo processamento", type="secondary", use_container_width=True, on_click=reset_app)
