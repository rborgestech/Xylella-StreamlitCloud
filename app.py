# -*- coding: utf-8 -*-
import streamlit as st
import tempfile, os, shutil, time
from pathlib import Path
from datetime import datetime
from xylella_processor import process_pdf, build_zip

st.set_page_config(page_title="Xylella Processor", page_icon="🧪", layout="centered")
st.title("🧪 Xylella Processor")
st.caption("Processa PDFs de requisições Xylella e gera automaticamente 1 Excel por requisição.")

# ————— CSS —————
st.markdown("""
<style>
.success-box{background:#E8F5E9;border-left:5px solid #2E7D32;padding:.7rem 1rem;border-radius:6px;margin:.35rem 0}
.warning-box{background:#FFF8E1;border-left:5px solid #FBC02D;padding:.7rem 1rem;border-radius:6px;margin:.35rem 0}
.info-box{background:#E3F2FD;border-left:5px solid #1E88E5;padding:.7rem 1rem;border-radius:6px;margin:.35rem 0}
.button-row{display:flex;gap:1rem;justify-content:center;margin-top:1rem}
.stDownloadButton button,.stButton button{background:#fff!important;border:1.5px solid #CA4300!important;color:#CA4300!important;font-weight:600!important;border-radius:8px!important;padding:.6rem 1.2rem!important}
.stDownloadButton button:hover,.stButton button:hover{background:#CA4300!important;color:#fff!important}
.st-processing-dots::after{content:' ';animation:dots 1.2s steps(4,end) infinite;color:#CA4300;font-weight:700;margin-left:.15rem}
@keyframes dots{0%,20%{content:''}40%{content:'.'}60%{content:'..'}80%,100%{content:'...'}}
</style>
""", unsafe_allow_html=True)

# ————— Estado —————
if "processing" not in st.session_state: st.session_state.processing = False
if "finished"   not in st.session_state: st.session_state.finished   = False
if "entries"    not in st.session_state: st.session_state.entries    = []
if "zip_bytes"  not in st.session_state: st.session_state.zip_bytes  = None
if "zip_name"   not in st.session_state: st.session_state.zip_name   = None

# ————— Ecrã inicial —————
if not st.session_state.processing and not st.session_state.finished:
    uploads = st.file_uploader("📂 Carrega um ou vários PDFs", type=["pdf"], accept_multiple_files=True)
    if uploads:
        if st.button(f"📄 Processar {len(uploads)} ficheiro(s) de Input"):
            st.session_state.processing = True
            st.session_state._uploads = uploads
            st.rerun()
    else:
        st.info("💡 Carrega ficheiros PDF para ativar o processamento.")

# ————— Helpers —————
def _norm_entry(item):
    """
    Normaliza um resultado de process_pdf para:
    {
        "path": str,
        "processed": int|None,
        "requested": int|None,
        "discrepancy": bool,
        "detail": (requested, processed) | None
    }
    Aceita tanto dicts (com várias chaves possíveis) como strings (apenas path).
    """
    if isinstance(item, str):
        return {"path": item, "processed": None, "requested": None, "discrepancy": False, "detail": None}

    if not isinstance(item, dict):
        return {"path": str(item), "processed": None, "requested": None, "discrepancy": False, "detail": None}

    # possíveis alias vindos do core/wrapper
    path = item.get("path") or item.get("file") or item.get("filepath") or item.get("excel") or ""
    processed = (
        item.get("processed") or item.get("samples") or item.get("amostras") or
        item.get("n_amostras") or item.get("count") or None
    )
    requested = (
        item.get("requested") or item.get("declared") or item.get("esperadas") or
        item.get("expected") or None
    )

    # discrepância pode vir como diff numérico, flag, ou deduzida
    diff = item.get("diff")
    discrepancy_flag = item.get("discrepancy")
    if diff is not None:
        discrepancy = bool(diff)
    else if discrepancy_flag is not None:
        discrepancy = bool(discrepancy_flag)
    else if processed is not None and requested is not None:
        discrepancy = (int(processed) != int(requested))
    else:
        discrepancy = False

    detail = None
    if discrepancy and (requested is not None) and (processed is not None):
        detail = (int(requested), int(processed))

    return {
        "path": str(path),
        "processed": None if processed is None else int(processed),
        "requested": None if requested is None else int(requested),
        "discrepancy": discrepancy,
        "detail": detail
    }

# ————— Processamento —————
else if st.session_state.processing:
    uploads = st.session_state._uploads
    total = len(uploads)

    st.markdown('<div class="info-box">⏳ A processar ficheiros... aguarde até o processo terminar.</div>', unsafe_allow_html=True)
    with st.expander("📄 Ficheiros em processamento", expanded=True):
        for up in uploads:
            st.markdown(f"- {up.name}")

    panel = st.expander("📄 Ficheiros gerados", expanded=True)
    progress = st.progress(0)
    status = st.empty()

    entries = []
    total_proc = 0
    discrep_count = 0

    session_dir = tempfile.mkdtemp(prefix="xylella_session_")
    try:
        for i, up in enumerate(uploads, start=1):
            status.markdown(
                f'<div class="info-box">📘 <b>A processar ficheiro {i}/{total}</b>'
                f'<span class="st-processing-dots"></span><br>{up.name}</div>',
                unsafe_allow_html=True
            )

            tmpdir = tempfile.mkdtemp(dir=session_dir)
            tmp_path = os.path.join(tmpdir, up.name)
            with open(tmp_path, "wb") as f:
                f.write(up.getbuffer())

            os.environ["OUTPUT_DIR"] = tmpdir
            res = process_pdf(tmp_path)

            if not res:
                panel.markdown(
                    f'<div class="warning-box">⚠️ Nenhum ficheiro gerado para <b>{up.name}</b>.</div>',
                    unsafe_allow_html=True
                )
            else:
                for raw in res:
                    e = _norm_entry(raw)

                    # nome base
                    base = Path(e["path"]).name
                    req  = e["requested"]
                    proc = e["processed"]

                    if e["discrepancy"]:
                        # amarelo com solicitadas/processadas
                        if e["detail"]:
                            a, b = e["detail"]
                            msg = f"🟡 <b>{base}</b>: ficheiro gerado. (<b>{a}</b> solicitadas / <b>{b}</b> processadas)"
                        elif req is not None and proc is not None:
                            msg = f"🟡 <b>{base}</b>: ficheiro gerado. (<b>{req}</b> solicitadas / <b>{proc}</b> processadas)"
                        else:
                            msg = f"🟡 <b>{base}</b>: ficheiro gerado. ⚠️ discrepância"
                        css = "warning-box"
                        discrep_count += 1
                    else:
                        if proc is not None:
                            msg = f"✅ <b>{base}</b>: ficheiro gerado. (<b>{proc}</b> amostra{'s' if proc!=1 else ''} OK)"
                        else:
                            msg = f"✅ <b>{base}</b>: ficheiro gerado."
                        css = "success-box"

                    panel.markdown(f'<div class="{css}">{msg}</div>', unsafe_allow_html=True)
                    entries.append(e)
                    if proc is not None:
                        total_proc += proc

            progress.progress(i / total)
            time.sleep(0.15)

        # Resumo dentro do painel
        panel.markdown(
            f'<div class="info-box"><b>📊 Resumo:</b><br>'
            f'🧪 Total de amostras processadas: {total_proc}<br>'
            f'🗂️ Total: {len(entries)} ficheiro(s) Excel<br>'
            f'🟡 {discrep_count} ficheiro(s) com discrepâncias</div>',
            unsafe_allow_html=True
        )

        status.empty()

        # ZIP e fim
        if entries:
            with st.spinner("🧩 A gerar ficheiro ZIP…"):
                zip_bytes = build_zip([e["path"] for e in entries])  # ← só paths!
            st.session_state.entries = entries
            st.session_state.zip_name = f"xylella_output_{datetime.now():%Y%m%d_%H%M%S}.zip"
            st.session_state.zip_bytes = zip_bytes
            st.session_state.processing = False
            st.session_state.finished = True
            st.rerun()
        else:
            st.warning("⚠️ Nenhum ficheiro Excel foi detetado.")

    except Exception as e:
        st.error(f"❌ Erro inesperado: {e}")
    finally:
        shutil.rmtree(session_dir, ignore_errors=True)

# ————— Final —————
else if st.session_state.finished and st.session_state.entries:
    total_proc = sum([(e.get("processed") or 0) for e in st.session_state.entries])
    num_files = len(st.session_state.entries)

    st.markdown(
        f'<div class="success-box" style="text-align:center">'
        f'<b>✅ Processamento concluído</b><br>'
        f'{num_files} ficheiro{"s" if num_files!=1 else ""} Excel gerado{"s" if num_files!=1 else ""} · '
        f'{total_proc} amostra{"s" if total_proc!=1 else ""} no total'
        f'</div>', unsafe_allow_html=True
    )

    st.markdown('<div class="button-row">', unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        st.download_button("⬇️ Descarregar resultados (ZIP)",
                           data=st.session_state.zip_bytes,
                           file_name=st.session_state.zip_name,
                           mime="application/zip",
                           key="zip_dl")
    with c2:
        if st.button("🔁 Novo processamento"):
            for k in list(st.session_state.keys()):
                del st.session_state[k]
            st.rerun()
