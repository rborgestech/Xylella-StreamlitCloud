# -*- coding: utf-8 -*-
import streamlit as st
import tempfile, os, shutil, time
from pathlib import Path
from datetime import datetime
from xylella_processor import process_pdf, build_zip

# ───────────────────────────────────────────────
# Configuração base
# ───────────────────────────────────────────────
st.set_page_config(page_title="Xylella Processor", page_icon="🧪", layout="centered")
st.title("🧪 Xylella Processor")
st.caption("Processa PDFs de requisições Xylella e gera automaticamente 1 Excel por requisição.")

# ───────────────────────────────────────────────
# CSS — tema SGS + animações
# ───────────────────────────────────────────────
st.markdown("""
<style>
.stButton > button[kind="primary"]{background:#CA4300!important;border:1px solid #CA4300!important;color:#fff!important;font-weight:600!important;border-radius:6px!important}
.stButton > button[kind="primary"]:hover{background:#A13700!important;border:1px solid #A13700!important}
[data-testid="stFileUploader"] > div:first-child{border:2px dashed #CA4300!important;border-radius:10px!important;padding:1rem!important}
.success-box{background:#E8F5E9;border-left:5px solid #2E7D32;padding:.7rem 1rem;border-radius:6px;margin-bottom:.4rem}
.warning-box{background:#FFF3E0;border-left:5px solid #F57C00;padding:.7rem 1rem;border-radius:6px;margin-bottom:.4rem}
.info-box{background:#E3F2FD;border-left:5px solid #1E88E5;padding:.7rem 1rem;border-radius:6px;margin-bottom:.4rem}
.st-processing-dots::after{content:' ';animation:dots 1.2s steps(4,end) infinite;color:#CA4300;font-weight:700;margin-left:.15rem}
@keyframes dots{0%,20%{content:''}40%{content:'.'}60%{content:'..'}80%,100%{content:'...'}}
.button-row{display:flex;justify-content:center;align-items:center;gap:1rem;margin-top:1.5rem}
.stDownloadButton button,.stButton button{background:#fff!important;border:1.5px solid #CA4300!important;color:#CA4300!important;font-weight:600!important;border-radius:8px!important;padding:.6rem 1.2rem!important;transition:all .2s}
.stDownloadButton button:hover,.stButton button:hover{background:#CA4300!important;color:#fff!important;border-color:#A13700!important}
.fade-in{animation:fadeIn .8s ease-in-out}
@keyframes fadeIn{from{opacity:0;transform:translateY(10px)}to{opacity:1;transform:translateY(0)}}
</style>
""", unsafe_allow_html=True)

# ───────────────────────────────────────────────
# Estado
# ───────────────────────────────────────────────
if "processing" not in st.session_state: st.session_state.processing = False
if "finished"   not in st.session_state: st.session_state.finished   = False
if "uploads"    not in st.session_state: st.session_state.uploads    = []
if "entries"    not in st.session_state: st.session_state.entries    = []
if "zip_bytes"  not in st.session_state: st.session_state.zip_bytes  = None
if "zip_name"   not in st.session_state: st.session_state.zip_name   = None

# ───────────────────────────────────────────────
# Ecrã inicial
# ───────────────────────────────────────────────
if not st.session_state.processing and not st.session_state.finished:
    uploads = st.file_uploader("📂 Carrega um ou vários PDFs", type=["pdf"], accept_multiple_files=True)
    if uploads:
        st.session_state.uploads = uploads
        if st.button(f"📄 Processar {len(uploads)} ficheiro(s) de Input", type="primary"):
            st.session_state.processing = True
            st.rerun()
    else:
        st.info("💡 Carrega ficheiros PDF para ativar o processamento.")

# ───────────────────────────────────────────────
# Ecrã de processamento
# ───────────────────────────────────────────────
elif st.session_state.processing:
    uploads = st.session_state.uploads
    total = len(uploads)

    st.markdown('<div class="info-box">⏳ A processar ficheiros... aguarde até o processo terminar.</div>', unsafe_allow_html=True)
    with st.expander("📄 Ficheiros em processamento", expanded=True):
        for up in uploads: st.markdown(f"- {up.name}")

    generated_panel = st.expander("📄 Ficheiros gerados", expanded=True)
    progress = st.progress(0)
    status_text = st.empty()
    all_entries = []
    session_dir = tempfile.mkdtemp(prefix="xylella_session_")

    try:
        for i, up in enumerate(uploads, start=1):
            status_text.markdown(
                f'<div class="info-box">📘 <b>A processar ficheiro {i}/{total}</b>'
                f'<span class="st-processing-dots"></span><br>{up.name}</div>',
                unsafe_allow_html=True
            )

            tmpdir = tempfile.mkdtemp(dir=session_dir)
            tmp_path = os.path.join(tmpdir, up.name)
            with open(tmp_path, "wb") as f: f.write(up.getbuffer())

            os.environ["OUTPUT_DIR"] = tmpdir
            entries = process_pdf(tmp_path)  # normalizado para lista de dicts

            if not entries:
                generated_panel.markdown(
                    f'<div class="warning-box">⚠️ Nenhum ficheiro gerado para <b>{up.name}</b>.</div>',
                    unsafe_allow_html=True
                )
            else:
                for e in entries:
                    all_entries.append(e)
                    base = Path(e["path"]).name
                    samples = e.get("samples")
                    disc = e.get("discrepancy")

                    if disc and disc != 0:
                        if isinstance(disc, (tuple, list)) and len(disc) == 2:
                            msg = f"⚠️ <b>{base}</b>: ficheiro gerado. ⚠️ discrepância ({disc[0]} vs {disc[1]})"
                        else:
                            msg = f"⚠️ <b>{base}</b>: ficheiro gerado. ⚠️ discrepância detectada ({disc})"
                        css = "warning-box"
                    else:
                        smp_txt = f"({samples} amostra{'s' if samples != 1 else ''} OK)" if samples else ""
                        msg = f"✅ <b>{base}</b>: ficheiro gerado. {smp_txt}"
                        css = "success-box"

                    generated_panel.markdown(f'<div class="{css}">{msg}</div>', unsafe_allow_html=True)

            progress.progress(i / total)
            time.sleep(0.2)

        # 🧾 Resumo final
        total_samples = sum([(e.get("samples") or 0) for e in all_entries])
        discrep_files = sum([1 for e in all_entries if e.get("discrepancy") not in (None, 0)])
        resumo = f"""
        <div style="background:#FFF; border:1px solid #E5E7EB; border-radius:10px; padding:12px; margin-top:8px;">
          <div style="font-weight:700; margin-bottom:6px;">🧾 RESUMO DE EXECUÇÃO</div>
          <div style="line-height:1.5;">
            {'<br>'.join([
                (f"⚠️ {Path(e['path']).name}: ficheiro gerado. ⚠️ discrepância ({e['discrepancy'][0]} vs {e['discrepancy'][1]})"
                 if isinstance(e['discrepancy'], (tuple, list)) and len(e['discrepancy']) == 2 else
                 f"⚠️ {Path(e['path']).name}: ficheiro gerado. ⚠️ discrepância detectada ({e['discrepancy']})")
                if e.get('discrepancy') not in (None, 0)
                else f"✅ {Path(e['path']).name}: ficheiro gerado. ({e.get('samples') or 0} amostras OK)"
                for e in all_entries
            ])}
          </div>
          <div style="height:10px;"></div>
          <div style="display:flex; gap:18px; align-items:center; flex-wrap:wrap; font-weight:600;">
            <span>🧪 Total de amostras processadas: {int(total_samples)}</span>
            <span>🗂️ Total: {len(all_entries)} ficheiro(s) Excel</span>
            <span>🟡 {discrep_files} ficheiro(s) com discrepâncias</span>
          </div>
        </div>
        """
        generated_panel.markdown(resumo, unsafe_allow_html=True)

        # ZIP final
        if all_entries:
            with st.spinner("🧩 A gerar ficheiro ZIP…"):
                st.session_state.entries = all_entries
                st.session_state.finished = True
                st.session_state.zip_name = f"xylella_output_{datetime.now():%Y%m%d_%H%M%S}.zip"
                st.session_state.zip_bytes = build_zip([e["path"] for e in all_entries])
        else:
            st.warning("⚠️ Nenhum ficheiro Excel foi detetado.")

    except Exception as e:
        st.error(f"❌ Erro inesperado: {e}")
    finally:
        shutil.rmtree(session_dir, ignore_errors=True)
        st.session_state.processing = False

# ───────────────────────────────────────────────
# Ecrã final
# ───────────────────────────────────────────────
if st.session_state.finished and st.session_state.entries:
    total_samples = sum([(e.get("samples") or 0) for e in st.session_state.entries])
    num_files = len(st.session_state.entries)

    st.markdown(f"""
    <div class="fade-in" style="background:#E8F5E9; border-left:6px solid #2E7D32; border-radius:10px;
         padding:1.2rem 1.6rem; margin-top:1.4rem; text-align:center;">
      <h4 style="color:#2E7D32; font-weight:600; margin:.2rem 0 .3rem 0;">✅ Processamento concluído</h4>
      <p style="color:#2E7D32; margin:.2rem 0 0 0;">
        {num_files} ficheiro{'s' if num_files>1 else ''} Excel gerado{'s' if num_files>1 else ''} ·
        <b>{int(total_samples)}</b> amostra{'s' if total_samples!=1 else ''} no total
      </p>
    </div>
    """, unsafe_allow_html=True)

    zip_name = st.session_state.zip_name
    zip_bytes = st.session_state.zip_bytes

    st.markdown('<div class="button-row">', unsafe_allow_html=True)
    c1, c2 = st.columns([1, 1])
    with c1:
        st.download_button("⬇️ Descarregar resultados (ZIP)", data=zip_bytes,
                           file_name=zip_name, mime="application/zip",
                           key="zip_download_final")
    with c2:
        if st.button("🔁 Novo processamento", key="btn_new_run"):
            with st.spinner("🔄 A reiniciar..."):
                for k in ["processing","finished","uploads","entries","zip_bytes","zip_name"]:
                    if k in st.session_state: del st.session_state[k]
                time.sleep(0.5)
                st.rerun()
    st.markdown("</div>", unsafe_allow_html=True)
