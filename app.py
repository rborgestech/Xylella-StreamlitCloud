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
@keyframes fadeIn{from{opacity:0;transform:translateY(10px)}to{opacity:1;transform:translateY(0)}}
.fade-in{animation:fadeIn .8s ease-in-out}
</style>
""", unsafe_allow_html=True)

# ───────────────────────────────────────────────
# Estado
# ───────────────────────────────────────────────
if "processing" not in st.session_state: st.session_state.processing = False
if "finished"   not in st.session_state: st.session_state.finished   = False
if "uploads"    not in st.session_state: st.session_state.uploads    = []
if "entries"    not in st.session_state: st.session_state.entries    = []  # lista por ficheiro: {path, samples, discrepancy}
if "zip_bytes"  not in st.session_state: st.session_state.zip_bytes  = None
if "zip_name"   not in st.session_state: st.session_state.zip_name   = None

# ───────────────────────────────────────────────
# Helpers
# ───────────────────────────────────────────────
def normalize_result(result):
    """
    Normaliza o retorno do process_pdf para uma lista de entradas por ficheiro:
    Cada entrada = {"path": <str>, "samples": <int|None>, "discrepancy": <None|int|tuple>}
    Compatibiliza vários formatos:
      1) [(path, samples, disc), ...]
      2) [{"path":..., "samples":..., "discrepancy":...}, ...]
      3) ( [paths], samples_map, discrepancy_map )
      4) ( [paths], total_samples, total_discrepancies )  # legacy — atribui None por ficheiro
    """
    entries = []
    # Caso 1/2: lista
    if isinstance(result, list):
        for item in result:
            if isinstance(item, dict):
                entries.append({
                    "path": item.get("path") or item.get("filepath") or item.get("file"),
                    "samples": item.get("samples") or item.get("amostras"),
                    "discrepancy": item.get("discrepancy") or item.get("discrepancias")
                })
            elif isinstance(item, (tuple, list)) and len(item) >= 1:
                p   = item[0]
                smp = item[1] if len(item) > 1 else None
                dsc = item[2] if len(item) > 2 else None
                entries.append({"path": p, "samples": smp, "discrepancy": dsc})
    # Caso 3/4: tuplo
    elif isinstance(result, tuple) and len(result) >= 1:
        paths = result[0] or []
        # dicionários por ficheiro
        if len(result) >= 3 and isinstance(result[1], dict) and isinstance(result[2], dict):
            samples_map = result[1]; disc_map = result[2]
            for p in paths:
                entries.append({
                    "path": p,
                    "samples": samples_map.get(p),
                    "discrepancy": disc_map.get(p)
                })
        else:
            # legacy: números agregados — não conseguimos por ficheiro
            total_samples = result[1] if len(result) > 1 and isinstance(result[1], (int, float)) else None
            # Se só existe 1 ficheiro, atribuímos ao único
            if len(paths) == 1:
                entries.append({"path": paths[0], "samples": total_samples, "discrepancy": None if len(result) < 3 else result[2]})
            else:
                for p in paths:
                    entries.append({"path": p, "samples": None, "discrepancy": None})
    # Limpa paths vazios
    entries = [e for e in entries if e.get("path")]
    return entries

def fmt_samples(n):
    if n is None: return ""
    return f"({n} amostra{'s' if n != 1 else ''} OK)"

def fmt_discrepancy(d):
    if d is None or d == 0: return None
    if isinstance(d, (tuple, list)) and len(d) == 2:
        return f"⚠️ discrepância ({d[0]} vs {d[1]})"
    return f"⚠️ discrepância detectada ({d})"

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
    all_entries = []  # lista de entradas por ficheiro
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
            result = process_pdf(tmp_path)

            entries = normalize_result(result)  # ← NORMALIZAÇÃO POR FICHEIRO
            if not entries:
                generated_panel.markdown(
                    f'<div class="warning-box">⚠️ Nenhum ficheiro gerado para <b>{up.name}</b>.</div>',
                    unsafe_allow_html=True
                )
            else:
                for e in entries:
                    all_entries.append(e)
                    base_name = Path(e["path"]).name
                    dmsg = fmt_discrepancy(e.get("discrepancy"))
                    if dmsg:
                        msg = f"⚠️ <b>{base_name}</b>: ficheiro gerado. <span style='color:#F57C00;'>{dmsg}</span>"
                        css = "warning-box"
                    else:
                        smsg = fmt_samples(e.get("samples"))
                        msg = f"✅ <b>{base_name}</b>: ficheiro gerado. {smsg}"
                        css = "success-box"
                    generated_panel.markdown(f'<div class="{css}">{msg}</div>', unsafe_allow_html=True)

            progress.progress(i / total)
            time.sleep(0.15)

        # Resumo dentro do painel
        status_text.empty()
        total_samples = sum([e["samples"] or 0 for e in all_entries])
        discrep_files = sum([1 for e in all_entries if e.get("discrepancy") not in (None, 0)])
        resumo_html = f"""
<pre style='background:#FAFAFA;border-radius:8px;padding:1rem;font-size:.95rem;border:1px solid #DDD;'>
🧾 <b>RESUMO DE EXECUÇÃO</b>
──────────────────────────────────────────
{chr(10).join([
  (f"⚠️ {Path(e['path']).name}: ficheiro gerado. " + fmt_discrepancy(e.get('discrepancy')))
   if e.get('discrepancy') not in (None,0)
   else (f"✅ {Path(e['path']).name}: ficheiro gerado. {fmt_samples(e.get('samples'))}")
  for e in all_entries
])}
──────────────────────────────────────────
📊 <b>Total:</b> {len(all_entries)} ficheiro(s) Excel
🧪 <b>Total de amostras processadas:</b> {total_samples}
⚠️ <b>{discrep_files}</b> ficheiro(s) com discrepâncias
──────────────────────────────────────────
</pre>
"""
        generated_panel.markdown(resumo_html, unsafe_allow_html=True)

        # Concluir e preparar ZIP
        with st.spinner("🧩 A gerar ficheiro ZIP… aguarde alguns segundos."):
            if all_entries:
                st.session_state.entries   = all_entries
                st.session_state.finished  = True
                st.session_state.zip_name  = f"xylella_output_{datetime.now():%Y%m%d_%H%M%S}.zip"
                st.session_state.zip_bytes = build_zip([e["path"] for e in all_entries])
            else:
                st.warning("⚠️ Nenhum ficheiro Excel foi detetado.")

    except Exception as e:
        st.error(f"❌ Erro inesperado: {e}")
    finally:
        shutil.rmtree(session_dir, ignore_errors=True)
        st.session_state.processing = False

# ───────────────────────────────────────────────
# Ecrã final — painel de sucesso + botões lado a lado
# ───────────────────────────────────────────────
if st.session_state.finished and st.session_state.entries:
    total_samples = sum([(e.get("samples") or 0) for e in st.session_state.entries])
    num_files = len(st.session_state.entries)

    st.markdown(
        f"""
        <div class="fade-in" style="
          background:#E8F5E9; border-left:6px solid #2E7D32; border-radius:10px;
          padding:1.2rem 1.6rem; margin-top:1.4rem; text-align:center;
        ">
          <h4 style="color:#2E7D32; font-weight:600; margin:.2rem 0 .3rem 0;">✅ Processamento concluído</h4>
          <p style="color:#2E7D32; margin:.2rem 0 0 0;">
            {num_files} ficheiro{'s' if num_files>1 else ''} Excel gerado{'s' if num_files>1 else ''} ·
            <b>{total_samples}</b> amostra{'s' if total_samples!=1 else ''} no total
          </p>
        </div>
        """,
        unsafe_allow_html=True
    )

    zip_name  = st.session_state.zip_name
    zip_bytes = st.session_state.zip_bytes

    st.markdown('<div class="button-row">', unsafe_allow_html=True)
    c1, c2 = st.columns([1,1])
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
