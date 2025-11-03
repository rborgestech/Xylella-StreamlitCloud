# -*- coding: utf-8 -*-
import os, time
import streamlit as st
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, as_completed
from xylella_processor import process_pdf

st.set_page_config(page_title="🧪 Xylella Processor", layout="centered")

st.title("🧪 Xylella Processor")
st.write("Processa PDFs de requisições Xylella e gera automaticamente 1 ficheiro Excel por requisição.")

MAX_PDF_WORKERS = int(os.getenv("MAX_PDF_WORKERS", "2"))

uploaded_files = st.file_uploader("Carrega ficheiros PDF", type=["pdf"], accept_multiple_files=True)
if uploaded_files and st.button("▶️ Processar ficheiros"):
    start_ts = time.time()
    results = []
    progress = st.progress(0)
    placeholders = []

    # Mostrar placeholders das boxes
    for i, f in enumerate(uploaded_files, start=1):
        ph = st.empty()
        html = f"<div class='file-box {'active' if i == 1 else ''}'><b>📄 {f.name}</b><br><small>em fila...</small></div>"
        ph.markdown(html, unsafe_allow_html=True)
        placeholders.append(ph)

    def _render_state(idx, kind, title, sub):
        placeholders[idx].markdown(
            f"<div class='file-box {kind}'><b>{title}</b><br><small>{sub}</small></div>",
            unsafe_allow_html=True
        )

    # Criar diretório de saída
    outdir = Path.cwd() / "output_final"
    outdir.mkdir(exist_ok=True)

    # Guardar PDFs temporariamente
    tmp_files = []
    for f in uploaded_files:
        p = outdir / f.name
        with open(p, "wb") as fp:
            fp.write(f.getbuffer())
        tmp_files.append(str(p))

    # Processamento paralelo com limite
    with ThreadPoolExecutor(max_workers=min(MAX_PDF_WORKERS, len(tmp_files))) as ex:
        future_map = {ex.submit(process_pdf, p): i for i, p in enumerate(tmp_files)}
        done = 0
        for fut in as_completed(future_map):
            idx = future_map[fut]
            name = uploaded_files[idx].name
            try:
                out_paths = fut.result()
                if not out_paths:
                    _render_state(idx, "error", f"📄 {name}", "❌ Nenhum ficheiro gerado.")
                else:
                    _render_state(idx, "success", f"📄 {name}", f"✅ {len(out_paths)} ficheiro(s) Excel gerado(s).")
            except Exception as e:
                _render_state(idx, "error", f"📄 {name}", f"❌ Erro: {e}")
            done += 1
            progress.progress(done / len(uploaded_files))

    total_time = time.time() - start_ts
    st.success(f"🏁 Processamento concluído em {total_time:.1f}s.")
else:
    st.info("💡 Carrega um ou mais ficheiros PDF para começar.")
