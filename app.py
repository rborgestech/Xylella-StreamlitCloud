# -*- coding: utf-8 -*-
import streamlit as st
import tempfile, os, shutil, time, io, zipfile, re
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
# CSS — estilo laranja limpo
# ───────────────────────────────────────────────
st.markdown("""
<style>
.stButton > button[kind="primary"] {
  background-color: #CA4300 !important;
  border: 1px solid #CA4300 !important;
  color: #fff !important;
  font-weight: 600 !important;
  border-radius: 6px !important;
}
.stButton > button[kind="primary"]:hover {
  background-color: #A13700 !important;
  border-color: #A13700 !important;
}
[data-testid="stFileUploader"] > div:first-child {
  border: 2px dashed #CA4300 !important;
  border-radius: 10px !important;
  padding: 1rem !important;
}
.clean-btn {
  background-color: #fff !important;
  border: 1px solid #ccc !important;
  color: #333 !important;
  font-weight: 600 !important;
  border-radius: 8px !important;
  padding: 0.5rem 1.2rem !important;
  transition: all 0.2s ease-in-out !important;
}
.clean-btn:hover { border-color: #999 !important; color: #000 !important; }
</style>
""", unsafe_allow_html=True)

# ───────────────────────────────────────────────
# Estado
# ───────────────────────────────────────────────
if "processing" not in st.session_state:
    st.session_state.processing = False
if "done" not in st.session_state:
    st.session_state.done = False

# ───────────────────────────────────────────────
# Placeholder principal (garante re-render)
# ───────────────────────────────────────────────
placeholder = st.empty()

# ───────────────────────────────────────────────
# Função principal
# ───────────────────────────────────────────────
def show_uploader():
    """Mostra o uploader de ficheiros, sempre reconstruído após refresh."""
    uploads = placeholder.file_uploader(
        "📂 Carrega um ou vários PDFs",
        type=["pdf"],
        accept_multiple_files=True,
        key=f"upload_{int(time.time())}"
    )

    if uploads:
        start = placeholder.button("📄 Processar ficheiros de Input", type="primary")
        if start:
            placeholder.empty()
            run_processing(uploads)
    else:
        st.info("💡 Carrega um ficheiro PDF e clica em **Processar ficheiros de Input**.")

# ───────────────────────────────────────────────
# Processamento
# ───────────────────────────────────────────────
def run_processing(uploads):
    st.session_state.processing = True
    st.info("⚙️ A processar... isto pode demorar alguns segundos.")
    all_excel = []

    final_dir = Path.cwd() / "output_final"
    final_dir.mkdir(exist_ok=True)
    progress = st.progress(0)
    total = len(uploads)

    start_time = time.time()

    for i, up in enumerate(uploads, start=1):
        st.markdown(f"### 📄 {up.name}")
        st.write(f"⏳ A processar ficheiro {i}/{total}...")
        tmpdir = tempfile.mkdtemp()
        tmp_path = os.path.join(tmpdir, up.name)
        with open(tmp_path, "wb") as f:
            f.write(up.getbuffer())

        os.environ["OUTPUT_DIR"] = tmpdir
        created = process_pdf(tmp_path)

        if not created:
            st.warning(f"⚠️ Nenhum ficheiro gerado para {up.name}")
        else:
            for fp in created:
                dest = final_dir / Path(fp).name
                shutil.copy(fp, dest)
                all_excel.append(str(dest))
                st.success(f"✅ {Path(fp).name} gravado")

        progress.progress(i / total)
        time.sleep(0.2)

    total_time = time.time() - start_time

    if all_excel:
        zip_name = f"xylella_output_{datetime.now():%Y%m%d_%H%M%S}.zip"
        zip_bytes = build_zip(all_excel)
        lisbon_time = datetime.now().strftime("%d/%m/%Y às %H:%M:%S")

        st.markdown(f"""
        <div style='text-align:center;margin-top:1.5rem;'>
            <h3>🏁 Processamento concluído!</h3>
            <p>Foram gerados <b>{len(all_excel)}</b> ficheiro(s) Excel.<br>
            Tempo total de execução: <b>{total_time:.1f} segundos</b>.<br>
            Executado em: <b>{lisbon_time}</b>.</p>
        </div>
        """, unsafe_allow_html=True)

        col1, col2 = st.columns(2)
        with col1:
            st.download_button(
                "⬇️ Descarregar resultados (ZIP)",
                data=zip_bytes,
                file_name=zip_name,
                mime="application/zip",
                use_container_width=True
            )
        with col2:
            if st.button("🔁 Novo processamento", type="secondary", use_container_width=True):
                # Faz refresh completo e reconstrói uploader
                st.markdown("""
                <script>
                setTimeout(function() {
                    window.location.reload(true);
                }, 300);
                </script>
                """, unsafe_allow_html=True)
                st.stop()

        st.session_state.done = True
        st.session_state.processing = False
    else:
        st.error("⚠️ Nenhum ficheiro Excel foi detetado.")
        st.session_state.processing = False

# ───────────────────────────────────────────────
# Execução inicial
# ───────────────────────────────────────────────
if not st.session_state.processing and not st.session_state.done:
    show_uploader()
