import streamlit as st
import tempfile, os, shutil, time, io, zipfile
from pathlib import Path
from datetime import datetime
from xylella_processor import process_pdf_with_stats

# ───────────────────────────────────────────────
# Configuração base
# ───────────────────────────────────────────────
st.set_page_config(page_title="Xylella Processor", page_icon="🧪", layout="centered")
st.title("🧪 Xylella Processor")
st.caption("Processa PDFs de requisições Xylella e gera automaticamente 1 Excel por requisição.")

# ───────────────────────────────────────────────
# CSS — Cores consistentes (laranja) e sem vermelhos
# ───────────────────────────────────────────────
st.markdown("""
<style>
/* Botão principal laranja */
.stButton > button[kind="primary"] {
  background-color: #CA4300 !important;
  border: 1px solid #CA4300 !important;
  color: #fff !important;
  font-weight: 600 !important;
  border-radius: 6px !important;
  transition: background-color 0.2s ease-in-out !important;
}

/* Hover, Focus, Active */
.stButton > button[kind="primary"]:hover,
.stButton > button[kind="primary"]:focus,
.stButton > button[kind="primary"]:active {
  background-color: #A13700 !important;
  border: 1px solid #A13700 !important;
  color: #fff !important;
  box-shadow: none !important;
  outline: none !important;
}

/* Estado desativado (cinzento) */
.stButton > button[kind="primary"][disabled],
.stButton > button[kind="primary"][disabled]:hover {
  background-color: #b3b3b3 !important;
  border: 1px solid #b3b3b3 !important;
  color: #f2f2f2 !important;
  cursor: not-allowed !important;
  box-shadow: none !important;
}

/* File uploader */
[data-testid="stFileUploader"] > div:first-child {
  border: 2px dashed #CA4300 !important;
  border-radius: 10px !important;
  padding: 1rem !important;
  transition: border-color 0.3s ease-in-out;
}

[data-testid="stFileUploader"] > div:first-child:hover {
  border-color: #A13700 !important;
}

[data-testid="stFileUploader"] > div:focus-within {
  border-color: #CA4300 !important;
  box-shadow: none !important;
}

/* Remover qualquer vermelho global */
:root {
  --primary-color: #CA4300 !important;
  --secondary-color: #CA4300 !important;
  --accent-color: #CA4300 !important;
  --text-selection-color: #CA4300 !important;
  --focus-ring: 0 0 0 0 rgba(0,0,0,0) !important;
}
</style>
""", unsafe_allow_html=True)

# ───────────────────────────────────────────────
# Função utilitária — criar ZIP com summary.txt
# ───────────────────────────────────────────────
def build_zip_with_summary(files: list[str], stats: list[dict]) -> bytes:
    """Cria um ZIP com todos os ficheiros Excel e um summary.txt."""
    mem = io.BytesIO()
    with zipfile.ZipFile(mem, "w", zipfile.ZIP_DEFLATED) as z:
        summary_lines = []
        for pdf_stats in stats:
            summary_lines.append(f"{pdf_stats['pdf']}: {pdf_stats['req_count']} requisições, {pdf_stats['samples_total']} amostras")
            for req in pdf_stats["per_req"]:
                line = f"  • Req {req['req']}: {req['samples']} amostras → {Path(req['file']).name}"
                if req.get("diff"):
                    sign = "+" if req['diff'] > 0 else ""
                    line += f" ⚠️ discrepância {sign}{req['diff']} (decl={req['expected']})"
                summary_lines.append(line)
            summary_lines.append("")

        summary = "\n".join(summary_lines)
        z.writestr("summary.txt", summary)

        for f in files:
            if os.path.exists(f):
                z.write(f, arcname=Path(f).name)

    mem.seek(0)
    return mem.getvalue()

# ───────────────────────────────────────────────
# Interface de Upload
# ───────────────────────────────────────────────
uploads = st.file_uploader("📂 Carrega um ou vários PDFs", type=["pdf"], accept_multiple_files=True)

if "processing" not in st.session_state:
    st.session_state.processing = False

start = st.button("📄 Processar ficheiros de Input", type="primary", disabled=st.session_state.processing)

# ───────────────────────────────────────────────
# Execução principal
# ───────────────────────────────────────────────
if start and uploads:
    st.session_state.processing = True
    try:
        st.info("⚙️ A processar... aguarda alguns segundos.")
        all_excel = []
        all_stats = []

        final_dir = Path.cwd() / "output_final"
        final_dir.mkdir(exist_ok=True)

        progress = st.progress(0)
        total = len(uploads)

        for i, up in enumerate(uploads, start=1):
            st.markdown(f"### 📄 {up.name}")
            st.write("⏳ Início de processamento...")

            # Diretório de trabalho persistente
            run_dir = final_dir / f"run_{datetime.now():%Y%m%d_%H%M%S}_{i}"
            run_dir.mkdir(parents=True, exist_ok=True)

            tmp_path = run_dir / up.name
            with open(tmp_path, "wb") as f:
                f.write(up.getbuffer())

            # Definir diretório de saída para o core
            os.environ["OUTPUT_DIR"] = str(run_dir)

            files, stats = process_pdf_with_stats(str(tmp_path))
            all_stats.append(stats)

            # Copiar ficheiros gerados
            for fp in files:
                if os.path.exists(fp):
                    dest = final_dir / Path(fp).name
                    shutil.copy(fp, dest)
                    all_excel.append(str(dest))

            # Mostrar resumo por PDF
            st.write(f"✅ {up.name}: {stats['req_count']} requisições, {stats['samples_total']} amostras.")
            for item in stats["per_req"]:
                msg = f" • Requisição {item['req']}: {item['samples']} amostras → {Path(item['file']).name}"
                if item.get("diff"):
                    sign = "+" if item["diff"] > 0 else ""
                    msg += f" ⚠️ discrepância {sign}{item['diff']} (decl={item['expected']})"
                st.write(msg)

            progress.progress(i / total)
            time.sleep(0.2)

        # 🔹 Criar ZIP final com summary.txt
        valid_files = [str(f) for f in all_excel if os.path.exists(f)]
        if valid_files:
            zip_name = f"xylella_output_{datetime.now():%Y%m%d_%H%M%S}.zip"
            zip_bytes = build_zip_with_summary(valid_files, all_stats)
            st.success(f"🏁 Processamento concluído ({len(valid_files)} ficheiros Excel gerados).")
            st.download_button(
                "⬇️ Descarregar resultados (ZIP)",
                data=zip_bytes,
                file_name=zip_name,
                mime="application/zip"
            )
            st.balloons()
        else:
            st.warning("⚠️ Nenhum ficheiro Excel foi detetado para incluir no ZIP.")

    finally:
        st.session_state.processing = False

else:
    st.info("💡 Carrega um ficheiro PDF e clica em **Processar ficheiros de Input**.")
