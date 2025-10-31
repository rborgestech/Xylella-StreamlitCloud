# -*- coding: utf-8 -*-
import streamlit as st
import tempfile, os, shutil, time, re, io, zipfile
from pathlib import Path
from datetime import datetime
from openpyxl import load_workbook
from xylella_processor import process_pdf

# ───────────────────────────────────────────────
# Configuração base
# ───────────────────────────────────────────────
st.set_page_config(page_title="Xylella Processor", page_icon="🧪", layout="centered")
st.title("🧪 Xylella Processor")
st.caption("Processa PDFs de requisições Xylella e gera automaticamente 1 Excel por requisição.")

# ───────────────────────────────────────────────
# CSS — laranja #CA4300 e sem vermelhos
# ───────────────────────────────────────────────
st.markdown("""
<style>
.stButton > button[kind="primary"] {
  background-color: #CA4300 !important;
  border: 1px solid #CA4300 !important;
  color: #fff !important;
  font-weight: 600 !important;
  border-radius: 6px !important;
  transition: background-color 0.2s ease-in-out !important;
}
.stButton > button[kind="primary"]:hover,
.stButton > button[kind="primary"]:focus,
.stButton > button[kind="primary"]:active {
  background-color: #A13700 !important;
  border: 1px solid #A13700 !important;
  color: #fff !important;
  box-shadow: none !important;
  outline: none !important;
}
.stButton > button[kind="primary"][disabled],
.stButton > button[kind="primary"][disabled]:hover {
  background-color: #b3b3b3 !important;
  border: 1px solid #b3b3b3 !important;
  color: #f2f2f2 !important;
  cursor: not-allowed !important;
  box-shadow: none !important;
}
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
:root {
  --primary-color: #CA4300 !important;
  --secondary-color: #CA4300 !important;
  --accent-color: #CA4300 !important;
  --text-selection-color: #CA4300 !important;
}
</style>
""", unsafe_allow_html=True)

# ───────────────────────────────────────────────
# Interface de Upload
# ───────────────────────────────────────────────
uploads = st.file_uploader("📂 Carrega um ou vários PDFs", type=["pdf"], accept_multiple_files=True)
if "processing" not in st.session_state:
    st.session_state.processing = False

start = st.button("📄 Processar ficheiros de Input", type="primary",
                  disabled=st.session_state.processing or not uploads)

# ───────────────────────────────────────────────
# Função auxiliar: ler contagens do template (E1)
# ───────────────────────────────────────────────
def read_e1_counts(xlsx_path: str):
    declared, processed = None, None
    try:
        wb = load_workbook(xlsx_path, data_only=False)
        ws = wb.worksheets[0]
        val = str(ws["E1"].value or "")
        m = re.search(r"(\d+)\s*/\s*(\d+)", val)
        if m:
            declared = int(m.group(1))
            processed = int(m.group(2))
    except Exception:
        pass
    return declared, processed

# ───────────────────────────────────────────────
# Execução principal
# ───────────────────────────────────────────────
if start and uploads:
    st.session_state.processing = True
    session_dir = tempfile.mkdtemp(prefix="xylella_session_")

    try:
        st.info("⚙️ A processar ficheiros... aguarda alguns segundos.")
        all_excel, debug_files = [], []
        summary_lines = []

        progress = st.progress(0)
        total = len(uploads)

        for i, up in enumerate(uploads, start=1):
            st.markdown(f"### 📄 {up.name}")
            st.write(f"⏳ A processar ficheiro {i}/{total}...")

            tmpdir = tempfile.mkdtemp(dir=session_dir)
            tmp_path = os.path.join(tmpdir, up.name)
            with open(tmp_path, "wb") as f:
                f.write(up.getbuffer())

            os.environ["OUTPUT_DIR"] = tmpdir
            created = process_pdf(tmp_path)

            if not created:
                st.warning(f"⚠️ Nenhum ficheiro gerado para {up.name}")
            else:
                req_count = len(created)
                total_samples, discrepancies = 0, []
                for fp in created:
                    all_excel.append(fp)
                    declared, processed = read_e1_counts(fp)
                    if declared and processed:
                        total_samples += processed
                        if declared != processed:
                            diff = processed - declared
                            discrepancies.append(f"{Path(fp).name}: Esperado {declared}, Processado {processed} (Δ {diff:+d})")
                    st.success(f"✅ {Path(fp).name} gravado")

                # Mensagem final do ficheiro
                if discrepancies:
                    st.warning(f"✅ {up.name}: {req_count} requisições, {total_samples} amostras (⚠️ discrepâncias: {', '.join(discrepancies)})")
                else:
                    st.success(f"✅ {up.name}: {req_count} requisições, {total_samples} amostras (sem discrepâncias)")

                summary_lines.append(f"{up.name}: {req_count} requisições, {total_samples} amostras.")

            # Ficheiros de debug
            for f in Path(tmpdir).glob("*_ocr_debug.txt"):
                debug_files.append(str(f))
            for logf in Path(tmpdir).glob("process_log.csv"):
                debug_files.append(str(logf))
            for summ in Path(tmpdir).glob("process_summary_*.txt"):
                debug_files.append(str(summ))

            progress.progress(i / total)
            time.sleep(0.2)

        # ───────────────────────────────────────────────
        # ZIP final com debug e summary
        # ───────────────────────────────────────────────
        if all_excel:
            summary_lines.append(f"\n📊 Total: {len(all_excel)} ficheiro(s) Excel gerado(s)")
            summary_text = "\n".join(summary_lines)

            zip_name = f"xylella_output_{datetime.now():%Y%m%d_%H%M%S}.zip"
            mem = io.BytesIO()
            with zipfile.ZipFile(mem, "w", zipfile.ZIP_DEFLATED) as z:
                # Excel
                for f in all_excel:
                    if os.path.exists(f):
                        z.write(f, arcname=os.path.basename(f))
                # Pasta debug
                for dbg in debug_files:
                    if os.path.exists(dbg):
                        z.write(dbg, arcname=f"debug/{os.path.basename(dbg)}")
                # Summary
                z.writestr("summary.txt", summary_text)
            mem.seek(0)

            st.success(f"🏁 Processamento concluído ({len(all_excel)} ficheiros Excel gerados).")
            st.download_button("⬇️ Descarregar resultados (ZIP)",
                               data=mem.read(),
                               file_name=zip_name,
                               mime="application/zip")
            st.balloons()

            # 🔹 Limpar ficheiros carregados automaticamente
            uploads = None
            st.session_state.processing = False

        else:
            st.error("⚠️ Nenhum ficheiro Excel foi detetado para incluir no ZIP.")

    except Exception as e:
        st.error(f"❌ Erro inesperado: {e}")

    finally:
        try:
            shutil.rmtree(session_dir, ignore_errors=True)
        except Exception as e:
            st.warning(f"Não foi possível limpar ficheiros temporários: {e}")
        st.session_state.processing = False

else:
    st.info("💡 Carrega um ficheiro PDF e clica em **Processar ficheiros de Input**.")
