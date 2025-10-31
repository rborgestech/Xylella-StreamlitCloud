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
# Estado
# ───────────────────────────────────────────────
if "processing" not in st.session_state:
    st.session_state.processing = False
if "uploader_key" not in st.session_state:
    st.session_state.uploader_key = 0  # usado para limpar o file_uploader

# ───────────────────────────────────────────────
# UI — Upload
# ───────────────────────────────────────────────
uploads = st.file_uploader(
    "📂 Carrega um ou vários PDFs",
    type=["pdf"],
    accept_multiple_files=True,
    key=f"uploader-{st.session_state.uploader_key}",
)

start = st.button(
    "📄 Processar ficheiros de Input",
    type="primary",
    disabled=st.session_state.processing or not uploads
)

# ───────────────────────────────────────────────
# Helpers
# ───────────────────────────────────────────────
def read_e1_counts(xlsx_path: str):
    """Lê 'Nº Amostras: X / Y' da E1 (declared/processed)."""
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

def collect_debug_files(root_dir: Path) -> list[str]:
    """Apanha logs e txt de debug recursivamente para o ZIP/debug/."""
    debug = []
    patterns = ["*_ocr_debug.txt", "process_log.csv", "process_summary_*.txt"]
    for pat in patterns:
        for f in root_dir.rglob(pat):
            debug.append(str(f))
    return debug

# ───────────────────────────────────────────────
# Execução principal
# ───────────────────────────────────────────────
if start and uploads:
    st.session_state.processing = True
    session_root = Path(tempfile.mkdtemp(prefix="xylella_session_"))

    try:
        top_info = st.info("⚙️ A processar ficheiros... aguarda alguns segundos.")
        all_excel: list[str] = []
        all_debug: list[str] = []
        summary_lines: list[str] = []

        progress = st.progress(0.0)
        total = len(uploads)

        # Validação rápida
        for up in uploads:
            if not up.name.lower().endswith(".pdf"):
                st.error(f"❌ Ficheiro inválido: {up.name} (apenas PDFs são permitidos)")
                st.session_state.processing = False
                st.stop()

        for i, up in enumerate(uploads, start=1):
            st.markdown(f"### 📄 {up.name}")
            step_msg = st.empty()
            step_msg.info(f"⏳ A processar ficheiro {i}/{total}...")

            tmpdir = session_root / f"job_{i:02d}"
            tmpdir.mkdir(parents=True, exist_ok=True)
            tmp_path = tmpdir / up.name
            with open(tmp_path, "wb") as f:
                f.write(up.getbuffer())

            # Isolar saída do core
            os.environ["OUTPUT_DIR"] = str(tmpdir)
            created = process_pdf(str(tmp_path))

            if not created:
                step_msg.warning(f"⚠️ Nenhum ficheiro gerado para {up.name}")
            else:
                req_count = len(created)
                total_samples = 0
                discrepancies_msgs = []

                for fp in created:
                    # contar amostras + discrepâncias
                    declared, processed = read_e1_counts(fp)
                    if processed:
                        total_samples += processed
                    if declared is not None and processed is not None and declared != processed:
                        diff = processed - declared
                        discrepancies_msgs.append(
                            f"{Path(fp).name}: Esperado {declared}, Processado {processed} (Δ {diff:+d})"
                        )

                    all_excel.append(fp)
                    st.success(f"✅ {Path(fp).name} gravado")

                # Mensagem final do ficheiro
                if discrepancies_msgs:
                    step_msg.warning(
                        f"✅ {up.name}: {req_count} requisições, {total_samples} amostras "
                        f"(⚠️ discrepâncias: {', '.join(discrepancies_msgs)})"
                    )
                else:
                    step_msg.success(
                        f"✅ {up.name}: {req_count} requisições, {total_samples} amostras (sem discrepâncias)"
                    )

                summary_lines.append(f"{up.name}: {req_count} requisições, {total_samples} amostras.")

            # recolha de debug (recursiva)
            all_debug.extend(collect_debug_files(tmpdir))

            progress.progress(i / total)
            time.sleep(0.15)

        # ZIP final
        if all_excel:
            summary_lines.append(f"\n📊 Total: {len(all_excel)} ficheiro(s) Excel gerado(s)")
            summary_text = "\n".join(summary_lines)

            zip_name = f"xylella_output_{datetime.now():%Y%m%d_%H%M%S}.zip"
            mem = io.BytesIO()
            with zipfile.ZipFile(mem, "w", zipfile.ZIP_DEFLATED) as z:
                # Excel (raiz)
                for f in all_excel:
                    if os.path.exists(f):
                        z.write(f, arcname=os.path.basename(f))
                # debug/
                for dbg in all_debug:
                    if os.path.exists(dbg):
                        z.write(dbg, arcname=f"debug/{os.path.basename(dbg)}")
                # summary.txt
                z.writestr("summary.txt", summary_text)
            mem.seek(0)

            top_info.success(f"🏁 Processamento concluído ({len(all_excel)} ficheiros Excel gerados).")

            # botão de download — se o utilizador clicar, limpamos uploads
            downloaded = st.download_button(
                "⬇️ Descarregar resultados (ZIP)",
                data=mem.read(),
                file_name=zip_name,
                mime="application/zip",
                type="primary",
                use_container_width=False,
            )

            if downloaded:
                # 🔹 Limpa a seleção do file_uploader e re-renderiza
                st.session_state.uploader_key += 1
                st.session_state.processing = False
                st.success("✅ Concluído. A lista de ficheiros foi limpa.")
                st.experimental_rerun()
        else:
            top_info.error("⚠️ Nenhum ficheiro Excel foi detetado para incluir no ZIP.")

    except Exception as e:
        st.error(f"❌ Erro inesperado: {e}")

    finally:
        st.session_state.processing = False
        # limpa o diretório temporário da sessão
        try:
            shutil.rmtree(session_root, ignore_errors=True)
        except Exception:
            pass

else:
    st.info("💡 Carrega um ficheiro PDF e clica em **Processar ficheiros de Input**.")
