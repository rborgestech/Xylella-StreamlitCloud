# -*- coding: utf-8 -*-
# app.py — versão final estável (Streamlit Cloud)
import streamlit as st
import tempfile, os, shutil, time, io, zipfile, re
from pathlib import Path
from datetime import datetime
from typing import List, Tuple
from openpyxl import load_workbook
from xylella_processor import process_pdf  # usamos só esta; o ZIP é construído aqui

# ───────────────────────────────────────────────
# Configuração base e estilo
# ───────────────────────────────────────────────
st.set_page_config(page_title="Xylella Processor", page_icon="🧪", layout="centered")
st.title("🧪 Xylella Processor")
st.caption("Processa PDFs de requisições Xylella e gera automaticamente 1 ficheiro Excel por requisição.")

# CSS — tons laranja (#CA4300) e sem vermelhos
st.markdown("""
<style>
/* Botão principal */
.stButton > button[kind="primary"] {
  background-color: #CA4300 !important;
  border: 1px solid #CA4300 !important;
  color: #fff !important;
  font-weight: 600 !important;
  border-radius: 6px !important;
  transition: background-color 0.2s ease-in-out !important;
}
/* Hover / Focus / Active */
.stButton > button[kind="primary"]:hover,
.stButton > button[kind="primary"]:focus,
.stButton > button[kind="primary"]:active {
  background-color: #A13700 !important;
  border: 1px solid #A13700 !important;
  color: #fff !important;
  box-shadow: none !important;
  outline: none !important;
}
/* Disabled */
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
[data-testid="stFileUploader"] > div:first-child:hover { border-color: #A13700 !important; }
[data-testid="stFileUploader"] > div:focus-within {
  border-color: #CA4300 !important;
  box-shadow: none !important;
}
/* Cores globais */
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
    st.session_state.uploader_key = "uploader_1"

# ───────────────────────────────────────────────
# Helpers
# ───────────────────────────────────────────────
def read_e1_counts(xlsx_path: str) -> Tuple[int | None, int | None]:
    """
    Lê a célula E1 do template e tenta extrair "Nº Amostras: X / Y".
    Retorna (expected, processed) quando possível.
    """
    try:
        wb = load_workbook(xlsx_path, data_only=True)
        ws = wb.worksheets[0]
        val = str(ws["E1"].value or "")
        m = re.search(r"(\d+)\s*/\s*(\d+)", val)
        if m:
            return int(m.group(1)), int(m.group(2))
    except Exception:
        pass
    return None, None

def collect_debug_files(output_dirs: List[Path]) -> List[str]:
    """
    Recolhe ficheiros de debug nos diretórios de OUTPUT_DIR usados durante o processamento.
    Inclui: *_ocr_debug.txt, process_log.csv, process_summary_*.txt
    """
    debug_files = []
    patterns = ["*_ocr_debug.txt", "process_log.csv", "process_summary_*.txt"]
    for outdir in output_dirs:
        try:
            for pat in patterns:
                for f in outdir.glob(pat):
                    if f.exists():
                        debug_files.append(str(f))
        except Exception:
            continue
    return debug_files

def build_zip_with_summary(excel_files: List[str], debug_files: List[str], summary_text: str) -> bytes:
    """
    Constrói um ZIP em memória com:
      • Todos os .xlsx (na raiz)
      • Pasta debug/ com ficheiros de debug
      • summary.txt na raiz (conteúdo textual passado)
    """
    mem = io.BytesIO()
    with zipfile.ZipFile(mem, "w", zipfile.ZIP_DEFLATED) as z:
        # XLSX na raiz
        for p in excel_files:
            if p and os.path.exists(p):
                z.write(p, arcname=os.path.basename(p))
        # Debug em /debug
        for d in debug_files:
            if d and os.path.exists(d):
                z.write(d, arcname=f"debug/{os.path.basename(d)}")
        # summary.txt
        z.writestr("summary.txt", summary_text or "")
    mem.seek(0)
    return mem.read()

# ───────────────────────────────────────────────
# Interface de Upload (oculta durante processamento)
# ───────────────────────────────────────────────
if not st.session_state.processing:
    uploads = st.file_uploader(
        "📂 Carrega um ou vários PDFs",
        type=["pdf"],
        accept_multiple_files=True,
        key=st.session_state.uploader_key
    )
    start_disabled = not uploads or st.session_state.processing
    start = st.button("📄 Processar ficheiros de Input", type="primary", disabled=start_disabled)
else:
    uploads = None
    start = False
    st.info("⚙️ A processar...")

# ───────────────────────────────────────────────
# Execução principal
# ───────────────────────────────────────────────
if start and uploads:
    st.session_state.processing = True

    # Diretório de trabalho da sessão (para isolar temporários)
    session_dir = tempfile.mkdtemp(prefix="xylella_session_")
    final_dir = Path.cwd() / "output_final"
    final_dir.mkdir(exist_ok=True)

    try:
        st.info("⚙️ A processar... isto pode demorar alguns segundos.")
        all_excel: List[str] = []
        outdirs_used: List[Path] = []
        summary_lines: List[str] = []

        # Validações rápidas
        for up in uploads:
            if not up.name.lower().endswith(".pdf"):
                st.error(f"❌ Ficheiro inválido: {up.name} (apenas PDFs).")
                st.session_state.processing = False
                st.stop()
            if up.size and up.size > 20 * 1024 * 1024:
                st.error(f"⚠️ {up.name} excede 20 MB.")
                st.session_state.processing = False
                st.stop()

        progress = st.progress(0.0)
        total = len(uploads)

        # Processamento sequencial (UI estável)
        for i, up in enumerate(uploads, start=1):
            st.markdown(f"### 📄 {up.name}")
            st.write("⏳ Início de processamento...")

            tmpdir = Path(tempfile.mkdtemp(dir=session_dir))
            tmp_pdf = tmpdir / up.name
            with open(tmp_pdf, "wb") as f:
                f.write(up.getbuffer())

            # OUTPUT_DIR isolado por ficheiro
            os.environ["OUTPUT_DIR"] = str(tmpdir)
            outdirs_used.append(tmpdir)

            # Processar → devolve lista de ficheiros .xlsx criados
            created = process_pdf(str(tmp_pdf))

            if not created:
                st.warning(f"⚠️ Nenhum ficheiro gerado para {up.name}")
                summary_lines.append(f"{up.name}: 0 requisições, 0 amostras (sem ficheiros).")
            else:
                # Copiar .xlsx para output_final + extrair contagens E1 e discrepâncias
                req_count = len(created)
                samples_total = 0
                per_req_msgs = []

                for idx, fp in enumerate(created, start=1):
                    dest = final_dir / Path(fp).name
                    shutil.copy(fp, dest)
                    all_excel.append(str(dest))

                    exp, proc = read_e1_counts(str(dest))
                    if exp is None and proc is None:
                        # Sem contagem detectada
                        per_req_msgs.append(f"• Requisição {idx}: {Path(dest).name} (sem contagem detectada)")
                    else:
                        samples_total += (proc or 0)
                        diff = (proc - exp) if (exp is not None and proc is not None) else None
                        if diff:
                            sign = "+" if diff > 0 else ""
                            per_req_msgs.append(
                                f"• Requisição {idx}: {proc} amostras → {Path(dest).name} "
                                f"⚠️ discrepância {sign}{diff} ({proc} processadas / {exp} declaradas)"
                            )
                        else:
                            per_req_msgs.append(
                                f"• Requisição {idx}: {proc} amostras → {Path(dest).name}"
                            )

                    # Mensagem “gravado” junto do detalhe da amostra
                    st.success(per_req_msgs[-1])

                # Cabeçalho com totals deste PDF
                st.success(f"✅ {up.name}: {req_count} requisição(ões), {samples_total} amostras.")
                for line in per_req_msgs:
                    # já mostrado acima; mantemos para garantir visibilidade
                    pass

                summary_lines.append(f"{up.name}: {req_count} requisição(ões), {samples_total} amostras.")

            # Atualizar progress bar
            progress.progress(i / total)
            time.sleep(0.2)

        # Construção do ZIP (com debug/ e summary.txt)
        if all_excel:
            # Recolher debug dos OUTPUT_DIR usados
            debug_files = collect_debug_files(outdirs_used)

            # Texto do summary.txt
            summary_text = "\n".join(summary_lines)
            summary_text += f"\n\n📊 Total: {len(all_excel)} ficheiro(s) Excel"

            zip_bytes = build_zip_with_summary(all_excel, debug_files, summary_text)
            zip_name = f"xylella_output_{datetime.now():%Y%m%d_%H%M%S}.zip"

            st.success(f"🏁 Processamento concluído ({len(all_excel)} ficheiros Excel gerados).")
            st.download_button("⬇️ Descarregar resultados (ZIP)", data=zip_bytes,
                               file_name=zip_name, mime="application/zip")

            # Auto-limpar o uploader (gera nova key) SEM rerun
            st.session_state.uploader_key = f"uploader_{datetime.now().timestamp()}"
        else:
            st.error("⚠️ Nenhum ficheiro Excel foi detetado para incluir no ZIP.")

    except Exception as e:
        st.error(f"❌ Erro inesperado: {e}")

    finally:
        # Limpeza de temporários
        try:
            shutil.rmtree(session_dir, ignore_errors=True)
        except Exception as e:
            st.warning(f"Não foi possível limpar ficheiros temporários: {e}")

        st.session_state.processing = False

# Sugestão quando inativo
if not st.session_state.processing:
    st.info("💡 Carrega um ficheiro PDF e clica em **Processar ficheiros de Input**.")
