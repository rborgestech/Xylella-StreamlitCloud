import streamlit as st
import tempfile, os, shutil, time, traceback
from pathlib import Path
from datetime import datetime
from xylella_processor import process_pdf, build_zip

# ───────────────────────────────────────────────
# Configuração base
# ───────────────────────────────────────────────
st.set_page_config(page_title="Xylella Processor", page_icon="🧪", layout="centered")
st.title("🧪 Xylella Processor (Cloud)")
st.caption("Faz upload de um ou vários PDFs. O sistema gera automaticamente um Excel por requisição.")

# ───────────────────────────────────────────────
# CSS — laranja (#CA4300), hover escuro, sem vermelhos
# ───────────────────────────────────────────────
st.markdown("""
<style>
/* Botão principal */
.stButton > button[kind="primary"]{
  background-color:#CA4300!important;border:1px solid #CA4300!important;color:#fff!important;
  font-weight:600!important;border-radius:6px!important;transition:background-color .2s ease-in-out!important;
}
.stButton > button[kind="primary"]:hover,
.stButton > button[kind="primary"]:focus,
.stButton > button[kind="primary"]:active{
  background-color:#A13700!important;border-color:#A13700!important;color:#fff!important;box-shadow:none!important;outline:none!important;
}
/* Disabled */
.stButton > button[kind="primary"][disabled],
.stButton > button[kind="primary"][disabled]:hover{
  background-color:#b3b3b3!important;border:1px solid #b3b3b3!important;color:#f2f2f2!important;cursor:not-allowed!important;box-shadow:none!important;
}
/* File uploader */
[data-testid="stFileUploader"] > div:first-child{
  border:2px dashed #CA4300!important;border-radius:10px!important;padding:1rem!important;
}
[data-testid="stFileUploader"] > div:first-child:hover{ border-color:#A13700!important; }
/* Remover tonalidades vermelhas globais */
:root{
  --primary-color:#CA4300!important;--secondary-color:#CA4300!important;--accent-color:#CA4300!important;
}
</style>
""", unsafe_allow_html=True)

# ───────────────────────────────────────────────
# Estado
# ───────────────────────────────────────────────
if "processing" not in st.session_state:
    st.session_state.processing = False

# ───────────────────────────────────────────────
# UI
# ───────────────────────────────────────────────
uploads = st.file_uploader("📂 Carrega um ou vários PDFs", type=["pdf"], accept_multiple_files=True)
start = st.button("📄 Processar ficheiros de Input", type="primary", disabled=(st.session_state.processing or not uploads))

# ───────────────────────────────────────────────
# Execução
# ───────────────────────────────────────────────
if start and uploads:
    st.session_state.processing = True
    try:
        st.info("⚙️ A processar... aguarda alguns segundos.")
        all_excel = []
        all_stats = []

        # Diretório persistente final
        final_dir = Path.cwd() / "output_final"
        final_dir.mkdir(exist_ok=True)

        progress = st.progress(0.0)
        total = len(uploads)

        for i, up in enumerate(uploads, start=1):
            st.markdown(f"### 🧾 {up.name}")
            st.write("⏳ Início de processamento...")

            tmpdir = tempfile.mkdtemp()
            in_path = os.path.join(tmpdir, up.name)
            with open(in_path, "wb") as f:
                f.write(up.getbuffer())

            # Onde o core vai gravar os .xlsx
            outdir = os.path.join(tmpdir, "out")
            os.makedirs(outdir, exist_ok=True)
            os.environ["OUTPUT_DIR"] = outdir

            # → Core (com contagens corretas)
            files, stats = process_pdf_with_stats(in_path)

            # Copiar para pasta final e registar
            for fp in files:
                if os.path.exists(fp):
                    dest = final_dir / Path(fp).name
                    shutil.copy(fp, dest)
                    all_excel.append(str(dest))

            all_stats.append(stats)

            # Resumo por PDF no ecrã
            st.success(f"✅ {up.name}: {stats['req_count']} requisições, {stats['samples_total']} amostras.")
            for item in stats["per_req"]:
                msg = f" • Requisição {item['req']}: {item['samples']} amostras → {Path(item['file']).name}"
                if item.get("expected") is not None:
                    diff = item['samples'] - (item['expected'] or 0)
                    sign = "+" if diff > 0 else ""
                    if diff != 0:
                        msg += f" ⚠️ discrepância {sign}{diff} (decl={item['expected']})"
                st.write(msg)

            progress.progress(i/total)
            time.sleep(0.2)

        # ZIP + summary
        if all_excel:
            zip_name = f"xylella_output_{datetime.now():%Y%m%d_%H%M%S}.zip"
            zip_bytes = build_zip_with_summary(all_excel, all_stats)
            st.success(f"🏁 Processamento concluído ({len(all_excel)} ficheiros Excel gerados).")
            st.download_button("⬇️ Descarregar resultados (ZIP)",
                               data=zip_bytes, file_name=zip_name, mime="application/zip")

            # Mostrar o mesmo summary no ecrã
            with st.expander("📋 Resumo do processamento"):
                for s in all_stats:
                    st.write(f"**📄 {s['pdf']}** — {s['req_count']} req · {s['samples_total']} amostras")
                    for r in s["per_req"]:
                        line = f"• Req {r['req']}: {r['samples']} amostras → {Path(r['file']).name}"
                        if r.get("expected") is not None:
                            diff = r['samples'] - (r['expected'] or 0)
                            sign = "+" if diff > 0 else ""
                            if diff != 0:
                                line += f"  ⚠️ ({sign}{diff})"
                        st.write(line)
        else:
            st.warning("⚠️ Nenhum ficheiro Excel foi detetado.")

    except Exception as e:
        st.error(f"❌ Erro inesperado: {e}")
        st.code(traceback.format_exc())
    finally:
        st.session_state.processing = False
else:
    st.info("💡 Carrega um ficheiro PDF e clica em **Processar ficheiros de Input**.")
