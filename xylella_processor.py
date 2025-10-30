# -*- coding: utf-8 -*-
"""
xylella_processor.py — integração Cloud com core_xylella

Responsável por:
  • Processar um PDF via core_xylella.process_pdf_sync()
  • Gerar ficheiros Excel por requisição
  • Criar um ZIP com todos os ficheiros + log de processamento
"""

import os, io, zipfile, traceback
from datetime import datetime
from pathlib import Path

# Importa as funções principais do core
try:
    import core_xylella as core
except ImportError as e:
    raise ImportError(
        f"❌ Não foi possível importar 'core_xylella'. "
        f"Verifica se o ficheiro está presente. Detalhe: {e}"
    )

# ───────────────────────────────────────────────
#  Função principal: processar 1 PDF
# ───────────────────────────────────────────────

def process_pdf(pdf_path: str) -> list[str]:
    """
    Processa um PDF completo (OCR + parser + geração Excel).
    Devolve a lista de ficheiros gerados (.xlsx).
    """
    print(f"\n🧪 Início de processamento: {os.path.basename(pdf_path)}")

    try:
        # OCR e parsing via core
        rows_per_req = core.process_pdf_sync(pdf_path)
        if not rows_per_req:
            print(f"⚠️ Nenhuma requisição válida em {pdf_path}")
            return []

        created_files = []
        base_name = Path(pdf_path).stem
        output_dir = Path(os.environ.get("OUTPUT_DIR", Path(__file__).parent / "Output"))
        output_dir.mkdir(exist_ok=True)

        # Gerar um ficheiro Excel por requisição
        for i, req_rows in enumerate(rows_per_req, start=1):
            if not req_rows:
                continue

            if len(rows_per_req) == 1:
                out_name = f"{base_name}.xlsx"
            else:
                out_name = f"{base_name}_req{i}.xlsx"

            out_path = output_dir / out_name
            core.write_to_template(
                req_rows,
                out_path,
                expected_count=len(req_rows),
                source_pdf=pdf_path
            )

            print(f"✅ Requisição {i}: {len(req_rows)} amostras gravadas → {out_path}")
            created_files.append(str(out_path))

        print(f"🏁 {pdf_path}: {len(created_files)} ficheiros Excel gerados.\n")
        return created_files

    except Exception as e:
        print(f"❌ Erro a processar {pdf_path}: {e}")
        traceback.print_exc()
        return []


# ───────────────────────────────────────────────
#  Função para criar ZIP com log
# ───────────────────────────────────────────────

def build_zip(file_paths: list[str]) -> bytes:
    """
    Gera um ZIP com todos os ficheiros .xlsx e adiciona um log_processamento.txt
    com o resumo do processamento.
    """
    if not file_paths:
        return b""

    base_dir = os.path.dirname(file_paths[0]) if file_paths else os.getcwd()
    log_path = os.path.join(base_dir, "log_processamento.txt")

    # Construir log detalhado
    with open(log_path, "w", encoding="utf-8") as f:
        f.write(f"📄 Log de Processamento — {datetime.now():%d/%m/%Y %H:%M}\n")
        f.write("──────────────────────────────────────────────\n\n")
        for fp in file_paths:
            if os.path.exists(fp):
                size_kb = os.path.getsize(fp) / 1024
                f.write(f"{os.path.basename(fp)} ({size_kb:.1f} KB)\n")
        f.write("\n✔️ Total de ficheiros: %d\n" % len(file_paths))

    file_paths.append(log_path)

    # Criar ZIP em memória
    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as z:
        for fpath in file_paths:
            if os.path.exists(fpath):
                z.write(fpath, os.path.basename(fpath))
    zip_buf.seek(0)

    print(f"📦 ZIP criado com {len(file_paths)} ficheiros (inclui log_processamento.txt)")
    return zip_buf.getvalue()
