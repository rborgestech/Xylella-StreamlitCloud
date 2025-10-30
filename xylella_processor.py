# -*- coding: utf-8 -*-
"""
xylella_processor.py — integração Cloud com core_xylella

Responsável por:
  • Processar PDFs via core_xylella
  • Gerar ficheiros Excel por requisição
  • Criar ZIP com todos os ficheiros + log
"""

import os, io, zipfile, traceback
from datetime import datetime
from pathlib import Path

# ───────────────────────────────────────────────
# Importar core_xylella
# ───────────────────────────────────────────────
try:
    import core_xylella as core
except ImportError as e:
    raise ImportError(
        f"❌ Não foi possível importar 'core_xylella'. "
        f"Verifica se o ficheiro está presente. Detalhe: {e}"
    )

# ───────────────────────────────────────────────
# Processar PDF + gerar ficheiros
# ───────────────────────────────────────────────
def process_pdf_with_stats(pdf_path: str):
    """
    Processa o PDF via core_xylella e devolve:
    - Lista de ficheiros gerados
    - Estatísticas de requisições e amostras
    """
    stats = {"pdf": os.path.basename(pdf_path), "req_count": 0, "samples_total": 0, "per_req": []}

    print(f"\n🧪 Início de processamento: {os.path.basename(pdf_path)}")
    try:
        rows_per_req = core.process_pdf_sync(pdf_path)
        if not rows_per_req:
            print(f"⚠️ Nenhuma requisição válida em {pdf_path}")
            return [], stats

        created_files = []
        base_name = Path(pdf_path).stem
        output_dir = Path(os.environ.get("OUTPUT_DIR", Path(__file__).parent / "Output"))
        output_dir.mkdir(exist_ok=True)
        stats["req_count"] = len(rows_per_req)

        for i, req_rows in enumerate(rows_per_req, start=1):
            if not req_rows:
                continue

            # nome do ficheiro (sem _req1 se for único)
            out_name = f"{base_name}.xlsx" if len(rows_per_req) == 1 else f"{base_name}_req{i}.xlsx"
            out_path = output_dir / out_name

            expected = len(req_rows)  # nunca None → evita “?”

            core.write_to_template(req_rows, out_path, expected_count=expected, source_pdf=pdf_path)
            created_files.append(str(out_path))
            stats["samples_total"] += len(req_rows)
            stats["per_req"].append({"req": i, "samples": len(req_rows), "file": str(out_path)})

            print(f"✅ Requisição {i}: {len(req_rows)} amostras gravadas → {out_path}")

        print(f"🏁 {pdf_path}: {len(created_files)} ficheiros Excel gerados.\n")
        return created_files, stats

    except Exception as e:
        print(f"❌ Erro a processar {pdf_path}: {e}")
        traceback.print_exc()
        return [], stats


# ───────────────────────────────────────────────
# Criar ZIP com log
# ───────────────────────────────────────────────
def build_zip(file_paths: list[str], log_lines: list[str] | None = None) -> bytes:
    """ZIP com .xlsx + log_processamento.txt"""
    if not file_paths and not log_lines:
        return b""

    base_dir = os.path.dirname(file_paths[0]) if file_paths else os.getcwd()
    log_path = os.path.join(base_dir, "log_processamento.txt")

    with open(log_path, "w", encoding="utf-8") as f:
        f.write(f"📄 Log de Processamento — {datetime.now():%d/%m/%Y %H:%M}\n")
        f.write("──────────────────────────────────────────────\n\n")
        if log_lines:
            for line in log_lines:
                f.write(line.rstrip() + "\n")
            f.write("\n")
        for fp in file_paths:
            if os.path.exists(fp):
                size_kb = os.path.getsize(fp) / 1024
                f.write(f"{os.path.basename(fp)} ({size_kb:.1f} KB)\n")
        f.write("\n✔️ Total de ficheiros: %d\n" % (len(file_paths)))

    all_paths = list(file_paths) + [log_path]

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for p in all_paths:
            if os.path.exists(p):
                z.write(p, os.path.basename(p))
    buf.seek(0)
    print(f"📦 ZIP criado com {len(all_paths)} ficheiros (inclui log_processamento.txt)")
    return buf.getvalue()
