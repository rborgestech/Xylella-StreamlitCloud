# -*- coding: utf-8 -*-
"""
xylella_processor.py — integração Cloud com core_xylella

Versão final:
  ✅ Gera ficheiros Excel por requisição
  ✅ Cria ZIP com TODOS os .xlsx + log detalhado
  ✅ Inclui discrepâncias e totais globais no log
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
    raise ImportError(f"❌ Erro a importar 'core_xylella': {e}")

# ───────────────────────────────────────────────
# Processar PDF + gerar ficheiros
# ───────────────────────────────────────────────
def process_pdf_with_stats(pdf_path: str):
    """
    Processa o PDF via core_xylella e devolve:
      - Lista de ficheiros gerados
      - Estatísticas detalhadas de requisições e amostras
    """
    import re

    stats = {"pdf": os.path.basename(pdf_path), "req_count": 0, "samples_total": 0, "per_req": []}
    print(f"\n🧪 Início de processamento: {os.path.basename(pdf_path)}")

    try:
        rows_per_req = core.process_pdf_sync(pdf_path)
        if not rows_per_req:
            print(f"⚠️ Nenhuma requisição válida em {pdf_path}")
            return [], stats

        created_files = []
        base_name = Path(pdf_path).stem

        # 🔒 Nome seguro (sem espaços nem acentos)
        safe_base_name = re.sub(r'[^\w\-_.]', '_', base_name)

        output_dir = Path(os.environ.get("OUTPUT_DIR", Path(__file__).parent / "Output"))
        output_dir.mkdir(exist_ok=True)

        stats["req_count"] = len(rows_per_req)

        for i, req_rows in enumerate(rows_per_req, start=1):
            if req_rows is None:
                print(f"⚠️ Requisição {i} ignorada (None)")
                continue

            out_name = f"{safe_base_name}.xlsx" if len(rows_per_req) == 1 else f"{safe_base_name}_req{i}.xlsx"
            out_path = output_dir / out_name

            expected = getattr(req_rows, "expected_count", None)
            if not expected:
                expected = len(req_rows)

            core.write_to_template(req_rows, out_path, expected_count=expected, source_pdf=pdf_path)

            if not out_path.exists():
                print(f"❌ Falha ao gravar: {out_path}")
                continue

            created_files.append(str(out_path))
            stats["samples_total"] += len(req_rows)

            discrepancy = None
            if expected != len(req_rows):
                discrepancy = expected - len(req_rows)

            stats["per_req"].append({
                "req": i,
                "samples": len(req_rows),
                "file": str(out_path),
                "expected": expected,
                "diff": discrepancy
            })

            print(f"✅ Requisição {i}: {len(req_rows)} amostras gravadas → {out_path}")

        print(f"🏁 {pdf_path}: {len(created_files)} ficheiros Excel gerados.\n")
        return created_files, stats

    except Exception as e:
        print(f"❌ Erro a processar {pdf_path}: {e}")
        traceback.print_exc()
        return [], stats


# ───────────────────────────────────────────────
# Criar ZIP com log detalhado
# ───────────────────────────────────────────────
def build_zip(file_paths: list[str], all_stats: list[dict]) -> bytes:
    """ZIP com .xlsx + log_processamento.txt detalhado."""
    if not file_paths:
        return b""

    # Criar log no diretório atual
    base_dir = Path.cwd()
    log_path = base_dir / "log_processamento.txt"

    # Cálculo de totais globais
    total_pdfs = len(all_stats)
    total_reqs = sum(s["req_count"] for s in all_stats)
    total_samples = sum(s["samples_total"] for s in all_stats)

    with open(log_path, "w", encoding="utf-8") as f:
        f.write(f"📄 Log de Processamento — {datetime.now():%d/%m/%Y %H:%M}\n")
        f.write("──────────────────────────────────────────────\n\n")

        for s in all_stats:
            f.write(f"📘 {s['pdf']}\n")
            f.write(f"   → {s['req_count']} requisições, {s['samples_total']} amostras.\n")

            for r in s["per_req"]:
                line = f"      Req {r['req']}: {r['samples']} amostras → {Path(r['file']).name}"
                if r["diff"]:
                    sign = "+" if r["diff"] > 0 else ""
                    line += f" ⚠️ discrepância {sign}{r['diff']} (decl={r['expected']})"
                f.write(line + "\n")
            f.write("\n")

        f.write("──────────────────────────────────────────────\n")
        f.write(f"📊 Total global: {total_pdfs} PDFs, {total_reqs} requisições, {total_samples} amostras.\n")

    # Criar ZIP com todos os ficheiros
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        # adicionar todos os ficheiros Excel
        for fp in file_paths:
            if os.path.exists(fp):
                z.write(fp, os.path.basename(fp))
        # adicionar o log no final
        z.write(log_path, os.path.basename(log_path))

    buf.seek(0)
    print(f"📦 ZIP criado com {len(file_paths)} ficheiros Excel + log_processamento.txt")
    return buf.getvalue()
