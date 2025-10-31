# -*- coding: utf-8 -*-
"""
xylella_processor.py — camada intermédia entre Streamlit (app.py) e core_xylella.py

Funções expostas:
  • process_pdf(pdf_path) → devolve lista de ficheiros Excel gerados (.xlsx)
  • build_zip(file_paths) → constrói ZIP em memória com .xlsx e logs
"""

import os, io, zipfile, importlib
from pathlib import Path
from typing import List, Dict, Any

# Import dinâmico do core
_CORE_MODULE_NAME = "core_xylella"
core = importlib.import_module(_CORE_MODULE_NAME)

# Diretório de saída
OUTPUT_DIR = Path(os.environ.get("OUTPUT_DIR", Path(__file__).parent / "Output"))
OUTPUT_DIR.mkdir(exist_ok=True)

# ───────────────────────────────────────────────
# Processar PDF
# ───────────────────────────────────────────────
def process_pdf(pdf_path: str) -> List[str]:
    """
    Processa um PDF via core_xylella e devolve a lista de caminhos .xlsx criados.
    Aguenta 3 formatos de retorno do core:
      A) List[List[Dict]]  -> escreve 1 xlsx por req
      B) List[Dict]        -> escreve 1 xlsx
      C) List[str]         -> já são caminhos xlsx -> devolve tal como estão
    """
    print(f"\n📄 A processar: {os.path.basename(pdf_path)}")
    base = os.path.splitext(os.path.basename(pdf_path))[0]

    req_results = core.process_pdf_sync(pdf_path)
    if not req_results:
        print(f"⚠️ Nenhuma requisição extraída de {base}.")
        return []

    # Caso C) já são ficheiros .xlsx (strings)
    if isinstance(req_results, list) and all(isinstance(x, str) for x in req_results):
        created_files = [p for p in req_results if os.path.exists(p)]
        print(f"🟢 Core devolveu {len(created_files)} ficheiros já criados.")
        return created_files

    created_files: List[str] = []

    def _write_one_req(rows: list, req_idx: int, total_reqs: int):
        """Escreve uma requisição (lista de dicts) no template e retorna o caminho."""
        if not rows or not isinstance(rows, list):
            return None
        if not all(isinstance(r, dict) for r in rows):
            # proteção extra: se por algum motivo vierem strings aqui, ignora
            print(f"⚠️ Req {req_idx}: formato inesperado (não é lista de dicts). Ignorado.")
            return None

        # tenta obter expected se vier embutido em cada row (compatibilidade futura)
        expected = None
        try:
            if rows and isinstance(rows[0], dict) and "expected" in rows[0]:
                expected = rows[0].get("expected")
        except Exception:
            expected = None

        out_name = f"{base}_req{req_idx}.xlsx" if total_reqs > 1 else f"{base}.xlsx"
        out_path = core.write_to_template(rows, out_name, expected_count=expected, source_pdf=pdf_path)
        if out_path and os.path.exists(out_path):
            print(f"✅ Requisição {req_idx}: {len(rows)} amostras → {os.path.basename(out_path)}")
            return out_path
        return None

    # Caso B) uma única requisição (lista de dicts)
    if isinstance(req_results, list) and req_results and all(isinstance(x, dict) for x in req_results):
        p = _write_one_req(req_results, 1, 1)
        return [p] if p else []

    # Caso A) várias requisições (lista de listas de dicts)
    if isinstance(req_results, list) and all(isinstance(x, list) for x in req_results):
        total = len(req_results)
        for i, rows in enumerate(req_results, start=1):
            p = _write_one_req(rows, i, total)
            if p:
                created_files.append(p)
        print(f"🏁 {base}: {len(created_files)} ficheiro(s) Excel criados.")
        return created_files

    # Formato desconhecido — não faz nada
    print(f"⚠️ Formato de retorno inesperado de core.process_pdf_sync para {base}.")
    return []



# ───────────────────────────────────────────────
# Gerar ZIP com resultados e logs
# ───────────────────────────────────────────────
def _extract_summary_info(xlsx_path: str):
    """Lê E1 ou F1 e tenta detetar nº de amostras declaradas/processadas (ex: 'Nº Amostras: 10 / 9')."""
    declared, processed = None, None
    try:
        wb = load_workbook(xlsx_path, data_only=False)
        ws = wb.worksheets[0]

        # ler valor literal da célula (sem fórmulas avaliadas)
        for cell_ref in ("E1", "F1"):
            val = str(ws[cell_ref].value or "").strip()
            if not val:
                continue

            # procurar padrões "Nº Amostras: 10 / 9" ou "10/9"
            m = re.search(r"(\d+)\s*/\s*(\d+)", val)
            if m:
                declared = int(m.group(1))
                processed = int(m.group(2))
                break

            # fallback para "Nº Amostras: 10"
            m2 = re.search(r"(\d+)", val)
            if m2 and declared is None:
                declared = int(m2.group(1))
                processed = None

        # se não encontrar nada em E1/F1, tentar ler texto completo das primeiras células
        if declared is None:
            for row in ws.iter_rows(min_row=1, max_row=2, max_col=6, values_only=True):
                row_text = " ".join([str(v) for v in row if v])
                m3 = re.search(r"(\d+)\s*/\s*(\d+)", row_text)
                if m3:
                    declared = int(m3.group(1))
                    processed = int(m3.group(2))
                    break

    except Exception as e:
        print(f"⚠️ Falha ao ler E1/F1 de {os.path.basename(xlsx_path)}: {e}")

    return declared, processed



def build_zip(file_paths: List[str]) -> bytes:
    """
    Constrói um ZIP em memória com:
      • ficheiros Excel processados (.xlsx)
      • pasta 'debug' com ficheiros OCR e logs
      • summary.txt detalhado na raiz
    """
    mem = io.BytesIO()
    summary_lines = []
    summary_lines.append("🧾 RESUMO DE EXECUÇÃO\n")
    summary_lines.append("──────────────────────────────────────────\n")

    total_files = 0
    total_samples = 0
    discrepancias = 0

    for fp in sorted(file_paths):
        if not os.path.exists(fp):
            continue

        total_files += 1
        name = os.path.basename(fp)
        declared, processed = _extract_summary_info(fp)
        line = f"{name}: ficheiro gerado."

        if processed:
            total_samples += processed
            if declared is not None:
                diff = processed - declared
                if diff != 0:
                    line += f" ⚠️ discrepância ({processed} vs {declared})"
                    discrepancias += 1
                else:
                    line += f" ({processed} amostras OK)"
            else:
                line += f" ({processed} amostras)"
        else:
            line += " (sem contagem detectada)"

        summary_lines.append(line)

    # Totais finais
    summary_lines.append("\n──────────────────────────────────────────")
    summary_lines.append(f"📊 Total: {total_files} ficheiro(s) Excel")
    summary_lines.append(f"🧪 Total de amostras processadas: {total_samples}")
    if discrepancias:
        summary_lines.append(f"⚠️ {discrepancias} ficheiro(s) com discrepâncias")
    summary_lines.append("──────────────────────────────────────────\n")

    # ZIP
    with zipfile.ZipFile(mem, "w", compression=zipfile.ZIP_DEFLATED) as z:
        # Excel
        for p in file_paths:
            if p and os.path.exists(p):
                z.write(p, arcname=os.path.basename(p))

        # Pasta debug
        for extra in OUTPUT_DIR.glob("*_ocr_debug.txt"):
            z.write(extra, arcname=f"debug/{os.path.basename(extra)}")
        for logf in OUTPUT_DIR.glob("process_log.csv"):
            z.write(logf, arcname=f"debug/{os.path.basename(logf)}")
        for summ in OUTPUT_DIR.glob("process_summary_*.txt"):
            z.write(summ, arcname=f"debug/{os.path.basename(summ)}")

        # Summary detalhado
        z.writestr("summary.txt", "\n".join(summary_lines))

    mem.seek(0)
    print(f"📦 ZIP criado: {total_files} ficheiros Excel, pasta debug, summary.txt.")
    return mem.read()
