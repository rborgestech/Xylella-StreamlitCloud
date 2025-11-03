# ───────────────────────────────────────────────
# API pública usada pela app Streamlit
# ───────────────────────────────────────────────
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import List, Dict, Any
import os
from pathlib import Path
from datetime import datetime

def process_pdf_sync(pdf_path: str) -> List[Dict[str, Any]]:
    """
    Executa o OCR Azure direto ao PDF e o parser Colab integrado, em paralelo por requisição.
    Devolve: lista de dicionários:
        [
            {"rows": [...], "declared": int},
            {"rows": [...], "declared": int},
            ...
        ]
    Cada elemento representa uma requisição.
    """
    base = os.path.basename(pdf_path)
    print(f"\n🧪 Início de processamento: {base}")

    # Diretório de output e ficheiro de debug
    OUTPUT_DIR = Path(os.getenv("OUTPUT_DIR", "/tmp"))
    txt_path = OUTPUT_DIR / f"{os.path.splitext(base)[0]}_ocr_debug.txt"

    # 1️⃣ OCR Azure direto
    result_json = azure_analyze_pdf(pdf_path)

    # 2️⃣ Guardar texto OCR global (debug)
    txt_path.write_text(extract_all_text(result_json), encoding="utf-8")
    print(f"📝 Texto OCR bruto guardado em: {txt_path}")

    # 3️⃣ Dividir em requisições a processar
    requisitions = parse_all_requisitions(result_json, pdf_path, str(txt_path))
    total_reqs = len(requisitions)
    print(f"🔍 {total_reqs} requisição(ões) detetada(s).")

    # 4️⃣ Processamento paralelo de cada requisição
    results: List[Dict[str, Any]] = []
    start_time = datetime.now()

    with ThreadPoolExecutor(max_workers=min(4, total_reqs)) as executor:
        futures = {
            executor.submit(_process_single_req, i, req, base, pdf_path): i
            for i, req in enumerate(requisitions, 1)
        }
        for future in as_completed(futures):
            i = futures[future]
            try:
                result_item = future.result()
                if result_item and result_item.get("rows"):
                    results.append(result_item)
            except Exception as e:
                print(f"❌ Erro na requisição {i}: {e}")

    # 5️⃣ Log de resumo
    total_amostras = sum(len(r["rows"]) for r in results)
    elapsed = (datetime.now() - start_time).total_seconds()
    print(f"✅ {base}: {len(results)} requisições processadas ({total_amostras} amostras) em {elapsed:.1f}s.")
    return results


def _process_single_req(i: int, req: Dict[str, Any], base: str, pdf_path: str) -> Dict[str, Any]:
    """
    Processa uma única requisição (subfunção auxiliar paralela).
    Retorna {"rows": [...], "declared": expected}
    """
    try:
        rows = req.get("rows", [])
        expected = req.get("expected", 0) or 0

        if not rows:
            print(f"⚠️ Requisição {i}: sem amostras — ignorada.")
            return {"rows": [], "declared": expected}

        diff = len(rows) - expected
        if expected and diff != 0:
            print(f"⚠️ Requisição {i}: {len(rows)} processadas vs {expected} declaradas ({diff:+d}).")
        else:
            print(f"✅ Requisição {i}: {len(rows)} amostras processadas (declaradas: {expected}).")

        # devolve estrutura padronizada para o xylella_processor
        return {"rows": rows, "declared": expected}

    except Exception as e:
        print(f"❌ Erro interno na requisição {i}: {e}")
        return {"rows": [], "declared": 0}
