# -*- coding: utf-8 -*-
"""
xylella_processor.py — versão estável e compatível com o app Streamlit.
- Garante criação da pasta de debug
- Suporta múltiplas requisições (req1, req2, ...)
- Normaliza o retorno em [(path, n_amostras, discrepancias), ...]
- Mantém compatibilidade com process_pdf_sync() original
"""

import io
import traceback
from pathlib import Path
from zipfile import ZipFile

# ⚠️ Importa o parser verdadeiro (ajusta se necessário)
from core_xylella import process_pdf_sync


# ───────────────────────────────────────────────
# Função principal — processa 1 PDF
# ───────────────────────────────────────────────
def process_pdf(pdf_path):
    """
    Processa um PDF e gera 1 ou mais ficheiros Excel.

    Retorna:
        [(path, n_amostras, discrepancias), ...]
    onde:
        - path → caminho absoluto do Excel gerado
        - n_amostras → nº de amostras encontradas (ou None)
        - discrepancias → nº ou tuplo (esperado, encontrado)
    """
    try:
        # Cria pasta de debug caso não exista
        DEBUG_DIR = Path("debug")
        DEBUG_DIR.mkdir(exist_ok=True)

        # Chama o parser verdadeiro (multi-requisição)
        result = process_pdf_sync(pdf_path)

        if not result:
            print(f"⚠️ Nenhum resultado devolvido para {pdf_path}")
            return []

        normalized = []

        # Caso 1 — lista de paths simples
        if isinstance(result, list) and all(isinstance(x, str) for x in result):
            for fp in result:
                normalized.append((str(Path(fp).resolve()), None, None))

        # Caso 2 — lista de tuplos (path, amostras, discrepancias)
        elif isinstance(result, list) and all(isinstance(x, tuple) for x in result):
            for item in result:
                fp, n_amostras, discrepancias = item + (None,) * (3 - len(item))
                normalized.append((str(Path(fp).resolve()), n_amostras, discrepancias))

        # Caso 3 — lista de dicionários (path, samples, discrepancies)
        elif isinstance(result, list) and all(isinstance(x, dict) for x in result):
            for r in result:
                normalized.append((
                    str(Path(r.get("path")).resolve()),
                    r.get("samples"),
                    r.get("discrepancies"),
                ))

        # Caso 4 — tuplo único com ([paths], info extra)
        elif isinstance(result, tuple) and len(result) >= 1:
            paths = result[0]
            n_amostras = None
            discrepancias = None
            if len(result) >= 3:
                n_amostras = result[1]
                discrepancias = result[2]
            elif len(result) == 2:
                n_amostras = result[1]
            for p in paths:
                normalized.append((str(Path(p).resolve()), n_amostras, discrepancias))

        # Caso 5 — único ficheiro (string isolada)
        elif isinstance(result, str):
            normalized.append((str(Path(result).resolve()), None, None))

        else:
            print(f"⚠️ Formato inesperado retornado por process_pdf_sync: {type(result)} → {result}")

        # Log no diretório de debug
        log_path = DEBUG_DIR / f"{Path(pdf_path).stem}_debug.log"
        with open(log_path, "w", encoding="utf-8") as logf:
            logf.write("🧾 RESULTADO NORMALIZADO:\n")
            for fp, n, d in normalized:
                logf.write(f"{fp} | amostras={n} | discrepâncias={d}\n")

        return normalized

    except Exception as e:
        print("❌ ERRO no process_pdf:", e)
        traceback.print_exc()
        return []


# ───────────────────────────────────────────────
# Função auxiliar — criação do ZIP final
# ───────────────────────────────────────────────
def build_zip(file_paths):
    """
    Cria um ZIP em memória com todos os ficheiros Excel.
    """
    zip_buffer = io.BytesIO()
    with ZipFile(zip_buffer, "w") as zip_file:
        for fp in file_paths:
            try:
                zip_file.write(fp, arcname=Path(fp).name)
            except Exception as e:
                print(f"⚠️ Erro ao adicionar {fp} ao ZIP: {e}")
    zip_buffer.seek(0)
    return zip_buffer.getvalue()


# ───────────────────────────────────────────────
# Execução direta para teste
# ───────────────────────────────────────────────
if __name__ == "__main__":
    import sys
    if len(sys.argv) < 2:
        print("Uso: python xylella_processor.py <ficheiro.pdf>")
        sys.exit(0)

    pdf_path = sys.argv[1]
    results = process_pdf(pdf_path)
    print("\n🧾 Resultado final:")
    for fp, n, d in results:
        print(f"  - {fp} ({n or '?'} amostras, discrepâncias: {d or '0'})")
    print("\n✅ Total ficheiros gerados:", len(results))
