# -*- coding: utf-8 -*-
"""
xylella_processor.py — versão final com:
✅ criação da pasta debug
✅ ficheiro summary.txt
✅ suporte a múltiplas requisições (req1, req2…)
✅ contagem de amostras (solicitadas/processadas)
"""

import io
import traceback
from pathlib import Path
from zipfile import ZipFile
from datetime import datetime
from core_xylella import process_pdf_sync  # garante que chama o parser real


def process_pdf(pdf_path):
    """
    Processa um PDF e gera um ou mais ficheiros Excel.
    Retorna lista de tuplos (path, solicitadas, processadas)
    """
    try:
        # 📁 Garante pasta debug
        debug_dir = Path(__file__).parent / "debug"
        debug_dir.mkdir(exist_ok=True)

        # 🧩 Executa parser real
        result = process_pdf_sync(pdf_path)
        normalized = []

        if not result:
            print(f"⚠️ Nenhum resultado devolvido para {pdf_path}")
            return []

        # 🧾 Normaliza todos os tipos de resultado
        for item in result:
            if isinstance(item, dict):
                fp = item.get("path")
                solicitadas = item.get("samples_requested") or item.get("samples") or 0
                processadas = item.get("samples_processed") or item.get("processed") or 0
            elif isinstance(item, tuple):
                fp, solicitadas, processadas = item + (0,) * (3 - len(item))
            else:
                fp, solicitadas, processadas = str(item), 0, 0

            normalized.append((str(Path(fp).resolve()), solicitadas, processadas))

        # ✏️ Cria ficheiro summary individual
        summary_path = debug_dir / f"{Path(pdf_path).stem}_summary.txt"
        with open(summary_path, "w", encoding="utf-8") as f:
            f.write(f"🧾 RESUMO DE EXECUÇÃO — {datetime.now():%Y-%m-%d %H:%M:%S}\n")
            f.write(f"PDF origem: {Path(pdf_path).name}\n\n")
            total_amostras = 0
            discrepantes = 0
            for fp, s, p in normalized:
                if s and p:
                    diff = "" if s == p else f" ⚠️ discrepância ({s} vs {p})"
                    if s != p:
                        discrepantes += 1
                    f.write(f"✅ {Path(fp).name}: {p} amostras processadas{diff}\n")
                    total_amostras += p
                else:
                    f.write(f"✅ {Path(fp).name}: ficheiro gerado (0 amostras)\n")

            f.write("\n──────────────────────────────────────────\n")
            f.write(f"📊 Total de ficheiros: {len(normalized)}\n")
            f.write(f"🧪 Total de amostras processadas: {total_amostras}\n")
            f.write(f"⚠️ {discrepantes} ficheiro(s) com discrepâncias\n")
            f.write("──────────────────────────────────────────\n")

        print(f"✅ Summary gravado em {summary_path}")

        return normalized

    except Exception as e:
        print("❌ ERRO no process_pdf:", e)
        traceback.print_exc()
        return []


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


# Teste direto (terminal)
if __name__ == "__main__":
    import sys
    if len(sys.argv) < 2:
        print("Uso: python xylella_processor.py <ficheiro.pdf>")
        sys.exit(0)

    pdf_path = sys.argv[1]
    results = process_pdf(pdf_path)
    print("\n🧾 Resultado final:")
    for fp, s, p in results:
        print(f"  - {fp} ({p} processadas de {s})")
