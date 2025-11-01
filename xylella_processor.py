# -*- coding: utf-8 -*-
import os
import shutil
from pathlib import Path
from datetime import datetime

try:
    from core_xylella import process_pdf_sync
except ImportError:
    process_pdf_sync = None


def process_pdf(pdf_path):
    """
    Executa o core_xylella.py no contexto real do Streamlit Cloud (/mount/src/xylella-streamlitcloud),
    garantindo a criação de debug/ e summary.
    """
    import subprocess, json, sys
    from pathlib import Path

    # Caminho real onde o core está a correr no Streamlit Cloud
    project_root = Path("/mount/src/xylella-streamlitcloud").resolve()
    pdf_path = Path(pdf_path).resolve()
    pdf_name = pdf_path.name
    stable_pdf = project_root / pdf_name

    # Copiar PDF carregado para o diretório real do projeto
    try:
        shutil.copy(pdf_path, stable_pdf)
    except Exception as e:
        print(f"⚠️ Erro ao copiar PDF: {e}")
        return []

    print(f"📄 Copiado para {stable_pdf}")
    print(f"📂 Working dir forçado: {project_root}")

    # Criar script temporário que chama o core, tal como no teste
    helper = project_root / "_run_core_wrapper.py"
    helper.write_text(f"""
    import json
    from core_xylella import process_pdf_sync
    res = process_pdf_sync(r"{stable_pdf}")
    print(json.dumps(res if isinstance(res, (list, dict)) else str(res)))
    """)
    
        # Executar o core dentro do contexto correto
        result = subprocess.run(
            [sys.executable, str(helper)],
            capture_output=True, text=True, cwd=project_root
        )
    
        if result.returncode != 0:
            print("❌ Erro ao executar core_xylella:")
            print(result.stderr)
            return []
    
        # Logar saída bruta
        print(result.stdout)
    
        # Normalizar resposta
        try:
            parsed = json.loads(result.stdout)
        except Exception:
            parsed = []
    
        entries = []
        if isinstance(parsed, list):
            for r in parsed:
                if isinstance(r, str):
                    entries.append({"path": r, "processed": 0, "discrepancy": False})
                elif isinstance(r, dict):
                    entries.append(r)
                elif isinstance(r, tuple):
                    entries.append({
                        "path": r[0],
                        "processed": r[1] if len(r) > 1 else 0,
                        "discrepancy": bool(r[2]) if len(r) > 2 else False
                    })
        return entries



def _normalize_result(result):
    """Normaliza diferentes formatos devolvidos pelo core."""
    entries = []
    if isinstance(result, list):
        for r in result:
            if isinstance(r, str):
                entries.append({"path": r, "processed": 0, "discrepancy": False})
            elif isinstance(r, dict):
                entries.append(r)
            elif isinstance(r, tuple):
                entries.append({
                    "path": r[0],
                    "processed": r[1] if len(r) > 1 else 0,
                    "discrepancy": bool(r[2]) if len(r) > 2 else False
                })
    elif isinstance(result, tuple):
        files, samples, discrepancies = result
        for i, f in enumerate(files):
            entries.append({
                "path": str(f),
                "processed": samples if isinstance(samples, int) else samples[i] if isinstance(samples, list) else 0,
                "discrepancy": discrepancies if isinstance(discrepancies, bool) else bool(discrepancies[i]) if isinstance(discrepancies, list) else False
            })
    return entries


def build_zip(paths):
    """Gera ZIP a partir de paths (strings ou dicts)."""
    import io, zipfile
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as z:
        for p in paths:
            if isinstance(p, dict):
                p = p.get("path")
            p = Path(p)
            if p.exists():
                z.write(p, arcname=p.name)
    zip_buffer.seek(0)
    return zip_buffer.getvalue()
