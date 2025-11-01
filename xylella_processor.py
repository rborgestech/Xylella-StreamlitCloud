# -*- coding: utf-8 -*-
"""
xylella_processor.py — wrapper robusto para o teu parser:
✅ NÃO mexe no teu process_pdf_sync
✅ Aceita qualquer formato de retorno
✅ Normaliza para lista de dicionários:
   [{"path": "...xlsx", "requested": int|None, "processed": int|None, "discrepancy": bool, "detail": (req, proc)|None}, ...]
✅ Cria pasta debug/ ao lado deste ficheiro e grava <pdf>_summary.txt
"""

from pathlib import Path
from datetime import datetime
import io, traceback
from zipfile import ZipFile

# IMPORTA o teu motor original (não alterado)
from core_xylella import process_pdf_sync


def _to_abs(p):
    return str(Path(p).resolve()) if p else None


def _norm_result(result):
    """
    Converte QUALQUER retorno do parser numa lista de dicts com:
      path, requested, processed, discrepancy, detail
    """
    entries = []

    if not result:
        return entries

    # Caso: lista homogénea
    if isinstance(result, list):
        for item in result:
            # dicionário
            if isinstance(item, dict):
                p = item.get("path") or item.get("filepath") or item.get("file")
                req = (
                    item.get("samples_requested")
                    or item.get("requested")
                    or item.get("solicitadas")
                    or None
                )
                proc = (
                    item.get("samples_processed")
                    or item.get("processed")
                    or item.get("amostras")
                    or item.get("samples")
                    or None
                )
                dsc = item.get("discrepancy") or item.get("discrepancias") or None
                # se dsc vier como (req,proc) usa; se vier bool/num, calcula pelo req/proc
                detail = None
                if isinstance(dsc, (tuple, list)) and len(dsc) == 2:
                    detail = (dsc[0], dsc[1])
                    dflag = (dsc[0] is not None and dsc[1] is not None and dsc[0] != dsc[1])
                else:
                    dflag = bool(dsc) if dsc not in (None, 0) else False
                    if req is not None and proc is not None and req != proc:
                        dflag = True
                        detail = (req, proc)
                entries.append({
                    "path": _to_abs(p),
                    "requested": req,
                    "processed": proc,
                    "discrepancy": dflag,
                    "detail": detail
                })

            # tuplo/lista (path, req?, proc? | path, processed?, discrepancy?)
            elif isinstance(item, (tuple, list)):
                p = item[0] if len(item) > 0 else None
                # heurística: se 3 elementos e o último for tuplo (req,proc), usa
                req = proc = None
                dflag = False
                detail = None
                if len(item) >= 3 and isinstance(item[2], (tuple, list)) and len(item[2]) == 2:
                    req, proc = item[2]
                    dflag = (req is not None and proc is not None and req != proc)
                    detail = (req, proc)
                else:
                    # tentar interpretar 2º e 3º elemento como req/proc
                    if len(item) >= 2:
                        # se só houver um número, assumimos processed
                        if isinstance(item[1], (int, float)) and (len(item) == 2 or not isinstance(item[2], (int, float))):
                            proc = int(item[1])
                        elif len(item) >= 3 and isinstance(item[1], (int, float)) and isinstance(item[2], (int, float)):
                            req = int(item[1]); proc = int(item[2])
                            dflag = (req != proc)
                            detail = (req, proc)
                    # se último for bool/num, usar como discrepância
                    if len(item) >= 3 and isinstance(item[2], (int, float, bool)) and detail is None:
                        dflag = bool(item[2])

                entries.append({
                    "path": _to_abs(p),
                    "requested": req,
                    "processed": proc,
                    "discrepancy": dflag,
                    "detail": detail
                })

            # string (apenas path)
            elif isinstance(item, str):
                entries.append({
                    "path": _to_abs(item),
                    "requested": None,
                    "processed": None,
                    "discrepancy": False,
                    "detail": None
                })

    # Caso: tuplo tipo ([paths], mapas/valores)
    elif isinstance(result, tuple) and len(result) >= 1:
        paths = result[0] or []
        samples_map = None
        proc_map = None
        disc_map = None

        # tentar desvendar segundo/terceiro elementos
        for extra in result[1:]:
            if isinstance(extra, dict):
                # heurística: se valores forem tuplos (req,proc)
                if all(isinstance(v, (tuple, list)) and len(v) == 2 for v in extra.values()):
                    proc_map = {k: v[1] for k, v in extra.items()}
                    samples_map = {k: v[0] for k, v in extra.items()}
                else:
                    # se os valores parecerem bool/num, assumir discrepâncias
                    if all(isinstance(v, (int, float, bool)) for v in extra.values()):
                        disc_map = extra
                    else:
                        # fallback: assumir processed
                        proc_map = extra
            elif isinstance(extra, (int, float, bool)):
                # valor agregado ignorado a nível de entrada individual
                pass

        for p in paths:
            req = samples_map.get(p) if samples_map else None
            proc = proc_map.get(p) if proc_map else None
            dflag = False
            detail = None
            if req is not None and proc is not None:
                dflag = (req != proc)
                detail = (req, proc)
            if disc_map and not dflag:
                dflag = bool(disc_map.get(p))
            entries.append({
                "path": _to_abs(p),
                "requested": req,
                "processed": proc,
                "discrepancy": dflag,
                "detail": detail
            })

    # filtra entradas sem path
    entries = [e for e in entries if e["path"]]
    return entries


def process_pdf(pdf_path):
    """
    Chama o teu parser e normaliza o output.
    Também cria debug/<nome>_summary.txt com o resumo desta execução.
    """
    debug_dir = Path(__file__).resolve().parent / "debug"
    debug_dir.mkdir(parents=True, exist_ok=True)

    try:
        raw = process_pdf_sync(pdf_path)
        entries = _norm_result(raw)

        # escrever summary
        summary_path = debug_dir / f"{Path(pdf_path).stem}_summary.txt"
        total_proc = sum([e["processed"] or 0 for e in entries])
        discrep = sum([1 for e in entries if e["discrepancy"]])
        with open(summary_path, "w", encoding="utf-8") as f:
            f.write(f"🧾 RESUMO — {datetime.now():%Y-%m-%d %H:%M:%S}\n")
            f.write(f"PDF origem: {Path(pdf_path).name}\n\n")
            for e in entries:
                base = Path(e["path"]).name
                req = e["requested"]; proc = e["processed"]
                if e["discrepancy"]:
                    if e["detail"]:
                        f.write(f"⚠️ {base}: discrepância ({e['detail'][0]} vs {e['detail'][1]})\n")
                    else:
                        f.write(f"⚠️ {base}: discrepância\n")
                else:
                    if proc is not None:
                        f.write(f"✅ {base}: {proc} amostras OK\n")
                    else:
                        f.write(f"✅ {base}: ficheiro gerado\n")
            f.write("\n")
            f.write(f"🗂️ Total ficheiros: {len(entries)}\n")
            f.write(f"🧪 Total amostras processadas: {total_proc}\n")
            f.write(f"🟡 Ficheiros com discrepâncias: {discrep}\n")

        return entries

    except Exception as e:
        # em caso de erro, deixa rasto no debug
        err_path = debug_dir / f"{Path(pdf_path).stem}_error.txt"
        with open(err_path, "w", encoding="utf-8") as f:
            f.write(f"{type(e).__name__}: {e}\n")
            f.write(traceback.format_exc())
        return []


def build_zip(paths_or_entries):
    """
    Aceita lista de strings (paths) OU lista de dict entries.
    """
    zip_buffer = io.BytesIO()
    with ZipFile(zip_buffer, "w") as zf:
        for it in paths_or_entries:
            p = it["path"] if isinstance(it, dict) else it
            try:
                zf.write(p, arcname=Path(p).name)
            except Exception:
                # ignora faltas pontuais
                pass
    zip_buffer.seek(0)
    return zip_buffer.getvalue()
