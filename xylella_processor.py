# xylella_processor.py
# -*- coding: utf-8 -*-
"""
Xylella Processor – Orquestrador OneDrive
Autor: Rosa Borges
Data: 30/10/2025

Função:
 - Liga-se ao OneDrive/SharePoint via Microsoft Graph
 - Faz download automático dos PDFs da pasta 'Input'
 - Executa o processamento OCR + Parser (core_xylella)
 - Faz upload automático dos Excel gerados para 'Amostras_Processadas'
 - Mostra estado de sincronização em tempo real
"""

import os
import asyncio
from one_drive_service import (
    get_token,
    list_pdfs,
    download_item,
    upload_file,
    INPUT_FOLDER,
    OUTPUT_FOLDER,
    resolve_site_id,
    resolve_drive_id,
    ensure_folder,
)
from core_xylella import process_folder_async

# ───────────────────────────────────────────────────────────────
# CONFIGURAÇÃO LOCAL
# ───────────────────────────────────────────────────────────────
LOCAL_INPUT = "Input"
LOCAL_OUTPUT = "Output"
os.makedirs(LOCAL_INPUT, exist_ok=True)
os.makedirs(LOCAL_OUTPUT, exist_ok=True)

# ───────────────────────────────────────────────────────────────
# PROCESSAMENTO PRINCIPAL
# ───────────────────────────────────────────────────────────────
async def main():
    print("🔐 A autenticar na Microsoft 365...")
    token = get_token()

    print("🌐 A conectar ao OneDrive / SharePoint...")
    try:
        site_id = resolve_site_id(token)
        drive_id = resolve_drive_id(token, site_id)
    except Exception:
        # fallback para OneDrive pessoal (sem site)
        site_id, drive_id = None, None
        print("⚙️  Modo OneDrive pessoal ativo.")

    print(f"📁 A garantir pastas remotas: {INPUT_FOLDER} e {OUTPUT_FOLDER}")
    if drive_id:
        ensure_folder(token, drive_id, INPUT_FOLDER)
        ensure_folder(token, drive_id, OUTPUT_FOLDER)

    print("\n🔄 A sincronizar ficheiros de entrada (Input)...")
    items = list_pdfs(token, drive_id, INPUT_FOLDER)
    if not items:
        print("ℹ️ Nenhum PDF encontrado no OneDrive.")
        return

    for it in items:
        name = it["name"]
        local_path = os.path.join(LOCAL_INPUT, name)
        print(f"⬇️  Download: {name}")
        download_item(token, it, local_path)

    print("\n⚙️  A iniciar processamento local...")
    await process_folder_async(LOCAL_INPUT)

    print("\n⬆️  A enviar resultados para OneDrive...")
    for f in os.listdir(LOCAL_OUTPUT):
        if f.lower().endswith(".xlsx"):
            local_path = os.path.join(LOCAL_OUTPUT, f)
            remote_path = f"{OUTPUT_FOLDER}/{f}"
            print(f"   ↳ {f}")
            upload_file(token, drive_id, local_path, remote_path)

    print("\n🏁 Sincronização concluída com sucesso!")

# ───────────────────────────────────────────────────────────────
# EXECUTAR
# ───────────────────────────────────────────────────────────────
if __name__ == "__main__":
    asyncio.run(main())
