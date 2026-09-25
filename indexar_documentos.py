"""
indexar_documentos.py
=====================
Script para indexar manualmente todos los documentos de la carpeta /documentos
en la base vectorial ChromaDB.

Uso:
    python indexar_documentos.py

Requiere:
    - OPENAI_API_KEY configurada en .env o como variable de entorno
    - pip install -r requirements.txt
"""

import os
import sys
from pathlib import Path

# ── Cargar .env si existe ──────────────────────────────────────────────────────
try:
    from dotenv import load_dotenv
    load_dotenv()
    print("✅ Variables de entorno cargadas desde .env")
except ImportError:
    print("⚠️  python-dotenv no instalado. Usando variables de entorno del sistema.")

# ── Verificar API Key ──────────────────────────────────────────────────────────
api_key = os.environ.get("OPENAI_API_KEY", "")
if not api_key:
    print("\n❌ ERROR: No se encontró OPENAI_API_KEY.")
    print("   Creá el archivo .env copiando .env.example y completando tu API key.")
    sys.exit(1)

print(f"✅ OPENAI_API_KEY: {api_key[:8]}...")

# ── Verificar documentos ───────────────────────────────────────────────────────
DOCS_DIR = Path(__file__).resolve().parent / "documentos"
extensiones = {".pdf", ".txt", ".docx", ".doc", ".xlsx", ".xls"}

if not DOCS_DIR.exists():
    print(f"\n❌ La carpeta 'documentos' no existe en {DOCS_DIR}")
    sys.exit(1)

archivos = [f for f in DOCS_DIR.iterdir() if f.suffix.lower() in extensiones]

if not archivos:
    print("\n⚠️  No hay documentos en la carpeta 'documentos'. Agregá PDFs, TXTs o Excel primero.")
    sys.exit(0)

print(f"\n📂 Documentos encontrados ({len(archivos)}):")
for f in archivos:
    print(f"   - {f.name} ({round(f.stat().st_size / 1024, 1)} KB)")

# ── Indexar ────────────────────────────────────────────────────────────────────
print("\n🔄 Iniciando indexación en ChromaDB...\n")

try:
    from bot_rag import indexar_todos_los_documentos, CHROMA_DIR
except ImportError as e:
    print(f"❌ No se pudo importar bot_rag: {e}")
    print("   Asegurate de correr este script desde la carpeta del proyecto.")
    sys.exit(1)

import logging
logging.basicConfig(level=logging.INFO, format="%(levelname)s | %(message)s")

total = indexar_todos_los_documentos()

print("\n" + "=" * 50)
print(f"✅ INDEXACIÓN COMPLETADA")
print(f"   Fragmentos indexados: {total}")
print(f"   Base vectorial en:    {CHROMA_DIR}")
print("=" * 50)
print("\nAhora podés usar el chat de IA en la aplicación web.")
