"""
bot_rag.py
==========
Motor de inteligencia artificial (RAG) para el bot de WhatsApp y analizador de fojas quirúrgicas.

Formatos soportados para análisis:
  - Imagen directa: JPG, PNG, WEBP → GPT-4o Vision
  - PDF con texto extraíble         → GPT-4o + RAG (nomencladores)
  - PDF escaneado (sin texto)       → convierte a imagen → GPT-4o Vision + RAG

DEPENDENCIAS: pip install langchain langchain-openai langchain-community chromadb PyMuPDF tiktoken openai
"""

import os
import logging
from pathlib import Path

# ============================================================
# CONFIGURACIÓN
# ============================================================

BASE_DIR        = Path(__file__).resolve().parent
DOCS_DIR        = BASE_DIR / "documentos"
CHROMA_DIR      = BASE_DIR / "chroma_db"

DOCS_DIR.mkdir(exist_ok=True)
CHROMA_DIR.mkdir(exist_ok=True)

CHUNK_SIZE      = 1000
CHUNK_OVERLAP   = 150
MODEL_NAME      = "gpt-4o-mini"        # Para respuestas de texto (más económico)
MODEL_VISION    = "gpt-4o"             # Para análisis de imágenes (visión completa)
EMBEDDING_MODEL = "text-embedding-3-small"

EXTENSIONES_IMAGEN = {".jpg", ".jpeg", ".png", ".webp", ".gif"}
MIME_TIPOS = {
    ".jpg":  "image/jpeg",
    ".jpeg": "image/jpeg",
    ".png":  "image/png",
    ".webp": "image/webp",
    ".gif":  "image/gif",
}

logger = logging.getLogger(__name__)


# ============================================================
# CARGA DE DOCUMENTOS (para indexar en ChromaDB)
# ============================================================

def _cargar_pdf(ruta: Path) -> list:
    from langchain_community.document_loaders import PyMuPDFLoader
    return PyMuPDFLoader(str(ruta)).load()


def _cargar_txt(ruta: Path) -> list:
    from langchain_community.document_loaders import TextLoader
    return TextLoader(str(ruta), encoding="utf-8").load()


def _cargar_docx(ruta: Path) -> list:
    from langchain_community.document_loaders import UnstructuredWordDocumentLoader
    return UnstructuredWordDocumentLoader(str(ruta)).load()


def _cargar_excel(ruta: Path) -> list:
    import openpyxl
    from langchain.schema import Document
    documentos = []
    # read_only=True es CLAVE para no saturar la memoria RAM
    wb = openpyxl.load_workbook(ruta, data_only=True, read_only=True)
    try:
        for nombre_hoja in wb.sheetnames:
            ws = wb[nombre_hoja]
            filas_texto = []
            for fila in ws.iter_rows(values_only=True):
                partes = [str(c) for c in fila if c is not None]
                if partes:
                    filas_texto.append(" | ".join(partes))
            if filas_texto:
                contenido = "=== Hoja: {} ===\n{}".format(nombre_hoja, "\n".join(filas_texto))
                documentos.append(Document(
                    page_content=contenido,
                    metadata={"source": str(ruta), "sheet": nombre_hoja},
                ))
    finally:
        wb.close()
    return documentos


def cargar_documento(ruta: Path) -> list:
    ext = ruta.suffix.lower()
    if ext == ".pdf":
        return _cargar_pdf(ruta)
    elif ext == ".txt":
        return _cargar_txt(ruta)
    elif ext in (".docx", ".doc"):
        return _cargar_docx(ruta)
    elif ext in (".xlsx", ".xls"):
        return _cargar_excel(ruta)
    else:
        logger.warning("Formato no soportado para indexar: %s", ruta.name)
        return []


# ============================================================
# BASE VECTORIAL (ChromaDB)
# ============================================================

def _obtener_embeddings():
    from langchain_openai import OpenAIEmbeddings
    api_key = os.environ.get("OPENAI_API_KEY")
    if not api_key:
        raise RuntimeError("Falta la variable de entorno OPENAI_API_KEY.")
    return OpenAIEmbeddings(model=EMBEDDING_MODEL, openai_api_key=api_key)


def _obtener_vectorstore():
    from langchain_community.vectorstores import Chroma
    return Chroma(
        persist_directory=str(CHROMA_DIR),
        embedding_function=_obtener_embeddings(),
        collection_name="ssspadron_docs",
    )


def agregar_documento_a_base(ruta: Path) -> int:
    from langchain.text_splitter import RecursiveCharacterTextSplitter
    documentos = cargar_documento(ruta)
    if not documentos:
        return 0
    splitter   = RecursiveCharacterTextSplitter(chunk_size=CHUNK_SIZE, chunk_overlap=CHUNK_OVERLAP)
    fragmentos = splitter.split_documents(documentos)
    _obtener_vectorstore().add_documents(fragmentos)
    logger.info("'%s' indexado: %d fragmentos.", ruta.name, len(fragmentos))
    return len(fragmentos)


def indexar_todos_los_documentos() -> int:
    try:
        vs = _obtener_vectorstore()
        # Prevent duplicate indexing if there are already documents
        if vs._collection.count() > 0:
            logger.info("La base de datos vectorial ya contiene documentos. Se omite indexación automática.")
            return 0
    except Exception as e:
        logger.warning(f"No se pudo verificar la colección: {e}")
        
    extensiones = {".pdf", ".txt", ".docx", ".doc", ".xlsx", ".xls"}
    total = 0
    for archivo in DOCS_DIR.iterdir():
        if archivo.suffix.lower() in extensiones:
            total += agregar_documento_a_base(archivo)
    return total


# ============================================================
# CADENA RAG (bot de WhatsApp)
# ============================================================

def _obtener_llm():
    from langchain_openai import ChatOpenAI
    api_key = os.environ.get("OPENAI_API_KEY")
    if not api_key:
        raise RuntimeError("Falta la variable de entorno OPENAI_API_KEY.")
    return ChatOpenAI(model=MODEL_NAME, temperature=0.2, openai_api_key=api_key)


def construir_cadena_qa():
    from langchain.chains import RetrievalQA
    from langchain.prompts import PromptTemplate

    retriever = _obtener_vectorstore().as_retriever(
        search_type="similarity", search_kwargs={"k": 5}
    )

    PROMPT_SISTEMA = (
        "Sos un asistente especializado en salud, facturación médica y nomencladores. "
        "Respondé SIEMPRE en español, de forma clara y precisa. "
        "Si la información no está en los documentos disponibles, decilo claramente. "
        "No inventes datos ni códigos que no estén en el contexto."
    )
    template = (
        "{ps}\n\nContexto de los documentos:\n{{context}}\n\nPregunta: {{question}}\nRespuesta:"
    ).format(ps=PROMPT_SISTEMA)

    prompt = PromptTemplate(input_variables=["context", "question"], template=template)
    return RetrievalQA.from_chain_type(
        llm=_obtener_llm(),
        chain_type="stuff",
        retriever=retriever,
        chain_type_kwargs={"prompt": prompt},
        return_source_documents=False,
    )


_cadena_qa = None


def responder(pregunta: str) -> str:
    global _cadena_qa
    if _cadena_qa is None:
        _cadena_qa = construir_cadena_qa()
    try:
        resultado = _cadena_qa.invoke({"query": pregunta})
        return resultado.get("result", "No pude generar una respuesta.")
    except Exception as e:
        logger.error("Error RAG: %s", e)
        return "Ocurrió un error al procesar tu consulta. Por favor intentá de nuevo."


# ============================================================
# ANÁLISIS DE FOJAS QUIRÚRGICAS
# ============================================================

PROMPT_FOJA_TEXTO = """Sos un experto en facturación médica quirúrgica argentina.
Se te proporciona el contenido de una foja quirúrgica y el contexto de nomencladores e instructivos.

Tu tarea:
1. Identificar los PROCEDIMIENTOS QUIRÚRGICOS realizados.
2. Determinar el CÓDIGO de nomenclador (NOMIVAC, PMO, AAM u otro de los instructivos).
3. Indicar la UNIDAD DE VALOR o MONTO si está disponible.
4. Señalar MATERIALES o GASTOS ADICIONALES facturables (prótesis, descartables, etc.).
5. Indicar REGLAS ESPECIALES (reducción por acto múltiple, códigos complementarios, etc.).
6. Si algo no está claro o sin código, decilo explícitamente.

Estructura tu respuesta así:

## Procedimientos identificados
[lista de procedimientos de la foja]

## Códigos de facturación
[código - descripción - unidad/monto por procedimiento]

## Materiales o gastos adicionales facturables
[si corresponde]

## Reglas y observaciones
[consideraciones especiales, reducciones, actos múltiples, etc.]

## Resumen para presentar
[texto listo para copiar y presentar a la obra social]

---
CONTEXTO DE NOMENCLADORES E INSTRUCTIVOS:
{context}

---
CONTENIDO DE LA FOJA QUIRÚRGICA:
{foja}
"""

PROMPT_FOJA_VISION = """Sos un experto en facturación médica quirúrgica argentina.
En la imagen adjunta se encuentra una FOJA QUIRÚRGICA o documento médico.
Leé e interpretá TODA la información visible.

Tenés acceso al siguiente contexto de nomencladores e instructivos de facturación:
---
{context}
---

Realizá el siguiente análisis CODIFICANDO ESTRICTAMENTE EN BASE A LOS INSTRUCTIVOS Y NOMENCLADORES PROPORCIONADOS:
1. Identificá los PROCEDIMIENTOS QUIRÚRGICOS realizados.
2. Determiná el CÓDIGO de nomenclador para cada procedimiento (utilizando únicamente los instructivos y nomencladores en contexto).
3. Indicá la UNIDAD DE VALOR o MONTO si está disponible en el nomenclador.
4. Señalá MATERIALES o GASTOS ADICIONALES facturables (prótesis, descartables, etc.).
5. Indicá REGLAS ESPECIALES (reducción por acto múltiple, códigos complementarios, etc.).
6. Si algo no está claro o sin código asignado, decilo explícitamente.

Estructurá tu respuesta así:

## Procedimientos identificados
[lista]

## Códigos de facturación
[código - descripción - unidad/monto]

## Materiales o gastos adicionales facturables
[si corresponde]

## Reglas y observaciones
[consideraciones especiales]

## Resumen para presentar
[texto listo para copiar y presentar a la obra social]
"""


def _obtener_contexto_rag() -> str:
    """Recupera fragmentos de nomencladores relevantes para cirugía.
    Si no hay documentos cargados, devuelve string vacío sin tocar ChromaDB.
    """
    # Si no hay documentos, no intentamos tocar ChromaDB (evita errores de colección vacía)
    extensiones = {".pdf", ".txt", ".docx", ".doc", ".xlsx", ".xls"}
    hay_docs = any(
        f.suffix.lower() in extensiones
        for f in DOCS_DIR.iterdir()
        if f.is_file()
    ) if DOCS_DIR.exists() else False

    if not hay_docs:
        return ""  # Sin contexto: GPT usará su conocimiento general

    try:
        retriever  = _obtener_vectorstore().as_retriever(
            search_type="similarity", search_kwargs={"k": 8}
        )
        fragmentos = retriever.invoke(
            "códigos facturación cirugía quirúrgico procedimiento nomenclador honorarios"
        )
        contexto = "\n\n".join(f.page_content for f in fragmentos)
        return contexto
    except Exception as e:
        logger.warning("No se pudo recuperar contexto RAG (omitido): %s", e)
        return ""


def _analizar_con_vision(imagen_bytes: bytes, mime_type: str, contexto: str) -> dict:
    """
    Envía la imagen a GPT-4o Vision con el contexto de nomencladores.
    Retorna {'ok': True, 'analisis': str} o {'ok': False, 'error': str}.
    """
    import base64
    from openai import OpenAI

    api_key = os.environ.get("OPENAI_API_KEY")
    if not api_key:
        return {"ok": False, "error": "Falta la variable de entorno OPENAI_API_KEY."}

    client    = OpenAI(api_key=api_key)
    b64_image = base64.b64encode(imagen_bytes).decode("utf-8")

    # Sección de contexto solo si hay documentos cargados
    if contexto.strip():
        seccion_contexto = (
            "Tenés acceso al siguiente contexto de nomencladores e instructivos de facturación:\n"
            "---\n"
            f"{contexto}\n"
            "---\n\n"
        )
    else:
        seccion_contexto = (
            "No hay nomencladores específicos cargados. Usá tu conocimiento general "
            "sobre nomencladores médicos argentinos (NOMIVAC, PMO, AAM, etc.).\n\n"
        )

    prompt = PROMPT_FOJA_VISION.format(context=seccion_contexto)

    try:
        respuesta = client.chat.completions.create(
            model=MODEL_VISION,
            messages=[
                {
                    "role": "user",
                    "content": [
                        {"type": "text", "text": prompt},
                        {
                            "type": "image_url",
                            "image_url": {
                                "url": f"data:{mime_type};base64,{b64_image}",
                                "detail": "auto",   # auto = más rápido, evita timeouts
                            },
                        },
                    ],
                }
            ],
            max_tokens=2000,
            timeout=55,   # 55 seg (Render free corta a 60s)
        )
        return {"ok": True, "analisis": respuesta.choices[0].message.content}
    except Exception as e:
        err_str = str(e).lower()
        # Reintentar una vez si la API está sobrecargada
        if "overloaded" in err_str or "rate limit" in err_str or "529" in err_str:
            import time
            logger.warning("API sobrecargada, reintentando en 8 segundos...")
            time.sleep(8)
            try:
                respuesta = client.chat.completions.create(
                    model=MODEL_VISION,
                    messages=[
                        {
                            "role": "user",
                            "content": [
                                {"type": "text", "text": prompt},
                                {
                                    "type": "image_url",
                                    "image_url": {
                                        "url": f"data:{mime_type};base64,{b64_image}",
                                        "detail": "auto",
                                    },
                                },
                            ],
                        }
                    ],
                    max_tokens=2000,
                    timeout=55,
                )
                return {"ok": True, "analisis": respuesta.choices[0].message.content}
            except Exception as e2:
                logger.error("Error en reintento GPT-4o Vision: %s", e2)
                return {
                    "ok": False,
                    "error": "La API de OpenAI está temporalmente sobrecargada. Esperá unos minutos e intentá de nuevo."
                }
        logger.error("Error en GPT-4o Vision: %s", e)
        return {"ok": False, "error": f"Error al analizar la imagen: {e}"}


def analizar_foja_quirurgica(ruta: Path) -> dict:
    """
    Analiza una foja quirúrgica. Acepta:
      - Imagen directa (JPG, PNG, WEBP)  → GPT-4o Vision
      - PDF con texto extraíble          → GPT-4o texto + RAG
      - PDF escaneado (sin texto)        → convierte a imagen → GPT-4o Vision

    Devuelve {'ok': bool, 'analisis': str} o {'ok': False, 'error': str}.
    """
    ext      = ruta.suffix.lower()
    contexto = _obtener_contexto_rag()

    # ════════════════════════════════════════════════════════
    # CASO A: imagen directa (JPG, PNG, WEBP, GIF)
    # ════════════════════════════════════════════════════════
    if ext in EXTENSIONES_IMAGEN:
        try:
            imagen_bytes = ruta.read_bytes()
            mime_type    = MIME_TIPOS.get(ext, "image/jpeg")
            return _analizar_con_vision(imagen_bytes, mime_type, contexto)
        except Exception as e:
            logger.error("Error al leer imagen: %s", e)
            return {"ok": False, "error": "No se pudo leer el archivo de imagen."}

    # ════════════════════════════════════════════════════════
    # CASO B: PDF
    # ════════════════════════════════════════════════════════
    if ext == ".pdf":
        import fitz  # PyMuPDF

        try:
            doc        = fitz.open(str(ruta))
            texto_foja = "\n".join(pagina.get_text() for pagina in doc).strip()
        except Exception as e:
            logger.error("Error al leer PDF: %s", e)
            return {"ok": False, "error": "No se pudo leer el PDF. Verificá que sea un archivo válido."}

        # ── B1: PDF con texto extraíble ──────────────────────
        if texto_foja:
            from langchain.schema import HumanMessage
            try:
                llm      = _obtener_llm()
                prompt   = PROMPT_FOJA_TEXTO.format(context=contexto, foja=texto_foja[:6000])
                response = llm.invoke([HumanMessage(content=prompt)])
                doc.close()
                return {"ok": True, "analisis": response.content}
            except Exception as e:
                doc.close()
                logger.error("Error en LLM con texto: %s", e)
                return {"ok": False, "error": "Error al generar el análisis. Verificá la OPENAI_API_KEY."}

        # ── B2: PDF escaneado → convertir a imagen con PyMuPDF ──
        logger.info("PDF sin texto. Usando Vision en la primera página.")
        try:
            pagina    = doc[0]
            mat       = fitz.Matrix(2.0, 2.0)  # resolución 2×
            pixmap    = pagina.get_pixmap(matrix=mat)
            img_bytes = pixmap.tobytes("png")
            doc.close()
            return _analizar_con_vision(img_bytes, "image/png", contexto)
        except Exception as e:
            doc.close()
            logger.error("Error al convertir PDF escaneado: %s", e)
            return {"ok": False, "error": "No se pudo convertir el PDF escaneado a imagen."}

    return {"ok": False, "error": f"Formato no soportado: {ext}. Usá PDF, JPG, PNG o WEBP."}


# ============================================================
# UTILIDADES
# ============================================================

def listar_documentos() -> list:
    """Lista los archivos en DOCS_DIR con metadata básica."""
    extensiones = {".pdf", ".txt", ".docx", ".doc", ".xlsx", ".xls"}
    docs = []
    for archivo in sorted(DOCS_DIR.iterdir()):
        if archivo.suffix.lower() in extensiones:
            docs.append({
                "nombre":    archivo.name,
                "tamaño_kb": round(archivo.stat().st_size / 1024, 1),
            })
    return docs
