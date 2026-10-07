import os
import re
import time
import uuid
import threading
import html
import unicodedata
from pathlib import Path
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed

import requests
from requests.adapters import HTTPAdapter
import urllib3
from urllib3.util.retry import Retry
import sqlite3
import queue
import json

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

import openpyxl
from flask import Flask, request, render_template, send_file, jsonify, redirect

from generador_anexos import (
    extraer_pacientes_afiliados_excel,
    generar_anexos_docx,
    generar_anexos_pdf,
    buscar_rnos_por_nombre,
    extraer_rnos_y_obrasocial,
    formatear_fecha,
    ordenar_pacientes_anexos,
    CATALOGO_HOSPITALES,
    HOSPITAL_DEFECTO,
    resolver_hospital,
)

# ── Módulos del bot de IA (opcionales: requieren pip install -r requirements.txt) ──
try:
    from bot_rag import (
        responder,
        agregar_documento_a_base,
        listar_documentos,
        analizar_foja_quirurgica,
        indexar_todos_los_documentos,
        DOCS_DIR,
    )
    from whatsapp_api import verificar_webhook, extraer_mensaje, enviar_texto, descargar_media, marcar_leido
    BOT_DISPONIBLE = True
except ImportError as _e:
    BOT_DISPONIBLE = False
    _bot_error = (
        "El módulo de IA no está disponible. "
        "Ejecutá: pip install -r requirements.txt\n"
        f"Detalle: {_e}"
    )
    # Funciones vacías para que los endpoints no exploten
    def responder(p):              return _bot_error
    def agregar_documento_a_base(r): return 0
    def listar_documentos():       return []
    def analizar_foja_quirurgica(r): return {"ok": False, "error": _bot_error}
    def indexar_todos_los_documentos(): return 0
    def verificar_webhook(a):      return "Bot no disponible", 503
    def extraer_mensaje(p):        return None
    def enviar_texto(n, t):        return False
    def descargar_media(m):        return None, None
    def marcar_leido(m):           pass
    DOCS_DIR = Path(__file__).resolve().parent / "documentos"
    DOCS_DIR.mkdir(exist_ok=True)

try:
    from sss_beneficiarios_hospitales.data import DataBeneficiariosSSSHospital
except ImportError:
    DataBeneficiariosSSSHospital = None


# ============================================================
# AUTO-INDEXING EN SEGUNDO PLANO
# ============================================================

def _iniciar_indexacion_background():
    if BOT_DISPONIBLE:
        def tarea():
            try:
                # Esto indexará los documentos si la colección está vacía
                indexar_todos_los_documentos()
            except Exception as e:
                print(f"Error en auto-indexación: {e}")
        
        t = threading.Thread(target=tarea, daemon=True)
        t.start()

# Disparar indexación al iniciar
_iniciar_indexacion_background()

# ============================================================
# CONFIGURACIÓN
# ============================================================

BASE_DIR = Path(__file__).resolve().parent

UPLOAD_DIR = BASE_DIR / "uploads"
RESULT_DIR = BASE_DIR / "resultados"

UPLOAD_DIR.mkdir(exist_ok=True)
RESULT_DIR.mkdir(exist_ok=True)

FILA_INICIO = 11
COLUMNA_DNI = 1
COLUMNA_OBRA_SOCIAL = 7

# Concurrencia con Pool de Sesiones Independientes y Base de Datos Local
NUM_SESSIONS = int(os.environ.get("SSS_NUM_SESSIONS", 4))
MAX_WORKERS = int(os.environ.get("SSS_MAX_WORKERS", 4))
ESPERA_ENTRE_CONSULTAS = float(os.environ.get("SSS_DELAY", 0.25))
MAX_CONSULTAS_POR_SESION = int(os.environ.get("SSS_MAX_REQS_PER_SESSION", 75))
CACHE_DNI = {}
CACHE_LOCK = threading.Lock()

# Base de datos SQLite persistente para resolver DNIs conocidos en 0 segundos
CACHE_DB_PATH = BASE_DIR / "padron_cache.sqlite"

def init_cache_db():
    try:
        with sqlite3.connect(CACHE_DB_PATH) as conn:
            conn.execute("""
                CREATE TABLE IF NOT EXISTS padron_cache (
                    dni TEXT PRIMARY KEY,
                    resultado TEXT,
                    fecha TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            """)
            conn.execute("CREATE INDEX IF NOT EXISTS idx_padron_dni ON padron_cache(dni)")
    except Exception as e:
        print(f"Error inicializando sqlite cache: {e}")

init_cache_db()

def obtener_cache_multiples(dnis):
    encontrados = {}
    if not dnis:
        return encontrados
    try:
        with sqlite3.connect(CACHE_DB_PATH, timeout=30.0) as conn:
            cur = conn.cursor()
            for i in range(0, len(dnis), 900):
                lote = dnis[i:i+900]
                q = f"SELECT dni, resultado FROM padron_cache WHERE dni IN ({','.join('?' for _ in lote)})"
                for row in cur.execute(q, lote):
                    if row[1] and row[1] != "ERROR CONSULTA":
                        encontrados[str(row[0])] = row[1]
    except Exception as e:
        print(f"Error leyendo cache sqlite: {e}")
    return encontrados

def guardar_cache_multiples(items):
    if not items:
        return
    try:
        with sqlite3.connect(CACHE_DB_PATH, timeout=30.0) as conn:
            conn.executemany(
                "INSERT OR REPLACE INTO padron_cache (dni, resultado) VALUES (?, ?)",
                [(str(d), str(r)) for d, r in items if r and r != "ERROR CONSULTA"]
            )
    except Exception as e:
        print(f"Error guardando cache sqlite: {e}")


app = Flask(__name__)

app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024


# ============================================================
# LIMPIEZA Y VALIDACIÓN DE DNI
# ============================================================

def limpiar_dni(val):
    """
    Normaliza y valida un DNI:
    - Extrae únicamente dígitos numéricos.
    - Elimina sufijo decimal '.0' proveniente de celdas float en Excel.
    - Ignora filas vacías, encabezados, pies de página o textos alfanuméricos.
    - Valida longitud típica de DNI argentino (entre 5 y 9 dígitos).
    """
    if val is None:
        return None
    val_str = str(val).strip()
    if not val_str or val_str.lower() in ["none", "null", "nan", "total", "totales"]:
        return None
    if val_str.endswith(".0"):
        val_str = val_str[:-2].strip()
    digitos = re.sub(r'\D', '', val_str)
    if 5 <= len(digitos) <= 9:
        return digitos
    return None




# ============================================================
# ESTADO DE PROCESAMIENTO
# ============================================================

procesos = {}
ESTADOS_PROCESOS_DIR = BASE_DIR / "estados_tmp"
ESTADOS_PROCESOS_DIR.mkdir(exist_ok=True)
ESTADO_LOCK = threading.Lock()

def guardar_estado_proceso(id_proceso, datos):
    with ESTADO_LOCK:
        procesos[id_proceso] = dict(datos)
    try:
        ruta = ESTADOS_PROCESOS_DIR / f"{id_proceso}.json"
        ruta.write_text(json.dumps(datos, ensure_ascii=False), encoding="utf-8")
    except Exception:
        pass

def obtener_estado_proceso(id_proceso):
    with ESTADO_LOCK:
        if id_proceso in procesos:
            return dict(procesos[id_proceso])
    ruta = ESTADOS_PROCESOS_DIR / f"{id_proceso}.json"
    if ruta.exists():
        try:
            d = json.loads(ruta.read_text(encoding="utf-8"))
            with ESTADO_LOCK:
                procesos[id_proceso] = d
            return dict(d)
        except Exception:
            pass
    return None


# ============================================================
# CLIENTE SSSALUD ULTRARRÁPIDO & PARSER
# ============================================================

def parsear_respuesta_sss(html_text):
    if not html_text:
        return "ERROR CONSULTA"

    texto = html.unescape(html_text)
    # Corregir posibles secuencias mojibake UTF-8 decodificadas como latin-1
    for malo, bueno in [
        ('Ã³', 'ó'), ('Ã', 'Ó'), ('Ã¡', 'á'), ('Ã', 'Á'),
        ('Ã©', 'é'), ('Ã', 'É'), ('Ã­', 'í'), ('Ã', 'Í'),
        ('Ãº', 'ú'), ('Ã', 'Ú'), ('Ã±', 'ñ'), ('Ã', 'Ñ')
    ]:
        texto = texto.replace(malo, bueno)

    texto_upper = texto.upper()

    # 1. NO AFILIADO
    indicadores_no_afiliado = [
        "NO SE REPORTAN DATOS PARA EL NUMERO DE DOCUMENTO",
        "NO SE REPORTAN DATOS",
        "NO AFILIADO",
        "NO REGISTRA COBERTURA",
        "NO POSEE COBERTURA",
        "SIN COBERTURA",
    ]
    es_no_afiliado = any(k in texto_upper for k in indicadores_no_afiliado)
    tiene_afiliacion = bool(
        re.search(r'DATOS\s+DE\s+AFILIACI(?:Ó|O)N\s+VIGENTE', texto, re.I)
        or "AFILIACION VIGENTE" in texto_upper
        or "AFILIACIÓN VIGENTE" in texto_upper
    )

    if es_no_afiliado and not tiene_afiliacion:
        return "NO AFILIADO"

    # 2. AFILIADO VIGENTE
    if tiene_afiliacion:
        codigo = None
        denominacion = None

        # Regex flexible para capturar Código de Obra Social / RNOS / RNAS
        m_cod = re.search(
            r'(?:C(?:ó|o)digo\s+(?:de\s+)?(?:Obra\s+Social|OS)|RNOS|RNAS).*?<td[^>]*>(?:<[^>]+>)*\s*([0-9\-\.]+)\s*<',
            texto,
            re.I | re.DOTALL
        )
        if m_cod:
            codigo = m_cod.group(1).strip()

        # Regex flexible para capturar Denominación
        m_den = re.search(
            r'Denominaci(?:ó|o)n\s+(?:de\s+)?(?:Obra\s+Social|OS)?.*?<td[^>]*>(?:<[^>]+>)*\s*([^<]+?)\s*<',
            texto,
            re.I | re.DOTALL
        )
        if m_den:
            denominacion = m_den.group(1).strip()

        # Fallback a BeautifulSoup si falta alguno
        if not (codigo and denominacion):
            try:
                from bs4 import BeautifulSoup
                soup = BeautifulSoup(texto, 'html.parser')
                for tr in soup.find_all('tr'):
                    tds = [td.get_text(strip=True) for td in tr.find_all(['td', 'th'])]
                    if len(tds) >= 2:
                        k_clean = re.sub(r'[^a-z0-9]', '', tds[0].lower())
                        v = tds[1].strip()
                        if any(x in k_clean for x in ['codigodeobrasocial', 'codigoobrasocial', 'rnos', 'rnas', 'codigoos']):
                            if not codigo and v:
                                codigo = v
                        elif any(x in k_clean for x in ['denominacionobrasocial', 'denominacionos', 'nombreobrasocial']):
                            if not denominacion and v:
                                denominacion = v
            except Exception:
                pass

        # Limpiar y normalizar código numérico
        if codigo:
            digs = re.sub(r'\D', '', codigo)
            if 4 <= len(digs) <= 8:
                codigo = digs

        # Si aún no tenemos código pero sí denominación, resolver por diccionario oficial
        if not codigo and denominacion:
            codigo = buscar_rnos_por_nombre(denominacion)

        if codigo and denominacion:
            return f"{codigo} - {denominacion}"
        elif denominacion:
            cod_fb = buscar_rnos_por_nombre(denominacion)
            if cod_fb:
                return f"{cod_fb} - {denominacion}"
            return str(denominacion)
        elif codigo:
            return str(codigo)
        return "AFILIADO - SIN OBRA SOCIAL IDENTIFICADA"

    if es_no_afiliado:
        return "NO AFILIADO"

    return "ERROR CONSULTA"


def parsear_datos_fallback(datos):
    afiliado = datos.get("afiliado")
    if afiliado is False:
        return "NO AFILIADO"
    if afiliado is True:
        tablas = datos.get("tablas", [])
        codigo_obra_social = None
        denominacion_obra_social = None
        for tabla in tablas:
            if str(tabla.get("name", "")).strip().upper() == "AFILIADO":
                data = tabla.get("data", {})
                codigo_obra_social = data.get("Código de Obra Social")
                denominacion_obra_social = data.get("Denominación Obra Social")
                break
        if codigo_obra_social and denominacion_obra_social:
            return f"{codigo_obra_social} - {denominacion_obra_social}"
        if denominacion_obra_social:
            return str(denominacion_obra_social)
        if codigo_obra_social:
            return str(codigo_obra_social)
        return "AFILIADO - SIN OBRA SOCIAL IDENTIFICADA"
    return "ERROR CONSULTA"


class SessionWorker:
    """Instancia de sesión HTTP independiente con su propia cookie PHPSESSID y auto-renovación"""
    LOGIN_URL = 'https://seguro.sssalud.gob.ar/login.php?b_publica=Acceso+Restringido+para+Hospitales&opc=bus650&user=HPGD'

    def __init__(self, user, password, idx):
        self.user = user
        self.password = password
        self.idx = idx
        self.session = None
        self.logged_in = False
        self.lock = threading.Lock()
        self.ultimo_request = 0.0
        self.total_consultas = 0
        self._init_session()

    def _init_session(self):
        if self.session is not None:
            try:
                self.session.close()
            except Exception:
                pass
        self.session = requests.Session()
        adapter = HTTPAdapter(
            pool_connections=4,
            pool_maxsize=4,
            max_retries=Retry(total=3, backoff_factor=0.4, status_forcelist=[500, 502, 503, 504])
        )
        self.session.mount('https://', adapter)
        self.session.mount('http://', adapter)
        self.session.headers.update({
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
            'Accept-Language': 'es-ES,es;q=0.9',
            'Connection': 'keep-alive',
        })

    def login(self, force=False):
        with self.lock:
            if self.logged_in and not force:
                return True
            for intento in range(2):
                try:
                    # Limpiar cookies y crear sesión nueva para evitar arrastrar PHPSESSID expirada
                    self._init_session()
                    try:
                        self.session.get(self.LOGIN_URL, verify=False, timeout=8)
                    except Exception:
                        pass

                    params = {
                        '_user_name_': self.user,
                        '_pass_word_': self.password,
                        'submitbtn': 'Ingresar'
                    }
                    res = self.session.post(self.LOGIN_URL, data=params, verify=False, timeout=8)
                    text = res.text
                    url_final = res.url.lower()

                    if (
                        'usuario_logueado' in text
                        or 'nro_doc' in text
                        or 'cat=consultas' in text
                        or 'indexss.php' in url_final
                        or 'pagina_consulta' in text
                    ):
                        self.logged_in = True
                        self.total_consultas = 0
                        return True
                    else:
                        print(f"Sesión {self.idx}: intento de login {intento+1} no reconocido por SSSalud (status {res.status_code})")
                except Exception as e:
                    print(f"Error en login sesión {self.idx} (intento {intento+1}): {e}")

                time.sleep(0.5 * (intento + 1))

            self.logged_in = False
            return False


class FastSSSaludClient:
    QUERY_URL = 'https://seguro.sssalud.gob.ar/indexss.php?opc=bus650&user=HPGD&cat=consultas'

    def __init__(self, user, password, num_sessions=4):
        self.user = user
        self.password = password
        self.num_sessions = max(2, num_sessions)
        self.session_pool = queue.Queue()
        for i in range(self.num_sessions):
            worker = SessionWorker(user, password, i + 1)
            self.session_pool.put(worker)

    def query(self, dni):
        dni_str = str(dni).strip()
        if not dni_str:
            return "ERROR CONSULTA"

        # 1. Memoria RAM
        with CACHE_LOCK:
            if dni_str in CACHE_DNI and CACHE_DNI[dni_str] != "ERROR CONSULTA":
                return CACHE_DNI[dni_str]

        params = {
            'pagina_consulta': '',
            'cuil_b': '',
            'nro_doc': dni_str,
            'B1': 'Consultar'
        }

        try:
            worker = self.session_pool.get(timeout=10)
        except Exception:
            return "ERROR CONSULTA"

        try:
            # Rotación preventiva de sesión antes de que el servidor la expire por límite de consultas
            if worker.total_consultas >= MAX_CONSULTAS_POR_SESION:
                worker.login(force=True)

            max_intentos = 2
            for intento in range(max_intentos):
                # Pacing por worker para evitar bloqueos por tasa de consultas de SSSalud
                ahora = time.time()
                tiempo_desde_ultimo = ahora - worker.ultimo_request
                if tiempo_desde_ultimo < ESPERA_ENTRE_CONSULTAS:
                    time.sleep(ESPERA_ENTRE_CONSULTAS - tiempo_desde_ultimo)

                if not worker.logged_in:
                    if not worker.login(force=True):
                        time.sleep(0.5 * (intento + 1))
                        continue

                try:
                    worker.ultimo_request = time.time()
                    res = worker.session.post(self.QUERY_URL, data=params, verify=False, timeout=8)
                    worker.total_consultas += 1
                    try:
                        if res.encoding and res.encoding.lower() == 'iso-8859-1':
                            raw_b = res.content
                            try:
                                text = raw_b.decode('utf-8')
                            except UnicodeDecodeError:
                                text = raw_b.decode('latin-1', errors='replace')
                        else:
                            text = res.text or ""
                    except Exception:
                        text = res.text or ""
                    url_final = res.url.lower()

                    # Comprobar si la sesión expiró o redirigió al login
                    es_sesion_caida = (
                        res.status_code in [401, 403, 429, 500, 502, 503, 504]
                        or 'login.php' in url_final
                        or 'b_publica' in url_final
                        or ('_user_name_' in text and ('submitbtn' in text or 'Ingresar' in text))
                        or any(k in text.lower() for k in [
                            'sesión caducada', 'sesion caducada',
                            'sesión expirada', 'sesion expirada',
                            'acceso restringido', 'debe identificarse',
                            'tiempo de espera agotado'
                        ])
                    )

                    if es_sesion_caida:
                        worker.logged_in = False
                        time.sleep(0.4 * (intento + 1))
                        worker.login(force=True)
                        continue

                    resultado = parsear_respuesta_sss(text)
                    if resultado != "ERROR CONSULTA":
                        with CACHE_LOCK:
                            CACHE_DNI[dni_str] = resultado
                        return resultado

                except Exception as e:
                    print(f"Error consultando DNI {dni_str} en worker {worker.idx} (intento {intento+1}): {e}")
                    worker.logged_in = False
                    time.sleep(0.5 * (intento + 1))
                    worker.login(force=True)

            return "ERROR CONSULTA"
        finally:
            self.session_pool.put(worker)


# ============================================================
# CONEXIÓN SSSALUD
# ============================================================

def crear_conexion_sss():
    usuario = os.environ.get("SSS_USER")
    password = os.environ.get("SSS_PASSWORD")

    if not usuario or not password:
        raise RuntimeError("Faltan las variables SSS_USER y SSS_PASSWORD.")

    client = FastSSSaludClient(
        user=usuario,
        password=password,
        num_sessions=NUM_SESSIONS
    )
    return client


# ============================================================
# CONSULTA SSSALUD
# ============================================================

def consultar_dni(sss, dni):
    if isinstance(sss, FastSSSaludClient):
        return sss.query(dni)
    try:
        resultado = sss.query(dni)
        ok = resultado.get("ok")
        datos = resultado.get("resultados", {})
        if not ok:
            return "ERROR CONSULTA"
        return parsear_datos_fallback(datos)
    except Exception as e:
        print("ERROR CONSULTA:", e)
        return "ERROR CONSULTA"


# ============================================================
# PROCESAMIENTO EN SEGUNDO PLANO (MULTI-SESIÓN + CACHÉ PERSISTENTE)
# ============================================================

def procesar_archivo(
    id_proceso,
    archivo_entrada,
    archivo_salida,
    hospital_nombre=None,
    hospital_refes=None,
    especialidad="CARDIOLOGIA"
):
    estado = obtener_estado_proceso(id_proceso) or {
        "estado": "iniciando",
        "total": 0,
        "procesadas": 0,
        "fila": 0,
        "dni": "",
        "consultas": 0,
        "afiliados": 0,
        "no_afiliados": 0,
        "errores": 0,
        "ultimo_resultado": "",
        "porcentaje": 0,
        "segundos": 0,
        "estimado_restante": 0,
        "archivo": "",
        "error": ""
    }

    try:
        # ----------------------------------------------------
        # 1. ABRIR EXCEL
        # ----------------------------------------------------
        estado["estado"] = "abriendo_excel"
        guardar_estado_proceso(id_proceso, estado)
        wb = openpyxl.load_workbook(archivo_entrada)
        ws = wb.active
        ultima_fila = ws.max_row

        # ----------------------------------------------------
        # 2. DETECCIÓN DINÁMICA DE ENCABEZADOS Y COLUMNAS
        # ----------------------------------------------------
        fila_encabezado = None
        col_dni = None
        col_rnos = None
        col_obra_social = None
        max_col_detectada = 1

        # Escanear filas 1 a 45 buscando la fila que tenga DNI / Documento
        for f in range(1, min(45, ultima_fila + 1)):
            for c in range(1, 40):
                val = str(ws.cell(f, c).value or "").strip().lower()
                if val:
                    max_col_detectada = max(max_col_detectada, c)
                    if not fila_encabezado and any(k in val for k in ["dni", "documento", "nro doc", "nro. doc", "doc"]):
                        fila_encabezado = f
                        col_dni = c

            if fila_encabezado:
                break

        # Si no se detectó fila por texto de encabezado, buscar la primera con DNI numérico
        if not fila_encabezado:
            for f in range(1, min(50, ultima_fila + 1)):
                d_val = limpiar_dni(ws.cell(f, COLUMNA_DNI).value)
                if d_val:
                    fila_encabezado = max(1, f - 1)
                    col_dni = COLUMNA_DNI
                    break

        fila_encabezado = fila_encabezado or 1
        col_dni = col_dni or COLUMNA_DNI
        fila_inicio_datos = fila_encabezado + 1

        # En la fila de encabezados, buscar si ya existen columnas para RNOS y Obra Social
        for c in range(1, 40):
            val_h = str(ws.cell(fila_encabezado, c).value or "").strip().lower()
            if not val_h:
                continue
            max_col_detectada = max(max_col_detectada, c)
            if any(k in val_h for k in ["rnos", "rnas", "código os", "codigo os"]):
                col_rnos = c
            elif any(k in val_h for k in ["obra social", "obrasocial", "cobertura", "prepaga", "o.s."]):
                col_obra_social = c

        # Si no existen, ubicarlas al final de la tabla (fuera de cualquier celda combinada)
        if not col_rnos and not col_obra_social:
            col_rnos = max_col_detectada + 1
            col_obra_social = max_col_detectada + 2
            try:
                ws.cell(fila_encabezado, col_rnos).value = "RNOS"
                ws.cell(fila_encabezado, col_obra_social).value = "OBRA SOCIAL"
            except Exception:
                pass
        elif not col_rnos:
            col_rnos = max(max_col_detectada, col_obra_social) + 1
            try:
                ws.cell(fila_encabezado, col_rnos).value = "RNOS"
            except Exception:
                pass
        elif not col_obra_social:
            col_obra_social = max(max_col_detectada, col_rnos) + 1
            try:
                ws.cell(fila_encabezado, col_obra_social).value = "OBRA SOCIAL"
            except Exception:
                pass

        # ----------------------------------------------------
        # 3. BUSCAR FILAS CON DNI VÁLIDO Y AGRUPAR
        # ----------------------------------------------------
        dni_a_filas = defaultdict(list)
        total_filas = 0

        for fila in range(fila_inicio_datos, ultima_fila + 1):
            dni_val = ws.cell(fila, col_dni).value
            dni_limpio = limpiar_dni(dni_val)
            if not dni_limpio:
                continue
            dni_a_filas[dni_limpio].append(fila)
            total_filas += 1

        dnis_unicos = list(dni_a_filas.keys())
        total = total_filas

        estado["total"] = total
        estado["procesadas"] = 0
        estado["estado"] = "iniciando"
        guardar_estado_proceso(id_proceso, estado)

        print("")
        print("====================================================")
        print("       PROCESAMIENTO SSSALUD ULTRARRÁPIDO")
        print("====================================================")
        print("Fila encabezado detectada:", fila_encabezado)
        print("Columna DNI:", col_dni)
        print("Columna RNOS:", col_rnos)
        print("Columna Obra Social:", col_obra_social)
        print("Última fila del Excel:", ultima_fila)
        print("Registros válidos de pacientes:", total)
        print("Total de DNIs únicos a consultar:", len(dnis_unicos))
        print("Columna destino Obra Social:", col_obra_social)
        print(f"Sesiones concurrentes: {NUM_SESSIONS} | Workers: {MAX_WORKERS}")
        print("====================================================")

        if total == 0:
            wb.save(archivo_salida)
            estado["estado"] = "terminado"
            estado["porcentaje"] = 100
            estado["archivo"] = Path(archivo_salida).name
            estado["segundos"] = 0
            estado["estimado_restante"] = 0
            guardar_estado_proceso(id_proceso, estado)
            return

        # ----------------------------------------------------
        # 4. RESOLUCIÓN INSTANTÁNEA DESDE BASE / CACHÉ LOCAL
        # ----------------------------------------------------
        inicio = time.time()
        resultados_dni = obtener_cache_multiples(dnis_unicos)

        # Poblar contadores con los ya conocidos de SQLite
        cant_resueltos_db = 0
        for d, res in resultados_dni.items():
            cant = len(dni_a_filas[d])
            cant_resueltos_db += cant
            if res == "NO AFILIADO":
                estado["no_afiliados"] += cant
            elif res == "ERROR CONSULTA":
                estado["errores"] += cant
            else:
                estado["afiliados"] += cant

        estado["procesadas"] = cant_resueltos_db
        if total > 0:
            estado["porcentaje"] = min(100.0, round((cant_resueltos_db / total) * 100, 1))
        guardar_estado_proceso(id_proceso, estado)

        dnis_pendientes = [d for d in dnis_unicos if d not in resultados_dni]
        print(f"DNIs resueltos al instante por Base de Datos/Caché: {len(resultados_dni)} ({cant_resueltos_db} filas)")
        print(f"DNIs pendientes de consulta remota: {len(dnis_pendientes)}")

        # ----------------------------------------------------
        # 5. SI HAY PENDIENTES, CONSULTAR CON SESIONES PARALELAS
        # ----------------------------------------------------
        nuevos_para_guardar = []

        if dnis_pendientes:
            estado["estado"] = "conectando_sssalud"
            guardar_estado_proceso(id_proceso, estado)
            sss = crear_conexion_sss()
            estado["estado"] = "consultando"
            guardar_estado_proceso(id_proceso, estado)

            lock_estado = threading.Lock()
            ultimo_guardado = [time.time()]

            def procesar_un_dni(dni):
                time.sleep(0.04)  # Espaciado suave para distribuir la carga entre hilos
                res = consultar_dni(sss, dni)
                filas = dni_a_filas[dni]
                cant = len(filas)

                with lock_estado:
                    r_cod, r_den = extraer_rnos_y_obrasocial(res)
                    if r_cod and r_den and r_cod not in r_den:
                        res_completo = f"{r_cod} - {r_den}"
                    else:
                        res_completo = res

                    resultados_dni[dni] = res_completo
                    nuevos_para_guardar.append((dni, res_completo))
                    # Guardar por lotes de 25 en SQLite para persistencia inmediata y evitar pérdidas
                    if len(nuevos_para_guardar) % 25 == 0:
                        guardar_cache_multiples(list(nuevos_para_guardar[-25:]))

                    estado["procesadas"] += cant
                    estado["consultas"] += 1
                    estado["fila"] = filas[-1]
                    estado["dni"] = dni
                    estado["ultimo_resultado"] = res

                    if res == "NO AFILIADO":
                        estado["no_afiliados"] += cant
                    elif res == "ERROR CONSULTA":
                        estado["errores"] += cant
                    else:
                        estado["afiliados"] += cant

                    transcurrido = time.time() - inicio
                    estado["segundos"] = round(transcurrido, 1)

                    if total > 0:
                        estado["porcentaje"] = min(100.0, round((estado["procesadas"] / total) * 100, 1))
                        pendientes_count = total - estado["procesadas"]
                        if estado["procesadas"] > cant_resueltos_db and transcurrido > 0:
                            velocidad = (estado["procesadas"] - cant_resueltos_db) / transcurrido
                            estado["estimado_restante"] = round(pendientes_count / max(0.1, velocidad), 1)

                    ahora = time.time()
                    if ahora - ultimo_guardado[0] >= 0.5:
                        guardar_estado_proceso(id_proceso, estado)
                        ultimo_guardado[0] = ahora

                return dni, res

            with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
                futuros = [executor.submit(procesar_un_dni, d) for d in dnis_pendientes]
                for f in as_completed(futuros):
                    try:
                        f.result()
                    except Exception as ex_hilo:
                        print(f"Error en hilo de consulta: {ex_hilo}")

            # Asegurar que ningún DNI quede sin registrar o pendiente
            for d in dnis_pendientes:
                if d not in resultados_dni:
                    resultados_dni[d] = "ERROR CONSULTA"
                    cant = len(dni_a_filas[d])
                    estado["errores"] += cant
                    estado["procesadas"] += cant

            estado["procesadas"] = total
            estado["porcentaje"] = 100
            estado["estado"] = "escribiendo_excel"
            estado["ultimo_resultado"] = "Escribiendo RNOS y Obras Sociales en el Excel..."
            guardar_estado_proceso(id_proceso, estado)

            # Guardar remanente en la base de datos persistente SQLite
            if nuevos_para_guardar:
                guardar_cache_multiples(nuevos_para_guardar)

        # ----------------------------------------------------
        # 6. ESCRIBIR RESULTADOS EN EXCEL (RNOS Y OBRA SOCIAL)
        # ----------------------------------------------------
        estado["estado"] = "escribiendo_excel"
        estado["ultimo_resultado"] = "Escribiendo resultados en memoria de Excel..."
        guardar_estado_proceso(id_proceso, estado)
        print("Escribiendo resultados en memoria...")

        for dni, filas in dni_a_filas.items():
            res_dni = resultados_dni.get(dni, "ERROR CONSULTA")
            rnos_val, os_val = extraer_rnos_y_obrasocial(res_dni)

            for fila in filas:
                # Escribir RNOS
                if col_rnos:
                    try:
                        cell_r = ws.cell(fila, col_rnos)
                        if type(cell_r).__name__ != "MergedCell":
                            cell_r.value = rnos_val
                    except Exception as e_r:
                        print(f"Error escribiendo RNOS fila {fila}: {e_r}")

                # Escribir Obra Social
                if col_obra_social:
                    try:
                        cell_o = ws.cell(fila, col_obra_social)
                        if type(cell_o).__name__ != "MergedCell":
                            cell_o.value = os_val
                    except Exception as e_o:
                        print(f"Error escribiendo Obra Social fila {fila}: {e_o}")

        # ----------------------------------------------------
        # 7. GUARDADO FINAL DE EXCEL ATÓMICO
        # ----------------------------------------------------
        estado["estado"] = "guardando"
        estado["ultimo_resultado"] = "Guardando archivo Excel en el servidor..."
        guardar_estado_proceso(id_proceso, estado)

        archivo_salida_path = Path(archivo_salida)
        archivo_salida_tmp = archivo_salida_path.with_suffix(f"{archivo_salida_path.suffix}.tmp_{uuid.uuid4().hex[:8]}")
        wb.save(archivo_salida_tmp)
        try:
            wb.close()
            del wb
        except Exception:
            pass
        import gc
        gc.collect()

        if archivo_salida_path.exists():
            try:
                archivo_salida_path.unlink()
            except Exception:
                pass
        archivo_salida_tmp.replace(archivo_salida_path)

        # El Excel ya está 100% guardado y disponible para descarga inmediata
        estado["archivo"] = archivo_salida_path.name
        estado["procesadas"] = total
        estado["porcentaje"] = 100
        estado["estado"] = "generando_anexos"
        estado["ultimo_resultado"] = "Generando y ordenando fojas de Anexo II en Word..."
        guardar_estado_proceso(id_proceso, estado)

        # ----------------------------------------------------
        # 8. GENERACIÓN AUTOMÁTICA DE ANEXOS II (WORD .DOCX OFICIAL)
        # ----------------------------------------------------
        archivo_anexos = RESULT_DIR / f"anexos_{id_proceso}.docx"
        cant_anexos = 0
        try:
            esp_base = especialidad or estado.get("especialidad", "CARDIOLOGIA")
            pacientes_afiliados = extraer_pacientes_afiliados_excel(archivo_salida, especialidad_defecto=esp_base)
            if pacientes_afiliados:
                h_nom = hospital_nombre or estado.get("hospital_nombre") or HOSPITAL_DEFECTO["nombre"]
                h_ref = hospital_refes or estado.get("hospital_refes") or HOSPITAL_DEFECTO["refes"]
                def progreso_anexo(actual, total_anx):
                    estado["ultimo_resultado"] = f"Generando foja {actual} de {total_anx} en Word..."
                    guardar_estado_proceso(id_proceso, estado)

                cant_anexos = generar_anexos_docx(
                    pacientes_afiliados,
                    archivo_anexos,
                    hospital_nombre=h_nom,
                    hospital_refes=h_ref,
                    progress_callback=progreso_anexo
                )
                print(f"Se generaron exitosamente {cant_anexos} Anexos II en Word (.docx) para {h_nom}")
        except Exception as e_anexos:
            print(f"Aviso al generar Anexos II automáticos en Word: {e_anexos}")

        # ----------------------------------------------------
        # 9. FINALIZADO
        # ----------------------------------------------------
        estado["estado"] = "terminado"
        estado["procesadas"] = total
        estado["porcentaje"] = 100
        estado["archivo"] = Path(archivo_salida).name
        estado["archivo_anexos"] = Path(archivo_anexos).name if cant_anexos > 0 else ""
        estado["archivo_anexos_pdf"] = ""
        estado["anexos_generados"] = cant_anexos
        estado["segundos"] = round(time.time() - inicio, 1)
        estado["estimado_restante"] = 0
        guardar_estado_proceso(id_proceso, estado)

        print("")
        print("====================================================")
        print("          PROCESAMIENTO TERMINADO EXITOSAMENTE")
        print("====================================================")
        print(f"Tiempo total: {estado['segundos']} segundos")
        print("Registros procesados:", estado["procesadas"])
        print("Afiliados:", estado["afiliados"])
        print("No afiliados:", estado["no_afiliados"])
        print("Errores:", estado["errores"])
        print(f"Anexos II Word generados: {cant_anexos}")
        print("====================================================")

    except Exception as e:
        print("")
        print("====================================================")
        print("ERROR GENERAL EN PROCESAMIENTO")
        print("====================================================")
        print(e)
        estado["estado"] = "error"
        estado["error"] = str(e)
        guardar_estado_proceso(id_proceso, estado)


# ============================================================
# PÁGINA PRINCIPAL
# ============================================================

@app.route("/", methods=["GET"])
def index():

    return render_template(
        "index.html"
    )


# ============================================================
# INICIAR PROCESAMIENTO
# ============================================================

@app.route(
    "/procesar",
    methods=["POST"]
)
def procesar():

    archivo = request.files.get(
        "archivo"
    )

    # --------------------------------------------------------
    # VALIDAR ARCHIVO
    # --------------------------------------------------------

    if not archivo or archivo.filename == "":

        return """
        <h2>No se seleccionó ningún archivo.</h2>
        <a href="/">Volver</a>
        """

    if not archivo.filename.lower().endswith(
        ".xlsx"
    ):

        return """
        <h2>El archivo debe ser .xlsx</h2>
        <a href="/">Volver</a>
        """

    # --------------------------------------------------------
    # ID ÚNICO
    # --------------------------------------------------------

    identificador = uuid.uuid4().hex

    archivo_entrada = (
        UPLOAD_DIR
        /
        f"{identificador}_{archivo.filename}"
    )

    archivo_salida = (
        RESULT_DIR
        /
        f"procesado_{identificador}_{archivo.filename}"
    )

    # --------------------------------------------------------
    # GUARDAR ARCHIVO
    # --------------------------------------------------------

    archivo.save(
        archivo_entrada
    )

    # --------------------------------------------------------
    # RESOLVER HOSPITAL Y ESPECIALIDAD SELECCIONADOS
    # --------------------------------------------------------

    hosp_id = request.form.get("hospital", "san_roque")
    hosp_custom_nom = request.form.get("hospital_nombre_custom", "").strip()
    hosp_custom_ref = request.form.get("hospital_refes_custom", "").strip()
    hosp_info = resolver_hospital(hosp_id, hosp_custom_nom, hosp_custom_ref)
    esp_req = request.form.get("especialidad", "").strip() or "CARDIOLOGIA"

    # --------------------------------------------------------
    # CREAR ESTADO PERSISTENTE
    # --------------------------------------------------------

    nuevo_estado = {
        "estado": "preparando",
        "total": 0,
        "procesadas": 0,
        "fila": 0,
        "dni": "",
        "consultas": 0,
        "afiliados": 0,
        "no_afiliados": 0,
        "errores": 0,
        "ultimo_resultado": "",
        "porcentaje": 0,
        "segundos": 0,
        "estimado_restante": 0,
        "archivo": "",
        "error": "",
        "hospital_id": hosp_info.get("id", "san_roque"),
        "hospital_nombre": hosp_info["nombre"],
        "hospital_refes": hosp_info["refes"],
        "especialidad": esp_req,
    }
    guardar_estado_proceso(identificador, nuevo_estado)

    # --------------------------------------------------------
    # INICIAR HILO
    # --------------------------------------------------------

    hilo = threading.Thread(
        target=procesar_archivo,
        args=(
            identificador,
            archivo_entrada,
            archivo_salida,
            hosp_info["nombre"],
            hosp_info["refes"],
            esp_req
        ),
        daemon=False
    )
    hilo.start()

    # --------------------------------------------------------
    # MOSTRAR PANTALLA DE PROGRESO
    # --------------------------------------------------------

    return render_template(
        "progreso.html",
        id_proceso=identificador
    )


# ============================================================
# CONSULTAR PROGRESO (SEGURO MULTI-PROCESO GUNICORN)
# ============================================================

@app.route(
    "/progreso/<id_proceso>"
)
def progreso(id_proceso):
    estado = obtener_estado_proceso(id_proceso)
    if not estado:
        return jsonify({
            "estado": "preparando",
            "porcentaje": 0,
            "procesadas": 0,
            "total": 0,
            "afiliados": 0,
            "no_afiliados": 0,
            "errores": 0
        })
    return jsonify(estado)


# ============================================================
# DESCARGAR RESULTADO
# ============================================================

@app.route(
    "/descargar/<nombre>"
)
def descargar(nombre):

    archivo = (
        RESULT_DIR
        /
        nombre
    )

    if not archivo.exists():

        return (
            "Archivo no encontrado",
            404
        )

    return send_file(
        archivo,
        as_attachment=True
    )


# ============================================================
# DESCARGAR EXCEL PROCESADO POR ID
# ============================================================

@app.route("/descargar_excel/<id_proceso>")
def descargar_excel_por_id(id_proceso):
    for _ in range(40):
        estado = obtener_estado_proceso(id_proceso)
        if estado and estado.get("archivo"):
            archivo = RESULT_DIR / estado["archivo"]
            if archivo.exists():
                return send_file(archivo, as_attachment=True)
        for p in RESULT_DIR.glob(f"*{id_proceso}*.xlsx"):
            if p.exists() and ".tmp_" not in p.name:
                return send_file(p, as_attachment=True)
        if estado and estado.get("estado") in ["guardando", "escribiendo_excel"]:
            time.sleep(0.5)
        else:
            break


    # Reconstrucción instantánea de contingencia desde SQLite
    for entrada in UPLOAD_DIR.glob(f"*{id_proceso}*"):
        if entrada.exists() and entrada.suffix.lower() == ".xlsx":
            try:
                salida_path = RESULT_DIR / f"procesado_{entrada.name}"
                if salida_path.exists():
                    return send_file(salida_path, as_attachment=True)
                wb = openpyxl.load_workbook(entrada)
                ws = wb.active
                fila_enc = 1
                col_d = COLUMNA_DNI
                for f in range(1, min(45, ws.max_row + 1)):
                    for c in range(1, 40):
                        v = str(ws.cell(f, c).value or "").strip().lower()
                        if any(k in v for k in ["dni", "documento", "nro doc", "doc"]):
                            fila_enc = f
                            col_d = c
                            break
                    if fila_enc > 1:
                        break
                max_c = ws.max_column
                col_r = max_c + 1
                col_os = max_c + 2
                ws.cell(fila_enc, col_r).value = "RNOS"
                ws.cell(fila_enc, col_os).value = "OBRA SOCIAL"
                dnis_map = {}
                for f in range(fila_enc + 1, ws.max_row + 1):
                    d = limpiar_dni(ws.cell(f, col_d).value)
                    if d:
                        dnis_map.setdefault(d, []).append(f)
                res_cache = obtener_cache_multiples(list(dnis_map.keys()))
                for d, filas in dnis_map.items():
                    res_val = res_cache.get(d, "NO AFILIADO")
                    r_val, o_val = extraer_rnos_y_obrasocial(res_val)
                    for fil in filas:
                        ws.cell(fil, col_r).value = r_val
                        ws.cell(fil, col_os).value = o_val
                wb.save(salida_path)
                try:
                    wb.close()
                except Exception:
                    pass
                return send_file(salida_path, as_attachment=True)
            except Exception as e_rec:
                print(f"Error reconstruyendo Excel de contingencia: {e_rec}")

    return "El archivo Excel procesado aún no está listo o en proceso de guardado.", 404


# ============================================================
# DESCARGAR ANEXOS II GENERADOS
# ============================================================

@app.route("/descargar_anexos/<id_proceso>")
def descargar_anexos(id_proceso):
    archivo_anexos = RESULT_DIR / f"anexos_{id_proceso}.docx"

    # Si se está generando en segundo plano en este momento, esperar unos segundos a que termine el empaquetado
    for _ in range(40):
        if archivo_anexos.exists():
            return send_file(
                archivo_anexos,
                as_attachment=True,
                download_name=f"Anexos_II_{id_proceso[:8]}.docx"
            )
        estado = obtener_estado_proceso(id_proceso)
        if estado and estado.get("estado") in ["guardando", "generando_anexos", "escribiendo_excel"]:
            time.sleep(0.5)
        else:
            break

    if archivo_anexos.exists():
        return send_file(
            archivo_anexos,
            as_attachment=True,
            download_name=f"Anexos_II_{id_proceso[:8]}.docx"
        )

    estado = obtener_estado_proceso(id_proceso)
    archivo_salida = None
    if estado and estado.get("archivo"):
        archivo_salida = RESULT_DIR / estado["archivo"]
    if not archivo_salida or not archivo_salida.exists():
        for p in RESULT_DIR.glob(f"*{id_proceso}*.xlsx"):
            if p.exists():
                archivo_salida = p
                break

    # Si aún no existe archivo_salida, buscar en UPLOAD_DIR
    if not archivo_salida or not archivo_salida.exists():
        for entrada in UPLOAD_DIR.glob(f"*{id_proceso}*"):
            if entrada.exists() and entrada.suffix.lower() == ".xlsx":
                salida_candidata = RESULT_DIR / f"procesado_{entrada.name}"
                if not salida_candidata.exists():
                    try:
                        descargar_excel_por_id(id_proceso)
                    except Exception:
                        pass
                if salida_candidata.exists():
                    archivo_salida = salida_candidata
                break

    if archivo_salida and archivo_salida.exists():
        try:
            esp_base = (estado and estado.get("especialidad")) or "CARDIOLOGIA"
            pacientes = extraer_pacientes_afiliados_excel(archivo_salida, especialidad_defecto=esp_base)
            if pacientes:
                h_nombre = (estado and estado.get("hospital_nombre")) or HOSPITAL_DEFECTO["nombre"]
                h_refes = (estado and estado.get("hospital_refes")) or HOSPITAL_DEFECTO["refes"]
                generar_anexos_docx(
                    pacientes,
                    archivo_anexos,
                    hospital_nombre=h_nombre,
                    hospital_refes=h_refes
                )
                if archivo_anexos.exists():
                    return send_file(
                        archivo_anexos,
                        as_attachment=True,
                        download_name=f"Anexos_II_{id_proceso[:8]}.docx"
                    )
        except Exception as e:
            return f"Error generando anexos: {e}", 500

    return "No se encontraron anexos para este proceso o no hay pacientes afiliados.", 404


# ============================================================
# DESCARGAR ANEXOS II EN PDF (MASIVO LISTO PARA IMPRIMIR)
# ============================================================

@app.route("/descargar_anexos_pdf/<id_proceso>")
def descargar_anexos_pdf(id_proceso):
    # Redirigir a la descarga oficial en Word (.docx)
    return descargar_anexos(id_proceso)


# ============================================================
# VISTA PREVIA E IMPRESIÓN DIRECTA EN NAVEGADOR (CTRL + P)
# ============================================================

@app.route("/imprimir_anexos/<id_proceso>")
def imprimir_anexos(id_proceso):
    estado = obtener_estado_proceso(id_proceso)
    if not estado or not estado.get("archivo"):
        return "Proceso no encontrado o aún no terminado.", 404

    archivo_salida = RESULT_DIR / estado["archivo"]
    if not archivo_salida.exists():
        return "Archivo procesado no encontrado.", 404

    try:
        esp_base = estado.get("especialidad", "CARDIOLOGIA")
        pacientes = extraer_pacientes_afiliados_excel(archivo_salida, especialidad_defecto=esp_base)
        if not pacientes:
            return "No se encontraron pacientes afiliados para imprimir.", 404

        h_nombre = estado.get("hospital_nombre") or HOSPITAL_DEFECTO["nombre"]
        h_refes = estado.get("hospital_refes") or HOSPITAL_DEFECTO["refes"]

        hoy = date.today()
        fecha_act_str = hoy.strftime("%d-%m-%Y")
        primer_dia_este_mes = hoy.replace(day=1)
        ultimo_dia_mes_ant = primer_dia_este_mes - timedelta(days=1)
        per_dec_str = ultimo_dia_mes_ant.strftime("%m-%Y")

        return render_template(
            "imprimir_anexos.html",
            pacientes=pacientes,
            total_pacientes=len(pacientes),
            hospital_nombre=h_nombre,
            hospital_refes=h_refes,
            fecha_actual=fecha_act_str,
            periodo_declarado=per_dec_str
        )
    except Exception as e:
        return f"Error cargando fojas para impresión: {e}", 500


# ============================================================
# GENERADOR DIRECTO DE ANEXOS II (DESDE EXCEL YA VERIFICADO)
# ============================================================

@app.route("/generar_anexos", methods=["GET", "POST"])
def generar_anexos_directo():
    if request.method == "GET":
        return redirect("/")

    try:
        archivo = request.files.get("archivo")
        if not archivo or not archivo.filename:
            return """
            <!DOCTYPE html>
            <html>
            <head><title>Aviso</title><meta charset="utf-8"></head>
            <body style="font-family: sans-serif; padding: 40px; background: #0f172a; color: #f8fafc; text-align: center;">
                <h2 style="color: #f59e0b;">⚠️ No se seleccionó ningún archivo Excel</h2>
                <p style="color: #94a3b8; max-width: 500px; margin: 15px auto;">Por favor seleccioná un archivo .xlsx antes de generar los anexos.</p>
                <a href="/" style="display: inline-block; padding: 10px 20px; background: #6366f1; color: white; border-radius: 6px; text-decoration: none; font-weight: bold;">⬅ Volver al inicio</a>
            </body>
            </html>
            """, 200

        ext = Path(archivo.filename).suffix.lower()
        if ext != ".xlsx":
            return f"""
            <!DOCTYPE html>
            <html>
            <head><title>Formato no compatible</title><meta charset="utf-8"></head>
            <body style="font-family: sans-serif; padding: 40px; background: #0f172a; color: #f8fafc; text-align: center;">
                <h2 style="color: #ef4444;">❌ El archivo debe ser un Excel (.xlsx)</h2>
                <p style="color: #94a3b8; max-width: 580px; margin: 15px auto; line-height: 1.5;">El archivo subido tiene extensión <code>{ext}</code>.<br>Si es un .xls antiguo, abrilo en Excel y guardalo como <strong>"Libro de Excel (*.xlsx)"</strong>.</p>
                <a href="/" style="display: inline-block; padding: 10px 20px; background: #6366f1; color: white; border-radius: 6px; text-decoration: none; font-weight: bold;">⬅ Volver al inicio</a>
            </body>
            </html>
            """, 200

        identificador = uuid.uuid4().hex
        temp_excel = UPLOAD_DIR / f"anexo_upload_{identificador}.xlsx"
        temp_docx = RESULT_DIR / f"anexos_II_{identificador}.docx"
        archivo.save(str(temp_excel))

        esp_req = request.form.get("especialidad", "").strip() or "CARDIOLOGIA"
        hosp_id = request.form.get("hospital", "san_roque")
        hosp_custom_nom = request.form.get("hospital_nombre_custom", "").strip()
        hosp_custom_ref = request.form.get("hospital_refes_custom", "").strip()
        hosp_info = resolver_hospital(hosp_id, hosp_custom_nom, hosp_custom_ref)

        pacientes = extraer_pacientes_afiliados_excel(str(temp_excel), especialidad_defecto=esp_req)
        if not pacientes:
            return """
            <!DOCTYPE html>
            <html>
            <head><title>Sin afiliados</title><meta charset="utf-8"></head>
            <body style="font-family: sans-serif; padding: 40px; background: #0f172a; color: #f8fafc; text-align: center;">
                <h2 style="color: #f59e0b;">⚠️ No se encontraron pacientes afiliados con Obra Social válida</h2>
                <p style="color: #94a3b8; max-width: 600px; margin: 15px auto; line-height: 1.6;">
                    No se detectaron registros con Obra Social identificada en el archivo Excel subido.<br>
                    Asegurate de que contenga columnas con <strong>DNI</strong>, <strong>Paciente</strong> y <strong>Obra Social</strong> verificada.
                </p>
                <a href="/" style="display: inline-block; padding: 10px 20px; background: #6366f1; color: white; border-radius: 6px; text-decoration: none; font-weight: bold;">⬅ Volver al inicio</a>
            </body>
            </html>
            """, 200

        generar_anexos_docx(
            pacientes,
            str(temp_docx),
            hospital_nombre=hosp_info["nombre"],
            hospital_refes=hosp_info["refes"]
        )

        # Liberar archivo temporal y memoria RAM
        try:
            if temp_excel.exists():
                temp_excel.unlink()
        except Exception:
            pass
        import gc
        gc.collect()

        # Nombre seguro y ASCII para encabezados HTTP
        raw_stem = Path(archivo.filename).stem
        stem_ascii = unicodedata.normalize('NFKD', raw_stem).encode('ASCII', 'ignore').decode('ASCII')
        stem_seguro = re.sub(r'[^a-zA-Z0-9_\-]', '_', stem_ascii).strip('_') or "pacientes"
        nombre_descarga = f"Anexos_II_{stem_seguro}.docx"

        return send_file(
            str(temp_docx),
            as_attachment=True,
            download_name=nombre_descarga,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    except Exception as e:
        import traceback
        traceback.print_exc()
        err_msg = html.escape(str(e))
        return f"""
        <!DOCTYPE html>
        <html>
        <head><title>Error</title><meta charset="utf-8"></head>
        <body style="font-family: sans-serif; padding: 40px; background: #0f172a; color: #f8fafc;">
            <h2 style="color: #ef4444;">❌ Error al procesar el archivo Excel para Anexos II</h2>
            <div style="background: #1e293b; padding: 15px; border-radius: 8px; border: 1px solid #334155; margin: 20px 0; font-family: monospace; font-size: 14px; color: #fca5a5;">
                {err_msg}
            </div>
            <p style="color: #94a3b8;">Revisá que el Excel tenga las columnas de DNI, Paciente y Obra Social o volvé a intentar.</p>
            <a href="/" style="display: inline-block; padding: 10px 20px; background: #6366f1; color: white; border-radius: 6px; text-decoration: none; font-weight: bold;">⬅ Volver al inicio</a>
        </body>
        </html>
        """, 200


# ============================================================
# GENERADOR DIRECTO DE ANEXOS II EN PDF (MASIVO IMPRIMIBLE)
# ============================================================

@app.route("/generar_anexos_pdf", methods=["POST"])
def generar_anexos_pdf_directo():
    # El usuario prefiere exclusivamente el formato oficial idéntico en Word (.docx)
    return generar_anexos_directo()


# ============================================================
# IMPRESIÓN DIRECTA DESDE EXCEL SUBIDO (VISTA WEB CTRL+P)
# ============================================================

@app.route("/imprimir_anexos_directo", methods=["POST"])
def imprimir_anexos_directo():
    try:
        archivo = request.files.get("archivo")
        if not archivo or not archivo.filename:
            return "No se subió ningún archivo.", 200
        if not archivo.filename.lower().endswith(".xlsx"):
            return "El archivo debe ser un Excel (.xlsx)", 200

        identificador = uuid.uuid4().hex
        temp_excel = UPLOAD_DIR / f"anexo_print_{identificador}.xlsx"
        archivo.save(str(temp_excel))

        esp_req = request.form.get("especialidad", "").strip() or "CARDIOLOGIA"
        hosp_id = request.form.get("hospital", "san_roque")
        hosp_custom_nom = request.form.get("hospital_nombre_custom", "").strip()
        hosp_custom_ref = request.form.get("hospital_refes_custom", "").strip()
        hosp_info = resolver_hospital(hosp_id, hosp_custom_nom, hosp_custom_ref)

        pacientes = extraer_pacientes_afiliados_excel(str(temp_excel), especialidad_defecto=esp_req)
        if not pacientes:
            return """
            <h2>No se encontraron pacientes afiliados con Obra Social válida en el Excel.</h2>
            <a href="/">Volver al inicio</a>
            """, 200

        hoy = date.today()
        fecha_act_str = hoy.strftime("%d-%m-%Y")
        primer_dia_este_mes = hoy.replace(day=1)
        ultimo_dia_mes_ant = primer_dia_este_mes - timedelta(days=1)
        per_dec_str = ultimo_dia_mes_ant.strftime("%m-%Y")

        return render_template(
            "imprimir_anexos.html",
            pacientes=pacientes,
            total_pacientes=len(pacientes),
            hospital_nombre=hosp_info["nombre"],
            hospital_refes=hosp_info["refes"],
            fecha_actual=fecha_act_str,
            periodo_declarado=per_dec_str
        )
    except Exception as e:
        import traceback
        traceback.print_exc()
        return f"<h2>Error preparando fojas para impresión:</h2><p>{html.escape(str(e))}</p><a href='/'>Volver</a>", 200


# ============================================================
# WEBHOOK WHATSAPP — VERIFICACIÓN (GET)
# ============================================================

@app.route("/webhook", methods=["GET"])
def webhook_verificar():

    challenge, status = verificar_webhook(request.args)
    return challenge, status


# ============================================================
# WEBHOOK WHATSAPP — MENSAJES ENTRANTES (POST)
# ============================================================

@app.route("/webhook", methods=["POST"])
def webhook_recibir():

    payload = request.get_json(silent=True) or {}

    mensaje = extraer_mensaje(payload)

    if mensaje:

        numero     = mensaje["numero"]
        tipo       = mensaje.get("tipo", "text")
        message_id = mensaje.get("message_id")

        if message_id:
            marcar_leido(message_id)

        # ── CASO 1: Mensaje de texto ────────────────────────
        if tipo == "text":
            texto = mensaje.get("texto", "")

            def responder_texto_async():
                respuesta = responder(texto)
                enviar_texto(numero, respuesta)

            threading.Thread(target=responder_texto_async, daemon=True).start()

        # ── CASO 2: Foja quirúrgica (imagen o PDF) ──────────
        elif tipo in ["image", "document"]:
            media_id = mensaje.get("media_id")
            filename = mensaje.get("filename") or ("foja.jpg" if tipo == "image" else "foja.pdf")

            def procesar_foja_whatsapp_async():
                try:
                    enviar_texto(
                        numero,
                        "⏳ Recibí la foja quirúrgica. La estoy analizando con el nomenclador oficial e instructivos..."
                    )

                    contenido_bytes, mime = descargar_media(media_id)
                    if not contenido_bytes:
                        enviar_texto(
                            numero,
                            "❌ No se pudo descargar el archivo de WhatsApp. Por favor volvé a enviarlo."
                        )
                        return

                    ext = Path(filename).suffix.lower()
                    if not ext:
                        ext = ".png" if tipo == "image" else ".pdf"

                    tmp_path = FOJAS_DIR / f"wa_{uuid.uuid4().hex}{ext}"
                    tmp_path.write_bytes(contenido_bytes)

                    try:
                        resultado = analizar_foja_quirurgica(tmp_path)
                        if resultado.get("ok"):
                            enviar_texto(
                                numero,
                                resultado.get("analisis", "Análisis completado.")
                            )
                        else:
                            enviar_texto(
                                numero,
                                f"❌ Error en análisis: {resultado.get('error', 'No se pudo interpretar la foja.')}"
                            )
                    finally:
                        tmp_path.unlink(missing_ok=True)

                except Exception as ex:
                    print(f"Error procesando foja WhatsApp: {ex}")
                    enviar_texto(
                        numero,
                        "❌ Ocurrió un error inesperado al analizar el documento."
                    )

            threading.Thread(target=procesar_foja_whatsapp_async, daemon=True).start()

    # Meta requiere siempre un 200 rápido
    return jsonify({"status": "ok"}), 200


# ============================================================
# Estado de indexado en memoria
_estado_indexado = {}


BOT_PASSWORD = "Lupeycoca2026"


def _verificar_clave_bot():
    """Verifica la contraseña del Bot. Retorna True si es válida."""
    clave = request.form.get("clave") or request.args.get("clave") or ""
    return clave == BOT_PASSWORD


@app.route("/api/upload_docs", methods=["POST"])
def upload_documento():

    if not _verificar_clave_bot():
        return jsonify({"ok": False, "error": "Contraseña incorrecta."}), 403

    archivo = request.files.get("documento")

    if not archivo or archivo.filename == "":
        return jsonify({"ok": False, "error": "No se recibió ningún archivo."}), 400

    extensiones_permitidas = {".pdf", ".txt", ".docx", ".doc", ".xlsx", ".xls"}
    ext = Path(archivo.filename).suffix.lower()

    if ext not in extensiones_permitidas:
        return jsonify({
            "ok": False,
            "error": "Formato no permitido. Usá: PDF, TXT, DOCX o XLSX."
        }), 400

    destino = DOCS_DIR / archivo.filename
    archivo.save(destino)

    nombre = archivo.filename
    _estado_indexado[nombre] = {"estado": "procesando", "fragmentos": 0}

    # Indexar en segundo plano para no bloquear la respuesta HTTP
    def _indexar():
        try:
            n = agregar_documento_a_base(destino)
            _estado_indexado[nombre] = {"estado": "listo", "fragmentos": n}
        except Exception as e:
            _estado_indexado[nombre] = {"estado": "error", "error": str(e)}

    threading.Thread(target=_indexar, daemon=True).start()

    # Respuesta inmediata — el cliente puede consultar /api/estado_indexado/<nombre>
    return jsonify({
        "ok":      True,
        "nombre":  nombre,
        "mensaje": "Archivo recibido. Indexando en segundo plano...",
    })


@app.route("/api/estado_indexado/<nombre>")
def estado_indexado(nombre):
    estado = _estado_indexado.get(nombre, {"estado": "desconocido"})
    return jsonify(estado)


# ============================================================
# LISTAR DOCUMENTOS DEL BOT
# ============================================================

@app.route("/api/docs", methods=["GET"])
def listar_docs():
    if not _verificar_clave_bot():
        return jsonify({"ok": False, "error": "Contraseña incorrecta."}), 403
    return jsonify(listar_documentos())


# ============================================================
# ELIMINAR DOCUMENTO DEL BOT
# ============================================================

@app.route("/api/delete_doc", methods=["POST"])
def eliminar_doc():
    if not _verificar_clave_bot():
        return jsonify({"ok": False, "error": "Contraseña incorrecta."}), 403

    nombre = request.form.get("nombre", "").strip()
    if not nombre:
        return jsonify({"ok": False, "error": "No se indicó el nombre del archivo."}), 400

    # Seguridad: evitar path traversal
    ruta = (DOCS_DIR / nombre).resolve()
    if not str(ruta).startswith(str(DOCS_DIR.resolve())):
        return jsonify({"ok": False, "error": "Nombre de archivo inválido."}), 400

    if not ruta.exists():
        return jsonify({"ok": False, "error": "Archivo no encontrado."}), 404

    try:
        ruta.unlink()
        return jsonify({"ok": True, "mensaje": f"{nombre} eliminado correctamente."})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


# ============================================================
# ANÁLISIS DE FOJA QUIRÚRGICA
# ============================================================

FOJAS_DIR = BASE_DIR / "fojas_tmp"
FOJAS_DIR.mkdir(exist_ok=True)

def guardar_estado_foja(job_id, datos):
    try:
        ruta = FOJAS_DIR / f"{job_id}.json"
        ruta.write_text(json.dumps(datos, ensure_ascii=False), encoding="utf-8")
    except Exception as e:
        print(f"Error guardando estado foja: {e}")

def obtener_estado_foja(job_id):
    ruta = FOJAS_DIR / f"{job_id}.json"
    if ruta.exists():
        try:
            return json.loads(ruta.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {"estado": "procesando"}


@app.route("/api/analizar_foja", methods=["POST"])
def analizar_foja():

    archivo = request.files.get("foja")

    if not archivo or archivo.filename == "":
        return jsonify({"ok": False, "error": "No se recibió ningún archivo."}), 400

    EXTENSIONES_OK = {".pdf", ".jpg", ".jpeg", ".png", ".webp", ".gif"}
    ext = Path(archivo.filename).suffix.lower()

    if ext not in EXTENSIONES_OK:
        return jsonify({
            "ok": False,
            "error": "Formato no soportado. Usá PDF, JPG, PNG o WEBP."
        }), 400

    # Guardar temporalmente con extensión original para que bot_rag la detecte
    job_id     = uuid.uuid4().hex
    nombre_tmp = f"{job_id}{ext}"
    ruta_tmp   = FOJAS_DIR / nombre_tmp
    archivo.save(ruta_tmp)

    guardar_estado_foja(job_id, {"estado": "procesando"})

    def _analizar():
        try:
            resultado = analizar_foja_quirurgica(ruta_tmp)
            guardar_estado_foja(job_id, {"estado": "listo", **resultado})
        except Exception as e:
            guardar_estado_foja(job_id, {"estado": "error", "ok": False, "error": str(e)})
        finally:
            try:
                ruta_tmp.unlink(missing_ok=True)
            except Exception:
                pass

    threading.Thread(target=_analizar, daemon=True).start()

    # Respuesta inmediata: el cliente hace polling con el job_id
    return jsonify({"ok": True, "job_id": job_id, "estado": "procesando"})


@app.route("/api/estado_foja/<job_id>")
def estado_foja(job_id):
    estado = obtener_estado_foja(job_id)
    return jsonify(estado)



# ============================================================
# CHAT CON RAG (CONSULTA A INSTRUCTIVOS Y DOCUMENTOS)
# ============================================================

@app.route("/api/chat", methods=["POST"])
def chat():
    datos = request.get_json(silent=True) or {}
    pregunta = datos.get("pregunta", "").strip()

    if not pregunta:
        return jsonify({"ok": False, "error": "La pregunta no puede estar vacía."}), 400

    try:
        respuesta_texto = responder(pregunta)
        return jsonify({"ok": True, "respuesta": respuesta_texto})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


# ============================================================
# MANEJADORES GLOBALES DE ERROR (PREVENIR PANTALLAS BLANCAS 500)
# ============================================================

@app.errorhandler(500)
@app.errorhandler(Exception)
def error_global_servidor(e):
    import traceback
    traceback.print_exc()
    tb_txt = html.escape(traceback.format_exc())
    err_str = html.escape(str(e))
    return f"""
    <!DOCTYPE html>
    <html>
    <head><title>Aviso del Sistema</title><meta charset="utf-8"></head>
    <body style="font-family: sans-serif; padding: 40px; background: #0f172a; color: #f8fafc;">
        <h2 style="color: #ef4444;">❌ Se produjo una incidencia en el servidor</h2>
        <div style="background: #1e293b; padding: 15px; border-radius: 8px; border: 1px solid #334155; margin: 20px 0; font-family: monospace; font-size: 14px; color: #fca5a5;">
            {err_str}
        </div>
        <details style="margin: 15px 0;">
            <summary style="cursor: pointer; color: #94a3b8; font-size: 13px;">Ver detalle técnico del error</summary>
            <pre style="background: #090d16; padding: 12px; border-radius: 6px; font-size: 12px; overflow-x: auto; color: #cbd5e1; margin-top: 10px;">{tb_txt}</pre>
        </details>
        <a href="/" style="display: inline-block; padding: 10px 20px; background: #6366f1; color: white; border-radius: 6px; text-decoration: none; font-weight: bold;">⬅ Volver al inicio</a>
    </body>
    </html>
    """, 200


# ============================================================
# EJECUTAR
# ============================================================

if __name__ == "__main__":

    app.run(

        host="127.0.0.1",

        port=5000,

        debug=False
    )