import os
import re
import time
import uuid
import threading
from pathlib import Path
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed

import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
import sqlite3
import queue

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

import openpyxl
from flask import Flask, request, render_template, send_file, jsonify

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
MAX_WORKERS = int(os.environ.get("SSS_MAX_WORKERS", 14))
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
        with sqlite3.connect(CACHE_DB_PATH) as conn:
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
        with sqlite3.connect(CACHE_DB_PATH) as conn:
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


# ============================================================
# CLIENTE SSSALUD ULTRARRÁPIDO & PARSER
# ============================================================

def parsear_respuesta_sss(html_text):
    if not html_text:
        return "ERROR CONSULTA"

    # 1. NO AFILIADO
    if (
        "No se reportan datos para el NUMERO DE DOCUMENTO" in html_text
        or "NO AFILIADO" in html_text.upper()
    ) and "DATOS DE AFILIACION VIGENTE" not in html_text:
        return "NO AFILIADO"

    # 2. AFILIADO VIGENTE
    if "DATOS DE AFILIACION VIGENTE" in html_text:
        codigo = None
        denominacion = None

        m_cod = re.search(
            r'C(?:ó|&oacute;|o)digo\s+de\s+Obra\s+Social.*?<td[^>]*>(?:<[^>]+>)*\s*([^<]+?)\s*<',
            html_text,
            re.I | re.DOTALL
        )
        if m_cod:
            codigo = m_cod.group(1).strip()

        m_den = re.search(
            r'Denominaci(?:ó|&oacute;|o)n\s+Obra\s+Social.*?<td[^>]*>(?:<[^>]+>)*\s*([^<]+?)\s*<',
            html_text,
            re.I | re.DOTALL
        )
        if m_den:
            denominacion = m_den.group(1).strip()

        # Fallback a BeautifulSoup si el regex no capturó ambos
        if not (codigo and denominacion):
            try:
                from bs4 import BeautifulSoup
                soup = BeautifulSoup(html_text, 'html.parser')
                for tr in soup.find_all('tr'):
                    tds = [td.get_text(strip=True) for td in tr.find_all(['td', 'th'])]
                    if len(tds) >= 2:
                        k, v = tds[0], tds[1]
                        k_lower = k.lower()
                        if "código de obra social" in k_lower or "codigo de obra social" in k_lower:
                            if not codigo and v:
                                codigo = v
                        elif "denominación obra social" in k_lower or "denominacion obra social" in k_lower:
                            if not denominacion and v:
                                denominacion = v
            except Exception:
                pass

        if codigo and denominacion:
            return f"{codigo} - {denominacion}"
        elif denominacion:
            return str(denominacion)
        elif codigo:
            return str(codigo)
        return "AFILIADO - SIN OBRA SOCIAL IDENTIFICADA"

    # Verificación secundaria
    if "No se reportan datos para el NUMERO DE DOCUMENTO" in html_text:
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
    """Instancia de sesión HTTP independiente con su propia cookie PHPSESSID"""
    LOGIN_URL = 'https://seguro.sssalud.gob.ar/login.php?b_publica=Acceso+Restringido+para+Hospitales&opc=bus650&user=HPGD'

    def __init__(self, user, password, idx):
        self.user = user
        self.password = password
        self.idx = idx
        self.session = requests.Session()
        adapter = HTTPAdapter(
            pool_connections=6,
            pool_maxsize=6,
            max_retries=Retry(total=2, backoff_factor=0.2, status_forcelist=[500, 502, 503, 504])
        )
        self.session.mount('https://', adapter)
        self.session.mount('http://', adapter)
        self.session.headers.update({
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
        })
        self.logged_in = False
        self.lock = threading.Lock()

    def login(self, force=False):
        with self.lock:
            if self.logged_in and not force:
                return True
            try:
                params = {
                    '_user_name_': self.user,
                    '_pass_word_': self.password,
                    'submitbtn': 'Ingresar'
                }
                res = self.session.post(self.LOGIN_URL, data=params, verify=False, timeout=12)
                if 'usuario_logueado' in res.text or 'nro_doc' in res.text or 'cat=consultas' in res.text:
                    self.logged_in = True
                    return True
            except Exception as e:
                print(f"Error en login sesión {self.idx}: {e}")
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
            worker.login()
            self.session_pool.put(worker)

        # Fallback usando la librería tradicional si está instalada
        self._fallback_sss = None
        if DataBeneficiariosSSSHospital is not None:
            try:
                self._fallback_sss = DataBeneficiariosSSSHospital(user=user, password=password)
                self._fallback_sss.pause_before_requests = 0
                self._fallback_sss._save_response = lambda filename, resp: None
            except Exception:
                pass

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

        worker = self.session_pool.get()
        try:
            for intento in range(2):
                if not worker.logged_in:
                    worker.login(force=True)

                try:
                    res = worker.session.post(self.QUERY_URL, data=params, verify=False, timeout=10)
                    text = res.text

                    if res.status_code in [401, 403, 429, 500, 502, 503, 504] or ('Ingresar' in text and '_user_name_' in text):
                        time.sleep(0.3 * (intento + 1))
                        worker.login(force=True)
                        continue

                    resultado = parsear_respuesta_sss(text)
                    if resultado != "ERROR CONSULTA":
                        with CACHE_LOCK:
                            CACHE_DNI[dni_str] = resultado
                        return resultado

                    # Fallback opcional si el parseo dio error
                    if self._fallback_sss:
                        try:
                            res_fb = self._fallback_sss.query(dni_str)
                            if res_fb.get("ok"):
                                resultado = parsear_datos_fallback(res_fb.get("resultados", {}))
                                if resultado != "ERROR CONSULTA":
                                    with CACHE_LOCK:
                                        CACHE_DNI[dni_str] = resultado
                                    return resultado
                        except Exception:
                            pass

                except Exception as e:
                    print(f"Error consultando DNI {dni_str} en worker {worker.idx}: {e}")
                    time.sleep(0.3)
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
    archivo_salida
):
    estado = procesos[id_proceso]

    try:
        # ----------------------------------------------------
        # 1. ABRIR EXCEL
        # ----------------------------------------------------
        estado["estado"] = "abriendo_excel"
        wb = openpyxl.load_workbook(archivo_entrada)
        ws = wb.active
        ultima_fila = ws.max_row

        # ----------------------------------------------------
        # 2. DETERMINAR COLUMNA DESTINO DE OBRA SOCIAL
        # ----------------------------------------------------
        col_obra_social = COLUMNA_OBRA_SOCIAL
        fila_encabezado = max(1, FILA_INICIO - 1)
        encontrada = False

        # Si en los encabezados ya existe una columna de obra social
        for c in range(1, 20):
            val_h = str(ws.cell(fila_encabezado, c).value or "").strip().lower()
            if any(k in val_h for k in ["obra social", "cobertura", "prepaga", "afiliacion", "o.s."]):
                col_obra_social = c
                encontrada = True
                break

        # Si no existe, y la columna 7 actual está ocupada (ej. 'Fecha ingreso'), agregar columna a la derecha
        if not encontrada:
            val_col7 = str(ws.cell(fila_encabezado, COLUMNA_OBRA_SOCIAL).value or "").strip().lower()
            if val_col7 and not any(k in val_col7 for k in ["obra social", "cobertura", "prepaga"]):
                max_col = max([c for c in range(1, 30) if ws.cell(fila_encabezado, c).value is not None] or [COLUMNA_OBRA_SOCIAL])
                col_obra_social = max_col + 1
                try:
                    ws.cell(fila_encabezado, col_obra_social).value = "Obra Social (SSSalud)"
                except Exception:
                    pass

        # ----------------------------------------------------
        # 3. BUSCAR FILAS CON DNI VÁLIDO Y AGRUPAR
        # ----------------------------------------------------
        dni_a_filas = defaultdict(list)
        total_filas = 0

        for fila in range(FILA_INICIO, ultima_fila + 1):
            dni_val = ws.cell(fila, COLUMNA_DNI).value
            dni_limpio = limpiar_dni(dni_val)
            if not dni_limpio:
                continue
            dni_a_filas[dni_limpio].append(fila)
            total_filas += 1

        # Fallback si las filas empezaban antes de FILA_INICIO
        if total_filas == 0 and ultima_fila >= 2:
            for fila in range(2, min(FILA_INICIO, ultima_fila + 1)):
                dni_val = ws.cell(fila, COLUMNA_DNI).value
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

        print("")
        print("====================================================")
        print("       PROCESAMIENTO SSSALUD ULTRARRÁPIDO")
        print("====================================================")
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

        dnis_pendientes = [d for d in dnis_unicos if d not in resultados_dni]
        print(f"DNIs resueltos al instante por Base de Datos/Caché: {len(resultados_dni)} ({cant_resueltos_db} filas)")
        print(f"DNIs pendientes de consulta remota: {len(dnis_pendientes)}")

        # ----------------------------------------------------
        # 5. SI HAY PENDIENTES, CONSULTAR CON SESIONES PARALELAS
        # ----------------------------------------------------
        nuevos_para_guardar = []

        if dnis_pendientes:
            estado["estado"] = "conectando_sssalud"
            sss = crear_conexion_sss()
            estado["estado"] = "consultando"

            lock_estado = threading.Lock()

            def procesar_un_dni(dni):
                time.sleep(0.03)  # Pequeño espaciado para distribuir la carga entre sesiones
                res = consultar_dni(sss, dni)
                filas = dni_a_filas[dni]
                cant = len(filas)

                with lock_estado:
                    resultados_dni[dni] = res
                    nuevos_para_guardar.append((dni, res))
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

                return dni, res

            with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
                futuros = [executor.submit(procesar_un_dni, d) for d in dnis_pendientes]
                for f in as_completed(futuros):
                    try:
                        f.result()
                    except Exception as ex_hilo:
                        print(f"Error en hilo de consulta: {ex_hilo}")

            # Guardar los nuevos en la base de datos persistente SQLite
            if nuevos_para_guardar:
                guardar_cache_multiples(nuevos_para_guardar)

        # ----------------------------------------------------
        # 6. ESCRIBIR RESULTADOS EN EXCEL (CON PROTECCIÓN MERGEDCELL)
        # ----------------------------------------------------
        print("Escribiendo resultados en memoria...")
        for dni, filas in dni_a_filas.items():
            res_dni = resultados_dni.get(dni, "ERROR CONSULTA")
            for fila in filas:
                try:
                    cell = ws.cell(fila, col_obra_social)
                    if type(cell).__name__ == "MergedCell":
                        continue
                    cell.value = res_dni
                except (AttributeError, Exception) as err_cell:
                    print(f"No se pudo escribir en fila {fila}: {err_cell}")

        # ----------------------------------------------------
        # 7. GUARDADO FINAL ÚNICO
        # ----------------------------------------------------
        estado["estado"] = "guardando"
        wb.save(archivo_salida)

        # ----------------------------------------------------
        # 8. FINALIZADO
        # ----------------------------------------------------
        estado["estado"] = "terminado"
        estado["procesadas"] = total
        estado["porcentaje"] = 100
        estado["archivo"] = Path(archivo_salida).name
        estado["segundos"] = round(time.time() - inicio, 1)
        estado["estimado_restante"] = 0

        print("")
        print("====================================================")
        print("          PROCESAMIENTO TERMINADO EXITOSAMENTE")
        print("====================================================")
        print(f"Tiempo total: {estado['segundos']} segundos")
        print("Registros procesados:", estado["procesadas"])
        print("Afiliados:", estado["afiliados"])
        print("No afiliados:", estado["no_afiliados"])
        print("Errores:", estado["errores"])
        print("====================================================")

    except Exception as e:
        print("")
        print("====================================================")
        print("ERROR GENERAL EN PROCESAMIENTO")
        print("====================================================")
        print(e)
        estado["estado"] = "error"
        estado["error"] = str(e)


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
    # CREAR ESTADO
    # --------------------------------------------------------

    procesos[identificador] = {

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

        "error": ""
    }

    # --------------------------------------------------------
    # INICIAR HILO
    # --------------------------------------------------------

    hilo = threading.Thread(

        target=procesar_archivo,

        args=(
            identificador,
            archivo_entrada,
            archivo_salida
        ),

        daemon=True
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
# IMPORTAR PADRÓN PUCO O BASE PREVIA (INSTANTÁNEO)
# ============================================================

@app.route("/importar_padron", methods=["POST"])
def importar_padron():
    archivo = request.files.get("archivo")
    if not archivo or not archivo.filename:
        return jsonify({"ok": False, "error": "No se seleccionó archivo"}), 400

    nombre = archivo.filename.lower()
    registros_guardados = 0

    try:
        if nombre.endswith(".xlsx"):
            wb = openpyxl.load_workbook(archivo, data_only=True)
            ws = wb.active
            fila_inicio = 1
            col_dni = 1
            col_os = 2

            for r in range(1, 6):
                for c in range(1, 15):
                    val = str(ws.cell(r, c).value or "").lower()
                    if "dni" in val or "documento" in val:
                        col_dni = c
                        fila_inicio = r + 1
                    elif "obra social" in val or "cobertura" in val or "prepaga" in val:
                        col_os = c

            items = []
            for r in range(fila_inicio, ws.max_row + 1):
                d = limpiar_dni(ws.cell(r, col_dni).value)
                os_val = str(ws.cell(r, col_os).value or "").strip()
                if d and os_val and os_val.lower() not in ["none", "null", ""]:
                    items.append((d, os_val))
                    if len(items) >= 1000:
                        guardar_cache_multiples(items)
                        registros_guardados += len(items)
                        items = []
            if items:
                guardar_cache_multiples(items)
                registros_guardados += len(items)

        elif nombre.endswith(".csv") or nombre.endswith(".txt"):
            contenido = archivo.read().decode("utf-8", errors="ignore")
            lineas = contenido.splitlines()
            items = []
            for l in lineas:
                partes = re.split(r'[,;\t|]', l)
                if len(partes) >= 2:
                    d = limpiar_dni(partes[0])
                    os_val = partes[1].strip()
                    if d and os_val and os_val.lower() not in ["none", "null", ""]:
                        items.append((d, os_val))
                        if len(items) >= 1000:
                            guardar_cache_multiples(items)
                            registros_guardados += len(items)
                            items = []
            if items:
                guardar_cache_multiples(items)
                registros_guardados += len(items)
        else:
            return jsonify({"ok": False, "error": "Formato no soportado. Usá .xlsx, .csv o .txt"}), 400

        return jsonify({
            "ok": True,
            "mensaje": f"Se importaron {registros_guardados} registros en la base local permanente. Las consultas de estos pacientes ahora serán instantáneas."
        })

    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


# ============================================================
# CONSULTAR PROGRESO
# ============================================================

@app.route(
    "/progreso/<id_proceso>"
)
def progreso(id_proceso):

    estado = procesos.get(
        id_proceso
    )

    if not estado:

        return jsonify({

            "estado": "error",

            "error": "Proceso no encontrado"

        }), 404

    return jsonify(
        estado
    )


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

# Estado en memoria de análisis de fojas (job_id -> resultado)
_estado_fojas = {}


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

    _estado_fojas[job_id] = {"estado": "procesando"}

    def _analizar():
        try:
            resultado = analizar_foja_quirurgica(ruta_tmp)
            _estado_fojas[job_id] = {"estado": "listo", **resultado}
        except Exception as e:
            _estado_fojas[job_id] = {"estado": "error", "ok": False, "error": str(e)}
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
    estado = _estado_fojas.get(job_id, {"estado": "no_encontrado"})
    return jsonify(estado)



# ============================================================
# EJECUTAR
# ============================================================

if __name__ == "__main__":

    app.run(

        host="127.0.0.1",

        port=5000,

        debug=False
    )