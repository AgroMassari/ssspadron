import re
import zipfile
import xml.sax.saxutils as saxutils
from pathlib import Path
from datetime import datetime, date, timedelta
import openpyxl

PLANTILLA_DOCX = Path(__file__).resolve().parent / "plantilla_anexo_ii.docx"


def escapar_xml(texto):
    if texto is None:
        return ""
    return saxutils.escape(str(texto))


# Catálogo de Hospitales HPGD configurados
CATALOGO_HOSPITALES = {
    "san_roque": {
        "id": "san_roque",
        "nombre": "NUEVO HOSPITAL SAN ROQUE",
        "refes": "10140142131256",
        "descripcion": "Nuevo Hospital San Roque (Córdoba)",
    },
    "pasteur": {
        "id": "pasteur",
        "nombre": "HOSPITAL REGIONAL LOUIS PASTEUR",
        "refes": "10140422131251",
        "descripcion": "Hospital Regional Louis Pasteur (Villa María)",
    },
}

HOSPITAL_DEFECTO = CATALOGO_HOSPITALES["san_roque"]


def resolver_hospital(hospital_id=None, custom_nombre=None, custom_refes=None):
    """
    Resuelve el nombre y código REFES del hospital a partir de un ID o valores personalizados.
    Por defecto utiliza Nuevo Hospital San Roque.
    """
    if hospital_id == "custom" and custom_nombre:
        return {
            "id": "custom",
            "nombre": str(custom_nombre).strip().upper(),
            "refes": str(custom_refes or "").strip(),
            "descripcion": str(custom_nombre).strip().upper(),
        }
    if hospital_id and hospital_id in CATALOGO_HOSPITALES:
        return CATALOGO_HOSPITALES[hospital_id]
    if custom_nombre:
        return {
            "id": "custom",
            "nombre": str(custom_nombre).strip().upper(),
            "refes": str(custom_refes or "").strip(),
            "descripcion": str(custom_nombre).strip().upper(),
        }
    return HOSPITAL_DEFECTO


# Mapeo oficial de las principales Obras Sociales de Argentina a su código RNOS / RNAS (6 dígitos)
TABLA_RNOS_OFICIAL = [
    # Metales / UOM
    ("126205", ["METALURGICA", "METALURGICO", "UOM", "UNION OBRERA METALURGICA", "OBRERA METALURGICA"]),
    # Comercio / OSECAC
    ("126304", ["COMERCIO", "OSECAC", "ACTIVIDADES CIVILES", "MERCANTILES", "MERCANTIL"]),
    # PAMI / INSSJP
    ("500807", ["PAMI", "INSSJP", "JUBILADOS Y PENSIONADOS", "SERVICIOS SOCIALES PARA JUBILADOS"]),
    # OSDE
    ("400800", ["OSDE", "ORGANIZACION DE SERVICIOS DIRECTOS"]),
    # Sanidad / OSPSA
    ("112000", ["SANIDAD", "OSPSA", "PERSONAL DE LA SANIDAD", "SANATORIOS"]),
    # Construcción / UOCRA
    ("105404", ["CONSTRUCCION", "CONSTRUCCIÓN", "UOCRA", "OBREROS DE LA CONSTRUCCION"]),
    # Camioneros / OSCHOCA
    ("104001", ["CAMIONEROS", "CAMIONERO", "CHOFERES DE CAMIONES", "OSCHOCA", "TRANSPORTE DE CARGAS"]),
    # Rurales / OSPRERA
    ("119701", ["RURAL", "RURALES", "ESTIBADORES", "OSPRERA", "UATRE"]),
    # Gastronómicos / OSUTHGRA
    ("108801", ["GASTRONOMICO", "GASTRONÓMICO", "TURISMO, HOTELERO", "HOTELEROS", "OSUTHGRA", "UTHGRA"]),
    # UPCN (Personal Civil de la Nación)
    ("119800", ["CIVIL DE LA NACION", "CIVIL DE LA NACIÓN", "UPCN", "PERSONAL CIVIL"]),
    # Unión Personal / Accord Salud
    ("128003", ["UNION PERSONAL", "UNIÓN PERSONAL", "ACCORD"]),
    # Mecánicos / SMATA
    ("113706", ["MECANICOS", "MECÁNICOS", "SMATA", "OSMATA"]),
    # Seguros / OSSEG
    ("121304", ["SEGUROS", "SEGURO", "OSSEG"]),
    # Casas Particulares / OSPACP (Empleadas domésticas)
    ("128508", ["CASAS PARTICULARES", "OSPACP", "DOMESTICA", "DOMÉSTICA"]),
    # Bancarios / OSBA
    ("121106", ["BANCARIA", "BANCARIO", "BANCARIOS", "OSBA"]),
    # OSDEPYM
    ("125103", ["OSDEPYM", "EMPRESARIOS, PROFESIONALES"]),
    # Alimentación / STIA / OSPIA
    ("101600", ["ALIMENTACION", "ALIMENTACIÓN", "OSPIA", "STIA"]),
    # Carne
    ("104407", ["CARNE", "INDUSTRIA DE LA CARNE"]),
    # Químicos
    ("117005", ["QUIMICOS", "QUÍMICOS", "PETROQUIMICOS", "PETROQUÍMICOS"]),
    # Plásticos / UOYEP
    ("115801", ["PLASTICO", "PLÁSTICO", "PLASTICOS", "PLÁSTICOS", "UOYEP"]),
    # Panaderos / FAUPPA
    ("114006", ["PANADEROS", "PANADERIA", "PANADERÍA", "FAUPPA"]),
    # Textiles / SETIA / AOT
    ("119909", ["TEXTIL", "TEXTILES", "SETIA", "AOT"]),
    # Madera
    ("111009", ["MADERA", "MADEREROS"]),
    # Aguas / SGBATOS
    ("100508", ["AGUA POTABLE", "AGUAS", "SGBATOS", "OSDAS"]),
    # Gráficos
    ("109002", ["GRAFICO", "GRÁFICO", "GRAFICOS", "GRÁFICOS"]),
    # Porteros / Edificios / SUTERH / OSPERYH
    ("107400", ["EDIFICIO", "EDIFICIOS", "SUTERH", "OSPERYH", "RENTA Y HORIZONTAL"]),
    # Ferroviarios / Trenes / OSFE
    ("120505", ["FERROVIARIO", "FERROVIARIOS", "OSFE"]),
    # UTA / Colectiveros / Automotor
    ("120208", ["TRANSPORTE AUTOMOTOR", "UTA", "COLECTIVEROS"]),
    # Luz y Fuerza / Electricistas
    ("108007", ["LUZ Y FUERZA", "ELECTRICISTAS"]),
    # Televisión / SATSAID
    ("119506", ["TELEVISION", "TELEVISIÓN", "SATSAID", "OSTV"]),
    # Telecomunicaciones / FOETRA
    ("119407", ["TELECOMUNICACIONES", "FOETRA", "OSTEL"]),
    # Petroleros
    ("115504", ["PETROLERO", "PETROLEROS", "PETROLEO", "PETRÓLEO", "SUPEH"]),
    # Maestranza / OSPM
    ("111306", ["MAESTRANZA", "OSPM"]),
    # Pasteleros y Confiteros
    ("106209", ["CONFITEROS", "PASTELEROS", "PIZZEROS", "ALFAJOREROS"]),
    # Calzado
    ("103804", ["CALZADO", "OSCRA"]),
    # Vidrio
    ("120703", ["VIDRIO", "SEIVARA"]),
    # Papeleros
    ("114204", ["PAPEL", "PAPELEROS"]),
    # Cerveceros
    ("104803", ["CERVECEROS", "CERVECERA"]),
    # Cuero
    ("107004", ["CUERO", "CURTIDORES"]),
    # Docentes Privados / SADOP / OSDOP
    ("107202", ["DOCENTES PARTICULARES", "DOCENTES PRIVADOS", "SADOP", "OSDOP"]),
    # Universidades / OSFATUN
    ("124007", ["UNIVERSIDADES", "NO DOCENTE", "OSFATUN"]),
    # Personal de Dirección / OSDE / ACCORD / LUIS PASTEUR
    ("400404", ["LUIS PASTEUR", "DIRECCION DE EMPRESAS", "PERSONAL DE DIRECCION"]),
    ("400107", ["ASE", "ACCION SOCIAL DE EMPRESARIOS"]),
    ("126007", ["SWISS MEDICAL", "DOCTHOS"]),
    ("126502", ["MEDICUS"]),
    ("126700", ["GALENO"]),
    ("127000", ["OMINT"]),
    ("127208", ["SANCOR SALUD", "SANCOR"]),
    # Sector Público
    ("500104", ["POLICIA FEDERAL", "POLICÍA FEDERAL", "BIENESTAR"]),
    ("500203", ["PODER JUDICIAL", "CORTE SUPREMA", "OSPJN"]),
    ("800109", ["IOSFA", "FUERZAS ARMADAS", "EJERCITO", "ARMADA", "GENDARMERIA", "PREFECTURA"]),
    ("900100", ["IOMA", "PROVINCIA DE BUENOS AIRES"]),
    ("900200", ["APROSS", "CORDOBA", "CÓRDOBA"]),
    ("900300", ["OSEP", "MENDOZA"]),
    ("900400", ["IPS"]),
    ("700100", ["OBSBA", "CIUDAD DE BUENOS AIRES"]),
]


def buscar_rnos_por_nombre(nombre_os):
    """
    Deduce el código RNOS / RNAS oficial de 6 dígitos a partir del nombre
    o denominación de la Obra Social.
    """
    if not nombre_os:
        return ""
    texto_limpio = str(nombre_os).upper()
    # Si ya contiene un código numérico tipo 1-2620-5 o 126205
    m = re.search(r'\b(\d{1}-\d{4}-\d{1}|\d{4,6})\b', texto_limpio)
    if m:
        digs = re.sub(r'\D', '', m.group(1))
        if 4 <= len(digs) <= 8:
            return digs
    for cod, keywords in TABLA_RNOS_OFICIAL:
        for kw in keywords:
            if kw in texto_limpio:
                return cod
    return ""


def extraer_rnos_y_obrasocial(res_dni):
    """
    Desglosa el resultado en tupla (rnos, obra_social).
    Resuelve el código RNOS / RNAS si viene con guiones o si solo se tiene el nombre.
    """
    if not res_dni:
        return "", ""
    res_str = str(res_dni).strip()
    if res_str.upper() in ["NO AFILIADO", "ERROR CONSULTA", "SIN DNI", "NONE", "NULL", ""]:
        return "", res_str

    rnos = ""
    obra_social = res_str

    # 1. Si viene con separador " - " (ej. "126205 - OBRA SOCIAL..." o "1-2620-5 - OBRA SOCIAL...")
    if " - " in res_str:
        partes = res_str.split(" - ", 1)
        cod_candidato = partes[0].strip()
        den_candidata = partes[1].strip()
        digs = re.sub(r'\D', '', cod_candidato)
        if 4 <= len(digs) <= 8:
            rnos = digs
            obra_social = den_candidata
        elif len(cod_candidato) <= 10:
            rnos = cod_candidato
            obra_social = den_candidata
    elif res_str.isdigit() and (4 <= len(res_str) <= 8):
        rnos = res_str
        obra_social = ""

    # 2. Si RNOS sigue vacío, resolver por nombre
    if not rnos and obra_social:
        rnos = buscar_rnos_por_nombre(obra_social)

    return rnos, obra_social


MESES_ES = {
    "ene": "01", "enero": "01",
    "feb": "02", "febrero": "02",
    "mar": "03", "marzo": "03",
    "abr": "04", "abril": "04",
    "may": "05", "mayo": "05",
    "jun": "06", "junio": "06",
    "jul": "07", "julio": "07",
    "ago": "08", "agosto": "08",
    "sep": "09", "septiembre": "09", "set": "09", "setiembre": "09",
    "oct": "10", "octubre": "10",
    "nov": "11", "noviembre": "11",
    "dic": "12", "diciembre": "12"
}


def desglosar_fecha(valor):
    """
    Desglosa una fecha en tupla (dia, mes, anio) como strings ("17", "07", "2026").
    Soporta datetime, date, números seriales de Excel, meses en texto español, DD/MM/YYYY, etc.
    """
    if valor is None:
        return "", "", ""
    if isinstance(valor, (datetime, date)):
        return f"{valor.day:02d}", f"{valor.month:02d}", f"{valor.year:04d}"

    val_str = str(valor).strip()
    if not val_str:
        return "", "", ""

    # Serial numérico de Excel (ej: 46220)
    try:
        val_float = float(val_str)
        if 30000 <= val_float <= 65000:
            dt = datetime(1899, 12, 30) + timedelta(days=int(val_float))
            return f"{dt.day:02d}", f"{dt.month:02d}", f"{dt.year:04d}"
    except (ValueError, OverflowError):
        pass

    clean_val = val_str.split()[0].split('T')[0].strip()
    original = val_str.split('T')[0].strip()

    # 1. YYYY-MM-DD o YYYY/MM/DD
    m_iso = re.match(r'^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})$', clean_val)
    if m_iso:
        y, m, d = int(m_iso.group(1)), int(m_iso.group(2)), int(m_iso.group(3))
        return f"{d:02d}", f"{m:02d}", f"{y:04d}"

    # 2. Con nombre de mes en español (ej. "17-jul-2026", "17 de julio de 2026", "6 ago 2026")
    m_mes = re.match(r'^(\d{1,2})\s*(?:de|[-/.])?\s*([a-zA-ZáéíóúÁÉÍÓÚ]+)\s*(?:de|[-/.])?\s*(\d{2,4})', original)
    if m_mes:
        d = int(m_mes.group(1))
        mes_txt = m_mes.group(2).lower()
        mes_pref = mes_txt[:3]
        if mes_pref in MESES_ES or mes_txt in MESES_ES:
            m_num = MESES_ES.get(mes_txt) or MESES_ES.get(mes_pref)
            y = int(m_mes.group(3))
            if y < 100:
                y = 2000 + y if y < 50 else 1900 + y
            return f"{d:02d}", m_num, f"{y:04d}"

    # 3. DD/MM/YYYY o DD-MM-YYYY (admite espacios y barras)
    m_lat = re.match(r'^(\d{1,2})\s*[-/.]\s*(\d{1,2})\s*[-/.]\s*(\d{2,4})', original)
    if m_lat:
        d, m = int(m_lat.group(1)), int(m_lat.group(2))
        y_str = m_lat.group(3)
        y = int(y_str)
        if len(y_str) == 2:
            y = 2000 + y if y < 50 else 1900 + y
        return f"{d:02d}", f"{m:02d}", f"{y:04d}"

    # 4. Dígitos separados por espacio (ej. "17 07 2026")
    partes = re.split(r'\s+', original)
    if len(partes) == 3 and all(p.isdigit() for p in partes):
        if len(partes[0]) == 4:
            return f"{int(partes[2]):02d}", f"{int(partes[1]):02d}", partes[0]
        else:
            y = partes[2]
            if len(y) == 2:
                y = f"20{y}"
            return f"{int(partes[0]):02d}", f"{int(partes[1]):02d}", y

    return val_str, "", ""


def formatear_fecha(valor):
    """
    Retorna la fecha formateada como 'DD/MM/YYYY' para tablas o textos.
    """
    d, m, y = desglosar_fecha(valor)
    if d and m and y:
        return f"{d}/{m}/{y}"
    elif d and m:
        return f"{d}/{m}"
    return str(valor or "").strip()


def _formatear_celda_fecha(cell_xml, texto):
    """
    Formatea una celda de recuadro de fecha (Día, Mes o Año) centrada y en negrita,
    preservando el ancho original y evitando cortes o wrap indeseado.
    """
    tcPr_m = re.search(r'<w:tcPr>[\s\S]*?</w:tcPr>', cell_xml)
    tcPr_xml = tcPr_m.group(0) if tcPr_m else ''
    if '<w:tcPr>' in cell_xml and '<w:noWrap/>' not in tcPr_xml:
        tcPr_xml = tcPr_xml.replace('</w:tcPr>', '<w:noWrap/></w:tcPr>', 1)
    p_xml = (
        f'<w:p><w:pPr><w:jc w:val="center"/></w:pPr>'
        f'<w:r><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi"/>'
        f'<w:b/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr>'
        f'<w:t>{escapar_xml(str(texto or ""))}</w:t></w:r></w:p>'
    )
    return f'<w:tc>{tcPr_xml}{p_xml}</w:tc>'


def _formatear_celda_texto(cell_xml, texto, font_sz="16"):
    """
    Formatea una celda con texto alineado a la izquierda, margen y fuente especificada.
    """
    tcPr_m = re.search(r'<w:tcPr>[\s\S]*?</w:tcPr>', cell_xml)
    tcPr_xml = tcPr_m.group(0) if tcPr_m else ''
    p_xml = (
        f'<w:p><w:pPr><w:ind w:left="60"/><w:jc w:val="left"/></w:pPr>'
        f'<w:r><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi"/>'
        f'<w:b/><w:sz w:val="{font_sz}"/><w:szCs w:val="{font_sz}"/></w:rPr>'
        f'<w:t>{escapar_xml(str(texto or ""))}</w:t></w:r></w:p>'
    )
    return f'<w:tc>{tcPr_xml}{p_xml}</w:tc>'


def _renderizar_tabla_paciente(tabla_template, paciente, hospital_nombre=None, hospital_refes=None):
    """
    Rellena una copia de la tabla Word del Anexo II con los datos de un paciente.
    Conserva los atributos de tabla (<w:tblPr> y <w:tblGrid>) intactos.
    """
    idx_primer_tr = tabla_template.find('<w:tr')
    if idx_primer_tr == -1:
        return tabla_template

    encabezado_tabla = tabla_template[:idx_primer_tr]

    filas = re.findall(r'<w:tr[\s\S]*?</w:tr>', tabla_template)
    if len(filas) < 22:
        return tabla_template

    # 0. R2 y R3: Denominación HPGD y Código REFES
    h_nom = hospital_nombre or paciente.get("hospital_nombre") or HOSPITAL_DEFECTO["nombre"]
    h_ref = hospital_refes or paciente.get("hospital_refes") or HOSPITAL_DEFECTO["refes"]

    if h_nom and len(filas) > 2:
        filas[2] = re.sub(
            r'<w:t>[^<]*(?:LOUIS PASTEUR|PASTEUR|SAN ROQUE|HOSPITAL)[^<]*</w:t>',
            f'<w:t>{escapar_xml(h_nom)}</w:t>',
            filas[2],
            flags=re.IGNORECASE
        )
    if h_ref and len(filas) > 3:
        filas[3] = re.sub(
            r'<w:t>\d{10,16}</w:t>',
            f'<w:t>{escapar_xml(h_ref)}</w:t>',
            filas[3]
        )

    # 1. R6: Datos Beneficiario (Apellidos y Nombres en c0, DNI en c1)
    r6_cells = re.findall(r'<w:tc[\s\S]*?</w:tc>', filas[6])
    if len(r6_cells) >= 2:
        # Nombre y Apellido
        c0_nuevo = re.sub(
            r'(<w:p[^>]*>)([\s\S]*?)(</w:p>)',
            rf'\1<w:pPr><w:jc w:val="left"/></w:pPr><w:r><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi"/><w:b/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr><w:t>{escapar_xml(paciente.get("nombre", ""))}</w:t></w:r>\3',
            r6_cells[0],
            count=1
        )
        # DNI
        c1_nuevo = re.sub(
            r'(<w:p[^>]*>)([\s\S]*?)(</w:p>)',
            rf'\1<w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr><w:t>{escapar_xml(paciente.get("dni", ""))}</w:t></w:r>\3',
            r6_cells[1],
            count=1
        )
        filas[6] = filas[6].replace(r6_cells[0], c0_nuevo, 1).replace(r6_cells[1], c1_nuevo, 1)

    # 2. R8: Checkboxes Beneficiario (Titular c1, Sexo F c15 / M c17, Edad c18)
    r8_cells = re.findall(r'<w:tc[\s\S]*?</w:tc>', filas[8])
    if len(r8_cells) >= 19:
        # Titular = X
        c1_titular = re.sub(
            r'(<w:p[^>]*>)([\s\S]*?)(</w:p>)',
            r'\1<w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="16"/></w:rPr><w:t>X</w:t></w:r>\3',
            r8_cells[1],
            count=1
        )
        filas[8] = filas[8].replace(r8_cells[1], c1_titular, 1)

        # Sexo
        sexo = str(paciente.get("sexo", "")).strip().upper()
        if sexo.startswith("F") and len(r8_cells) > 15:
            c15_sexo = re.sub(
                r'(<w:p[^>]*>)([\s\S]*?)(</w:p>)',
                r'\1<w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="16"/></w:rPr><w:t>X</w:t></w:r>\3',
                r8_cells[15],
                count=1
            )
            filas[8] = filas[8].replace(r8_cells[15], c15_sexo, 1)
        elif sexo.startswith("M") and len(r8_cells) > 17:
            c17_sexo = re.sub(
                r'(<w:p[^>]*>)([\s\S]*?)(</w:p>)',
                r'\1<w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="16"/></w:rPr><w:t>X</w:t></w:r>\3',
                r8_cells[17],
                count=1
            )
            filas[8] = filas[8].replace(r8_cells[17], c17_sexo, 1)

        # Edad
        edad = str(paciente.get("edad", "")).strip()
        if edad and len(r8_cells) > 18:
            c18_edad = re.sub(
                r'(<w:p[^>]*>)([\s\S]*?)(</w:p>)',
                rf'\1<w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="16"/></w:rPr><w:t>{escapar_xml(edad)}</w:t></w:r>\3',
                r8_cells[18],
                count=1
            )
            filas[8] = filas[8].replace(r8_cells[18], c18_edad, 1)

    # 3. R1 y R11: Fecha de atención desglosada en los 3 recuadros oficiales:
    #    Recuadro 1: DÍA (DD)
    #    Recuadro 2: MES (MM)
    #    Recuadro 3: AÑO (YYYY)
    fecha = paciente.get("fecha", "")
    dia_str, mes_str, anio_str = desglosar_fecha(fecha)

    # Cabecera (R1: celdas 1, 2 y 3 bajo 'Fecha')
    r1_cells = re.findall(r'<w:tc[\s\S]*?</w:tc>', filas[1])
    if len(r1_cells) >= 4:
        c1_dia = _formatear_celda_fecha(r1_cells[1], dia_str)
        c1_mes = _formatear_celda_fecha(r1_cells[2], mes_str)
        c1_anio = _formatear_celda_fecha(r1_cells[3], anio_str)

        filas[1] = (filas[1]
                    .replace(r1_cells[1], c1_dia, 1)
                    .replace(r1_cells[2], c1_mes, 1)
                    .replace(r1_cells[3], c1_anio, 1))

    # Fecha de prestación (R11: celdas 1, 2 y 3 bajo 'Fecha de prestación')
    r11_cells = re.findall(r'<w:tc[\s\S]*?</w:tc>', filas[11])
    if len(r11_cells) >= 4:
        c11_dia = _formatear_celda_fecha(r11_cells[1], dia_str)
        c11_mes = _formatear_celda_fecha(r11_cells[2], mes_str)
        c11_anio = _formatear_celda_fecha(r11_cells[3], anio_str)

        filas[11] = (filas[11]
                     .replace(r11_cells[1], c11_dia, 1)
                     .replace(r11_cells[2], c11_mes, 1)
                     .replace(r11_cells[3], c11_anio, 1))

    # 4. R12: Tipo de atención (Consulta c1 = X, Especialidad c3 = CARDIOLOGIA)
    r12_cells = re.findall(r'<w:tc[\s\S]*?</w:tc>', filas[12])
    if len(r12_cells) >= 4:
        c1_cons = re.sub(
            r'(<w:p[^>]*>)([\s\S]*?)(</w:p>)',
            r'\1<w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="16"/></w:rPr><w:t>X</w:t></w:r>\3',
            r12_cells[1],
            count=1
        )
        # Especialidad médica siempre CARDIOLOGIA
        c3_serv = _formatear_celda_texto(r12_cells[3], "CARDIOLOGIA", font_sz="16")
        filas[12] = filas[12].replace(r12_cells[1], c1_cons, 1).replace(r12_cells[3], c3_serv, 1)

    # 5. R13: Diagnóstico en c3
    r13_cells = re.findall(r'<w:tc[\s\S]*?</w:tc>', filas[13])
    if len(r13_cells) >= 4:
        diagnostico = str(paciente.get("diagnostico", "")).strip()
        c3_diag = _formatear_celda_texto(r13_cells[3], diagnostico, font_sz="15")
        filas[13] = filas[13].replace(r13_cells[3], c3_diag, 1)

    # 6. R21: Obra Social en c0, RNAS en c1
    r21_cells = re.findall(r'<w:tc[\s\S]*?</w:tc>', filas[21])
    if len(r21_cells) >= 3:
        obra_social = paciente.get("obra_social", "")
        c0_os = re.sub(
            r'(<w:p[^>]*>)([\s\S]*?)(</w:p>)',
            rf'\1<w:pPr><w:ind w:left="60"/><w:jc w:val="left"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="16"/></w:rPr><w:t>{escapar_xml(obra_social)}</w:t></w:r>\3',
            r21_cells[0],
            count=1
        )
        rnos = paciente.get("rnos", "")
        if not rnos and obra_social:
            rnos = buscar_rnos_por_nombre(obra_social)

        # Reemplazar celda de RNAS con un único párrafo centrado y limpio
        tcPr_m = re.search(r'<w:tcPr>[\s\S]*?</w:tcPr>', r21_cells[1])
        tcPr_xml = tcPr_m.group(0) if tcPr_m else '<w:tcPr><w:tcW w:w="1204" w:type="dxa"/><w:gridSpan w:val="4"/><w:vAlign w:val="center"/></w:tcPr>'
        c1_rnos = f'<w:tc>{tcPr_xml}<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi"/><w:b/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr><w:t>{escapar_xml(rnos)}</w:t></w:r></w:p></w:tc>'

        filas[21] = filas[21].replace(r21_cells[0], c0_os, 1).replace(r21_cells[1], c1_rnos, 1)

    return f"{encabezado_tabla}{''.join(filas)}</w:tbl>"


def extraer_pacientes_afiliados_excel(excel_path, especialidad_defecto="CARDIOLOGIA"):
    """
    Lee un Excel (original o ya verificado) y extrae todos los registros
    con Obra Social válida para generar los Anexos II.
    Garantiza captura de diagnóstico en todas las columnas y ordenamiento cronológico de menor a mayor.
    """
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    ws = wb.active
    max_row = ws.max_row

    # 1. Detectar fila de encabezados reales
    fila_encabezado = None
    col_dni = None
    col_paciente = None
    col_fecha = None
    cols_diagnostico = []
    col_sexo = None
    col_edad = None
    col_rnos = None
    col_obra_social = None

    for f in range(1, min(45, max_row + 1)):
        row_vals = {}
        for c in range(1, 40):
            val = str(ws.cell(f, c).value or "").strip()
            if val:
                row_vals[c] = val.lower()

        # Chequear si contiene DNI / Documento
        for c, t in row_vals.items():
            if any(k in t for k in ["dni", "documento", "nro doc", "nro. doc", "doc"]):
                fila_encabezado = f
                col_dni = c
                break

        if fila_encabezado:
            for c, t in row_vals.items():
                if any(k in t for k in ["paciente", "nombre", "apellido"]):
                    col_paciente = c
                elif any(k in t for k in ["cie10", "cie-10", "cie 10", "cie", "diag", "diagnóstico", "diagnostico", "motivo", "patolog", "afeccion", "cuadro"]):
                    if c not in cols_diagnostico:
                        cols_diagnostico.append(c)
                elif "sexo" in t:
                    col_sexo = c
                elif "edad" in t:
                    col_edad = c
                elif any(k in t for k in ["rnos", "rnas", "código os", "codigo os"]):
                    col_rnos = c
                elif any(k in t for k in ["obra social", "cobertura", "prepaga", "o.s."]):
                    col_obra_social = c

            # Detección inteligente de la columna de fecha de atención/prestación (evitando nacimiento/carga/vencimiento)
            for c, t in row_vals.items():
                if any(ex in t for ex in ["nacim", "nac", "venc", "emis", "carg", "alta", "baja"]):
                    continue
                if any(k in t for k in ["fecha de atenc", "fecha atenc", "fecha de prest", "fecha prest", "fecha turno", "fecha consult", "f. atenc", "f. prest"]):
                    col_fecha = c
                    break

            if not col_fecha:
                for c, t in row_vals.items():
                    if any(ex in t for ex in ["nacim", "nac", "venc", "emis", "carg", "alta", "baja"]):
                        continue
                    if any(k in t for k in ["ingreso", "egreso", "atencion", "prestacion"]) and "fecha" in t:
                        col_fecha = c
                        break

            if not col_fecha:
                for c, t in row_vals.items():
                    if any(ex in t for ex in ["nacim", "nac", "venc", "emis", "carg", "alta", "baja"]):
                        continue
                    if "fecha" in t:
                        col_fecha = c
                        break

            break

    if not fila_encabezado:
        fila_encabezado = 1
        fila_datos = 2
    else:
        fila_datos = fila_encabezado + 1

    col_dni = col_dni or 1
    col_paciente = col_paciente or 3
    col_fecha = col_fecha or 11

    # Buscar columna de Obra Social si no se detectó por encabezado
    if not col_obra_social:
        for c in range(1, 40):
            val_h = str(ws.cell(fila_encabezado, c).value or "").strip().lower()
            if any(k in val_h for k in ["obra social", "cobertura", "prepaga"]):
                col_obra_social = c
                break

    # Si tampoco se detectó col_rnos por encabezado
    if not col_rnos:
        for c in range(1, 40):
            val_h = str(ws.cell(fila_encabezado, c).value or "").strip().lower()
            if any(k in val_h for k in ["rnos", "rnas", "código os", "codigo os"]):
                col_rnos = c
                break

    pacientes = []

    for fila in range(fila_datos, max_row + 1):
        dni_val = ws.cell(fila, col_dni).value
        if not dni_val:
            continue
        dni_limpio = re.sub(r'\D', '', str(dni_val).split('.')[0].strip())
        if not (5 <= len(dni_limpio) <= 9):
            continue

        # Obtener Obra Social y RNOS
        os_val = ""
        rnos_val = ""

        if col_obra_social:
            raw_os = str(ws.cell(fila, col_obra_social).value or "").strip()
            if raw_os and raw_os.upper() not in ["NO AFILIADO", "SIN DNI", "ERROR CONSULTA", "NONE", ""]:
                if " - " in raw_os:
                    partes = raw_os.split(" - ", 1)
                    digs = re.sub(r'\D', '', partes[0])
                    if 4 <= len(digs) <= 8:
                        rnos_val = digs
                        os_val = partes[1].strip()
                    else:
                        os_val = raw_os
                else:
                    os_val = raw_os

        if col_rnos and not rnos_val:
            raw_rnos = str(ws.cell(fila, col_rnos).value or "").strip()
            digs_r = re.sub(r'\D', '', raw_rnos)
            if 4 <= len(digs_r) <= 8:
                rnos_val = digs_r
            elif raw_rnos and raw_rnos.upper() not in ["NONE", "NULL", ""]:
                rnos_val = raw_rnos

        # Si aún no tenemos RNOS, resolver por nombre de la Obra Social
        if not rnos_val and os_val:
            rnos_val = buscar_rnos_por_nombre(os_val)

        # Solo pacientes que tengan Obra Social identificada
        if not os_val or os_val.upper() in ["NO AFILIADO", "SIN DNI", "ERROR CONSULTA"]:
            continue

        nombre_val = str(ws.cell(fila, col_paciente).value or "").strip()
        fecha_val = formatear_fecha(ws.cell(fila, col_fecha).value)

        # Especialidad fijada en CARDIOLOGIA
        serv_val = "CARDIOLOGIA"

        # Captura exhaustiva de diagnóstico en todas las columnas candidatas
        diag_val = ""
        for c_diag in cols_diagnostico:
            v = str(ws.cell(fila, c_diag).value or "").strip()
            if v and v.upper() not in ["NONE", "NULL", "-", "0", "CARDIOLOGIA", "CONSULTA", ""]:
                if not diag_val:
                    diag_val = v
                elif len(v) > len(diag_val):
                    diag_val = v

        # Fallback si las columnas candidatas vinieron vacías para este paciente
        if not diag_val:
            for c_scan in range(1, min(ws.max_column + 1, 35)):
                if c_scan in [col_dni, col_paciente, col_fecha, col_sexo, col_edad, col_rnos, col_obra_social]:
                    continue
                v = str(ws.cell(fila, c_scan).value or "").strip()
                if v and len(v) >= 3 and v.upper() not in ["NONE", "NULL", "-", "0", "CARDIOLOGIA", "CONSULTA", "TITULAR", "M", "F", "NO AFILIADO"]:
                    diag_val = v
                    break

        sexo_val = str(ws.cell(fila, col_sexo).value or "").strip() if col_sexo else ""
        edad_val = str(ws.cell(fila, col_edad).value or "").strip() if col_edad else ""

        pacientes.append({
            "fila": fila,
            "nombre": nombre_val,
            "dni": dni_limpio,
            "fecha": fecha_val,
            "rnos": rnos_val,
            "obra_social": os_val,
            "servicio": serv_val,
            "diagnostico": diag_val,
            "sexo": sexo_val,
            "edad": edad_val
        })

    return ordenar_pacientes_anexos(pacientes)


def parsear_fecha_para_orden(fecha_str):
    """
    Convierte cualquier formato de fecha a 'YYYY-MM-DD' para ordenamiento cronológico.
    """
    if not fecha_str:
        return "9999-99-99"
    d, m, y = desglosar_fecha(fecha_str)
    if d and m and y and len(y) == 4 and len(m) == 2 and len(d) == 2:
        return f"{y}-{m}-{d}"
    return str(fecha_str).strip()


def ordenar_pacientes_anexos(pacientes):
    """
    Ordena la lista de pacientes agrupando primero por código RNOS (u Obra Social),
    luego cronológicamente por Fecha de prestación de menor a mayor (antigua a reciente, de un mes al otro),
    y finalmente por Apellido y Nombre.
    """
    if not pacientes:
        return []

    def clave_orden(p):
        rnos = str(p.get("rnos") or "").strip()
        os_nom = str(p.get("obra_social") or "").strip().upper()
        if not rnos and os_nom:
            rnos = buscar_rnos_por_nombre(os_nom)
        # Normalizar código RNOS a dígitos para agrupar exactamente la misma obra social
        rnos_digs = re.sub(r'\D', '', rnos)
        cod_agrupacion = rnos_digs if rnos_digs else f"ZZZ_{os_nom}"
        fecha_ord = parsear_fecha_para_orden(p.get("fecha"))
        nombre = str(p.get("nombre") or "").strip().upper()
        return (cod_agrupacion, fecha_ord, nombre)

    return sorted(pacientes, key=clave_orden)


def generar_anexos_docx(pacientes, ruta_salida_docx, plantilla_path=None, hospital_nombre=None, hospital_refes=None):
    """
    Genera un documento Word único (.docx) con todas las fojas de Anexo II,
    una página completa (encabezado oficial + tabla del paciente) por cada paciente afiliado,
    organizadas por Obra Social (RNOS) y cronológicamente por fecha, separadas por saltos de página.
    """
    if not pacientes:
        raise ValueError("No se encontraron pacientes afiliados para generar anexos.")

    # Asegurar orden agrupado por RNOS y fecha
    pacientes = ordenar_pacientes_anexos(pacientes)

    plantilla = Path(plantilla_path) if plantilla_path else PLANTILLA_DOCX
    if not plantilla.exists():
        raise FileNotFoundError(f"No se encontró la plantilla de Anexo II en: {plantilla}")

    h_nom = hospital_nombre or HOSPITAL_DEFECTO["nombre"]
    h_ref = hospital_refes or HOSPITAL_DEFECTO["refes"]

    # Leer plantilla como ZIP
    with zipfile.ZipFile(plantilla, 'r') as z_in:
        archivos_zip = {name: z_in.read(name) for name in z_in.namelist()}

    doc_xml = archivos_zip['word/document.xml'].decode('utf-8')

    idx_body_inicio = doc_xml.find('<w:body>')
    idx_body_fin = doc_xml.find('</w:body>')
    if idx_body_inicio == -1 or idx_body_fin == -1:
        raise ValueError("No se encontró <w:body> en la plantilla Word.")

    body_content = doc_xml[idx_body_inicio + 8 : idx_body_fin]

    idx_tbl_inicio = body_content.find('<w:tbl>')
    idx_tbl_fin = body_content.find('</w:tbl>')
    if idx_tbl_inicio == -1 or idx_tbl_fin == -1:
        raise ValueError("No se encontró la tabla en la plantilla Word.")
    idx_tbl_fin += 8  # incluir el cierre </w:tbl>

    encabezado_pagina = body_content[:idx_tbl_inicio]
    tabla_template = body_content[idx_tbl_inicio:idx_tbl_fin]
    despues_tabla = body_content[idx_tbl_fin:]

    # Generar una foja por cada paciente
    paginas = []
    for pac in pacientes:
        tabla_pac = _renderizar_tabla_paciente(
            tabla_template,
            pac,
            hospital_nombre=h_nom,
            hospital_refes=h_ref
        )
        paginas.append(f"{encabezado_pagina}{tabla_pac}")

    # Separar cada anexo con un salto de página
    salto_pagina = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'
    cuerpo_combinado = salto_pagina.join(paginas) + despues_tabla

    nuevo_doc_xml = doc_xml[:idx_body_inicio + 8] + cuerpo_combinado + doc_xml[idx_body_fin:]

    archivos_zip['word/document.xml'] = nuevo_doc_xml.encode('utf-8')

    # Guardar nuevo .docx
    Path(ruta_salida_docx).parent.mkdir(exist_ok=True, parents=True)
    with zipfile.ZipFile(ruta_salida_docx, 'w', compression=zipfile.ZIP_DEFLATED) as z_out:
        for name, data in archivos_zip.items():
            z_out.writestr(name, data)

    return len(pacientes)


def generar_anexos_pdf(pacientes, ruta_salida_pdf, hospital_nombre=None, hospital_refes=None):
    """
    Genera un archivo PDF único masivo con todas las fojas de Anexo II,
    una página por paciente afiliado, listo para imprimir directamente.
    """
    if not pacientes:
        raise ValueError("No se encontraron pacientes afiliados para generar anexos.")

    pacientes = ordenar_pacientes_anexos(pacientes)
    h_nom = hospital_nombre or HOSPITAL_DEFECTO["nombre"]
    h_ref = hospital_refes or HOSPITAL_DEFECTO["refes"]

    try:
        from reportlab.lib.pagesizes import A4
        from reportlab.lib import colors
        from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, PageBreak
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    except ImportError:
        raise RuntimeError("ReportLab no está disponible. Ejecute: pip install reportlab")

    Path(ruta_salida_pdf).parent.mkdir(exist_ok=True, parents=True)

    # A4 = 595.27 x 841.89 pt. Márgenes de 20 pt a cada lado -> ancho útil = 555.27 pt
    doc = SimpleDocTemplate(
        str(ruta_salida_pdf),
        pagesize=A4,
        leftMargin=20,
        rightMargin=20,
        topMargin=20,
        bottomMargin=20,
    )

    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        'AnexoTitulo',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=12,
        leading=14,
        alignment=1,
        spaceAfter=2,
    )

    subtitle_style = ParagraphStyle(
        'AnexoSubtitulo',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=9,
        leading=11,
        alignment=1,
        spaceAfter=6,
    )

    def P(texto, font='Helvetica', size=7.5, leading=9, align=0, bold=False):
        fName = 'Helvetica-Bold' if bold else font
        p_style = ParagraphStyle(
            f'P_{fName}_{size}_{align}',
            parent=styles['Normal'],
            fontName=fName,
            fontSize=size,
            leading=leading,
            alignment=align,
        )
        return Paragraph(str(texto or ''), p_style)

    col_widths = [150, 155, 145, 105.27]
    elements = []

    for idx, pac in enumerate(pacientes):
        nombre = pac.get("nombre", "")
        dni = pac.get("dni", "")
        fecha = formatear_fecha(pac.get("fecha", ""))
        servicio = str(pac.get("servicio", "")).strip()
        if not servicio or servicio.upper() in ["CONSULTA", "CONSULTAS", "CONSULTA MEDICA", "AMBULATORIO", "AMBULATORIA", "GUARDIA", "CONSULTA EXTERNA"]:
            servicio = "CARDIOLOGIA"
        diagnostico = pac.get("diagnostico", "")
        obra_social = pac.get("obra_social", "")
        rnos = pac.get("rnos", "")
        if not rnos and obra_social:
            rnos = buscar_rnos_por_nombre(obra_social)
        sexo = str(pac.get("sexo", "")).strip().upper()
        edad = str(pac.get("edad", "")).strip()

        es_f = "X" if sexo.startswith("F") else ""
        es_m = "X" if sexo.startswith("M") else ""

        # Encabezado oficial
        elements.append(Paragraph("ANEXO II", title_style))
        elements.append(Paragraph("<u>CERTIFICACIÓN DE PRÁCTICA MÉDICA Y ADMINISTRATIVA</u>", subtitle_style))

        tabla_data = [
            # 0
            [P("CERTIFICACIÓN DE PRÁCTICA MÉDICA Y ADMINISTRATIVA", bold=True, size=8), "", "", P("Fecha", bold=True, size=8, align=1)],
            # 1
            [P("Comprobante de atención médica y administrativa HPGD", size=7), "", "", P(fecha, bold=True, size=8.5, align=1)],
            # 2
            [P(f"Denominación del HPGD: <b>{escapar_xml(h_nom)}</b>", size=7.5), "", "", P(f"Código REFES: <b>{escapar_xml(h_ref)}</b>", size=7.5, align=1)],
            # 3: Banner Beneficiario
            [P("DATOS DEL BENEFICIARIO", bold=True, size=8, align=1), "", "", ""],
            # 4
            [P("Apellidos y Nombres", bold=True, size=7), "", "", P("N° de Documento", bold=True, size=7, align=1)],
            # 5
            [P(nombre, bold=True, size=9), "", "", P(dni, bold=True, size=9, align=1)],
            # 6
            [P("Tipo de Beneficiario", bold=True, size=7), P("Parentesco", bold=True, size=7), P("Sexo", bold=True, size=7, align=1), P("Edad", bold=True, size=7, align=1)],
            # 7
            [
                P("Titular <b>[X]</b>  Fam <b>[ ]</b>  Adh <b>[ ]</b>", size=6.8),
                P("Cónyuge <b>[ ]</b>  Hijo <b>[ ]</b>  Otro <b>[ ]</b>", size=6.8),
                P(f"F <b>[{es_f or ' '}]</b>   M <b>[{es_m or ' '}]</b>", size=7.5, align=1),
                P(edad, bold=True, size=8, align=1)
            ],
            # 8: Banner Atención
            [P("DATOS DE LA ATENCIÓN", bold=True, size=8, align=1), "", "", ""],
            # 9
            [P("Tipo de atención", bold=True, size=7), "", "", P("Fecha de prestación", bold=True, size=7, align=1)],
            # 10
            [P("Ambulatoria / Consulta médica programada", size=7), "", "", P(fecha, bold=True, size=8.5, align=1)],
            # 11
            [P("Consulta <b>[X]</b>", size=7.5), P(f"Especialidad: <b>{servicio}</b>", size=7.5), "", ""],
            # 12
            [P(f"Diagnóstico: <b>{diagnostico}</b>", size=7.5), "", "", ""],
            # 13
            [P("Práctica", bold=True, size=7), "", P("Código", bold=True, size=7), ""],
            # 14
            ["", "", "", ""],
            # 15
            ["", "", "", ""],
            # 16
            [P("Internación <b>[ ]</b>", size=7), P("Diagnóstico de Egreso CIE 10:", size=7), P("Cód. Principal:", size=7), P("Otros Códigos:", size=7)],
            # 17
            [P("CIE 10 Clasificación Internacional de Enfermedades", size=6.5, align=1), "", "", ""],
            # 18: Firma Médico
            [P("Firma del Médico y sello con N° de Matrícula", size=7), "", "", ""],
            # 19: Obra Social Header
            [P("NOMBRE DEL AGENTE DE SEGURO DE SALUD", bold=True, size=7, align=1), "", P("RNAS", bold=True, size=7, align=1), P("FIRMA DEL DIRECTOR O SUBDIRECTOR", bold=True, size=6.5, align=1)],
            # 20: Obra Social Values
            [P(obra_social, bold=True, size=8.5), "", P(rnos, bold=True, size=9, align=1), ""]
        ]

        t_style = [
            ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('TOPPADDING', (0, 0), (-1, -1), 2),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
            ('LEFTPADDING', (0, 0), (-1, -1), 4),
            ('RIGHTPADDING', (0, 0), (-1, -1), 4),

            ('SPAN', (0, 0), (2, 0)),
            ('SPAN', (0, 1), (2, 1)),
            ('SPAN', (0, 2), (2, 2)),
            ('SPAN', (0, 3), (3, 3)),
            ('SPAN', (0, 4), (2, 4)),
            ('SPAN', (0, 5), (2, 5)),
            ('SPAN', (0, 8), (3, 8)),
            ('SPAN', (0, 9), (2, 9)),
            ('SPAN', (0, 10), (2, 10)),
            ('SPAN', (1, 11), (3, 11)),
            ('SPAN', (0, 12), (3, 12)),
            ('SPAN', (0, 13), (1, 13)),
            ('SPAN', (2, 13), (3, 13)),
            ('SPAN', (0, 14), (1, 14)),
            ('SPAN', (2, 14), (3, 14)),
            ('SPAN', (0, 15), (1, 15)),
            ('SPAN', (2, 15), (3, 15)),
            ('SPAN', (0, 17), (3, 17)),
            ('SPAN', (0, 18), (3, 18)),
            ('SPAN', (0, 19), (1, 19)),
            ('SPAN', (0, 20), (1, 20)),

            ('BACKGROUND', (0, 3), (3, 3), colors.HexColor('#E5E7EB')),
            ('BACKGROUND', (0, 8), (3, 8), colors.HexColor('#E5E7EB')),
            ('BACKGROUND', (0, 0), (3, 0), colors.HexColor('#F3F4F6')),
            ('BACKGROUND', (0, 4), (3, 4), colors.HexColor('#F3F4F6')),
            ('BACKGROUND', (0, 6), (3, 6), colors.HexColor('#F3F4F6')),
            ('BACKGROUND', (0, 9), (3, 9), colors.HexColor('#F3F4F6')),
            ('BACKGROUND', (0, 13), (3, 13), colors.HexColor('#F3F4F6')),
            ('BACKGROUND', (0, 19), (3, 19), colors.HexColor('#F3F4F6')),

            ('BOTTOMPADDING', (0, 18), (3, 18), 35),
            ('BOTTOMPADDING', (0, 20), (3, 20), 30),
        ]

        row_heights = [
            18, 16, 18, 16, 14, 20, 14, 18, 16, 14, 18, 18, 20, 14, 14, 14, 16, 13, 50, 14, 45
        ]

        t = Table(tabla_data, colWidths=col_widths, rowHeights=row_heights, style=TableStyle(t_style))
        elements.append(t)

        if idx < len(pacientes) - 1:
            elements.append(PageBreak())

    doc.build(elements)
    return len(pacientes)

