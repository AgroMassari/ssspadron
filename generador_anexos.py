import re
import zipfile
import xml.sax.saxutils as saxutils
from pathlib import Path
from datetime import datetime, date
import openpyxl

PLANTILLA_DOCX = Path(__file__).resolve().parent / "plantilla_anexo_ii.docx"


def escapar_xml(texto):
    if texto is None:
        return ""
    return saxutils.escape(str(texto))


def formatear_fecha(valor):
    if valor is None:
        return ""
    if isinstance(valor, (datetime, date)):
        return valor.strftime("%d/%m/%Y")
    val_str = str(valor).strip()
    # Si viene como YYYY/MM/DD o YYYY-MM-DD
    m = re.match(r'^(\d{4})[-/](\d{1,2})[-/](\d{1,2})', val_str)
    if m:
        return f"{int(m.group(3)):02d}/{int(m.group(2)):02d}/{m.group(1)}"
    return val_str


def _renderizar_tabla_paciente(tabla_template, paciente):
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

    # 3. R1 y R11: Fecha de atención (ingreso y egreso va la misma desde-hasta)
    fecha = paciente.get("fecha", "")
    if fecha:
        # Cabecera (R1 celda 3)
        r1_cells = re.findall(r'<w:tc[\s\S]*?</w:tc>', filas[1])
        if len(r1_cells) >= 4:
            c_fec = re.sub(
                r'(<w:p[^>]*>)([\s\S]*?)(</w:p>)',
                rf'\1<w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="16"/></w:rPr><w:t>{escapar_xml(fecha)}</w:t></w:r>\3',
                r1_cells[3],
                count=1
            )
            filas[1] = filas[1].replace(r1_cells[3], c_fec, 1)

        # Fecha de prestación (R11 celda 3)
        r11_cells = re.findall(r'<w:tc[\s\S]*?</w:tc>', filas[11])
        if len(r11_cells) >= 4:
            c_fec2 = re.sub(
                r'(<w:p[^>]*>)([\s\S]*?)(</w:p>)',
                rf'\1<w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="16"/></w:rPr><w:t>{escapar_xml(fecha)}</w:t></w:r>\3',
                r11_cells[3],
                count=1
            )
            filas[11] = filas[11].replace(r11_cells[3], c_fec2, 1)

    # 4. R12: Tipo de atención (Consulta c1 = X, Especialidad c3)
    r12_cells = re.findall(r'<w:tc[\s\S]*?</w:tc>', filas[12])
    if len(r12_cells) >= 4:
        c1_cons = re.sub(
            r'(<w:p[^>]*>)([\s\S]*?)(</w:p>)',
            r'\1<w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="16"/></w:rPr><w:t>X</w:t></w:r>\3',
            r12_cells[1],
            count=1
        )
        servicio = paciente.get("servicio", "CONSULTA")
        c3_serv = re.sub(
            r'(<w:p[^>]*>)([\s\S]*?)(</w:p>)',
            rf'\1<w:pPr><w:ind w:left="60"/><w:jc w:val="left"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="16"/></w:rPr><w:t>{escapar_xml(servicio)}</w:t></w:r>\3',
            r12_cells[3],
            count=1
        )
        filas[12] = filas[12].replace(r12_cells[1], c1_cons, 1).replace(r12_cells[3], c3_serv, 1)

    # 5. R13: Diagnóstico en c3
    r13_cells = re.findall(r'<w:tc[\s\S]*?</w:tc>', filas[13])
    if len(r13_cells) >= 4:
        diagnostico = paciente.get("diagnostico", "")
        if diagnostico:
            c3_diag = re.sub(
                r'(<w:p[^>]*>)([\s\S]*?)(</w:p>)',
                rf'\1<w:pPr><w:ind w:left="60"/><w:jc w:val="left"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="15"/></w:rPr><w:t>{escapar_xml(diagnostico)}</w:t></w:r>\3',
                r13_cells[3],
                count=1
            )
            filas[13] = filas[13].replace(r13_cells[3], c3_diag, 1)

    # 6. R21: Obra Social en c0, RNOS en c1
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
        c1_rnos = re.sub(
            r'(<w:p[^>]*>)([\s\S]*?)(</w:p>)',
            rf'\1<w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="16"/></w:rPr><w:t>{escapar_xml(rnos)}</w:t></w:r>\3',
            r21_cells[1],
            count=1
        )
        filas[21] = filas[21].replace(r21_cells[0], c0_os, 1).replace(r21_cells[1], c1_rnos, 1)

    return f"{encabezado_tabla}{''.join(filas)}</w:tbl>"


def extraer_pacientes_afiliados_excel(excel_path):
    """
    Lee un Excel (original o ya verificado) y extrae todos los registros
    con Obra Social válida para generar los Anexos II.
    """
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    ws = wb.active
    max_row = ws.max_row

    # 1. Detectar fila de encabezados reales
    fila_encabezado = None
    col_dni = None
    col_paciente = None
    col_fecha = None
    col_prestacion = None
    col_diagnostico = None
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
                elif any(k in t for k in ["fecha turno", "fecha de", "fecha"]):
                    col_fecha = c
                elif any(k in t for k in ["prestación", "prestacion", "servicio", "práctica", "practica"]):
                    col_prestacion = c
                elif any(k in t for k in ["cie10", "cie-10", "diag", "diagnóstico", "diagnostico"]):
                    col_diagnostico = c
                elif "sexo" in t:
                    col_sexo = c
                elif "edad" in t:
                    col_edad = c
                elif any(k in t for k in ["rnos", "rnas", "código os", "codigo os"]):
                    col_rnos = c
                elif any(k in t for k in ["obra social", "cobertura", "prepaga", "o.s."]):
                    col_obra_social = c
            break

    if not fila_encabezado:
        fila_encabezado = 1
        fila_datos = 2
    else:
        fila_datos = fila_encabezado + 1

    col_dni = col_dni or 1
    col_paciente = col_paciente or 3
    col_fecha = col_fecha or 11
    col_prestacion = col_prestacion or 12
    col_diagnostico = col_diagnostico or 13

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
                    if partes[0].strip().isdigit():
                        rnos_val = partes[0].strip()
                        os_val = partes[1].strip()
                    else:
                        os_val = raw_os
                else:
                    os_val = raw_os

        if col_rnos and not rnos_val:
            raw_rnos = str(ws.cell(fila, col_rnos).value or "").strip()
            if raw_rnos and raw_rnos.isdigit():
                rnos_val = raw_rnos

        # Solo pacientes que tengan Obra Social identificada
        if not os_val or os_val.upper() in ["NO AFILIADO", "SIN DNI", "ERROR CONSULTA"]:
            continue

        nombre_val = str(ws.cell(fila, col_paciente).value or "").strip()
        fecha_val = formatear_fecha(ws.cell(fila, col_fecha).value)
        serv_val = str(ws.cell(fila, col_prestacion).value or "CONSULTA").strip()
        diag_val = str(ws.cell(fila, col_diagnostico).value or "").strip()
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

    return pacientes


def generar_anexos_docx(pacientes, ruta_salida_docx, plantilla_path=None):
    """
    Genera un documento Word único (.docx) con todas las fojas de Anexo II,
    una página completa (encabezado oficial + tabla del paciente) por cada paciente afiliado,
    separadas por saltos de página.
    """
    if not pacientes:
        raise ValueError("No se encontraron pacientes afiliados para generar anexos.")

    plantilla = Path(plantilla_path) if plantilla_path else PLANTILLA_DOCX
    if not plantilla.exists():
        raise FileNotFoundError(f"No se encontró la plantilla de Anexo II en: {plantilla}")

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
        tabla_pac = _renderizar_tabla_paciente(tabla_template, pac)
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
