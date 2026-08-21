import os
import requests

DNI = "35963384"

print("")
print("====================================================")
print("       PRUEBA 3 FUENTES DE COBERTURA")
print("====================================================")
print("")
print("DNI:", DNI)
print("")


# ============================================================
# 1. SSSALUD
# ============================================================

print("====================================================")
print("1. SSSALUD")
print("====================================================")

try:

    from sss_beneficiarios_hospitales.data import DataBeneficiariosSSSHospital

    usuario = os.environ["SSS_USER"]
    password = os.environ["SSS_PASSWORD"]

    sss = DataBeneficiariosSSSHospital(
        user=usuario,
        password=password
    )

    sss._save_response = lambda filename, resp: None

    respuesta_sss = sss.query(DNI)

    print("")
    print("RESPUESTA SSSALUD:")
    print(respuesta_sss)

except Exception as e:

    print("")
    print("ERROR SSSALUD:")
    print(str(e))


# ============================================================
# 2. ANSES / CODEM
# ============================================================

print("")
print("====================================================")
print("2. ANSES / CODEM")
print("====================================================")

try:

    session = requests.Session()

    url_anses = "https://servicioswww.anses.gob.ar/ooss2/"

    respuesta_anses = session.get(
        url_anses,
        timeout=30
    )

    print("")
    print("HTTP:", respuesta_anses.status_code)
    print("URL FINAL:", respuesta_anses.url)
    print("TAMAÑO:", len(respuesta_anses.text))

    print("")
    print("RESPUESTA ANSES:")
    print("--------------------------------------------")
    print(respuesta_anses.text[:5000])
    print("--------------------------------------------")

except Exception as e:

    print("")
    print("ERROR ANSES:")
    print(str(e))


# ============================================================
# 3. PUCO / SISA
# ============================================================

print("")
print("====================================================")
print("3. PUCO / SISA")
print("====================================================")

try:

    url_sisa = "https://sisa.msal.gov.ar/sisa/"

    respuesta_sisa = requests.get(
        url_sisa,
        timeout=30
    )

    print("")
    print("HTTP:", respuesta_sisa.status_code)
    print("URL FINAL:", respuesta_sisa.url)
    print("TAMAÑO:", len(respuesta_sisa.text))

    print("")
    print("RESPUESTA SISA:")
    print("--------------------------------------------")
    print(respuesta_sisa.text[:5000])
    print("--------------------------------------------")

except Exception as e:

    print("")
    print("ERROR SISA:")
    print(str(e))


print("")
print("====================================================")
print("             FIN DE LA PRUEBA")
print("====================================================")