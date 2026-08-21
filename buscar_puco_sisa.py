import requests
import re
import urllib3
from urllib.parse import urljoin

urllib3.disable_warnings()

BASE = "https://sisa.msal.gov.ar/sisa/"

print("====================================================")
print("       BÚSQUEDA PUCO DENTRO DE SISA")
print("====================================================")
print()

# --------------------------------------------------------
# Descargar página principal
# --------------------------------------------------------

r = requests.get(
    BASE,
    verify=False,
    timeout=60
)

print("Página principal")
print("HTTP:", r.status_code)
print("Tamaño:", len(r.text))
print()

html = r.text

# --------------------------------------------------------
# Buscar todos los JS
# --------------------------------------------------------

scripts = re.findall(
    r'<script[^>]+src=["\']([^"\']+)["\']',
    html,
    re.I
)

print("========== SCRIPTS ENCONTRADOS ==========")

for script in scripts:
    print(script)

print()

# --------------------------------------------------------
# Descargar cada JS
# --------------------------------------------------------

for script in scripts:

    url = urljoin(BASE, script)

    print("----------------------------------------------------")
    print("JS:", script)
    print("URL:", url)

    try:

        respuesta = requests.get(
            url,
            verify=False,
            timeout=60
        )

        print("HTTP:", respuesta.status_code)
        print("Tamaño:", len(respuesta.text))

        if respuesta.status_code != 200:
            continue

        nombre = script.split("/")[-1].split("?")[0]

        if not nombre:
            continue

        archivo = "sisa_" + nombre

        with open(
            archivo,
            "w",
            encoding="utf-8",
            errors="ignore"
        ) as f:
            f.write(respuesta.text)

        print("Guardado:", archivo)

        texto = respuesta.text.lower()

        palabras = [
            "puco",
            "padron",
            "padrón",
            "cobertura",
            "beneficiario",
            "afiliado",
            "documento",
            "dni",
            "obra social"
        ]

        encontradas = []

        for palabra in palabras:

            if palabra.lower() in texto:
                encontradas.append(palabra)

        if encontradas:

            print("")
            print(">>> PALABRAS ENCONTRADAS:")

            for palabra in encontradas:
                print("   ", palabra)

    except Exception as e:

        print("ERROR:", e)


print()
print("====================================================")
print("                    FIN")
print("====================================================")