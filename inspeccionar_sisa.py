import requests
import re
import urllib3

urllib3.disable_warnings()

URL = "https://sisa.msal.gov.ar/sisa/"

print("====================================================")
print("        INSPECCIÓN SISA / PUCO")
print("====================================================")
print()

r = requests.get(
    URL,
    verify=False,
    timeout=30
)

print("HTTP:", r.status_code)
print("URL:", r.url)
print("TAMAÑO:", len(r.text))
print()

html = r.text

print("========== ENLACES ENCONTRADOS ==========")

for x in re.findall(r'href=["\']([^"\']+)["\']', html, re.I):
    print(x)

print()
print("========== SCRIPTS ==========")

for x in re.findall(r'<script[^>]+src=["\']([^"\']+)["\']', html, re.I):
    print(x)

print()
print("========== PALABRAS RELACIONADAS ==========")

palabras = [
    "PUCO",
    "padrón",
    "padron",
    "cobertura",
    "beneficiario",
    "afiliado",
    "DNI",
    "documento",
    "consulta"
]

for palabra in palabras:

    if palabra.lower() in html.lower():
        print("ENCONTRADO:", palabra)

print()
print("========== FIN ==========")