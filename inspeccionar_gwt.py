import re

ARCHIVO = "sisa_nocache_real.js"

print("=" * 60)
print("       BUSQUEDA DE ARCHIVOS GWT SISA")
print("=" * 60)

with open(ARCHIVO, "r", encoding="utf-8", errors="ignore") as f:
    texto = f.read()

patrones = [
    r'["\']([^"\']*\.cache\.js[^"\']*)["\']',
    r'["\']([^"\']*\.cache[^"\']*)["\']',
    r'["\']([^"\']*sisa[^"\']*)["\']',
    r'["\']([^"\']*\.js)["\']'
]

encontrados = set()

for patron in patrones:

    resultados = re.findall(patron, texto, re.I)

    for resultado in resultados:
        encontrados.add(resultado)

print()
print("ARCHIVOS / RUTAS ENCONTRADAS:")
print()

for x in sorted(encontrados):

    if (
        ".cache" in x.lower()
        or "sisa" in x.lower()
        or ".js" in x.lower()
    ):
        print(x)

print()
print("=" * 60)
print("FIN")
print("=" * 60)