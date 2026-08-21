import re

ARCHIVO = "sisa_nocache_real.js"

print("=" * 60)
print("       BUSQUEDA DE ARCHIVOS GWT SISA")
print("=" * 60)

try:
    with open(ARCHIVO, "r", encoding="utf-8", errors="ignore") as f:
        texto = f.read()
except FileNotFoundError:
    print()
    print("ERROR: No existe", ARCHIVO)
    print("Primero ejecutá:")
    print("python inspeccionar_gwt.py")
    input("\nPresioná ENTER para salir...")
    raise SystemExit

print()
print("Tamaño del archivo:", len(texto))
print()

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

print("========== ARCHIVOS / RUTAS ENCONTRADAS ==========")
print()

if not encontrados:
    print("No se encontraron rutas directas.")
else:
    for x in sorted(encontrados):
        print(x)

print()
print("========== PALABRAS IMPORTANTES ==========")
print()

palabras = [
    "padron",
    "padrón",
    "ciudadano",
    "documento",
    "dni",
    "cobertura",
    "obra social",
    "beneficiario",
    "afiliado"
]

for palabra in palabras:

    cantidad = texto.lower().count(palabra.lower())

    print(f"{palabra}: {cantidad}")

print()
print("=" * 60)
print("                    FIN")
print("=" * 60)

input("\nPresioná ENTER para cerrar...")