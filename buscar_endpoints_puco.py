import os
import re

print("====================================================")
print("      BUSQUEDA DE PUCO / PADRON / COBERTURA")
print("====================================================")
print("")

archivos = [
    x for x in os.listdir(".")
    if x.startswith("sisa_") and x.endswith(".js")
]

palabras = [
    "puco",
    "padron",
    "padrón",
    "cobertura",
    "beneficiario",
    "afiliado",
    "obra social",
    "documento",
    "dni"
]

for archivo in archivos:

    try:

        with open(
            archivo,
            "r",
            encoding="utf-8",
            errors="ignore"
        ) as f:
            texto = f.read()

    except Exception:
        continue

    texto_lower = texto.lower()

    coincidencias = []

    for palabra in palabras:

        if palabra.lower() in texto_lower:
            coincidencias.append(palabra)

    if not coincidencias:
        continue

    print("")
    print("====================================================")
    print("ARCHIVO:", archivo)
    print("COINCIDENCIAS:", ", ".join(coincidencias))
    print("====================================================")

    # Mostrar fragmentos alrededor de cada coincidencia
    for palabra in coincidencias:

        patron = re.compile(
            re.escape(palabra),
            re.IGNORECASE
        )

        encontrados = list(patron.finditer(texto))

        print("")
        print("----", palabra, "----")

        for match in encontrados[:10]:

            inicio = max(0, match.start() - 300)
            fin = min(len(texto), match.end() + 500)

            fragmento = texto[inicio:fin]

            print(fragmento)
            print("")
            print("--------------------------------------------")


print("")
print("====================================================")
print("                    FIN")
print("====================================================")