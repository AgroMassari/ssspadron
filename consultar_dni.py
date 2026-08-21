import os
from pprint import pprint
from sss_beneficiarios_hospitales.data import DataBeneficiariosSSSHospital

usuario = os.environ["SSS_USER"]
password = os.environ["SSS_PASSWORD"]
dni = os.environ["SSS_DNI"]

sss = DataBeneficiariosSSSHospital(
    user=usuario,
    password=password
)

# Evita el problema de codificación al guardar las respuestas
sss._save_response = lambda filename, resp: None

print("Consultando DNI...")
print("")

try:
    resultado = sss.query(dni)

    print("Consulta ejecutada.")
    print("OK:", resultado.get("ok"))

    if resultado.get("ok"):
        print("")
        print("RESULTADOS:")
        pprint(resultado.get("resultados"))
    else:
        print("")
        print("ERROR:")
        pprint(resultado)

except Exception as e:
    print("")
    print("ERROR:")
    print(type(e).__name__)
    print(str(e))
