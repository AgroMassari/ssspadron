import os
from sss_beneficiarios_hospitales.data import DataBeneficiariosSSSHospital

usuario = os.environ["SSS_USER"]
password = os.environ["SSS_PASSWORD"]
dni = "25045507"

sss = DataBeneficiariosSSSHospital(
    user=usuario,
    password=password
)

sss._save_response = lambda filename, resp: None

resultado = sss.query(dni)

datos = resultado.get("resultados", {})

print("OK:", resultado.get("ok"))
print("Afiliado:", datos.get("afiliado"))
print("Título:", datos.get("title"))
print("")
print("TABLAS ENCONTRADAS:")

for tabla in datos.get("tablas", []):
    print("")
    print("Nombre:", tabla.get("name"))
    print("Campos:", list(tabla.get("data", {}).keys()))
