import os
from sss_beneficiarios_hospitales.data import DataBeneficiariosSSSHospital

usuario = os.environ["SSS_USER"]
password = os.environ["SSS_PASSWORD"]

sss = DataBeneficiariosSSSHospital(
    user=usuario,
    password=password
)

print("Intentando iniciar sesión en SSSalud...")

try:
    resultado = sss.login()

    if resultado:
        print("LOGIN CORRECTO")
        print("La cuenta pudo autenticarse correctamente.")
    else:
        print("LOGIN FALLIDO")
        print("Las credenciales no fueron aceptadas o el sitio no respondió como esperaba.")

except Exception as e:
    print("ERROR:")
    print(type(e).__name__)
    print(e)
