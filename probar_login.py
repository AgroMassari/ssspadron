import os
from sss_beneficiarios_hospitales.data import DataBeneficiariosSSSHospital

usuario = os.environ["SSS_USER"]
password = os.environ["SSS_PASSWORD"]

sss = DataBeneficiariosSSSHospital(
    user=usuario,
    password=password
)

# Evitamos temporalmente el problema de codificación
# al guardar login.html / login.json.
sss._save_response = lambda filename, resp: None

print("Intentando iniciar sesión en SSSalud...")

try:
    resultado = sss.login()

    print("")
    print("Resultado:", resultado)

    if resultado:
        print("LOGIN CORRECTO")
        print("La autenticación fue aceptada por SSSalud.")
    else:
        print("LOGIN FALLIDO")
        print("La respuesta no fue reconocida como una sesión autenticada.")

except Exception as e:
    print("")
    print("ERROR:")
    print(type(e).__name__)
    print(str(e))
