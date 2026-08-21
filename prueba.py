import requests

url = "https://seguro.sssalud.gob.ar/"

try:
    respuesta = requests.get(url, timeout=15)

    print("Código HTTP:", respuesta.status_code)
    print("Conexión exitosa")
    print("Servidor:", respuesta.url)

except requests.RequestException as e:
    print("Error de conexión:")
    print(e)