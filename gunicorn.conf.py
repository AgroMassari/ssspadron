import os

# Configuración de Gunicorn optimizada para Render (512 MB RAM Free Tier)
# y archivos masivos de Excel/Word

port = os.environ.get("PORT", "5000")
bind = f"0.0.0.0:{port}"

# 1 solo worker para evitar multiplicar el consumo de memoria RAM por proceso
workers = 1

# 4 threads para permitir concurrencia liviana (consultar progreso, etc.) sin overhead
threads = 4

# Timeout extendido a 300 segundos (5 minutos) para procesar subidas de Excels masivos sin WORKER TIMEOUT
timeout = 300

# Keep-alive para conexiones HTTP
keepalive = 5

# Reciclado de worker preventivo para evitar fugas de memoria
max_requests = 100
max_requests_jitter = 10

# Nivel de log
loglevel = "info"
