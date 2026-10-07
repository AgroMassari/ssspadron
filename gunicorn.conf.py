import os

# Configuración de Gunicorn optimizada para Render (512 MB RAM Free Tier)
# y archivos masivos de Excel/Word

port = os.environ.get("PORT", "5000")
bind = f"0.0.0.0:{port}"

# 1 solo worker para evitar multiplicar el consumo de memoria RAM por proceso
workers = 1

# 4 threads para permitir concurrencia liviana (consultar progreso, etc.) sin overhead
threads = 4

# Timeout extendido a 600 segundos (10 minutos) para procesar subidas y consultas masivas
timeout = 600
graceful_timeout = 60

# Keep-alive para conexiones HTTP
keepalive = 5

# Desactivar max_requests (0 = sin límite) para que Gunicorn no mate el worker ni los hilos en segundo plano
max_requests = 0
max_requests_jitter = 0

# Nivel de log
loglevel = "info"
