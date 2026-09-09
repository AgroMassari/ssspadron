"""
whatsapp_api.py
===============
Módulo auxiliar para interactuar con la API Oficial de WhatsApp (Meta Cloud API).

Funcionalidades:
- Verificación del webhook (handshake de Meta)
- Extracción de mensajes entrantes del payload JSON
- Envío de mensajes de texto de vuelta al usuario
"""

import os
import logging
import requests

logger = logging.getLogger(__name__)

# ============================================================
# CONFIGURACIÓN (variables de entorno)
# ============================================================

def _get_whatsapp_token() -> str:
    token = os.environ.get("WHATSAPP_TOKEN")
    if not token:
        raise RuntimeError("Falta la variable de entorno WHATSAPP_TOKEN.")
    return token


def _get_phone_number_id() -> str:
    phone_id = os.environ.get("WHATSAPP_PHONE_NUMBER_ID")
    if not phone_id:
        raise RuntimeError("Falta la variable de entorno WHATSAPP_PHONE_NUMBER_ID.")
    return phone_id


def _get_verify_token() -> str:
    return os.environ.get("WHATSAPP_VERIFY_TOKEN", "ssspadron_verify_token")


# ============================================================
# VERIFICACIÓN DEL WEBHOOK (GET)
# ============================================================

def verificar_webhook(args: dict) -> tuple:
    """
    Valida el handshake que hace Meta al configurar el webhook.
    Retorna (challenge, status_code).
    """
    mode      = args.get("hub.mode")
    token     = args.get("hub.verify_token")
    challenge = args.get("hub.challenge")

    if mode == "subscribe" and token == _get_verify_token():
        logger.info("Webhook verificado correctamente.")
        return challenge, 200

    logger.warning("Verificación de webhook fallida. Token incorrecto.")
    return "Verificación fallida", 403


# ============================================================
# PARSEAR MENSAJES ENTRANTES (POST)
# ============================================================

def extraer_mensaje(payload: dict) -> dict | None:
    """
    Extrae el mensaje (texto, imagen o documento PDF) y el número de teléfono del payload JSON de WhatsApp.
    Retorna un dict con { "numero": str, "tipo": str, ... } o None si no hay mensaje válido.
    """
    try:
        entry   = payload.get("entry", [{}])[0]
        changes = entry.get("changes", [{}])[0]
        value   = changes.get("value", {})
        mensaje = value.get("messages", [{}])[0]

        if not mensaje:
            return None

        numero     = mensaje.get("from")
        message_id = mensaje.get("id")
        tipo       = mensaje.get("type")

        if not numero:
            return None

        if tipo == "text":
            texto = mensaje.get("text", {}).get("body", "").strip()
            if not texto:
                return None
            return {
                "numero":     numero,
                "tipo":       "text",
                "texto":      texto,
                "message_id": message_id,
            }

        elif tipo == "image":
            img_obj   = mensaje.get("image", {})
            media_id  = img_obj.get("id")
            mime_type = img_obj.get("mime_type", "image/jpeg")
            caption   = img_obj.get("caption", "").strip()
            return {
                "numero":     numero,
                "tipo":       "image",
                "media_id":   media_id,
                "mime_type":  mime_type,
                "filename":   "foja.jpg",
                "texto":      caption,
                "message_id": message_id,
            }

        elif tipo == "document":
            doc_obj   = mensaje.get("document", {})
            media_id  = doc_obj.get("id")
            mime_type = doc_obj.get("mime_type", "application/pdf")
            filename  = doc_obj.get("filename", "foja.pdf")
            caption   = doc_obj.get("caption", "").strip()
            return {
                "numero":     numero,
                "tipo":       "document",
                "media_id":   media_id,
                "mime_type":  mime_type,
                "filename":   filename,
                "texto":      caption,
                "message_id": message_id,
            }

        return None

    except (IndexError, KeyError, TypeError):
        return None


# ============================================================
# DESCARGAR ARCHIVO MULTIMEDIA (IMAGEN O PDF)
# ============================================================

def descargar_media(media_id: str) -> tuple[bytes | None, str | None]:
    """
    Descarga el contenido multimedia de WhatsApp (imagen o PDF) mediante la Graph API de Meta.
    Retorna (bytes_contenido, mime_type) o (None, None) en caso de fallo.
    """
    if not media_id:
        return None, None

    try:
        token    = _get_whatsapp_token()
        headers  = {"Authorization": f"Bearer {token}"}
        url_info = f"https://graph.facebook.com/v20.0/{media_id}"

        # 1. Obtener la URL temporal de descarga desde Meta
        r_info = requests.get(url_info, headers=headers, timeout=15)
        if r_info.status_code != 200:
            logger.error("Error al consultar URL del medio en Meta (%s): %s", r_info.status_code, r_info.text)
            return None, None

        datos        = r_info.json()
        download_url = datos.get("url")
        mime_type    = datos.get("mime_type")

        if not download_url:
            logger.error("No se encontró URL de descarga en la respuesta de Meta: %s", datos)
            return None, None

        # 2. Descargar el archivo binario
        r_file = requests.get(download_url, headers=headers, timeout=30)
        if r_file.status_code == 200:
            return r_file.content, mime_type
        else:
            logger.error("Error al descargar archivo binario de Meta (%s)", r_file.status_code)
            return None, None

    except Exception as e:
        logger.error("Excepción al descargar multimedia de WhatsApp: %s", e)
        return None, None


# ============================================================
# ENVIAR MENSAJE DE TEXTO
# ============================================================

def enviar_texto(numero: str, texto: str) -> bool:
    """
    Envía un mensaje de texto al número indicado a través de la API de Meta.
    Devuelve True si fue exitoso, False en caso contrario.
    """
    token    = _get_whatsapp_token()
    phone_id = _get_phone_number_id()

    url     = f"https://graph.facebook.com/v20.0/{phone_id}/messages"
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type":  "application/json",
    }
    payload = {
        "messaging_product": "whatsapp",
        "recipient_type":    "individual",
        "to":                numero,
        "type":              "text",
        "text":              {"body": texto},
    }

    try:
        respuesta = requests.post(url, headers=headers, json=payload, timeout=15)
        if respuesta.status_code == 200:
            logger.info("Mensaje enviado a %s.", numero)
            return True
        else:
            logger.error(
                "Error al enviar mensaje. Status: %s | Body: %s",
                respuesta.status_code,
                respuesta.text,
            )
            return False
    except requests.RequestException as e:
        logger.error("Excepción al enviar mensaje: %s", e)
        return False


# ============================================================
# MARCAR MENSAJE COMO LEÍDO (opcional, mejora la UX)
# ============================================================

def marcar_leido(message_id: str) -> None:
    """Marca el mensaje de WhatsApp como leído (doble tilde azul)."""
    try:
        token    = _get_whatsapp_token()
        phone_id = _get_phone_number_id()

        url = f"https://graph.facebook.com/v20.0/{phone_id}/messages"
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type":  "application/json",
        }
        payload = {
            "messaging_product": "whatsapp",
            "status":            "read",
            "message_id":        message_id,
        }
        requests.post(url, headers=headers, json=payload, timeout=10)
    except Exception:
        pass  # No crítico
