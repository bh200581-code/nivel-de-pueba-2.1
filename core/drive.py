"""
core/drive.py — Respaldo automático en Google Drive (OAuth 2.0 de Escritorio)
Sube documentos generados (Word/Excel/PDF) a la carpeta central institucional.
• Autenticación OAuth 2.0 (compatible con políticas de empresa que bloquean service account keys).
• Genera token.json persistente tras autorizar una sola vez en el navegador.
• Asigna permisos públicos para enlaces y códigos QR.
"""
import io
from pathlib import Path
import streamlit as st

# ═══ Carpeta central de Google Drive ═══
DRIVE_CARPETA_ID = "1xZiEK49z3P-y1a4OAP3MEMih7sWnHSXD"

MIME_WORD = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
MIME_EXCEL = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
MIME_PDF = "application/pdf"

_base = Path(__file__).resolve().parent.parent
CRED_FILE = _base / "credentials.json"
TOKEN_FILE = _base / "token.json"

SCOPES = ["https://www.googleapis.com/auth/drive"]


def _servicio():
    """Construye el servicio de Drive mediante OAuth 2.0 persistente."""
    try:
        from google.oauth2.credentials import Credentials
        from google_auth_oauthlib.flow import InstalledAppFlow
        from google.auth.transport.requests import Request
        from googleapiclient.discovery import build
    except Exception:
        raise RuntimeError(
            "Faltan librerías de Google. Instala en tu terminal:\n"
            "pip install google-api-python-client google-auth-httplib2 google-auth-oauthlib"
        )

    creds = None
    if TOKEN_FILE.exists():
        creds = Credentials.from_authorized_file(str(TOKEN_FILE), SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not CRED_FILE.exists():
                raise RuntimeError(
                    "No existe credentials.json. Descárgalo desde Google Cloud Console "
                    "(OAuth Client ID → Desktop app) y colócalo junto a main.py."
                )
            flow = InstalledAppFlow.from_client_secrets_file(str(CRED_FILE), SCOPES)
            creds = flow.run_local_server(port=0)
        
        with open(TOKEN_FILE, "w") as f:
            f.write(creds.to_json())

    return build("drive", "v3", credentials=creds)


def subir_buffer(nombre: str, buffer_o_bytes, mime: str):
    """Sube un archivo a la carpeta de Drive y devuelve (ok, webViewLink)."""
    try:
        from googleapiclient.http import MediaIoBaseUpload
        data = buffer_o_bytes.getvalue() if hasattr(buffer_o_bytes, "getvalue") else buffer_o_bytes
        service = _servicio()
        media = MediaIoBaseUpload(io.BytesIO(data), mimetype=mime, resumable=True)
        meta = {"name": nombre, "parents": [DRIVE_CARPETA_ID]}

        archivo = service.files().create(
            body=meta, media_body=media, fields="id, webViewLink"
        ).execute()

        # Otorga permisos de lectura para enlace y QR
        try:
            service.permissions().create(
                fileId=archivo.get("id"),
                body={"type": "anyone", "role": "reader"}
            ).execute()
        except Exception:
            pass

        return True, archivo.get("webViewLink", "")
    except Exception as e:
        return False, str(e)


def auto_subir(nombre: str, buffer, mime: str):
    """Intenta subir a Drive emitiendo notificación de estado."""
    try:
        ok, info = subir_buffer(nombre, buffer, mime)
        if ok:
            st.toast(f"☁️ Guardado en Drive: {nombre}", icon="✅")
        else:
            st.toast(f"⚠️ No se subió a Drive: {info}", icon="⚠️")
    except Exception as e:
        st.toast(f"⚠️ Drive no configurado: {e}", icon="⚠️")