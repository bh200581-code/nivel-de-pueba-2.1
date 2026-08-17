"""
main.py — Núcleo del Sistema de Gestión Docente ETP (v2026.08 · Azul Metálico)
• Login unificado vía core.auth + credenciales de coordinador en SQLite.
• Panel lateral con degradado azul metálico claro y texto visible.
• Pantalla de acceso profesional con fondo degradado.
• Nombre del usuario propagado al campo Docente de las herramientas.
• Resaltado azul llamativo de la herramienta activa + grupo Herramientas Comunes.
"""
import os
import base64 as _b64
import hashlib
import secrets as _secrets
import sqlite3
import streamlit as st
import core.auth as auth
import core.ia as ia

# ═══════════════════════════════════════════════════════════════════════════
# 1. CONFIGURACIÓN GLOBAL
# ═══════════════════════════════════════════════════════════════════════════
# ── Logo del sistema (favicon de la pestaña del navegador) ──
_LOGO_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logo_bh.png")

st.set_page_config(
    page_title="Sistema de Gestión Docente - ETP",
    page_icon=_LOGO_PATH if os.path.exists(_LOGO_PATH) else "🏫",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ═══════════════════════════════════════════════════════════════════════════
# 2. SESIÓN
# ═══════════════════════════════════════════════════════════════════════════
_SESSION_DEFAULTS = {
    "api_key_global": "",
    "proveedor_ia_global": "Google Gemini",
    "modelo_global": "gemini-3.5-flash",
    "modelo_custom_text": "",
    "usar_modelo_custom": False,
    "coordinador_autenticado": False,
    "docente_autenticado": False,
    "nombre_docente": "",
    "coordinador_nombre": "Ing. Bernardo Hernández",
    "coordinador_usuario": "coordinacion",
    "aviso_credenciales_legacy": False,
    "show_welcome": "",
    "usuario_display_nombre": "",
}
for _k, _v in _SESSION_DEFAULTS.items():
    st.session_state.setdefault(_k, _v)

# ═══════════════════════════════════════════════════════════════════════════
# 3. BASE DE DATOS + CREDENCIALES DE COORDINADOR
# ═══════════════════════════════════════════════════════════════════════════
DB_NAME = "gestion_etp.db"


def _hash_clave(clave: str, salt: bytes = None) -> str:
    if salt is None:
        salt = _secrets.token_bytes(16)
    dk = hashlib.pbkdf2_hmac("sha256", clave.encode("utf-8"), salt, 100000)
    return salt.hex() + "$" + dk.hex()


def _verificar_clave(clave: str, almacenado: str) -> bool:
    try:
        salt_hex, hash_hex = almacenado.split("$", 1)
        salt = bytes.fromhex(salt_hex)
        dk = hashlib.pbkdf2_hmac("sha256", clave.encode("utf-8"), salt, 100000)
        return dk.hex() == hash_hex
    except Exception:
        return False


def inicializar_db():
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute('''CREATE TABLE IF NOT EXISTS docentes (
        docente TEXT, modulo TEXT, seccion TEXT, password TEXT DEFAULT '1234', usuario TEXT
    )''')
    try:
        cursor.execute("ALTER TABLE docentes ADD COLUMN usuario TEXT")
        conn.commit()
    except Exception:
        pass
    try:
        cursor.execute("SELECT rowid, docente FROM docentes WHERE usuario IS NULL OR usuario = ''")
        for rowid, doc_name in cursor.fetchall():
            if doc_name:
                usuario_generado = str(doc_name).strip().split()[0].lower()
                cursor.execute("UPDATE docentes SET usuario=? WHERE rowid=?", (usuario_generado, rowid))
        conn.commit()
    except Exception:
        pass
    cursor.execute('''CREATE TABLE IF NOT EXISTS config_coordinador (
        id INTEGER PRIMARY KEY CHECK (id = 1),
        usuario TEXT NOT NULL,
        clave_hash TEXT NOT NULL,
        nombre TEXT,
        actualizado_en TEXT
    )''')
    conn.commit()
    conn.close()


def obtener_credenciales_coordinador():
    try:
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("SELECT usuario, clave_hash, nombre FROM config_coordinador WHERE id = 1")
        row = cursor.fetchone()
        conn.close()
        return row if row else None
    except Exception:
        return None


def guardar_credenciales_coordinador(usuario: str, clave: str, nombre: str = None):
    import datetime
    clave_hash = _hash_clave(clave)
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute('''
    INSERT INTO config_coordinador (id, usuario, clave_hash, nombre, actualizado_en)
    VALUES (1, ?, ?, ?, ?)
    ON CONFLICT(id) DO UPDATE SET
    usuario=excluded.usuario, clave_hash=excluded.clave_hash,
    nombre=excluded.nombre, actualizado_en=excluded.actualizado_en
    ''', (usuario, clave_hash, nombre, datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
    conn.commit()
    conn.close()


def autenticar_coordinador_local(usuario: str, clave: str):
    creds = obtener_credenciales_coordinador()
    if creds:
        db_usuario, db_clave_hash, db_nombre = creds
        if usuario.strip().lower() == db_usuario.strip().lower() and _verificar_clave(clave, db_clave_hash):
            return True, "sqlite", db_nombre or "Coordinación"
    ok, origen = auth.autenticar_coordinador(usuario, clave)
    if ok:
        return True, origen, "Coordinación"
    return False, None, None


inicializar_db()
auth.asegurar_esquema()

# ═══════════════════════════════════════════════════════════════════════════
# 4. HELPERS
# ═══════════════════════════════════════════════════════════════════════════
def _iniciales(nombre: str) -> str:
    partes = str(nombre).replace("Ing.", "").replace("Lic.", "").replace("M.A", "").split()
    partes = [p for p in partes if p]
    if len(partes) >= 2:
        return (partes[0][0] + partes[1][0]).upper()
    elif partes:
        return partes[0][:2].upper()
    return "??"


def _get_usuario_display() -> str:
    """Devuelve el nombre del usuario que accedió (para auto-llenar herramientas)."""
    if st.session_state.coordinador_autenticado:
        return st.session_state.coordinador_nombre
    elif st.session_state.docente_autenticado:
        return st.session_state.nombre_docente
    return ""

# ═══════════════════════════════════════════════════════════════════════════
# 5. CSS GLOBAL — AZUL METÁLICO CLARO + RESALTADO ACTIVO LLAMATIVO
# ═══════════════════════════════════════════════════════════════════════════
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800;900&display=swap');
html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
.block-container { padding-top: 1.5rem; padding-bottom: 0.5rem; max-width: 1200px; }

/* ═══ SIDEBAR — AZUL METÁLICO CLARO ═══ */
[data-testid="stSidebar"] {
background: linear-gradient(160deg, #EAF3FC 0%, #DCEAF9 25%, #CFE2F6 50%, #C3D9F3 75%, #B8D2F0 100%);
border-right: 1px solid #BFDBFE;
}
[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] { color: #1E3A5F; }
[data-testid="stSidebar"] p, [data-testid="stSidebar"] label,
[data-testid="stSidebar"] span:not([class*="badge"]) { color: #1E3A5F; }

/* Tarjeta de perfil */
.perfil-card {
background: linear-gradient(135deg, #FFFFFF 0%, #F0F7FF 100%);
border: 1px solid #BFDBFE;
border-radius: 14px;
padding: 16px;
margin-bottom: 16px;
display: flex;
align-items: center;
gap: 12px;
box-shadow: 0 4px 12px rgba(37,99,235,0.10);
}
.perfil-avatar {
width: 48px; height: 48px;
border-radius: 50%;
background: linear-gradient(135deg, #2563EB, #1D4ED8);
display: flex; align-items: center; justify-content: center;
font-weight: 800; font-size: 1.2rem; color: #FFF;
flex-shrink: 0;
box-shadow: 0 4px 10px rgba(37,99,235,0.30);
}
.perfil-info { flex: 1; min-width: 0; }
.perfil-nombre { font-weight: 700; font-size: 0.95rem; color: #0F2A5F; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.perfil-rol { font-size: 0.75rem; color: #2563EB; text-transform: uppercase; letter-spacing: 0.05em; font-weight: 600; }

/* Encabezados de sección del menú */
.sidebar-header {
font-size: 0.72rem; font-weight: 800; color: #1D4ED8;
text-transform: uppercase; letter-spacing: 0.08em;
margin: 1.2rem 0 0.4rem 0; padding-bottom: 6px;
border-bottom: 2px solid #3B82F6;
}

/* Links de página */
[data-testid="stSidebar"] [data-testid="stPageLink"] {
border-radius: 10px;
padding: 6px 10px;
margin-bottom: 2px;
color: #1E3A5F !important;
transition: all 0.2s ease;
border-left: 3px solid transparent;
}
[data-testid="stSidebar"] [data-testid="stPageLink"]:hover {
background: rgba(59,130,246,0.12);
color: #1D4ED8 !important;
}

/* ═══ RESALTADO AZUL LLAMATIVO — HERRAMIENTA ACTIVA ═══ */
[data-testid="stSidebar"] [data-testid="stPageLink"][aria-current="page"] {
background: linear-gradient(90deg, #2563EB 0%, #3B82F6 60%, #60A5FA 100%) !important;
color: #FFFFFF !important;
font-weight: 800;
border-left: 5px solid #93C5FD;
border-radius: 10px;
box-shadow: 0 4px 16px rgba(59,130,246,0.55);
}
[data-testid="stSidebar"] [data-testid="stPageLink"][aria-current="page"] p {
color: #FFFFFF !important;
font-weight: 800;
}
[data-testid="stSidebar"] [data-testid="stPageLink"][aria-current="page"] svg {
color: #FFFFFF !important;
}

/* Ayuda API key */
.api-key-help { font-size: 0.78rem; color: #334E68; line-height: 1.4; margin-top: -0.6rem; margin-bottom: 0.6rem; }
.api-key-help a { color: #1D4ED8; font-weight: 600; text-decoration: none; }
.api-key-help a:hover { text-decoration: underline; }

/* ═══ DASHBOARD ═══ */
.hero-box {
background: linear-gradient(135deg, #0A1F44 0%, #12356B 45%, #1D4ED8 100%);
border-radius: 18px; padding: 2.6rem 2rem 2rem 2rem; margin-bottom: 1.3rem;
box-shadow: 0 25px 50px rgba(10,31,68,0.25); text-align: center;
position: relative; overflow: hidden;
}
.hero-box::before {
content: ''; position: absolute; top: -50%; left: -50%; width: 200%; height: 200%;
background: radial-gradient(circle, rgba(147,197,253,0.15) 0%, transparent 60%);
animation: heroPulse 8s ease-in-out infinite;
}
@keyframes heroPulse { 0%,100%{transform:scale(1)} 50%{transform:scale(1.05)} }
.hero-title { font-size: 2.4rem; font-weight: 800; color: #FFF; letter-spacing: -0.03em; line-height: 1.15; position: relative; }
.hero-sub { font-size: 1.05rem; color: #BFDBFE; font-weight: 400; margin-top: 0.5rem; margin-bottom: 1rem; position: relative; }
.hero-badge { display:inline-block; background: rgba(147,197,253,0.20); border:1px solid rgba(147,197,253,0.40);
color:#DBEAFE; padding:.35rem .9rem; border-radius:999px; font-size:.8rem; font-weight:700; position: relative; }

.tool-card {
background: #FFFFFF; border: 2px solid #DBEAFE; border-radius: 12px;
padding: 1.1rem 1.2rem 0.6rem 1.2rem; transition: all 0.25s ease;
box-shadow: 0 2px 8px rgba(37,99,235,0.05); position: relative; height: 100%;
}
.tool-card::after {
content: '→'; position: absolute; bottom: 10px; right: 14px;
color: #93C5FD; font-weight: 700; opacity: 0; transition: all 0.2s ease;
}
.tool-icon { font-size: 1.8rem; margin-bottom: 6px; }
.tool-name { font-size: 1rem; font-weight: 700; color: #0F2A5F; margin-bottom: 3px; }
.tool-tag { display: inline-block; font-size: 0.7rem; font-weight: 600; padding: 2px 8px; border-radius: 10px; margin-bottom: 6px; }
.tag-etp { background: #DBEAFE; color: #1D4ED8; }
.tag-coord { background: #EDE9FE; color: #6D28D9; }
.tag-docente { background: #D1FAE5; color: #047857; }
.tag-common { background: #FEF3C7; color: #B45309; }
.tool-desc { font-size: 0.82rem; color: #475569; line-height: 1.45; margin-bottom: 8px; }

div[class*="st-key-card_"] { display: grid !important; margin-bottom: 0.6rem; }
div[class*="st-key-card_"] > div { grid-area: 1 / 1 / 2 / 2 !important; }
div[class*="st-key-card_"] > div:nth-child(1) { z-index: 10 !important; }
div[class*="st-key-card_"] > div:nth-child(1) div[data-testid="stButton"] { height: 100% !important; }
div[class*="st-key-card_"] > div:nth-child(1) button {
width: 100% !important; height: 100% !important; min-height: 100px !important;
opacity: 0.01 !important; cursor: pointer !important;
background: transparent !important; border: none !important;
}
div[class*="st-key-card_"] > div:nth-child(2) { z-index: 1 !important; pointer-events: none !important; }
div[class*="st-key-card_"]:has(button:hover) .tool-card {
border-color: #3B82F6 !important; box-shadow: 0 8px 24px rgba(59,130,246,0.15) !important;
transform: translateY(-3px) !important;
}
div[class*="st-key-card_"]:has(button:hover) .tool-card::after { opacity: 1 !important; transform: translateX(3px) !important; }

.app-footer { text-align:center; color:#64748B; font-size:.78rem; margin:2rem 0 1rem 0; }
</style>
""", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════
# 6. PANTALLA DE LOGIN (profesional)
# ═══════════════════════════════════════════════════════════════════════════
def pantalla_login():
    st.markdown("""
    <style>
    [data-testid='stSidebar'] {display: none !important;}
    .stApp { background: linear-gradient(135deg, #0A1F44 0%, #12356B 45%, #1D4ED8 100%); }
    [data-testid="stForm"] {
    background: rgba(255,255,255,0.98) !important;
    border-radius: 22px !important;
    padding: 38px 34px !important;
    box-shadow: 0 30px 70px rgba(10,31,68,0.45) !important;
    border: none !important;
    }
    </style>
    """, unsafe_allow_html=True)
        # ── Logo institucional (con respaldo al emoji si no existe el archivo) ──
    _logo_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logo_bh.png")
    if os.path.exists(_logo_path):
        with open(_logo_path, "rb") as _f:
            _logo_b64 = _b64.b64encode(_f.read()).decode()
        logo_html = (
            f'<img src="data:image/png;base64,{_logo_b64}" '
            'style="width:150px; height:auto; margin:0 auto 12px auto; display:block; '
            'filter: drop-shadow(0 0 20px rgba(96,165,250,0.6));" />'
        )
    else:
        logo_html = '<div style="font-size:3.6rem;">🏫</div>'

    st.markdown(f"""
     <div style="text-align:center; margin-top:5vh; margin-bottom:1.2rem;">
         {logo_html}
         <div style="font-size:1.7rem; font-weight:800; color:#FFFFFF; letter-spacing:-0.02em;">Sistema de Gestión Docente ETP</div>
         <div style="font-size:1.05rem; color:#BFDBFE; font-weight:600;">Bernardo Hernandez M.A</div>
         <div style="font-size:0.85rem; color:#93C5FD; font-style:italic;">Master en Gestión de Tecnología Educativa</div>
     </div>
     """, unsafe_allow_html=True)
    _, col_central, _ = st.columns([1, 2, 1])
    with col_central:
        with st.form("login_form_bloqueo"):
            st.markdown("<h3 style='text-align:center;color:#0A1F44;'>🔒 Digite sus credenciales</h3>", unsafe_allow_html=True)
            st.markdown(
                "<p style='text-align:center;color:#64748B;font-size:0.88rem;'>"
                "👔 Coordinación: usuario de coordinación · 🧑‍🏫 Docentes: su usuario y contraseña personal.</p>",
                unsafe_allow_html=True,
            )
            usuario_input = st.text_input("Usuario", placeholder="Digite su Usuario")
            clave_input = st.text_input("Contraseña", type="password", placeholder="Digite su contraseña")
            submit_btn = st.form_submit_button("INICIAR SESIÓN", width="stretch")
            if submit_btn:
                if not usuario_input or not clave_input:
                    st.error("⚠️ Por favor complete ambos campos.")
                else:
                    ok_coord, origen, nombre_coord = autenticar_coordinador_local(usuario_input, clave_input)
                    if ok_coord:
                        st.session_state.coordinador_autenticado = True
                        st.session_state.coordinador_usuario = usuario_input.strip().lower()
                        if nombre_coord and nombre_coord != "Coordinación":
                            st.session_state.coordinador_nombre = nombre_coord
                        st.session_state.usuario_display_nombre = st.session_state.coordinador_nombre
                        st.session_state.aviso_credenciales_legacy = (origen == "legacy")
                        st.session_state.show_welcome = "👔 Modo coordinación activado"
                        auth.registrar_evento("Login coordinador", usuario_input.strip().lower(), f"origen={origen}")
                        st.rerun()
                    else:
                        docente = auth.autenticar_docente(usuario_input, clave_input)
                        if docente:
                            st.session_state.docente_autenticado = True
                            st.session_state.nombre_docente = docente
                            st.session_state.usuario_display_nombre = docente
                            st.session_state.show_welcome = f"👋 Bienvenido/a, {docente}"
                            auth.registrar_evento("Login docente", docente, "pantalla principal")
                            st.rerun()
                        else:
                            st.error("❌ Usuario o contraseña incorrectos. Verifique e intente de nuevo.")

# ═══════════════════════════════════════════════════════════════════════════
# 7. CONTROL DE ACCESO
# ═══════════════════════════════════════════════════════════════════════════
if st.query_params.get("prueba"):
    _pagina_prueba = st.Page("prueba_diagnostica.py", title="Prueba Diagnóstica", icon="📝")
    st.navigation([_pagina_prueba]).run()
    st.stop()

if st.query_params.get("juego"):
    _pagina_juego = st.Page("juegos_interactivos.py", title="Juego Interactivo", icon="🎮")
    st.navigation([_pagina_juego]).run()
    st.stop()

# ── NUEVA RUTA PARA LAS EVALUACIONES DEL MAESTRO MERITORIO SIN LOGIN ──
if st.query_params.get("eval"):
    _pagina_eval = st.Page("maestro_meritorio.py", title="Evaluación Meritorio", icon="🏆")
    st.navigation([_pagina_eval]).run()
    st.stop()

if not (st.session_state.coordinador_autenticado or st.session_state.docente_autenticado):
    pg_login = st.Page(pantalla_login, title="Autenticación Requerida", icon="🔒")
    st.navigation([pg_login]).run()
    st.stop()

# ═══════════════════════════════════════════════════════════════════════════
# 8. SIDEBAR — AZUL METÁLICO CLARO
# ═══════════════════════════════════════════════════════════════════════════
_img_path = "Gemini_Generated_Image_5ck0tc5ck0tc5ck0.png"
if os.path.exists(_img_path):
    with open(_img_path, "rb") as _f:
        _img_b64 = _b64.b64encode(_f.read()).decode()
    st.markdown(f"""
    <style>
    [data-testid="stSidebarContent"]::before {{
    content: ''; display: block; height: 85px;
    background-image: url("data:image/png;base64,{_img_b64}");
    background-size: contain; background-repeat: no-repeat;
    background-position: center; margin: 8px 12px 4px 12px;
    border-radius: 10px; box-shadow: 0 4px 12px rgba(37,99,235,0.15);
    }}
    </style>
    """, unsafe_allow_html=True)

with st.sidebar:
    # ── TARJETA DE PERFIL ──
    _nombre_perfil = _get_usuario_display()
    _rol_perfil = "👔 Coordinador ETP" if st.session_state.coordinador_autenticado else "🧑‍🏫 Docente"
    st.markdown(f"""
    <div class="perfil-card">
    <div class="perfil-avatar">{_iniciales(_nombre_perfil)}</div>
    <div class="perfil-info">
    <div class="perfil-nombre">{_nombre_perfil}</div>
    <div class="perfil-rol">{_rol_perfil}</div>
    </div>
    </div>
    """, unsafe_allow_html=True)
    # ── CONFIGURACIÓN DE IA ──
    st.markdown('<p class="sidebar-header">⚙️ Configuración de IA</p>', unsafe_allow_html=True)
    PROVEEDORES = ["Google Gemini", "OpenAI (ChatGPT)", "Anthropic (Claude)"]
    prov_actual = st.session_state.proveedor_ia_global
    proveedor = st.selectbox(
        "🧠 Proveedor", PROVEEDORES,
        index=PROVEEDORES.index(prov_actual) if prov_actual in PROVEEDORES else 0,
        key="prov_global_nav",
    )
    st.session_state.proveedor_ia_global = proveedor
    catalogo = ia.catalogo_modelos(proveedor)
    modelo_actual = st.session_state.modelo_global
    if not st.session_state.usar_modelo_custom and modelo_actual not in catalogo:
        st.session_state.modelo_global = catalogo[0]
        modelo_actual = catalogo[0]
    usar_custom = st.checkbox(
        "✏️ Usar modelo personalizado", value=st.session_state.usar_modelo_custom,
        key="usar_custom_nav", help="Escribe cualquier modelo de los 3 proveedores.",
    )
    st.session_state.usar_modelo_custom = usar_custom
    if usar_custom:
        custom_txt = st.text_input(
            "Nombre exacto del modelo", value=st.session_state.modelo_custom_text, key="modelo_custom_nav",
        )
        st.session_state.modelo_custom_text = custom_txt
        if custom_txt.strip():
            st.session_state.modelo_global = custom_txt.strip()
    else:
        st.session_state.modelo_custom_text = ""
        modelo = st.selectbox(
            "🤖 Modelo", catalogo,
            index=catalogo.index(modelo_actual) if modelo_actual in catalogo else 0,
            key="modelo_global_nav",
        )
        st.session_state.modelo_global = modelo
    api_key = st.text_input("API Key:", type="password", value=st.session_state.api_key_global, key="api_key_global_nav")
    st.session_state.api_key_global = api_key
    _AYUDAS = {
        "Google Gemini": ("Google Gemini", "https://aistudio.google.com/app/apikey", "Google AI Studio"),
        "OpenAI (ChatGPT)": ("OpenAI (ChatGPT)", "https://platform.openai.com/api-keys", "OpenAI Platform"),
        "Anthropic (Claude)": ("Anthropic (Claude)", "https://console.anthropic.com/settings/keys", "Anthropic Console"),
    }
    nombre, url, sitio = _AYUDAS[proveedor]
    st.markdown(
        f'<div class="api-key-help">🔑 Clave que autoriza el uso de <b>{nombre}</b> en esta app. '
        f'No se comparte ni se guarda fuera de tu sesión. '
        f'<a href="{url}" target="_blank">Generar mi API Key en {sitio} →</a></div>',
        unsafe_allow_html=True,
    )
    # ── CONFIGURACIÓN DEL COORDINADOR ──
    if st.session_state.coordinador_autenticado:
        st.markdown('<p class="sidebar-header">🔐 Mi Cuenta de Coordinador</p>', unsafe_allow_html=True)
        with st.expander("👤 Cambiar usuario / contraseña", expanded=False):
            st.caption("Actualiza tus credenciales de acceso. Se guardan de forma segura (hash PBKDF2).")
            nuevo_usuario_coord = st.text_input("Nuevo usuario", value=st.session_state.coordinador_usuario, key="nuevo_usuario_coord")
            nueva_clave_coord = st.text_input("Nueva contraseña", type="password", key="nueva_clave_coord")
            confirma_clave_coord = st.text_input("Confirmar contraseña", type="password", key="confirma_clave_coord")
            if st.button("💾 Guardar credenciales", width="stretch", key="btn_guardar_cred_coord"):
                if not nuevo_usuario_coord.strip():
                    st.error("⚠️ El usuario no puede estar vacío.")
                elif len(nueva_clave_coord) < 6:
                    st.error("⚠️ La contraseña debe tener al menos 6 caracteres.")
                elif nueva_clave_coord != confirma_clave_coord:
                    st.error("⚠️ Las contraseñas no coinciden.")
                else:
                    guardar_credenciales_coordinador(nuevo_usuario_coord.strip(), nueva_clave_coord, st.session_state.coordinador_nombre)
                    st.session_state.coordinador_usuario = nuevo_usuario_coord.strip().lower()
                    auth.registrar_evento("Cambio credenciales coordinador", nuevo_usuario_coord.strip().lower(), "actualizado por el usuario")
                    st.success("✅ Credenciales actualizadas. Úsalas en tu próximo login.")
    # ── CERRAR SESIÓN ──
    st.markdown("---")
    if st.button("🚪 Cerrar Sesión", width="stretch", key="btn_cerrar_sesion"):
        st.session_state.coordinador_autenticado = False
        st.session_state.docente_autenticado = False
        st.session_state.nombre_docente = ""
        st.session_state.usuario_display_nombre = ""
        st.rerun()

# ═══════════════════════════════════════════════════════════════════════════
# 9. DASHBOARD
# ═══════════════════════════════════════════════════════════════════════════
def _metrics_coordinacion():
    try:
        docentes = len(auth.listar_docentes())
    except Exception:
        docentes = 0
    try:
        asignaciones = len(auth.listar_asignaciones())
    except Exception:
        asignaciones = 0

    def _count(tabla, where=""):
        try:
            conn = sqlite3.connect(DB_NAME)
            cur = conn.cursor()
            cur.execute(f"SELECT COUNT(*) FROM {tabla} {where}")
            n = cur.fetchone()[0]
            conn.close()
            return int(n)
        except Exception:
            return 0
    return docentes, asignaciones, _count("alertas", "WHERE estado IN ('Roja','Crítico')"), _count("acuerdos")


def tarjeta_clickable(icon, name, tag_class, tag_label, desc, page, key):
    with st.container(key=f"card_{key}"):
        if st.button(" ", key=f"btn_{key}", width="stretch"):
            st.switch_page(page)
        st.markdown(f"""
        <div class="tool-card">
        <div class="tool-icon">{icon}</div>
        <div class="tool-name">{name}</div>
        <div class="tool-tag {tag_class}">{tag_label}</div>
        <div class="tool-desc">{desc}</div>
        </div>
        """, unsafe_allow_html=True)


def _filtra(herramientas, texto):
    if not texto:
        return herramientas
    t = texto.lower()
    return [h for h in herramientas if t in h["name"].lower() or t in h["desc"].lower()]


def pagina_inicio():
    if st.session_state.show_welcome:
        st.toast(st.session_state.show_welcome, icon="👋")
        st.session_state.show_welcome = ""
    st.markdown(f"""
    <div class="hero-box">
    <div class="hero-title">Sistema de Gestión Docente ETP</div>
    <div class="hero-sub">Automatización pedagógica con Inteligencia Artificial — Alineado al MINERD</div>
    <div class="hero-badge">🚀 Potenciado por IA &nbsp;·&nbsp; 🏫 Modalidad Técnico Profesional</div>
    </div>
    """, unsafe_allow_html=True)
    busqueda = st.text_input(
        "🔎 Buscar herramienta", placeholder="Ej: ponderación, actas, simulador, calificaciones…",
        label_visibility="collapsed", key="buscador_herramientas",
    )
    if st.session_state.coordinador_autenticado and not busqueda:
        d, a, al, ac = _metrics_coordinacion()
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("🧑‍ Docentes", d)
        m2.metric("📚 Asignaciones", a)
        m3.metric("🚨 Alertas rojas", al)
        m4.metric("🤝 Acuerdos", ac)
    if st.session_state.docente_autenticado:
        st.markdown("#### 🎓 Mi Portal Privado")
        col_doc1, col_doc2, col_doc3 = st.columns(3)
        with col_doc1:
            tarjeta_clickable(
                "📝", "Mis Calificaciones ETP", "tag-docente", "DOCENTE",
                "Gestiona tus Resultados de Aprendizaje y notas.",
                "portal_docente.py", key="portal_docente",
            )
        st.markdown("<br>", unsafe_allow_html=True)
    if st.session_state.coordinador_autenticado:
        herramientas_coord = [
            {"icon": "🎛️", "name": "Sala de Situación", "page": "dashboard_coordinacion.py", "desc": "Panel gerencial con métricas clave de coordinación."},
            {"icon": "👁️", "name": "Auditor de Calificaciones", "page": "visor_calificaciones.py", "desc": "Monitoreo y descarga de sábanas de notas de todos los docentes."},
            {"icon": "🔐", "name": "Gestor de Accesos", "page": "gestor_accesos.py", "desc": "Creación de usuarios y reseteo de contraseñas de docentes."},
            {"icon": "✅", "name": "Auditor de Planificaciones", "page": "evaluador_planes.py", "desc": "Evaluación y retroalimentación de planes de unidad y diarios."},
            {"icon": "🧾", "name": "Auditor de Matriz Curricular", "page": "auditor_matriz.py", "desc": "Evaluación de la matriz curricular para asegurar calidad y coherencia."},
            {"icon": "📋", "name": "Acompañamiento Docente", "page": "acompanamiento.py", "desc": "Soporte y orientación continua para el desarrollo profesional del docente."},
            {"icon": "🗓️", "name": "Cronograma de Acompañamiento", "page": "cronograma_acompanamiento.py", "desc": "Gestión de visitas y monitoreo para docentes ETP."},
            {"icon": "⚖️", "name": "Gestor de Acuerdos", "page": "acuerdos.py", "desc": "Gestión y seguimiento de acuerdos institucionales."},
            {"icon": "🧑‍🏫", "name": "Directorio de Docentes", "page": "directorio_docentes.py", "desc": "Acceso al directorio completo de carga horaria y docentes."},
            {"icon": "☁️", "name": "Repositorio Drive", "page": "repositorio_drive.py", "desc": "Acceso directo a la documentación y actas institucionales en la nube."},
            {"icon": "🚨", "name": "Registro de Incidencias", "page": "incidencias.py", "desc": "Registro y gestión de incidencias académicas y administrativas."},
            {"icon": "🏆", "name": "Selección Maestro Meritorio", "page": "maestro_meritorio.py", "desc": "Evaluación y selección de docentes meritorios."},
        ]
        lista = _filtra(herramientas_coord, busqueda)
        if lista:
            st.markdown("#### 👔 Herramientas de Coordinación Pedagógica ETP")
            cols = st.columns(3)
            for i, t in enumerate(lista):
                with cols[i % 3]:
                    tarjeta_clickable(t["icon"], t["name"], "tag-coord", "COORD", t["desc"], t["page"], key=t["page"])
            st.markdown("<br>", unsafe_allow_html=True)
    herramientas_etp = [
        {"icon": "📊", "name": "Ponderación RA", "page": "ponderacionra.py", "desc": "Distribución porcentual y temporal de Resultados de Aprendizaje."},
        {"icon": "🚀", "name": "Planificación Modular", "page": "planifiacionra.py", "desc": "Matriz de planificación por RA con Elementos de Capacidad."},
        {"icon": "📅", "name": "Plan Diario ETP", "page": "pladiario.py", "desc": "Planificación de clase diaria en 50 minutos con lista de cotejo."},
        {"icon": "📚", "name": "Generador de Contenidos", "page": "contenido.py", "desc": "Contenido anclado, progresión Bloom y rúbrica multinivel."},
        {"icon": "📖", "name": "Generador de Libros/Guías", "page": "generador_guia.py", "desc": "Creación de libros y guías educativas."},
        {"icon": "✒️", "name": "Redactor Profundo", "page": "redactor_capitulos.py", "desc": "Redacción técnica profunda de capítulos y temas con analogías."},
        {"icon": "💻", "name": "Fábrica de Simuladores", "page": "simuladores.py", "desc": "Simuladores web interactivos personalizados."},
        {"icon": "📝", "name": "Banco de Ítems", "page": "bancoitems.py", "desc": "Pruebas diversificadas: opción múltiple, completar, clasificación."},
        {"icon": "🚨", "name": "Recuperación R.A o Pedagógica", "page": "alerta.py", "desc": "Diagnóstico de severidad y plan de recuperación."},
        
    ]
    lista_etp = _filtra(herramientas_etp, busqueda)
    if lista_etp:
        st.markdown("#### 🔧 ETP — Talleres y Módulos Formativos")
        cols_etp = st.columns(3)
        for i, t in enumerate(lista_etp):
            with cols_etp[i % 3]:
                tarjeta_clickable(t["icon"], t["name"], "tag-etp", "ETP", t["desc"], t["page"], key=t["page"])
    # Herramientas comunes (coordinación + docentes)
    herramientas_comunes = [
        {"icon": "🩺", "name": "Pruebas Diagnósticas", "page": "prueba_diagnostica.py", "desc": "Evaluación diagnóstica inicial para estudiantes."},
        {"icon": "🎨", "name": "Fábrica Visual IA", "page": "presentaciones_ia.py", "desc": "Presentaciones PowerPoint e infografías generadas con IA."},
        {"icon": "🎮", "name": "Fábrica de Juegos", "page": "juegos_interactivos.py","desc": "Juegos interactivos generados por IA para compartir con estudiantes."},
        {"icon": "⭐", "name": "Feedback del Portal", "page": "feedback.py", "desc": "Valoración y comentarios sobre el portal."},

    ]
    lista_comunes = _filtra(herramientas_comunes, busqueda)
    if lista_comunes:
        st.markdown("#### 🤝 Herramientas Comunes")
        cols_com = st.columns(3)
        for i, t in enumerate(lista_comunes):
            with cols_com[i % 3]:
                tarjeta_clickable(t["icon"], t["name"], "tag-common", "COMÚN", t["desc"], t["page"], key=t["page"])
    if busqueda and not lista_etp and not lista_comunes and not (
        st.session_state.coordinador_autenticado and _filtra(herramientas_coord, busqueda)
    ):
        st.info("🔎 Sin resultados para tu búsqueda. Prueba con otro término.")
    st.markdown('<div class="app-footer">© 2026 Bernardo Hernández — Arquitecto de Futuros Digitales · Software & IA</div>', unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════
# 10. NAVEGACIÓN DINÁMICA
# ═══════════════════════════════════════════════════════════════════════════
inicio = st.Page(pagina_inicio, title="Inicio", icon="🏠", default=True)
# Herramientas ETP (docentes)
pagina_ponderacion = st.Page("ponderacionra.py", title="Ponderación RA", icon="📊")
pagina_planificacion = st.Page("planifiacionra.py", title="Planificación Modular", icon="🚀")
pagina_pladiario = st.Page("pladiario.py", title="Plan Diario ETP", icon="📅")
pagina_contenido = st.Page("contenido.py", title="Generador de Contenidos", icon="📚")
pagina_libro = st.Page("generador_guia.py", title="Generador de Libros/Guías", icon="📖")
pagina_redactor = st.Page("redactor_capitulos.py", title="Redactor Profundo", icon="✒️")
pagina_simuladores = st.Page("simuladores.py", title="Fábrica de Simuladores", icon="💻")
pagina_banco = st.Page("bancoitems.py", title="Banco de Ítems", icon="📝")
pagina_alerta = st.Page("alerta.py", title="Recuperación R.A", icon="🚨")
# Herramientas de coordinación
pagina_dashboard = st.Page("dashboard_coordinacion.py", title="Sala de Situación", icon="🎛️")
pagina_acompanamiento = st.Page("acompanamiento.py", title="Acompañamiento", icon="📋")
pagina_auditor_matriz = st.Page("auditor_matriz.py", title="Auditor de Matriz", icon="🧾")
pagina_evaluador = st.Page("evaluador_planes.py", title="Auditor de Plan", icon="✅")
pagina_cronograma = st.Page("cronograma_acompanamiento.py", title="Cronograma", icon="🗓️")
pagina_acuerdos = st.Page("acuerdos.py", title="Gestor de Acuerdos", icon="⚖️")
pagina_incidencias = st.Page("incidencias.py", title="Registro Incidencias", icon="🚨")
pagina_docentes = st.Page("directorio_docentes.py", title="Directorio Docentes", icon="🧑‍🏫")
pagina_visor_calif = st.Page("visor_calificaciones.py", title="Auditor Calificaciones", icon="👁️")
pagina_gestor_accesos = st.Page("gestor_accesos.py", title="Gestor de Accesos", icon="🔐")
pagina_drive = st.Page("repositorio_drive.py", title="Repositorio Drive", icon="☁️")
pagina_meritorio = st.Page("maestro_meritorio.py", title="Maestro Meritorio", icon="🏆")
# Pruebas y portal docente
pagina_prueba_diagnostica = st.Page("prueba_diagnostica.py", title="Pruebas Diagnósticas", icon="🩺")
pagina_portal_docente = st.Page("portal_docente.py", title="Mis Calificaciones ETP", icon="📝")
# Áreas académicas
pagina_academicas = st.Page("academicas.py", title="Plan Unidad (Acad)", icon="📖")
pagina_diario_acad = st.Page("diario_academico.py", title="Plan Diario (Acad)", icon="🗓️")
# Herramientas comunes (NUEVO)
pagina_presentaciones = st.Page("presentaciones_ia.py", title="Fábrica Visual IA", icon="🎨")
pagina_juegos = st.Page("juegos_interactivos.py", title="Fábrica de Juegos", icon="🎮")
pagina_feedback = st.Page("feedback.py", title="Feedback", icon="⭐")

# ═══ MENÚ ORGANIZADO EN 4 GRUPOS + HERRAMIENTAS COMUNES ═══
diccionario_menu = {"🏠 Principal": [inicio]}
if st.session_state.coordinador_autenticado:
    diccionario_menu["👔 Coordinación Pedagógica ETP"] = [
        pagina_dashboard, pagina_visor_calif, pagina_gestor_accesos,
        pagina_evaluador, pagina_auditor_matriz, pagina_acompanamiento,
        pagina_cronograma, pagina_acuerdos, pagina_incidencias,
        pagina_docentes, pagina_drive, pagina_meritorio,
    ]
if st.session_state.docente_autenticado:
    diccionario_menu["🎓 Mi Portal Docente"] = [pagina_portal_docente]
diccionario_menu["🔧 Docentes — Talleres y Módulos"] = [
    pagina_ponderacion, pagina_planificacion, pagina_pladiario,
    pagina_contenido, pagina_libro, pagina_redactor, pagina_simuladores,
    pagina_banco, pagina_alerta,
]
diccionario_menu["📚 Áreas Académicas"] = [pagina_academicas, pagina_diario_acad]
diccionario_menu["🤝 Herramientas Comunes"] = [pagina_prueba_diagnostica, pagina_presentaciones, pagina_juegos, pagina_feedback]

todas_las_paginas = [pagina for grupo in diccionario_menu.values() for pagina in grupo]
menu = st.navigation(todas_las_paginas, position="hidden")


def _bloque_menu(titulo: str, paginas: list) -> None:
    st.markdown(f'<p class="sidebar-header">{titulo}</p>', unsafe_allow_html=True)
    for pagina in paginas:
        st.page_link(pagina)


with st.sidebar:
    for titulo, paginas in diccionario_menu.items():
        _bloque_menu(titulo, paginas)

menu.run()