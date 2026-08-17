"""
repositorio_drive.py — Repositorio Institucional Drive (MÁXIMO NIVEL)
Explorador de documentos institucionales con gestión dinámica de carpetas,
auditoría de accesos, vistas Grid/Lista y guía de buenas prácticas.
• Persistencia SQLite (tabla drive_carpetas).
• Auditoría vía core.auth (un acceso por carpeta por sesión).
• Interfaz consistente con el resto del sistema.
"""
import re
import sqlite3
from datetime import datetime

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

# Auditoría opcional (no rompe si core.auth no está disponible)
try:
    import core.auth as auth
    AUDITORIA_DISPONIBLE = True
except Exception:
    auth = None
    AUDITORIA_DISPONIBLE = False

# ═══════════════════════════════════════════════════════════════════════════
# BASE DE DATOS — Gestión de carpetas
# ═══════════════════════════════════════════════════════════════════════════
DB_NAME = "gestion_etp.db"
ID_CARPETA_PRINCIPAL = "1kH1UNEth2VGqWnlOxnmsjuxX3HkvhUSi"
URL_APPS_SCRIPT = "https://script.google.com/macros/s/AKfycbxNIXgc6a431999vhWd_3LV8p4J96QW0UNG_7M0bZjSPedhixKs6ZkzhZNaR2hspwo8/exec"

CATEGORIAS_SUGERIDAS = [
    "General", "Actas", "Planificaciones", "Directorio",
    "Evaluaciones", "Acuerdos", "Incidencias", "Cronogramas",
]


def init_drive_db():
    conn = sqlite3.connect(DB_NAME, check_same_thread=False)
    cursor = conn.cursor()
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS drive_carpetas (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nombre TEXT NOT NULL,
        drive_id TEXT NOT NULL,
        descripcion TEXT,
        categoria TEXT
    )
    ''')
    conn.commit()
    cursor.execute("SELECT COUNT(*) FROM drive_carpetas")
    if cursor.fetchone()[0] == 0:
        cursor.execute('''
        INSERT INTO drive_carpetas (nombre, drive_id, descripcion, categoria)
        VALUES (?, ?, ?, ?)
        ''', (
            "Repositorio General",
            ID_CARPETA_PRINCIPAL,
            "Repositorio completo de documentación, actas y archivos institucionales",
            "General",
        ))
        conn.commit()
    return conn


conn = init_drive_db()


def obtener_carpetas():
    cursor = conn.cursor()
    cursor.execute("SELECT id, nombre, drive_id, descripcion, categoria FROM drive_carpetas ORDER BY categoria, nombre")
    rows = cursor.fetchall()
    return [dict(zip(["id", "nombre", "drive_id", "descripcion", "categoria"], r)) for r in rows]


def agregar_carpeta(nombre, drive_id, descripcion, categoria):
    cursor = conn.cursor()
    cursor.execute('''
    INSERT INTO drive_carpetas (nombre, drive_id, descripcion, categoria)
    VALUES (?, ?, ?, ?)
    ''', (nombre, drive_id, descripcion, categoria))
    conn.commit()


def eliminar_carpeta(carpeta_id):
    cursor = conn.cursor()
    cursor.execute("DELETE FROM drive_carpetas WHERE id = ?", (carpeta_id,))
    conn.commit()


def validar_drive_id(drive_id):
    """Valida que el ID tenga el formato típico de Google Drive (25-60 caracteres alfanuméricos)."""
    drive_id = drive_id.strip()
    if not drive_id:
        return False, "El ID no puede estar vacío."
    if len(drive_id) < 25 or len(drive_id) > 60:
        return False, "El ID debe tener entre 25 y 60 caracteres."
    if not re.match(r'^[a-zA-Z0-9_-]+$', drive_id):
        return False, "El ID solo puede contener letras, números, guiones y guiones bajos."
    return True, "ID válido."


def contar_accesos_repositorio():
    if not AUDITORIA_DISPONIBLE:
        return 0
    try:
        c = sqlite3.connect(DB_NAME)
        cur = c.cursor()
        cur.execute("SELECT COUNT(*) FROM accesos_log WHERE accion = 'Acceso Repositorio'")
        n = cur.fetchone()[0]
        c.close()
        return n
    except Exception:
        return 0


def registrar_acceso(nombre_carpeta):
    if not AUDITORIA_DISPONIBLE:
        return
    try:
        actor = (
            st.session_state.get("coordinador_nombre")
            or st.session_state.get("nombre_docente")
            or "Usuario"
        )
        auth.registrar_evento("Acceso Repositorio", actor, f"Carpeta abierta: {nombre_carpeta}")
    except Exception:
        pass


def url_embed(drive_id, vista="grid"):
    return f"https://drive.google.com/embeddedfolderview?id={drive_id}#{vista}"


def url_externa(drive_id):
    return f"https://drive.google.com/drive/folders/{drive_id}"


# ═══════════════════════════════════════════════════════════════════════════
# ESTILOS
# ═══════════════════════════════════════════════════════════════════════════
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800;900&display=swap');

html, body, [class*="css"] {
    font-family: 'Inter', sans-serif;
    background-color: #F0F4F8;
    color: #1E293B;
}

.drive-hero {
    background: linear-gradient(135deg, #0F172A 0%, #0E7490 45%, #06B6D4 80%, #67E8F9 100%);
    color: #fff;
    padding: 2.2rem;
    border-radius: 20px;
    margin-bottom: 1.5rem;
    box-shadow: 0 25px 50px rgba(8, 145, 178, 0.3);
    position: relative;
    overflow: hidden;
}

.drive-hero::before {
    content: '';
    position: absolute;
    top: -50%; left: -50%; width: 200%; height: 200%;
    background: radial-gradient(circle, rgba(255,255,255,0.12) 0%, transparent 60%);
    animation: drivePulse 6s ease-in-out infinite;
}

@keyframes drivePulse {
    0%, 100% { transform: scale(1); opacity: 0.5; }
    50% { transform: scale(1.1); opacity: 0.8; }
}

.drive-hero-title { font-size: 2.4rem; font-weight: 900; letter-spacing: -0.03em; margin-bottom: 0.4rem; position: relative; }
.drive-hero-sub { font-size: 1.05rem; opacity: 0.9; line-height: 1.5; position: relative; }

.drive-hero-badge {
    display: inline-block;
    background: rgba(255,255,255,0.15);
    border: 1px solid rgba(255,255,255,0.25);
    border-radius: 8px;
    padding: 4px 12px;
    font-size: 0.8rem;
    font-weight: 600;
    margin-top: 0.8rem;
    margin-right: 8px;
    position: relative;
}

.drive-section-title {
    color: #0E7490;
    font-weight: 700;
    font-size: 1.12rem;
    border-bottom: 2px solid #CFFAFE;
    padding-bottom: 8px;
    margin: 1.2rem 0 0.9rem 0;
}

.drive-kpi-card {
    background: #fff;
    border-radius: 14px;
    padding: 1.2rem;
    text-align: center;
    border: 1px solid #E2E8F0;
    box-shadow: 0 4px 12px rgba(0,0,0,0.06);
    transition: all 0.3s ease;
}

.drive-kpi-card:hover { transform: translateY(-4px); box-shadow: 0 12px 24px rgba(0,0,0,0.12); }
.drive-kpi-icon { font-size: 2rem; margin-bottom: 0.5rem; }
.drive-kpi-value { font-size: 2rem; font-weight: 800; color: #0E7490; }
.drive-kpi-label { font-size: 0.75rem; font-weight: 700; color: #94A3B8; text-transform: uppercase; letter-spacing: 0.05em; margin-top: 0.3rem; }

.drive-folder-card {
    background: #fff;
    border: 2px solid #E2E8F0;
    border-radius: 14px;
    padding: 1.2rem;
    transition: all 0.25s ease;
    height: 100%;
    box-shadow: 0 2px 8px rgba(0,0,0,0.04);
}

.drive-folder-card:hover {
    border-color: #06B6D4;
    transform: translateY(-3px);
    box-shadow: 0 8px 20px rgba(6, 182, 212, 0.15);
}

.drive-folder-icon { font-size: 2rem; margin-bottom: 0.5rem; }
.drive-folder-name { font-weight: 700; font-size: 1rem; color: #0F172A; margin-bottom: 0.3rem; }
.drive-folder-cat {
    display: inline-block;
    background: #CFFAFE;
    color: #0E7490;
    border-radius: 6px;
    padding: 2px 8px;
    font-size: 0.7rem;
    font-weight: 700;
    margin-bottom: 0.5rem;
}
.drive-folder-desc { font-size: 0.82rem; color: #64748B; line-height: 1.4; }

.drive-container {
    background-color: #fff;
    padding: 10px;
    border-radius: 12px;
    border: 1px solid #E2E8F0;
    box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05);
}

.drive-link-btn {
    display: block;
    text-align: center;
    padding: 12px;
    border-radius: 8px;
    font-weight: 600;
    text-decoration: none;
    color: #fff;
    transition: all 0.2s ease;
}

.drive-link-btn:hover { transform: translateY(-2px); opacity: 0.92; color: #fff; }
</style>
""", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════
# ENCABEZADO
# ═══════════════════════════════════════════════════════════════════════════
st.markdown("""
<div class="drive-hero">
    <div class="drive-hero-title">☁️ Repositorio Institucional Drive</div>
    <div class="drive-hero-sub">
        Acceso centralizado a documentación, actas, planificaciones y archivos de la Coordinación Pedagógica ETP
    </div>
    <div>
        <span class="drive-hero-badge">📂 Carpetas Dinámicas</span>
        <span class="drive-hero-badge">🔍 Vista Grid / Lista</span>
        <span class="drive-hero-badge">🧾 Auditoría de Accesos</span>
        <span class="drive-hero-badge">📖 Guía de Buenas Prácticas</span>
    </div>
</div>
""", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════
# KPIs
# ═══════════════════════════════════════════════════════════════════════════
carpetas = obtener_carpetas()
total_accesos = contar_accesos_repositorio()

k1, k2, k3 = st.columns(3)
with k1:
    st.markdown(f"""
    <div class="drive-kpi-card">
        <div class="drive-kpi-icon">📂</div>
        <div class="drive-kpi-value">{len(carpetas)}</div>
        <div class="drive-kpi-label">Carpetas Configuradas</div>
    </div>
    """, unsafe_allow_html=True)
with k2:
    st.markdown(f"""
    <div class="drive-kpi-card">
        <div class="drive-kpi-icon">👁️</div>
        <div class="drive-kpi-value">{total_accesos}</div>
        <div class="drive-kpi-label">Accesos Registrados</div>
    </div>
    """, unsafe_allow_html=True)
with k3:
    estado = "🟢 Conectado" if carpetas else "🔴 Sin carpetas"
    st.markdown(f"""
    <div class="drive-kpi-card">
        <div class="drive-kpi-icon">🔗</div>
        <div class="drive-kpi-value" style="font-size:1.3rem;">{estado}</div>
        <div class="drive-kpi-label">Estado del Repositorio</div>
    </div>
    """, unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════
# TABS
# ═══════════════════════════════════════════════════════════════════════════
tab_explorar, tab_gestion, tab_accesos, tab_guia = st.tabs([
    "📂 Explorador", "🗂️ Gestión de Carpetas", "🔗 Accesos Directos", "📖 Guía y Convenciones",
])

# ═══════════════════════════════════════════════════════════════════════════
# TAB 1: EXPLORADOR
# ═══════════════════════════════════════════════════════════════════════════
with tab_explorar:
    st.markdown('<div class="drive-section-title">📂 Explorador de Archivos Institucional</div>', unsafe_allow_html=True)

    if not carpetas:
        st.warning("⚠️ No hay carpetas configuradas. Ve a la pestaña '🗂️ Gestión de Carpetas' para agregar una.")
    else:
        col_sel, col_vista = st.columns([2, 1])
        with col_sel:
            opciones = {f"{c['nombre']} ({c['categoria']})": c for c in carpetas}
            sel_label = st.selectbox("Carpeta institucional", list(opciones.keys()))
            carpeta_sel = opciones[sel_label]
        with col_vista:
            vista = st.radio("Vista", ["Grid", "Lista"], horizontal=True)

        # Registrar acceso (una vez por carpeta por sesión)
        clave_audit = f"drive_audit_{carpeta_sel['id']}"
        if not st.session_state.get(clave_audit):
            st.session_state[clave_audit] = True
            registrar_acceso(carpeta_sel["nombre"])

        altura = st.slider("🔍 Altura del explorador (px)", 400, 900, 600, step=50)

        if carpeta_sel.get("descripcion"):
            st.caption(f"📌 {carpeta_sel['descripcion']}")

        url = url_embed(carpeta_sel["drive_id"], "grid" if vista == "Grid" else "list")

        st.info("💡 Haz doble clic en las carpetas para abrirlas o clic en los archivos para visualizarlos. "
                "Si el explorador no carga, verifica que la carpeta esté compartida como 'Cualquier persona con el enlace'.")

        st.markdown('<div class="drive-container">', unsafe_allow_html=True)
        components.iframe(url, height=altura, scrolling=True)
        st.markdown('</div>', unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)
        col_b1, col_b2 = st.columns(2)
        with col_b1:
            st.markdown(f'''
            <a href="{url_externa(carpeta_sel['drive_id'])}" target="_blank" class="drive-link-btn"
               style="background-color: #2563EB;">
               ☁️ Abrir carpeta externa en Google Drive
            </a>
            ''', unsafe_allow_html=True)
        with col_b2:
            st.markdown(f'''
            <a href="{URL_APPS_SCRIPT}" target="_blank" class="drive-link-btn"
               style="background-color: #0F172A;">
               ⚙️ Ejecutar Apps Script Externo
            </a>
            ''', unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════
# TAB 2: GESTIÓN DE CARPETAS
# ═══════════════════════════════════════════════════════════════════════════
with tab_gestion:
    st.markdown('<div class="drive-section-title">🗂️ Gestión de Carpetas Institucionales</div>', unsafe_allow_html=True)

    es_coordinador = st.session_state.get("coordinador_autenticado", False)
    if not es_coordinador:
        st.info("ℹ️ Solo Coordinación puede agregar o eliminar carpetas. Puedes ver el listado existente.")

    # Listado actual
    if carpetas:
        df_carpetas = pd.DataFrame(carpetas).rename(columns={
            "nombre": "Nombre", "drive_id": "ID de Drive",
            "descripcion": "Descripción", "categoria": "Categoría",
        })[["Nombre", "Categoría", "ID de Drive", "Descripción"]]
        st.dataframe(df_carpetas, use_container_width=True, hide_index=True)
    else:
        st.info("No hay carpetas configuradas todavía.")

    if es_coordinador:
        st.markdown("---")
        col_add, col_del = st.columns([1.4, 1])

        with col_add:
            st.markdown("#### ➕ Agregar carpeta")
            with st.form("form_nueva_carpeta", clear_on_submit=True):
                nuevo_nombre = st.text_input("Nombre de la carpeta", placeholder="Ej: Actas 2026-2027")
                nueva_categoria = st.selectbox("Categoría", CATEGORIAS_SUGERIDAS)
                nuevo_drive_id = st.text_input(
                    "ID de Google Drive",
                    placeholder="Ej: 1kH1UNEth2VGqWnlOxnmsjuxX3HkvhUSi",
                    help="Es la parte de la URL después de /folders/. Ej: drive.google.com/drive/folders/<ID>",
                )
                nueva_desc = st.text_input("Descripción (opcional)", placeholder="Ej: Actas del período actual")
                if st.form_submit_button("➕ Agregar carpeta", type="primary", use_container_width=True):
                    if not nuevo_nombre.strip():
                        st.error("⚠️ El nombre es obligatorio.")
                    else:
                        valido, msg = validar_drive_id(nuevo_drive_id)
                        if not valido:
                            st.error(f"❌ {msg}")
                        else:
                            agregar_carpeta(nuevo_nombre.strip(), nuevo_drive_id.strip(), nueva_desc, nueva_categoria)
                            st.success(f"✅ Carpeta '{nuevo_nombre.strip()}' agregada correctamente.")
                            st.rerun()

        with col_del:
            st.markdown("#### 🗑️ Eliminar carpeta")
            if carpetas:
                opciones_del = {f"{c['nombre']} ({c['categoria']})": c["id"] for c in carpetas}
                sel_del = st.selectbox("Carpeta a eliminar", list(opciones_del.keys()), key="del_carpeta_sel")
                confirmar = st.checkbox("Confirmo eliminar esta carpeta", key="conf_del_carpeta")
                if st.button("🗑️ Eliminar", type="primary", use_container_width=True, disabled=not confirmar):
                    eliminar_carpeta(opciones_del[sel_del])
                    st.success("✅ Carpeta eliminada.")
                    st.rerun()
            else:
                st.info("No hay carpetas para eliminar.")

# ═══════════════════════════════════════════════════════════════════════════
# TAB 3: ACCESOS DIRECTOS
# ═══════════════════════════════════════════════════════════════════════════
with tab_accesos:
    st.markdown('<div class="drive-section-title">🔗 Accesos Directos por Carpeta</div>', unsafe_allow_html=True)

    if not carpetas:
        st.info("No hay carpetas configuradas.")
    else:
        cols = st.columns(3)
        for i, c in enumerate(carpetas):
            with cols[i % 3]:
                st.markdown(f"""
                <div class="drive-folder-card">
                    <div class="drive-folder-icon">📁</div>
                    <div class="drive-folder-name">{c['nombre']}</div>
                    <div class="drive-folder-cat">{c['categoria']}</div>
                    <div class="drive-folder-desc">{c.get('descripcion') or 'Sin descripción'}</div>
                    <div style="margin-top:0.8rem;">
                        <a href="{url_externa(c['drive_id'])}" target="_blank" class="drive-link-btn"
                           style="background-color:#2563EB; padding:8px; font-size:0.85rem;">
                           ☁️ Abrir en Drive
                        </a>
                    </div>
                </div>
                """, unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown('<div class="drive-section-title">⚙️ Automatización</div>', unsafe_allow_html=True)
        st.markdown(f'''
        <a href="{URL_APPS_SCRIPT}" target="_blank" class="drive-link-btn"
           style="background-color:#0F172A; max-width:400px; margin:0 auto;">
           ⚙️ Ejecutar Apps Script Externo
        </a>
        ''', unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════
# TAB 4: GUÍA Y CONVENCIONES
# ═══════════════════════════════════════════════════════════════════════════
with tab_guia:
    st.markdown('<div class="drive-section-title">📖 Guía de Uso y Buenas Prácticas</div>', unsafe_allow_html=True)

    with st.expander("📐 Convención de nombres de archivos", expanded=True):
        st.markdown("""
        Para mantener el repositorio ordenado, usa esta convención:

        **Formato general:**
        ```
        [TIPO]_[AÑO]_[MES]_[Descripción]_[Versión]
        ```

        **Ejemplos:**
        - `ACTA_2026_05_ReuniónCoordinación_v1.pdf`
        - `PLAN_2026_05_PlanDiario_Matemáticas4B_v2.docx`
        - `DIR_2026_08_DirectorioDocentes_v1.xlsx`

        **Tipos recomendados:** `ACTA`, `PLAN`, `DIR`, `EVAL`, `ACUERDO`, `INCID`, `CRONO`, `INFORME`.
        """)

    with st.expander("⬆️ Cómo subir archivos al repositorio"):
        st.markdown("""
        1. Abre la carpeta correspondiente desde **🔗 Accesos Directos**.
        2. Haz clic en **"Nuevo" → "Subir archivo"** (o arrastra el archivo a la ventana).
        3. Aplica la **convención de nombres** antes de subir.
        4. Si es un documento de Google (Docs/Sheets/Slides), usa **"Nuevo → Google Docs/Sheets/Slides"**.
        5. Verifica que el archivo aparezca en el explorador integrado.
        """)

    with st.expander("🔐 Requisitos de permisos"):
        st.markdown("""
        Para que el explorador integrado funcione, la carpeta de Drive debe estar compartida así:

        1. Clic derecho en la carpeta → **"Compartir"**.
        2. En "Acceso general", selecciona **"Cualquier persona con el enlace"**.
        3. Rol: **"Lector"** (o "Editor" si el equipo debe editar).
        4. Copia el **ID** de la URL: `drive.google.com/drive/folders/<ID>`.
        5. Agrégalo en **🗂️ Gestión de Carpetas**.
        """)

    with st.expander("🧾 Auditoría de accesos"):
        st.markdown(f"""
        Cada vez que un usuario abre una carpeta en el explorador, se registra un evento
        de auditoría (una vez por carpeta por sesión).

        - **Acciones registradas:** `Acceso Repositorio`
        - **Total de accesos hasta ahora:** {total_accesos}

        Puedes consultar el detalle en **🔐 Gestor de Accesos → Auditoría**.
        """)

    st.markdown("---")
    st.caption("💡 Este repositorio está integrado al Sistema de Gestión Docente ETP y usa la misma base de datos para la gestión de carpetas y la auditoría.")