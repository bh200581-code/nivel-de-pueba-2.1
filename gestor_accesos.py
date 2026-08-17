"""
gestor_accesos.py — Gestor de Accesos y Usuarios (Standalone)
Capa de presentación sincronizada directamente con la tabla 'docentes'.

MEJORAS:
• Sincronizado en tiempo real con main.py y directorio_docentes.py.
• No depende de archivos externos ni módulos core.
• Contraseñas temporales seguras autogeneradas.
• Auditoría centralizada en accesos_log.
"""
import secrets
import string
import sqlite3
import datetime
import pandas as pd
import streamlit as st

DB_NAME = "gestion_etp.db"

# ═══════════════════════════════════════════════════════════════════════════
# FUNCIONES DE BASE DE DATOS (Sincronizadas con main.py)
# ═══════════════════════════════════════════════════════════════════════════
def asegurar_tablas():
    conn = sqlite3.connect(DB_NAME)
    cur = conn.cursor()
    cur.execute('''CREATE TABLE IF NOT EXISTS docentes (
        docente TEXT, modulo TEXT, seccion TEXT, password TEXT DEFAULT '1234', usuario TEXT
    )''')
    cur.execute("""CREATE TABLE IF NOT EXISTS accesos_log (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        fecha TEXT, accion TEXT, usuario TEXT, detalle TEXT
    )""")
    cur.execute("PRAGMA table_info(docentes)")
    cols = [info[1] for info in cur.fetchall()]
    if 'area_tecnica' not in cols: cur.execute("ALTER TABLE docentes ADD COLUMN area_tecnica TEXT")
    if 'horas' not in cols: cur.execute("ALTER TABLE docentes ADD COLUMN horas TEXT")
    conn.commit()
    conn.close()

def registrar_evento(accion: str, usuario: str, detalle: str = "") -> None:
    try:
        fecha = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        conn = sqlite3.connect(DB_NAME)
        conn.execute("INSERT INTO accesos_log (fecha, accion, usuario, detalle) VALUES (?,?,?,?)",
                     (fecha, accion, usuario, detalle))
        conn.commit()
        conn.close()
    except Exception: pass

def obtener_log() -> list:
    conn = sqlite3.connect(DB_NAME)
    cur = conn.cursor()
    cur.execute("SELECT fecha, accion, usuario, detalle FROM accesos_log ORDER BY id DESC")
    rows = [list(r) for r in cur.fetchall()]
    conn.close()
    return rows

def obtener_usuarios_resumen() -> list:
    conn = sqlite3.connect(DB_NAME)
    cur = conn.cursor()
    cur.execute("""
        SELECT docente, MAX(usuario),
               CASE WHEN MAX(password) = '1234' OR MAX(password) IS NULL OR MAX(password) = '' THEN 'Sin configurar' ELSE 'Configurada' END,
               SUM(CASE WHEN modulo IS NOT NULL AND modulo != '' THEN 1 ELSE 0 END),
               GROUP_CONCAT(CASE WHEN modulo IS NOT NULL AND modulo != '' THEN modulo || ' | ' || seccion ELSE NULL END, '; ')
        FROM docentes
        WHERE docente IS NOT NULL AND docente != ''
        GROUP BY docente
        ORDER BY docente
    """)
    rows = [[r[0], r[1] or "", r[2], r[3], r[4] or ""] for r in cur.fetchall()]
    conn.close()
    return rows

def crear_usuario(nombre, usuario, password=""):
    conn = sqlite3.connect(DB_NAME)
    cur = conn.cursor()
    cur.execute("SELECT 1 FROM docentes WHERE docente=?", (nombre,))
    if cur.fetchone():
        conn.close()
        raise ValueError("Ese docente ya existe.")
    
    usu = usuario if usuario else nombre.split()[0].lower()
    pwd = password if password else "1234"
    cur.execute("INSERT INTO docentes (docente, modulo, seccion, password, usuario, area_tecnica, horas) VALUES (?, '', '', ?, ?, '', '')", 
                (nombre, pwd, usu))
    conn.commit()
    conn.close()
    return usu

def eliminar_usuario(docente):
    conn = sqlite3.connect(DB_NAME)
    cur = conn.cursor()
    cur.execute("DELETE FROM docentes WHERE docente=?", (docente,))
    conn.commit()
    conn.close()

def obtener_modulos_usuario(docente):
    conn = sqlite3.connect(DB_NAME)
    cur = conn.cursor()
    cur.execute("SELECT rowid as id, modulo, seccion FROM docentes WHERE docente=? AND modulo IS NOT NULL AND modulo != ''", (docente,))
    rows = [list(r) for r in cur.fetchall()]
    conn.close()
    return rows

def asignar_modulo(docente, modulo, seccion):
    conn = sqlite3.connect(DB_NAME)
    cur = conn.cursor()
    cur.execute("SELECT rowid FROM docentes WHERE docente=? AND (modulo IS NULL OR modulo = '')", (docente,))
    row = cur.fetchone()
    if row:
        cur.execute("UPDATE docentes SET modulo=?, seccion=? WHERE rowid=?", (modulo, seccion, row[0]))
    else:
        cur.execute("SELECT password, usuario, area_tecnica, horas FROM docentes WHERE docente=? LIMIT 1", (docente,))
        info = cur.fetchone()
        if info:
            cur.execute("INSERT INTO docentes (docente, modulo, seccion, password, usuario, area_tecnica, horas) VALUES (?,?,?,?,?,?,?)",
                        (docente, modulo, seccion, info[0], info[1], info[2], info[3]))
    conn.commit()
    conn.close()

def quitar_modulo(id_modulo):
    conn = sqlite3.connect(DB_NAME)
    cur = conn.cursor()
    cur.execute("SELECT docente FROM docentes WHERE rowid=?", (id_modulo,))
    doc = cur.fetchone()
    if doc:
        cur.execute("SELECT count(*) FROM docentes WHERE docente=?", (doc[0],))
        count = cur.fetchone()[0]
        if count == 1:
            cur.execute("UPDATE docentes SET modulo='', seccion='' WHERE rowid=?", (id_modulo,))
        else:
            cur.execute("DELETE FROM docentes WHERE rowid=?", (id_modulo,))
    conn.commit()
    conn.close()

def restablecer_password(docente, nueva_pass):
    if len(str(nueva_pass)) < 4:
        raise ValueError("La contraseña debe tener al menos 4 caracteres.")
    conn = sqlite3.connect(DB_NAME)
    cur = conn.cursor()
    cur.execute("UPDATE docentes SET password=? WHERE docente=?", (nueva_pass, docente))
    conn.commit()
    conn.close()

def actualizar_usuario(docente, nuevo_usuario):
    conn = sqlite3.connect(DB_NAME)
    cur = conn.cursor()
    cur.execute("UPDATE docentes SET usuario=? WHERE docente=?", (nuevo_usuario, docente))
    conn.commit()
    conn.close()


# ═══════════════════════════════════════════════════════════════════════════
# ESTILOS
# ═══════════════════════════════════════════════════════════════════════════
st.markdown(
    """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800;900&display=swap');
html, body, [class*="css"] { font-family: 'Inter', sans-serif; background-color: #F0F4F8; color: #1E293B; }
.accesos-hero { background: linear-gradient(135deg, #0F172A 0%, #1E40AF 40%, #3B82F6 70%, #60A5FA 100%);
color: #fff; padding: 2.2rem; border-radius: 20px; margin-bottom: 1.5rem;
box-shadow: 0 25px 50px rgba(30,64,175,0.3); position: relative; overflow: hidden; }
.accesos-hero::before { content: ''; position: absolute; top: -50%; left: -50%; width: 200%; height: 200%;
background: radial-gradient(circle, rgba(255,255,255,0.1) 0%, transparent 60%); animation: heroPulse 6s ease-in-out infinite; }
@keyframes heroPulse { 0%,100%{transform:scale(1);opacity:0.5} 50%{transform:scale(1.1);opacity:0.8} }
.accesos-hero-title { font-size: 2.4rem; font-weight: 900; letter-spacing: -0.03em; margin-bottom: 0.4rem; position: relative; }
.accesos-hero-sub { font-size: 1.05rem; opacity: 0.9; line-height: 1.5; position: relative; }
.accesos-hero-badge { display: inline-block; background: rgba(255,255,255,0.15); border: 1px solid rgba(255,255,255,0.25);
border-radius: 8px; padding: 4px 12px; font-size: 0.8rem; font-weight: 600; margin-top: 0.8rem; margin-right: 8px; position: relative; }
.accesos-section-title { color: #1E40AF; font-weight: 700; font-size: 1.12rem; border-bottom: 2px solid #DBEAFE;
padding-bottom: 8px; margin: 1.2rem 0 0.9rem 0; }
.accesos-kpi-card { background: #fff; border-radius: 14px; padding: 1.2rem; text-align: center;
border: 1px solid #E2E8F0; box-shadow: 0 4px 12px rgba(0,0,0,0.06); transition: all 0.3s ease; }
.accesos-kpi-card:hover { transform: translateY(-4px); box-shadow: 0 12px 24px rgba(0,0,0,0.12); }
.accesos-kpi-icon { font-size: 2rem; margin-bottom: 0.5rem; }
.accesos-kpi-value { font-size: 2rem; font-weight: 800; color: #1E40AF; }
.accesos-kpi-label { font-size: 0.75rem; font-weight: 700; color: #94A3B8; text-transform: uppercase; letter-spacing: 0.05em; margin-top: 0.3rem; }
.accesos-user-card { background: #fff; border: 2px solid #E2E8F0; border-radius: 14px; padding: 1.2rem;
transition: all 0.25s ease; box-shadow: 0 2px 8px rgba(0,0,0,0.04); }
.accesos-user-card:hover { border-color: #3B82F6; transform: translateY(-3px); box-shadow: 0 8px 20px rgba(59,130,246,0.15); }
.accesos-user-avatar { width: 48px; height: 48px; border-radius: 50%; background: linear-gradient(135deg, #3B82F6, #8B5CF6);
color: #fff; display: flex; align-items: center; justify-content: center; font-weight: 800; font-size: 1.2rem; margin-right: 12px; }
.accesos-user-info { flex: 1; }
.accesos-user-name { font-weight: 700; font-size: 1rem; color: #0F172A; }
.accesos-user-username { font-size: 0.82rem; color: #64748B; }
.accesos-user-status { display: inline-block; padding: 3px 10px; border-radius: 6px; font-size: 0.72rem;
font-weight: 700; text-transform: uppercase; letter-spacing: 0.03em; }
.accesos-status-configurada { background: #D1FAE5; color: #065F46; }
.accesos-status-sin-configurar { background: #FEF3C7; color: #92400E; }
.accesos-progress-bar { height: 8px; border-radius: 4px; background: #E2E8F0; overflow: hidden; margin-top: 6px; }
.accesos-progress-fill { height: 100%; border-radius: 4px; background: linear-gradient(90deg, #3B82F6, #8B5CF6); transition: width 0.8s ease; }
.accesos-timeline-item { display: flex; align-items: flex-start; padding: 12px 0; border-bottom: 1px solid #F1F5F9; }
.accesos-timeline-icon { width: 36px; height: 36px; border-radius: 50%; display: flex; align-items: center;
justify-content: center; font-size: 1rem; margin-right: 12px; flex-shrink: 0; }
.accesos-timeline-content { flex: 1; }
.accesos-timeline-action { font-weight: 600; font-size: 0.92rem; color: #0F172A; }
.accesos-timeline-detail { font-size: 0.82rem; color: #64748B; margin-top: 2px; }
.accesos-timeline-date { font-size: 0.72rem; color: #94A3B8; margin-top: 4px; }
.accesos-strength-bar { height: 6px; border-radius: 3px; background: #E2E8F0; overflow: hidden; margin-top: 8px; }
.accesos-strength-fill { height: 100%; border-radius: 3px; transition: width 0.5s ease, background 0.5s ease; }
.accesos-module-row { display: flex; align-items: center; padding: 10px 0; border-bottom: 1px solid #F1F5F9; }
.accesos-module-name { flex: 1; font-size: 0.88rem; color: #334155; }
.accesos-module-count { font-weight: 700; font-size: 0.92rem; color: #1E40AF; margin-left: 12px; }
.accesos-flash { padding: 12px 16px; border-radius: 10px; margin-bottom: 12px; font-weight: 500; animation: slideIn 0.3s ease; }
@keyframes slideIn { from {opacity:0; transform:translateY(-10px);} to {opacity:1; transform:translateY(0);} }
.accesos-flash-success { background: #D1FAE5; color: #065F46; border-left: 4px solid #10B981; }
.accesos-flash-error { background: #FEE2E2; color: #991B1B; border-left: 4px solid #EF4444; }
.accesos-note { font-size: 0.88rem; color: #64748B; margin-top: 0.75rem; }
</style>
""",
    unsafe_allow_html=True,
)

# ═══════════════════════════════════════════════════════════════════════════
# GUARDIA + ESQUEMA
# ═══════════════════════════════════════════════════════════════════════════
if not st.session_state.get("coordinador_autenticado", False):
    st.error("🔒 Esta página es exclusiva de Coordinación.")
    st.stop()

asegurar_tablas()

# ═══════════════════════════════════════════════════════════════════════════
# HELPERS GRAFICOS
# ═══════════════════════════════════════════════════════════════════════════
df_usuarios = pd.DataFrame(
    obtener_usuarios_resumen(),
    columns=["docente", "usuario", "estado_password", "total_modulos", "modulos"],
)

def _actor() -> str:
    return st.session_state.get("coordinador_nombre", "coordinación")

def _iniciales(nombre: str) -> str:
    partes = str(nombre).strip().split()
    if len(partes) >= 2:
        return (partes[0][0] + partes[1][0]).upper()
    return nombre[:2].upper()

def _generar_password_temporal(longitud: int = 10) -> str:
    alfabeto = string.ascii_letters + string.digits + "!@#$%&*"
    while True:
        pwd = "".join(secrets.choice(alfabeto) for _ in range(longitud))
        if (
            any(c.isupper() for c in pwd)
            and any(c.islower() for c in pwd)
            and any(c.isdigit() for c in pwd)
        ):
            return pwd

def _calcular_fuerza_password(password: str) -> tuple:
    if not password:
        return 0, "#E2E8F0", "Sin contraseña"
    puntos = 0
    if len(password) >= 6: puntos += 1
    if len(password) >= 8: puntos += 1
    if len(password) >= 12: puntos += 1
    if any(c.isupper() for c in password): puntos += 1
    if any(c.islower() for c in password): puntos += 1
    if any(c.isdigit() for c in password): puntos += 1
    if any(c in "!@#$%^&*()_+-=[]{}|;:',.<>?" for c in password): puntos += 1
    porcentaje = min(int((puntos / 7) * 100), 100)
    if porcentaje >= 80: return porcentaje, "#10B981", "Fuerte"
    elif porcentaje >= 50: return porcentaje, "#F59E0B", "Media"
    else: return porcentaje, "#EF4444", "Débil"

def _icono_evento(accion: str) -> tuple:
    a = accion.lower()
    if "login" in a: return "🔑", "#DBEAFE"
    elif "crear" in a: return "➕", "#D1FAE5"
    elif "eliminar" in a: return "🗑️", "#FEE2E2"
    elif "módulo" in a or "modulo" in a: return "📚", "#FEF3C7"
    elif "contraseña" in a or "password" in a: return "🔐", "#FCE7F3"
    elif "usuario" in a: return "👤", "#E0E7FF"
    else: return "📋", "#F1F5F9"

# ═══════════════════════════════════════════════════════════════════════════
# HERO + FLASH
# ═══════════════════════════════════════════════════════════════════════════
st.markdown(
    """
    <div class="accesos-hero">
        <div class="accesos-hero-title">🔐 Gestor de Accesos y Usuarios</div>
        <div class="accesos-hero-sub">Un docente · Un usuario · Contraseñas seguras · Módulos vinculados</div>
        <div>
            <span class="accesos-hero-badge">🚫 Sin cuenta demo</span>
            <span class="accesos-hero-badge">📚 Módulos Vinculados</span>
            <span class="accesos-hero-badge">🧾 Auditoría Completa</span>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

if "flash" in st.session_state:
    st.markdown(f'<div class="accesos-flash accesos-flash-success">✅ {st.session_state.pop("flash")}</div>', unsafe_allow_html=True)
if "flash_error" in st.session_state:
    st.markdown(f'<div class="accesos-flash accesos-flash-error">❌ {st.session_state.pop("flash_error")}</div>', unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════
# KPIs
# ═══════════════════════════════════════════════════════════════════════════
if df_usuarios.empty:
    total_usuarios = con_modulos = sin_modulos = total_modulos = 0
else:
    total_usuarios = len(df_usuarios)
    con_modulos = int((df_usuarios["total_modulos"] > 0).sum())
    sin_modulos = int((df_usuarios["total_modulos"] == 0).sum())
    total_modulos = int(df_usuarios["total_modulos"].sum())

col_k1, col_k2, col_k3, col_k4 = st.columns(4)
with col_k1:
    st.markdown(f'<div class="accesos-kpi-card"><div class="accesos-kpi-icon">👤</div><div class="accesos-kpi-value">{total_usuarios}</div><div class="accesos-kpi-label">Usuarios Únicos</div></div>', unsafe_allow_html=True)
with col_k2:
    st.markdown(f'<div class="accesos-kpi-card"><div class="accesos-kpi-icon">🧑🏫</div><div class="accesos-kpi-value">{con_modulos}</div><div class="accesos-kpi-label">Con Módulos</div></div>', unsafe_allow_html=True)
with col_k3:
    st.markdown(f'<div class="accesos-kpi-card"><div class="accesos-kpi-icon">⚠️</div><div class="accesos-kpi-value">{sin_modulos}</div><div class="accesos-kpi-label">Sin Módulos</div></div>', unsafe_allow_html=True)
with col_k4:
    st.markdown(f'<div class="accesos-kpi-card"><div class="accesos-kpi-icon">📚</div><div class="accesos-kpi-value">{total_modulos}</div><div class="accesos-kpi-label">Módulos Asignados</div></div>', unsafe_allow_html=True)

tab_usuarios, tab_modulos, tab_seguridad, tab_auditoria = st.tabs(["👤 Usuarios", "📚 Módulos", "🔐 Seguridad", "🧾 Auditoría"])

# ═══════════════════════════════════════════════════════════════════════════
# TAB 1: USUARIOS
# ═══════════════════════════════════════════════════════════════════════════
with tab_usuarios:
    st.markdown('<div class="accesos-section-title">👤 Usuarios Únicos</div>', unsafe_allow_html=True)
    if df_usuarios.empty:
        st.info("No hay usuarios registrados.")
    else:
        col_buscar, col_filtro = st.columns([2, 1])
        with col_buscar:
            busqueda = st.text_input("🔍 Buscar", placeholder="Docente, usuario o módulo...")
        with col_filtro:
            filtro_estado = st.selectbox("Filtrar por estado", ["Todos", "Con módulos", "Sin módulos", "Contraseña configurada", "Sin contraseña"])

        df_filtrado = df_usuarios.copy()
        if busqueda.strip():
            texto = busqueda.strip().lower()
            df_filtrado = df_filtrado[df_filtrado.apply(
                lambda row: texto in " ".join(str(v).lower() for v in row if pd.notna(v) and str(v).strip()), axis=1)]
        if filtro_estado == "Con módulos": df_filtrado = df_filtrado[df_filtrado["total_modulos"] > 0]
        elif filtro_estado == "Sin módulos": df_filtrado = df_filtrado[df_filtrado["total_modulos"] == 0]
        elif filtro_estado == "Contraseña configurada": df_filtrado = df_filtrado[df_filtrado["estado_password"] == "Configurada"]
        elif filtro_estado == "Sin contraseña": df_filtrado = df_filtrado[df_filtrado["estado_password"] != "Configurada"]

        st.markdown(f"**{len(df_filtrado)}** usuario(s) encontrado(s)")
        for _, row in df_filtrado.iterrows():
            estado_class = "accesos-status-configurada" if row["estado_password"] == "Configurada" else "accesos-status-sin-configurar"
            estado_texto = "Configurada" if row["estado_password"] == "Configurada" else "Sin configurar"
            modulos_txt = str(row["modulos"]) if pd.notna(row["modulos"]) and str(row["modulos"]).strip() else "Sin módulos"
            st.markdown(f"""
            <div class="accesos-user-card" style="display:flex;align-items:center;margin-bottom:12px;">
                <div class="accesos-user-avatar">{_iniciales(row['docente'])}</div>
                <div class="accesos-user-info">
                    <div class="accesos-user-name">{row['docente']}</div>
                    <div class="accesos-user-username">@{row['usuario']}</div>
                    <div style="margin-top:6px;">
                        <span class="accesos-user-status {estado_class}">{estado_texto}</span>
                        <span style="margin-left:8px;font-size:0.78rem;color:#64748B;">📚 {row['total_modulos']} módulo(s)</span>
                    </div>
                    <div style="font-size:0.75rem;color:#94A3B8;margin-top:4px;">{modulos_txt}</div>
                    <div class="accesos-progress-bar"><div class="accesos-progress-fill" style="width:{min(row['total_modulos']*20,100)}%;"></div></div>
                </div>
            </div>
            """, unsafe_allow_html=True)

        st.markdown('<div class="accesos-note">🔒 Al crear un usuario se genera una <b>contraseña temporal segura</b> que se muestra una sola vez.</div>', unsafe_allow_html=True)

        csv = df_usuarios[["docente", "usuario", "estado_password", "total_modulos"]].to_csv(index=False).encode("utf-8-sig")
        st.download_button("📥 Exportar listado (.csv)", data=csv, file_name="usuarios_etp.csv", mime="text/csv")

        st.markdown("---")
        col_form, col_delete = st.columns([1.2, 1])
        with col_form:
            st.markdown("#### ➕ Crear Usuario Único")
            with st.form("form_crear_usuario", clear_on_submit=True):
                nuevo_nombre = st.text_input("Nombre completo del docente", placeholder="Ej: Docente Ejemplo")
                nuevo_usuario = st.text_input("Usuario (opcional)", placeholder="Se genera automáticamente si se deja vacío")
                pass_inicial = st.text_input("Contraseña inicial (opcional)", type="password",
                                             placeholder="Si se deja vacío, se genera una temporal segura")
                if st.form_submit_button("➕ Crear Usuario", type="primary", use_container_width=True):
                    try:
                        usuario_creado = crear_usuario(nuevo_nombre, nuevo_usuario)
                        password_final = pass_inicial.strip() if pass_inicial.strip() else _generar_password_temporal()
                        restablecer_password(nuevo_nombre, password_final)
                        registrar_evento("Crear usuario", _actor(), f"{nuevo_nombre} (@{usuario_creado})")
                        st.session_state.flash = (f"Usuario creado: {usuario_creado}. Contraseña inicial: {password_final} (compártela una sola vez).")
                        st.rerun()
                    except ValueError as e:
                        st.error(f"⚠️ {e}")
                    except Exception as e:
                        st.error(f"❌ Error inesperado: {e}")
        with col_delete:
            st.markdown("#### 🗑️ Eliminar Usuario")
            if df_usuarios.empty:
                st.info("No hay usuarios para eliminar.")
            else:
                opciones_eliminar = {f"{row.docente} (@{row.usuario})": row.docente for row in df_usuarios.itertuples()}
                opcion_eliminar = st.selectbox("Usuario a eliminar", list(opciones_eliminar.keys()), key="eliminar_usuario_select")
                confirmar_eliminacion = st.checkbox("Confirmo eliminar este usuario y sus módulos", key="confirmar_eliminar_usuario")
                if st.button("🗑️ Eliminar Usuario", type="primary", use_container_width=True, disabled=not confirmar_eliminacion):
                    try:
                        eliminar_usuario(opciones_eliminar[opcion_eliminar])
                        registrar_evento("Eliminar usuario", _actor(), opcion_eliminar)
                        st.session_state.flash = "Usuario eliminado correctamente."
                        st.rerun()
                    except ValueError as e:
                        st.error(f"⚠️ {e}")
                    except Exception as e:
                        st.error(f"❌ Error inesperado: {e}")

# ═══════════════════════════════════════════════════════════════════════════
# TAB 2: MÓDULOS
# ═══════════════════════════════════════════════════════════════════════════
with tab_modulos:
    st.markdown('<div class="accesos-section-title">📚 Módulos por Usuario</div>', unsafe_allow_html=True)
    if df_usuarios.empty:
        st.info("Primero crea un usuario.")
    else:
        st.markdown("#### 📊 Distribución de Módulos por Docente")
        df_mod_dist = df_usuarios[["docente", "total_modulos"]].sort_values("total_modulos", ascending=False)
        max_modulos = df_mod_dist["total_modulos"].max() if not df_mod_dist.empty else 1
        for _, row in df_mod_dist.iterrows():
            porcentaje = (row["total_modulos"] / max_modulos * 100) if max_modulos > 0 else 0
            st.markdown(f"""
            <div class="accesos-module-row">
                <div class="accesos-module-name">{row['docente']}</div>
                <div style="flex:2;margin:0 12px;"><div class="accesos-progress-bar"><div class="accesos-progress-fill" style="width:{porcentaje}%;"></div></div></div>
                <div class="accesos-module-count">{row['total_modulos']}</div>
            </div>
            """, unsafe_allow_html=True)
        st.markdown("---")
        opciones_usuario = {f"{row.docente} (@{row.usuario})": row.docente for row in df_usuarios.itertuples()}
        usuario_sel_label = st.selectbox("Selecciona el usuario", list(opciones_usuario.keys()), key="usuario_modulos_select")
        docente_sel = opciones_usuario[usuario_sel_label]
        df_modulos_usuario = pd.DataFrame(obtener_modulos_usuario(docente_sel), columns=["id", "modulo", "seccion"])
        if df_modulos_usuario.empty:
            st.info("Este usuario aún no tiene módulos asignados.")
        else:
            st.dataframe(df_modulos_usuario.rename(columns={"modulo": "Módulo", "seccion": "Sección"}).drop(columns=["id"], errors="ignore"), use_container_width=True, hide_index=True)
        st.markdown("---")
        col_add, col_remove = st.columns(2)
        with col_add:
            st.markdown("#### ➕ Asignar Módulo")
            with st.form("form_asignar_modulo", clear_on_submit=True):
                nuevo_modulo = st.text_input("Módulo", placeholder="Ej: Ofimática")
                nueva_seccion = st.text_input("Sección", placeholder="Ej: 4to A")
                if st.form_submit_button("➕ Asignar Módulo", type="primary", use_container_width=True):
                    try:
                        asignar_modulo(docente_sel, nuevo_modulo, nueva_seccion)
                        registrar_evento("Asignar módulo", _actor(), f"{docente_sel} -> {nuevo_modulo} | {nueva_seccion}")
                        st.session_state.flash = "Módulo asignado correctamente."
                        st.rerun()
                    except ValueError as e:
                        st.error(f"⚠️ {e}")
                    except Exception as e:
                        st.error(f"❌ Error inesperado: {e}")
        with col_remove:
            st.markdown("#### ➖ Quitar Módulo")
            if df_modulos_usuario.empty:
                st.info("No hay módulos para quitar.")
            else:
                opciones_modulo = {f"{row.modulo} | {row.seccion}": row.id for row in df_modulos_usuario.itertuples()}
                modulo_a_quitar_label = st.selectbox("Módulo a quitar", list(opciones_modulo.keys()), key="quitar_modulo_select")
                confirmar_quitar = st.checkbox("Confirmo quitar este módulo", key="confirmar_quitar_modulo")
                if st.button("➖ Quitar Módulo", use_container_width=True, disabled=not confirmar_quitar):
                    try:
                        quitar_modulo(opciones_modulo[modulo_a_quitar_label])
                        registrar_evento("Quitar módulo", _actor(), f"{docente_sel} <- {modulo_a_quitar_label}")
                        st.session_state.flash = "Módulo eliminado correctamente."
                        st.rerun()
                    except ValueError as e:
                        st.error(f"⚠️ {e}")
                    except Exception as e:
                        st.error(f"❌ Error inesperado: {e}")

# ═══════════════════════════════════════════════════════════════════════════
# TAB 3: SEGURIDAD 
# ═══════════════════════════════════════════════════════════════════════════
with tab_seguridad:
    st.markdown('<div class="accesos-section-title">🔐 Seguridad: Modificar Usuario y Contraseña</div>', unsafe_allow_html=True)
    if df_usuarios.empty:
        st.info("No hay usuarios registrados.")
    else:
        opciones_seguridad = {f"{row.docente} (@{row.usuario})": row.docente for row in df_usuarios.itertuples()}
        usuario_seguridad_label = st.selectbox("Selecciona el usuario a modificar", list(opciones_seguridad.keys()), key="usuario_seguridad_select")
        docente_seguridad = opciones_seguridad[usuario_seguridad_label]
        usuario_actual = df_usuarios.loc[df_usuarios["docente"] == docente_seguridad, "usuario"].iloc[0]
        estado_actual = df_usuarios.loc[df_usuarios["docente"] == docente_seguridad, "estado_password"].iloc[0]

        st.markdown(f"""
        <div class="accesos-user-card" style="display:flex;align-items:center;margin-bottom:16px;">
            <div class="accesos-user-avatar">{_iniciales(docente_seguridad)}</div>
            <div class="accesos-user-info">
                <div class="accesos-user-name">{docente_seguridad}</div>
                <div class="accesos-user-username">@{usuario_actual} · Contraseña: {estado_actual}</div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        col_pass, col_user = st.columns(2)
        with col_pass:
            st.markdown("#### 🔑 Modificar Contraseña")
            pass_custom = st.checkbox("Asignar contraseña personalizada", key="chk_pass_custom")
            if pass_custom:
                nueva_pass = st.text_input("Nueva contraseña", type="password", key="nueva_pass")
                confirma_pass = st.text_input("Confirmar nueva contraseña", type="password", key="confirma_pass")
                if nueva_pass:
                    porcentaje, color, nivel = _calcular_fuerza_password(nueva_pass)
                    st.markdown(f"""
                    <div style="margin-top:8px;">
                        <div style="display:flex;justify-content:space-between;font-size:0.78rem;color:#64748B;">
                            <span>Fuerza: {nivel}</span><span>{porcentaje}%</span>
                        </div>
                        <div class="accesos-strength-bar"><div class="accesos-strength-fill" style="width:{porcentaje}%;background:{color};"></div></div>
                    </div>
                    """, unsafe_allow_html=True)
                confirmar_cambio = st.checkbox("Confirmo el cambio de contraseña", key="confirmar_cambio_pass")
                if st.button("💾 Guardar Contraseña", type="primary", use_container_width=True, disabled=not confirmar_cambio):
                    if nueva_pass != confirma_pass:
                        st.error("⚠️ Las contraseñas no coinciden.")
                    elif len(nueva_pass) < 4:
                        st.error("⚠️ La contraseña debe tener al menos 4 caracteres.")
                    else:
                        try:
                            restablecer_password(docente_seguridad, nueva_pass)
                            registrar_evento("Modificar contraseña", _actor(), usuario_seguridad_label)
                            st.session_state.flash = "Contraseña actualizada correctamente."
                            st.rerun()
                        except ValueError as e:
                            st.error(f"⚠️ {e}")
                        except Exception as e:
                            st.error(f"❌ Error inesperado: {e}")
            else:
                st.write("Genera una **contraseña temporal segura** y compártela una sola vez.")
                confirmar_reset = st.checkbox("Confirmar generación de contraseña temporal", key="confirmar_reset_password")
                if st.button("🔑 Generar Contraseña Temporal", type="primary", use_container_width=True, disabled=not confirmar_reset):
                    try:
                        temp = _generar_password_temporal()
                        restablecer_password(docente_seguridad, temp)
                        registrar_evento("Generar contraseña temporal", _actor(), usuario_seguridad_label)
                        st.session_state.flash = f"Contraseña temporal para @{usuario_actual}: {temp}"
                        st.rerun()
                    except ValueError as e:
                        st.error(f"⚠️ {e}")
                    except Exception as e:
                        st.error(f"❌ Error inesperado: {e}")
        with col_user:
            st.markdown("#### 👤 Modificar Usuario (Login)")
            nuevo_usuario_seguridad = st.text_input("Nuevo usuario", placeholder=f"Actual: {usuario_actual}",
                                                    key="nuevo_usuario_seguridad",
                                                    help="Minúsculas, números, punto, guion o guion bajo (3–30 caracteres).")
            if st.button("👤 Actualizar Usuario", use_container_width=True):
                try:
                    actualizar_usuario(docente_seguridad, nuevo_usuario_seguridad)
                    registrar_evento("Modificar usuario", _actor(), f"{usuario_seguridad_label} -> {nuevo_usuario_seguridad}")
                    st.session_state.flash = "Usuario actualizado correctamente."
                    st.rerun()
                except ValueError as e:
                    st.error(f"⚠️ {e}")
                except Exception as e:
                    st.error(f"❌ Error inesperado: {e}")

        # ── PURGA DE CUENTAS DEMO ──
        st.markdown("---")
        st.markdown("#### 🧹 Limpieza de Cuentas de Demostración")
        demo_mask = (
            df_usuarios["usuario"].str.lower().isin(["demo", "demostracion", "test"])
            | df_usuarios["docente"].str.lower().str.contains("demo", na=False)
        )
        demo_list = df_usuarios[demo_mask]
        if demo_list.empty:
            st.caption("No se detectaron cuentas de demostración. ✅")
        else:
            st.warning(f"Se detectaron {len(demo_list)} cuenta(s) de demostración.")
            for row in demo_list.itertuples():
                st.write(f"- {row.docente} (@{row.usuario})")
            if st.button("🗑️ Eliminar cuentas demo", type="primary"):
                for row in demo_list.itertuples():
                    eliminar_usuario(row.docente)
                    registrar_evento("Eliminar cuenta demo", _actor(), f"{row.docente} (@{row.usuario})")
                st.session_state.flash = "Cuentas de demostración eliminadas."
                st.rerun()

# ═══════════════════════════════════════════════════════════════════════════
# TAB 4: AUDITORÍA
# ═══════════════════════════════════════════════════════════════════════════
with tab_auditoria:
    st.markdown('<div class="accesos-section-title">🧾 Auditoría — Timeline de Eventos</div>', unsafe_allow_html=True)
    df_log = pd.DataFrame(obtener_log(), columns=["fecha", "accion", "usuario", "detalle"])
    if df_log.empty:
        st.info("Aún no hay eventos registrados.")
    else:
        col_f1, col_f2, col_f3 = st.columns(3)
        with col_f1:
            filtro_accion = st.multiselect("Filtrar por acción", df_log["accion"].unique().tolist())
        with col_f2:
            filtro_usuario = st.multiselect("Filtrar por usuario", sorted(df_log["usuario"].unique().tolist()))
        with col_f3:
            limite = st.slider("Eventos a mostrar", 10, 100, 50, step=10)
        df_filtrado_log = df_log.copy()
        if filtro_accion:
            df_filtrado_log = df_filtrado_log[df_filtrado_log["accion"].isin(filtro_accion)]
        if filtro_usuario:
            df_filtrado_log = df_filtrado_log[df_filtrado_log["usuario"].isin(filtro_usuario)]
        df_filtrado_log = df_filtrado_log.head(limite)

        for _, row in df_filtrado_log.iterrows():
            icono, color_fondo = _icono_evento(row["accion"])
            st.markdown(f"""
            <div class="accesos-timeline-item">
                <div class="accesos-timeline-icon" style="background:{color_fondo};">{icono}</div>
                <div class="accesos-timeline-content">
                    <div class="accesos-timeline-action">{row['accion']}</div>
                    <div class="accesos-timeline-detail">{row['detalle']}</div>
                    <div class="accesos-timeline-date">👤 {row['usuario']} · 🕐 {row['fecha']}</div>
                </div>
            </div>
            """, unsafe_allow_html=True)
        st.markdown('<div class="accesos-note">Se registran creación de usuarios, asignación/eliminación de módulos, cambios de contraseñas, cambios de usuario, purga de demos y logins.</div>', unsafe_allow_html=True)
        st.markdown("---")
        with st.expander("📋 Ver como tabla"):
            st.dataframe(df_filtrado_log, use_container_width=True, hide_index=True)