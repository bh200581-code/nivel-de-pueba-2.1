"""
juegos_interactivos.py — Fábrica de Juegos Interactivos ETP
• La IA genera juegos HTML de un solo archivo (sin dependencias externas):
  Trivia/Quiz, Memoria, Verdadero/Falso, Emparejar, Ruleta y Ahorcado.
• Vista previa en la app + descarga .html para usar offline en el aula.
• Compartir con estudiantes vía enlace ?juego=ID (SIN login), como las pruebas.
• Galería SQLite para reutilizar juegos.
"""
import re
import sqlite3
from datetime import datetime

import streamlit as st
import streamlit.components.v1 as components

from core import ia
import core.auth as auth

DB_NAME = "gestion_etp.db"

TIPOS_JUEGO = {
    "🧠 Trivia / Quiz": "Preguntas de opción múltiple con 4 alternativas, retroalimentación inmediata con explicación, puntaje y pantalla final con reintento.",
    "🃏 Memoria (parejas)": "Tarjetas volteables para encontrar parejas de concepto–definición o símbolo–término.",
    "🎯 Verdadero o Falso": "Afirmaciones técnicas para juzgar verdadero/falso con retroalimentación y puntaje.",
    "🧩 Emparejar columnas": "Dos columnas (conceptos vs definiciones) que se unen haciendo clic; valida aciertos.",
    "🎡 Ruleta de preguntas": "Rueda que gira al azar y muestra preguntas para responder con temporizador.",
    "✍️ Ahorcado técnico": "Ahorcado con términos clave del tema; la pista es la definición; vidas y puntaje.",
}
DIFICULTADES = ["Fácil", "Media", "Difícil"]

# ═══════════════════════════════════════════════════════════════════════════
# BASE DE DATOS
# ═══════════════════════════════════════════════════════════════════════════
def _conn():
    return sqlite3.connect(DB_NAME, check_same_thread=False)

def asegurar_tabla_juegos():
    conn = _conn()
    conn.execute('''CREATE TABLE IF NOT EXISTS juegos_generados (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        fecha TEXT, docente TEXT, modulo TEXT, tema TEXT, tipo TEXT,
        dificultad TEXT, html TEXT
    )''')
    conn.commit()
    conn.close()

def guardar_juego(docente, modulo, tema, tipo, dificultad, html):
    conn = _conn()
    cur = conn.cursor()
    cur.execute('''INSERT INTO juegos_generados
        (fecha, docente, modulo, tema, tipo, dificultad, html)
        VALUES (?,?,?,?,?,?,?)''',
        (datetime.now().strftime("%Y-%m-%d %H:%M"), docente, modulo, tema, tipo, dificultad, html))
    conn.commit()
    juego_id = cur.lastrowid
    conn.close()
    return juego_id

def listar_juegos():
    conn = _conn()
    cur = conn.cursor()
    cur.execute("SELECT id, fecha, docente, modulo, tema, tipo, dificultad FROM juegos_generados ORDER BY id DESC")
    cols = ["id", "fecha", "docente", "modulo", "tema", "tipo", "dificultad"]
    rows = [dict(zip(cols, r)) for r in cur.fetchall()]
    conn.close()
    return rows

def obtener_juego(id_juego):
    conn = _conn()
    cur = conn.cursor()
    cur.execute("SELECT id, tema, tipo, modulo, html FROM juegos_generados WHERE id=?", (id_juego,))
    row = cur.fetchone()
    conn.close()
    if not row:
        return None
    return {"id": row[0], "tema": row[1], "tipo": row[2], "modulo": row[3], "html": row[4]}

def eliminar_juego(id_juego):
    conn = _conn()
    conn.execute("DELETE FROM juegos_generados WHERE id=?", (id_juego,))
    conn.commit()
    conn.close()

asegurar_tabla_juegos()

# ═══════════════════════════════════════════════════════════════════════════
# UTILIDADES HTML
# ═══════════════════════════════════════════════════════════════════════════
VOID_ELEMENTS = frozenset({'area','base','br','col','embed','hr','img','input',
                           'link','meta','param','source','track','wbr'})

def limpiar_html(texto):
    texto = (texto or "").strip()
    texto = re.sub(r'^```(?:html)?\s*', '', texto, flags=re.IGNORECASE)
    texto = re.sub(r'\s*```$', '', texto)
    return texto.strip()

def balancear_html(codigo):
    if not codigo:
        return codigo
    tag_pattern = re.compile(r'<(/?)([a-zA-Z][a-zA-Z0-9]*)[^>]*>', re.IGNORECASE)
    stack = []
    for match in tag_pattern.finditer(codigo):
        cierre, tag = match.group(1), match.group(2).lower()
        if tag in VOID_ELEMENTS:
            continue
        if cierre:
            if stack and stack[-1] == tag:
                stack.pop()
        else:
            stack.append(tag)
    if stack:
        codigo += "\n" + "\n".join(f"</{t}>" for t in reversed(stack))
    return codigo

def _link_compartir(id_juego):
    try:
        base = st.context.page_url.split("?")[0]
    except Exception:
        base = ""
    return f"{base}?juego={id_juego}"

def construir_prompt_juego(tipo, desc_tipo, tema, modulo, dificultad, num_items):
    return f"""Actúa como un Desarrollador de Juegos Educativos y Docente ETP del MINERD (República Dominicana).
Crea un JUEGO EDUCATIVO INTERACTIVO completo en UN SOLO ARCHIVO (HTML + CSS + JS incrustados).

TIPO DE JUEGO: {tipo}
MECÁNICA: {desc_tipo}
TEMA: {tema}
MÓDULO FORMATIVO: {modulo}
DIFICULTAD: {dificultad}
CANTIDAD DE ÍTEMS/PREGUNTAS: {num_items}

CONTENIDO: genera tú mismo las preguntas/ítems correctos, veraces y pedagógicos, apropiados al tema técnico indicado.

DISEÑO: juvenil, colorido y moderno para estudiantes; gradientes vivos, botones grandes,
barra de puntaje y progreso, pantalla final con calificación y mensaje motivador, botón "Jugar de nuevo",
campo opcional para escribir el nombre del estudiante al inicio.

REGLAS TÉCNICAS ESTRICTAS:
- CERO dependencias externas (sin CDN, sin librerías, sin fuentes de Google).
- Todo en un único archivo HTML con <style> y <script> internos.
- JavaScript vanilla sin alert()/prompt()/confirm(); usa el DOM para mensajes.
- Responsivo (media query para móviles) y con targets táctiles grandes.
- Incluye <meta viewport> y título del juego con el tema.

Devuelve ÚNICAMENTE el código HTML completo. Nada de explicaciones ni markdown.
"""

# ═══════════════════════════════════════════════════════════════════════════
# MODO ESTUDIANTE: enlace ?juego=ID sin login
# ═══════════════════════════════════════════════════════════════════════════
_juego_param = st.query_params.get("juego")
_logueado = bool(st.session_state.get("coordinador_autenticado") or st.session_state.get("docente_autenticado"))

if _juego_param and not _logueado:
    try:
        _id = int(_juego_param)
    except (TypeError, ValueError):
        _id = None
    juego = obtener_juego(_id) if _id else None
    if juego:
        st.markdown(f"# 🎮 {juego['tema']}")
        st.caption(f"{juego['tipo']} · {juego['modulo']} · ¡Demuestra lo que sabes! 🚀")
        components.html(juego["html"], height=820, scrolling=True)
    else:
        st.error("🔒 El juego no existe o el enlace no es válido.")
    st.stop()

if not _logueado:
    st.error("🔒 Debes iniciar sesión."); st.stop()

# ═══════════════════════════════════════════════════════════════════════════
# INTERFAZ DOCENTE
# ═══════════════════════════════════════════════════════════════════════════
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');
html, body, [class*="css"] { font-family: 'Inter', sans-serif; background-color: #F0F4F8; color: #1E293B; }
.juegos-hero { background: linear-gradient(135deg, #0F172A 0%, #7C3AED 45%, #EC4899 80%, #F59E0B 100%);
color: #fff; padding: 2rem; border-radius: 20px; margin-bottom: 1.4rem; box-shadow: 0 25px 50px rgba(124,58,237,0.3); }
.juegos-hero-title { font-size: 2.2rem; font-weight: 900; letter-spacing: -0.03em; }
.juegos-hero-sub { font-size: 1rem; opacity: .92; margin-top: .3rem; }
.section-title { color: #7C3AED; font-weight: 700; font-size: 1.1rem; border-bottom: 2px solid #EDE9FE;
padding-bottom: 8px; margin: 1.1rem 0 .9rem 0; }
</style>
""", unsafe_allow_html=True)

ia.panel_sidebar_ia("Fábrica de Juegos Interactivos")
st.markdown("""
<div class="juegos-hero">
    <div class="juegos-hero-title">🎮 Fábrica de Juegos Interactivos</div>
    <div class="juegos-hero-sub">Genera juegos con IA · juégalos en clase · compártelos con tus estudiantes sin login</div>
</div>
""", unsafe_allow_html=True)

if "juego_html" not in st.session_state: st.session_state.juego_html = None
if "juego_id" not in st.session_state: st.session_state.juego_id = None

tab_gen, tab_gal = st.tabs(["🎮 Generar Juego", "🕹️ Mis Juegos y Compartir"])

# ═══ TAB 1: GENERAR ═══
with tab_gen:
    st.markdown('<div class="section-title">🧙 Configura tu juego</div>', unsafe_allow_html=True)
    with st.form("form_juego"):
        c1, c2 = st.columns(2)
        with c1:
            docente = st.text_input("Docente",
                value=st.session_state.get("usuario_display_nombre", "") or "Ing. Bernardo Hernández")
            modulo = st.text_input("Módulo Formativo", placeholder="Ej: MF 358-3 Impuestos al Consumo")
            tema = st.text_input("Tema del juego", placeholder="Ej: Ley de Ohm, Subredes IP, Ofimática...")
        with c2:
            tipo = st.selectbox("Tipo de juego", list(TIPOS_JUEGO.keys()))
            dificultad = st.selectbox("Dificultad", DIFICULTADES, index=1)
            num_items = st.slider("Cantidad de ítems/preguntas", 5, 20, 10)
        max_tokens, temperature = ia.control_avanzado(default_tokens=16384, tope=32000, default_temp=0.35)
        btn = st.form_submit_button("🎲 Generar Juego con IA", type="primary", width="stretch")

    if btn:
        if not tema or not modulo:
            st.warning("⚠️ Completa el tema y el módulo.")
        else:
            with st.spinner("🧠 Construyendo tu juego interactivo..."):
                try:
                    prompt = construir_prompt_juego(tipo, TIPOS_JUEGO[tipo], tema, modulo, dificultad, num_items)
                    texto, flags = ia.solicitar_texto(prompt, max_tokens=max_tokens,
                                                      temperature=temperature, modulo="juegos_interactivos")
                    html_juego = balancear_html(limpiar_html(texto))
                    juego_id = guardar_juego(docente, modulo, tema, tipo, dificultad, html_juego)
                    st.session_state.juego_html = html_juego
                    st.session_state.juego_id = juego_id
                    st.toast("✅ Juego generado y guardado en la galería.", icon="🎮")
                    st.rerun()
                except Exception as e:
                    ia.render_error_ia(e)

    if st.session_state.juego_html:
        st.markdown('<div class="section-title">🕹️ Vista previa (pruébalo aquí mismo)</div>', unsafe_allow_html=True)
        components.html(st.session_state.juego_html, height=700, scrolling=True)
        c1, c2 = st.columns(2)
        with c1:
            st.download_button("📥 Descargar juego (.html)", data=st.session_state.juego_html,
                               file_name=f"Juego_{ia.sanear_nombre_archivo('tema')}.html",
                               mime="text/html", type="primary", width="stretch", key="dl_juego")
        with c2:
            link = _link_compartir(st.session_state.juego_id)
            st.markdown("**🔗 Enlace para estudiantes (sin login):**")
            st.code(link)
            st.caption("💡 Copia este enlace y compártelo (WhatsApp/Classroom). Los estudiantes juegan sin contraseña.")

# ═══ TAB 2: GALERÍA + COMPARTIR ═══
with tab_gal:
    st.markdown('<div class="section-title">🕹️ Mis juegos generados</div>', unsafe_allow_html=True)
    juegos = listar_juegos()
    if not juegos:
        st.info("Aún no has generado juegos. Crea el primero en la pestaña 🎮.")
    else:
        opciones = {f"#{j['id']} · {j['tipo']} · {j['tema']} ({j['fecha']})": j["id"] for j in juegos}
        sel = st.selectbox("Selecciona un juego", list(opciones.keys()))
        id_sel = opciones[sel]
        juego = obtener_juego(id_sel)
        if juego:
            c1, c2, c3 = st.columns(3)
            with c1:
                if st.button("▶️ Jugar / Previsualizar", width="stretch"):
                    st.session_state.juego_html = juego["html"]
                    st.session_state.juego_id = juego["id"]
                    st.rerun()
            with c2:
                st.download_button("📥 Descargar (.html)", data=juego["html"],
                                   file_name=f"Juego_{ia.sanear_nombre_archivo(juego['tema'])}.html",
                                   mime="text/html", width="stretch", key=f"dl_{id_sel}")
            with c3:
                if st.button("🗑️ Eliminar", width="stretch", key=f"del_{id_sel}"):
                    eliminar_juego(id_sel)
                    st.toast("Juego eliminado.", icon="🗑️")
                    st.rerun()
            st.markdown("**🔗 Enlace para compartir con estudiantes (sin login):**")
            st.code(_link_compartir(id_sel))
            st.caption("Los estudiantes abren este enlace y juegan directamente, sin usuario ni contraseña.")