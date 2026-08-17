"""
simuladores.py — Fábrica de Simuladores ETP (CORREGIDO + GALERÍA)
Genera simuladores web interactivos (single-file HTML) en 2 fases:
  Fase 1: Estructura visual (HTML + CSS) con tema y complejidad elegidos.
  Fase 2: Lógica JavaScript completa, insertada en el HTML de Fase 1.
Modo Refinamiento: ajusta un simulador existente sin romper lo que funciona.
• FIX: preview con components.html (el srcdoc crasheaba) y sin ternarias sueltas.
• NUEVO: persistencia SQLite + pestaña Galería (descargar/cargar/eliminar).
• IA vía core/ia (solicitar_texto, reintento por truncamiento, auditoría).
"""
import base64
import re
import datetime
import sqlite3

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

from core import ia

# ═══════════════════════════════════════════════════════════════════════════
# PERSISTENCIA SQLite (Galería de simuladores)
# ═══════════════════════════════════════════════════════════════════════════
DB_NAME = "gestion_etp.db"


def init_db():
    conn = sqlite3.connect(DB_NAME, check_same_thread=False)
    cursor = conn.cursor()
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS simuladores (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        fecha TEXT, tema TEXT, modulo TEXT,
        tema_visual TEXT, nivel TEXT, html TEXT
    )
    ''')
    conn.commit()
    return conn


conn = init_db()


def guardar_simulador(tema, modulo, tema_visual, nivel, html):
    cursor = conn.cursor()
    cursor.execute('''
    INSERT INTO simuladores (fecha, tema, modulo, tema_visual, nivel, html)
    VALUES (?, ?, ?, ?, ?, ?)
    ''', (datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
          tema, modulo, tema_visual, nivel, html))
    conn.commit()
    return cursor.lastrowid


def listar_simuladores():
    cursor = conn.cursor()
    cursor.execute("SELECT id, fecha, tema, modulo, tema_visual, nivel "
                   "FROM simuladores ORDER BY id DESC")
    cols = ["id", "fecha", "tema", "modulo", "tema_visual", "nivel"]
    return [dict(zip(cols, r)) for r in cursor.fetchall()]


def obtener_simulador(id_sim):
    cursor = conn.cursor()
    cursor.execute("SELECT html FROM simuladores WHERE id=?", (id_sim,))
    row = cursor.fetchone()
    return row[0] if row else None


def eliminar_simulador(id_sim):
    cursor = conn.cursor()
    cursor.execute("DELETE FROM simuladores WHERE id=?", (id_sim,))
    conn.commit()


# ═══════════════════════════════════════════════════════════════════════════
# TEMAS VISUALES Y NIVELES
# ═══════════════════════════════════════════════════════════════════════════
TEMAS_VISUALES = {
    "Corporativo Moderno": """
- Fondo general: gradiente sutil de #F1F5F9 a #E2E8F0.
- Tarjetas: blancas (#FFFFFF), border-radius 12px, box-shadow 0 10px 25px rgba(0,0,0,0.08).
- Color de acento primario: #2563EB (azul). Botones primarios con ese color y hover #1D4ED8.
- Texto principal: #0F172A. Texto secundario: #64748B.
- Tipografía: system-ui, -apple-system, sans-serif.""",
    "Oscuro Tech (Dark Mode)": """
- Fondo general: #0B1120 con un leve gradiente radial a #111827.
- Tarjetas: #1E293B, border-radius 14px, borde 1px solid #334155, sombra 0 10px 30px rgba(0,0,0,0.4).
- Color de acento primario: #22D3EE (cian) para bordes activos, botones y resultados destacados.
- Texto principal: #E2E8F0. Texto secundario: #94A3B8.
- Añade un ligero efecto glow (box-shadow con el color de acento) en inputs enfocados (:focus).""",
    "Educativo Vibrante": """
- Fondo general: gradiente diagonal de #FFF7ED a #FEF3C7 muy suave.
- Tarjetas: blancas con border-radius 16px y sombra suave, bordes superiores de 4px con color de acento.
- Color de acento primario: #F97316 (naranja) combinado con #10B981 (verde) para estados de éxito.
- Texto principal: #1F2937. Usa emojis o iconos SVG simples para reforzar contexto educativo sin saturar.
- Tipografía con buen tamaño (mínimo 16px) pensada para estudiantes.""",
    "Minimalista Suizo": """
- Fondo general: blanco puro (#FFFFFF) o gris muy claro (#FAFAFA).
- Tarjetas: bordes finos 1px solid #E5E5E5, sin sombras pronunciadas, border-radius 4px.
- Color de acento primario: #111111 (negro) o un único color de acento (#DC2626 rojo) usado con moderación.
- Tipografía protagonista: jerarquía tipográfica clara (títulos grandes en negrita, cuerpo ligero).
- Mucho espacio en blanco (whitespace), grid limpio, cero decoración innecesaria.""",
}

NIVELES_COMPLEJIDAD = {
    "Básico (cálculo directo)": "El simulador debe resolver una lógica directa y sencilla. Prioriza la claridad sobre la cantidad de features. Un único flujo principal, sin pasos intermedios.",
    "Intermedio (con validaciones)": "El simulador debe incluir validación de inputs en tiempo real, mensajes de error claros, y un botón de 'Reiniciar' que limpie el formulario y los resultados.",
    "Avanzado (multi-paso / interactivo)": "El simulador debe sentirse como una mini-aplicación: puede tener pasos o pestañas internas, actualización de resultados en tiempo real (evento 'input', no solo 'click' cuando aplique), animaciones de transición entre estados, y un resumen final destacado.",
}

# ═══════════════════════════════════════════════════════════════════════════
# UTILIDADES HTML (CORREGIDAS)
# ═══════════════════════════════════════════════════════════════════════════
VOID_ELEMENTS = frozenset({
    'area', 'base', 'br', 'col', 'embed', 'hr', 'img', 'input',
    'link', 'meta', 'param', 'source', 'track', 'wbr'
})


def limpiar_respuesta_html(texto):
    """Limpia fences de markdown y texto extra de la respuesta de la IA."""
    if not texto:
        return ""
    texto = texto.strip()
    texto = re.sub(r'^```(?:html)?\s*', '', texto, flags=re.IGNORECASE)
    texto = re.sub(r'\s*```$', '', texto)
    return texto.strip()


def balancear_html(codigo):
    """Cierra etiquetas HTML que quedaron abiertas por truncamiento."""
    if not codigo:
        return codigo
    tag_pattern = re.compile(r'<(/?)([a-zA-Z][a-zA-Z0-9]*)[^>]*>', re.IGNORECASE)
    stack = []
    for match in tag_pattern.finditer(codigo):
        is_closing = match.group(1) == '/'
        tag_name = match.group(2).lower()
        if tag_name in VOID_ELEMENTS:
            continue
        if is_closing:
            if stack and stack[-1] == tag_name:
                stack.pop()
        else:
            stack.append(tag_name)
    if not stack:
        return codigo
    closing_tags = '\n'.join(f'</{tag}>' for tag in reversed(stack))
    body_close = codigo.lower().rfind('</body>')
    if body_close != -1:
        codigo = codigo[:body_close] + closing_tags + '\n' + codigo[body_close:]
    else:
        codigo += '\n' + closing_tags
    return codigo


def extraer_script(codigo):
    """Extrae el contenido del bloque <script> del HTML."""
    match = re.search(r'<script[^>]*>([\s\S]*?)</script>', codigo, re.IGNORECASE)
    if match:
        return match.group(1).strip(), False
    match = re.search(r'<script[^>]*>([\s\S]*)$', codigo, re.IGNORECASE)
    if match:
        return match.group(1).strip(), True
    return "", False


def insertar_script(html, script):
    """Inserta el script antes de </body>, eliminando scripts previos."""
    html = re.sub(r'<script[^>]*>[\s\S]*?</script>', '', html, flags=re.IGNORECASE)
    script_block = f'\n<script>\n{script}\n</script>\n'
    body_close = html.lower().rfind('</body>')
    if body_close != -1:
        return html[:body_close] + script_block + html[body_close:]
    html_close = html.lower().rfind('</html>')
    if html_close != -1:
        return html[:html_close] + script_block + html[html_close:]
    return html + script_block


def generar_data_uri(html):
    """Genera un data URI para abrir en nueva pestaña."""
    b64 = base64.b64encode(html.encode('utf-8')).decode('utf-8')
    return f"data:text/html;base64,{b64}"


def evaluar_calidad(html):
    """Evalúa la calidad del HTML generado. Devuelve lista de (ok, mensaje)."""
    checks = []
    checks.append(('<!doctype html' in html.lower(), 'DOCTYPE HTML5'))
    checks.append(('viewport' in html.lower() and '<meta' in html.lower(), 'Meta viewport (responsivo)'))
    checks.append(('</html>' in html.lower(), 'Cierre </html>'))
    checks.append(('</body>' in html.lower(), 'Cierre </body>'))
    checks.append(('<script' in html.lower(), 'JavaScript incluido'))
    checks.append(('alert(' not in html, 'Sin alert()'))
    checks.append(('confirm(' not in html, 'Sin confirm()'))
    checks.append(('prompt(' not in html, 'Sin prompt()'))
    external = re.findall(r'(?:src|href)=["\']https?://[^"\']+["\']', html)
    checks.append((len(external) == 0, 'Sin dependencias externas'))
    size_kb = len(html.encode('utf-8')) / 1024
    checks.append((size_kb < 500, f'Tamaño OK ({size_kb:.0f} KB)'))
    return checks


# ═══════════════════════════════════════════════════════════════════════════
# PROMPTS
# ═══════════════════════════════════════════════════════════════════════════
def construir_prompt_estructura(politecnico, docente, modulo, tema, descripcion, tema_visual, nivel):
    tokens_diseno = TEMAS_VISUALES[tema_visual]
    guia_nivel = NIVELES_COMPLEJIDAD[nivel]
    return f"""Actúa como un Diseñador UI/UX Senior especializado en simuladores educativos.
ESTA ES LA FASE 1 DE 2: SOLO ESTRUCTURA VISUAL (HTML + CSS). NO escribas JavaScript todavía, la lógica se generará en un segundo paso por separado.

CONTEXTO EDUCATIVO:
Institución: {politecnico}
Docente: {docente}
Asignatura: {modulo}
Tema: {tema}

REQUERIMIENTOS FUNCIONALES QUE EL SIMULADOR DEBERÁ CUMPLIR (úsalos solo para decidir qué inputs, botones y contenedores de resultado necesitas crear; NO los implementes en JS ahora):
{descripcion}

NIVEL DE COMPLEJIDAD ESPERADO (afecta cuántos campos/pasos/controles debes crear):
{guia_nivel}

SISTEMA DE DISEÑO A APLICAR (tema visual: "{tema_visual}"):
{tokens_diseno}
- Escala de espaciado consistente (múltiplos de 4px u 8px).
- Al menos 2 niveles de jerarquía tipográfica claros.
- Estados :hover / :focus con transición suave (transition: all 0.2s ease).
- Layout responsivo con CSS Grid o Flexbox y media query para max-width: 600px.

REGLAS INQUEBRANTABLES:
- NO uses la etiqueta <script>. Nada de JavaScript en esta fase.
- Todo el CSS debe estar en un único bloque <style> dentro del <head>.
- NO uses CDN externos ni fuentes de Google Fonts. Usa system-ui o fuentes del sistema.
- Incluye <meta name="viewport" content="width=device-width, initial-scale=1.0">.
- Usa HTML5 semántico: <header>, <main>, <section>, <footer>.
- Cada input DEBE tener un <label> asociado con el atributo "for".
- Incluye un contenedor visible para mostrar resultados (id="resultado").
- Incluye un botón de tipo reset (id="btn-reset") además del botón principal.
- Añade estilos para estados de error (.error) y éxito (.exito) en los inputs.
- Incluye un @media print básico que oculte botones.
- Todos los targets táctiles deben tener mínimo 44px de alto/ancho.
- Asegura contraste de color WCAG AA (mínimo 4.5:1 para texto normal).

Devuelve ÚNICAMENTE el código HTML completo. Nada de explicaciones, nada de markdown, nada de texto antes o después.
"""


def construir_prompt_logica(codigo_html, descripcion, nivel):
    guia_nivel = NIVELES_COMPLEJIDAD[nivel]
    return f"""Actúa como un Desarrollador Frontend Senior especializado en simuladores educativos.
ESTA ES LA FASE 2 DE 2: SOLO LÓGICA JAVASCRIPT. El HTML y CSS ya fueron generados en la fase anterior.

HTML YA GENERADO (analiza los IDs y clases existentes para escribir el JS correcto):
{codigo_html}

REQUERIMIENTOS FUNCIONALES DEL SIMULADOR:
{descripcion}

NIVEL DE COMPLEJIDAD:
{guia_nivel}

REGLAS INQUEBRANTABLES:
- Devuelve ÚNICAMENTE el código JavaScript. Nada de <script>, nada de HTML, nada de explicaciones.
- NO uses librerías externas. Todo debe ser JavaScript vanilla.
- NO uses alert(), confirm() ni prompt(). Usa elementos del DOM para mostrar mensajes.
- Envuelve todo el código en un IIFE: (() => {{ ... }})();
- Usa 'use strict'; al inicio del IIFE.
- Valida TODOS los inputs antes de procesar. Muestra mensajes de error claros en el DOM.
- El botón de reset (id="btn-reset") debe limpiar todos los campos y el área de resultado.
- Maneja errores con try/catch y muestra el error en el DOM, no en consola.
- Usa addEventListener, nunca atributos onclick inline.
- Asegúrate de que todos los IDs referenciados en el JS existan en el HTML proporcionado.
- Si el nivel es Avanzado, añade localStorage para persistir datos entre sesiones.
- Optimiza para rendimiento: minimiza reflows, usa DocumentFragment si insertas muchos nodos.
- Añade transiciones suaves al mostrar/ocultar resultados (usa classList.add/remove con clases CSS).
"""


def construir_prompt_refinamiento(codigo_actual, instrucciones_ajuste):
    return f"""Eres el mismo Diseñador/Desarrollador Senior que generó el siguiente simulador web (Single-File HTML). El usuario quiere AJUSTES puntuales sin romper lo que ya funciona.

CÓDIGO HTML ACTUAL:
{codigo_actual}

AJUSTES SOLICITADOS POR EL USUARIO:
{instrucciones_ajuste}

REGLAS:
- Conserva toda la funcionalidad existente que el usuario no pidió cambiar; no borres lógica que ya funcionaba.
- Sigue respetando: cero dependencias externas, cero <script src>, cero alert(), todo el CSS en un único <style>.
- Devuelve ÚNICAMENTE el código HTML completo y funcional con los ajustes aplicados. Nada de explicaciones ni markdown.
- Mantén la coherencia visual con el tema original.
- Si el ajuste requiere nuevos elementos HTML, añádelos con el mismo estilo visual existente.
"""


# ═══════════════════════════════════════════════════════════════════════════
# SUPER INTERFAZ
# ═══════════════════════════════════════════════════════════════════════════
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');
html, body, [class*="css"] { font-family: 'Inter', sans-serif; background-color: #F8FAFC; color: #1E293B; }
.sim-hero { background: linear-gradient(135deg, #0F172A 0%, #7C3AED 55%, #A78BFA 100%); color: #fff;
padding: 1.8rem; border-radius: 18px; margin-bottom: 1.25rem; box-shadow: 0 18px 35px rgba(15,23,42,.18); }
.sim-title { font-size: 2rem; font-weight: 800; letter-spacing: -0.02em; }
.sim-sub { opacity: .88; font-size: 1rem; margin-top: .3rem; }
.section-title { color: #7C3AED; font-weight: 700; font-size: 1.12rem; border-bottom: 2px solid #EDE9FE;
padding-bottom: 8px; margin: 1.2rem 0 .9rem 0; }
.quality-ok { color: #059669; font-weight: 600; }
.quality-fail { color: #DC2626; font-weight: 600; }
.metric-card { background:#fff; border:1px solid #E2E8F0; border-top:4px solid #7C3AED; border-radius:12px;
padding:14px 16px; box-shadow:0 4px 12px rgba(15,23,42,.06); text-align:center; }
.metric-value { font-size:1.9rem; font-weight:800; color:#0F172A; }
.metric-label { font-size:.78rem; font-weight:800; color:#64748B; text-transform:uppercase; letter-spacing:.05em; }
</style>
""", unsafe_allow_html=True)

ia.panel_sidebar_ia("Fábrica de Simuladores")

st.markdown("""
<div class="sim-hero">
    <div class="sim-title">💻 Fábrica de Simuladores ETP</div>
    <div class="sim-sub">Simuladores web interactivos single-file · 2 fases · 4 temas · 3 niveles · Galería permanente</div>
</div>
""", unsafe_allow_html=True)

# ── Estado de sesión ──
if "sim_html" not in st.session_state:
    st.session_state.sim_html = None
if "sim_metadata" not in st.session_state:
    st.session_state.sim_metadata = {}

tab_generar, tab_refinar, tab_galeria = st.tabs(["🏗️ Generar Simulador", "🔧 Refinar Simulador", "🗂️ Galería"])

# ═══════════════════════════════════════════════════════════════════════════
# TAB 1: GENERAR
# ═══════════════════════════════════════════════════════════════════════════
with tab_generar:
    with st.form("form_simulador", clear_on_submit=False):
        st.markdown('<div class="section-title">📋 1. Contexto Educativo</div>', unsafe_allow_html=True)
        col1, col2 = st.columns(2)
        with col1:
            politecnico = st.text_input("Institución", value="Politécnico Salesiano Arquides Calderón")
            docente = st.text_input("Docente", value="Ing. Bernardo Antonio Hernández Batista")
            modulo = st.text_input("Asignatura / Módulo", placeholder="Ej: Física, Matemática, Redes LAN")
        with col2:
            tema = st.text_input("Tema del Simulador", placeholder="Ej: Ley de Ohm, Subredes IP")
            tema_visual = st.selectbox("🎨 Tema Visual", list(TEMAS_VISUALES.keys()))
            nivel = st.selectbox("⚙️ Nivel de Complejidad", list(NIVELES_COMPLEJIDAD.keys()))

        st.markdown('<div class="section-title">📝 2. Requerimientos Funcionales</div>', unsafe_allow_html=True)
        descripcion = st.text_area(
            "Describe qué debe hacer el simulador (inputs, cálculos, resultados esperados):",
            height=120,
            placeholder="Ej: El estudiante ingresa voltaje (V) y resistencia (Ω). El simulador calcula la corriente (I) usando la Ley de Ohm y muestra el resultado con unidades.",
        )
        max_tokens, temperature = ia.control_avanzado(default_tokens=16384, tope=32000, default_temp=0.15)
        st.markdown("<br>", unsafe_allow_html=True)
        submit_btn = st.form_submit_button("⚡ Generar Simulador (2 Fases)", type="primary", width="stretch")

    if submit_btn:
        cfg = ia.config_ia()
        if not cfg["api_key"]:
            st.error("🔒 Configura tu API Key en la página de Inicio.")
        elif not tema or not descripcion:
            st.warning("📝 Completa el tema y los requerimientos funcionales.")
        else:
            with st.spinner(f"🧠 Fase 1/2: Diseñando estructura visual con {cfg['modelo']}..."):
                texto_html = None
                texto_js = None
                try:
                    # ── FASE 1: ESTRUCTURA HTML + CSS ──
                    prompt_f1 = construir_prompt_estructura(
                        politecnico, docente, modulo, tema, descripcion, tema_visual, nivel
                    )
                    texto_html, flags1 = ia.solicitar_texto(
                        prompt_f1, max_tokens=max_tokens, temperature=0.15, modulo="simuladores"
                    )
                    texto_html = limpiar_respuesta_html(texto_html)
                    texto_html = balancear_html(texto_html)
                    if flags1.get("truncado"):
                        st.warning("⚠️ Fase 1: la estructura HTML se cortó por tokens. Se aplicó balanceo automático.")

                    # ── FASE 2: LÓGICA JAVASCRIPT ──
                    st.toast("✅ Fase 1 completada. Generando lógica...", icon="🔧")
                    prompt_f2 = construir_prompt_logica(texto_html, descripcion, nivel)
                    texto_js, flags2 = ia.solicitar_texto(
                        prompt_f2, max_tokens=max_tokens, temperature=0.15, modulo="simuladores"
                    )
                    texto_js = limpiar_respuesta_html(texto_js)
                    texto_js = re.sub(r'^```(?:javascript|js)?\s*', '', texto_js, flags=re.IGNORECASE)
                    texto_js = re.sub(r'\s*```$', '', texto_js)
                    if flags2.get("truncado"):
                        st.warning("⚠️ Fase 2: el JavaScript se cortó por tokens. El simulador puede estar incompleto.")

                    # ── COMBINAR + CALIDAD ──
                    html_final = insertar_script(texto_html, texto_js)
                    html_final = balancear_html(html_final)
                    calidad = evaluar_calidad(html_final)
                    ok_count = sum(1 for ok, _ in calidad if ok)
                    size_kb = len(html_final.encode('utf-8')) / 1024

                    # ── GUARDAR EN SESIÓN + GALERÍA ──
                    st.session_state.sim_html = html_final
                    st.session_state.sim_metadata = {
                        "tema": tema, "modulo": modulo, "tema_visual": tema_visual,
                        "nivel": nivel, "fecha": datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
                        "size_kb": size_kb,
                    }
                    guardar_simulador(tema, modulo, tema_visual, nivel, html_final)

                    # ── MÉTRICAS ──
                    st.toast("✅ ¡Simulador generado y guardado en la Galería!", icon="💻")
                    m1, m2, m3, m4 = st.columns(4)
                    with m1:
                        st.markdown(f'<div class="metric-card"><div class="metric-label">Calidad</div><div class="metric-value">{ok_count}/10</div></div>', unsafe_allow_html=True)
                    with m2:
                        st.markdown(f'<div class="metric-card"><div class="metric-label">Tamaño</div><div class="metric-value">{size_kb:.0f} KB</div></div>', unsafe_allow_html=True)
                    with m3:
                        st.markdown(f'<div class="metric-card"><div class="metric-label">Tema</div><div class="metric-value" style="font-size:1rem;">{tema_visual.split(" ")[0]}</div></div>', unsafe_allow_html=True)
                    with m4:
                        st.markdown(f'<div class="metric-card"><div class="metric-label">Nivel</div><div class="metric-value" style="font-size:1rem;">{nivel.split(" ")[0]}</div></div>', unsafe_allow_html=True)

                    with st.expander(f"🔍 Checklist de Calidad ({ok_count}/10)", expanded=False):
                        for ok, msg in calidad:
                            icon = "✅" if ok else "❌"
                            cls = "quality-ok" if ok else "quality-fail"
                            st.markdown(f'<span class="{cls}">{icon} {msg}</span>', unsafe_allow_html=True)

                    if flags1.get("reintento") or flags2.get("reintento"):
                        st.info("ℹ️ Se aplicó reintento automático por truncamiento en al menos una fase.")
                except ValueError as ve:
                    ia.render_error_ia(ve, texto_html or texto_js)
                except Exception as e:
                    ia.render_error_ia(e, texto_html or texto_js)

    # ── RESULTADO (fuera del form) — FIX: components.html ──
    if st.session_state.sim_html:
        st.markdown('<div class="section-title">👁️ Vista Previa del Simulador</div>', unsafe_allow_html=True)
        components.html(st.session_state.sim_html, height=600, scrolling=True)
        col_dl, col_open = st.columns(2)
        with col_dl:
            nombre = ia.sanear_nombre_archivo(st.session_state.sim_metadata.get("tema", "simulador"))
            st.download_button(
                label="📥 Descargar Simulador (.html)",
                data=st.session_state.sim_html,
                file_name=f"{nombre}.html",
                mime="text/html",
                type="primary",
                width="stretch",
            )
        with col_open:
            data_uri = generar_data_uri(st.session_state.sim_html)
            st.markdown(
                f'<a href="{data_uri}" target="_blank" style="display:block;text-align:center;'
                f'padding:10px;background:#7C3AED;color:#fff;border-radius:6px;text-decoration:none;'
                f'font-weight:600;">🌐 Abrir en Nueva Pestaña</a>',
                unsafe_allow_html=True,
            )

# ═══════════════════════════════════════════════════════════════════════════
# TAB 2: REFINAR
# ═══════════════════════════════════════════════════════════════════════════
with tab_refinar:
    if not st.session_state.sim_html:
        st.info("📝 Primero genera un simulador en '🏗️ Generar' o cárgalo desde la '🗂️ Galería'.")
    else:
        st.markdown(f'**Simulador actual:** {st.session_state.sim_metadata.get("tema", "N/A")} '
                    f'({st.session_state.sim_metadata.get("size_kb", 0):.0f} KB)')
        instrucciones = st.text_area(
            "🔧 Describe los ajustes que deseas (sin romper lo que funciona):",
            height=100,
            placeholder="Ej: Cambiar el color de los botones a verde, añadir un campo para la temperatura...",
        )
        if st.button("🔧 Aplicar Ajustes", type="primary", width="stretch"):
            if not instrucciones.strip():
                st.warning("⚠️ Describe los ajustes que deseas.")
            else:
                with st.spinner("🧠 Aplicando ajustes al simulador..."):
                    try:
                        prompt_ref = construir_prompt_refinamiento(
                            st.session_state.sim_html, instrucciones
                        )
                        texto_ref, flags_ref = ia.solicitar_texto(
                            prompt_ref, max_tokens=32000, temperature=0.15, modulo="simuladores"
                        )
                        texto_ref = limpiar_respuesta_html(texto_ref)
                        texto_ref = balancear_html(texto_ref)
                        calidad_ref = evaluar_calidad(texto_ref)
                        ok_ref = sum(1 for ok, _ in calidad_ref if ok)
                        st.session_state.sim_html = texto_ref
                        st.session_state.sim_metadata["size_kb"] = len(texto_ref.encode('utf-8')) / 1024
                        st.toast(f"✅ Ajustes aplicados. Calidad: {ok_ref}/10", icon="🔧")
                        st.rerun()
                    except ValueError as ve:
                        ia.render_error_ia(ve, None)
                    except Exception as e:
                        ia.render_error_ia(e, None)

        if st.session_state.sim_html:
            st.markdown('<div class="section-title">👁️ Vista Previa (Actualizada)</div>', unsafe_allow_html=True)
            components.html(st.session_state.sim_html, height=600, scrolling=True)
            nombre = ia.sanear_nombre_archivo(st.session_state.sim_metadata.get("tema", "simulador"))
            st.download_button(
                label="📥 Descargar Simulador Refinado (.html)",
                data=st.session_state.sim_html,
                file_name=f"{nombre}_refinado.html",
                mime="text/html",
                type="primary",
                width="stretch",
            )

# ═══════════════════════════════════════════════════════════════════════════
# TAB 3: GALERÍA (persistencia SQLite) — SIN TERNARIAS SUELTAS
# ═══════════════════════════════════════════════════════════════════════════
with tab_galeria:
    st.markdown('<div class="section-title">🗂️ Galería de Simuladores Guardados</div>', unsafe_allow_html=True)
    sims = listar_simuladores()

    if not sims:
        st.info("Aún no hay simuladores guardados. Genera uno en la pestaña '🏗️ Generar Simulador'.")
    else:
        opciones = {f"#{s['id']} · {s['tema']} ({s['fecha']})": s["id"] for s in sims}
        sel = st.selectbox("Selecciona un simulador", list(opciones.keys()))
        id_sel = opciones[sel]
        html_sel = obtener_simulador(id_sel)
        info_sel = next(s for s in sims if s["id"] == id_sel)

        col_a, col_b, col_c = st.columns(3)
        with col_a:
            st.download_button(
                "📥 Descargar (.html)",
                data=html_sel,
                file_name=f"{ia.sanear_nombre_archivo(info_sel['tema'])}.html",
                mime="text/html",
                width="stretch",
            )
        with col_b:
            if st.button("🔧 Cargar en Refinar", width="stretch"):
                st.session_state.sim_html = html_sel
                st.session_state.sim_metadata = {
                    "tema": info_sel["tema"], "modulo": info_sel["modulo"],
                    "tema_visual": info_sel["tema_visual"], "nivel": info_sel["nivel"],
                    "size_kb": len(html_sel.encode("utf-8")) / 1024,
                }
                st.toast("✅ Simulador cargado en Refinar.", icon="🔧")
                st.rerun()
        with col_c:
            if st.button("🗑️ Eliminar", width="stretch"):
                eliminar_simulador(id_sel)
                st.toast("🗑️ Simulador eliminado.", icon="🗑️")
                st.rerun()

        st.markdown("---")
        df_sims = pd.DataFrame(sims).drop(columns=["id"])
        st.dataframe(df_sims, width="stretch", hide_index=True)