"""
core/ia.py — Motor de IA unificado del Sistema de Gestión ETP (v2.2 · corregido)
═══════════════════════════════════════════════════════════════════════════════
• Unificación: marcadores, parseo robusto, reintentos y dispatch en UN solo lugar.
• Tres proveedores: Google Gemini, OpenAI y Anthropic (Claude).
• Errores en español claro + expander estándar "ver respuesta cruda".
• Límite de tokens y temperatura configurables por módulo.
• Mini-auditoría de IA: módulo, proveedor/modelo, tokens, truncamiento y reintento.
• Soporte schema_json para salida estructurada (Gemini response_schema).
"""
from __future__ import annotations

import json
import logging
import re
import unicodedata
from typing import Any, Dict, Optional, Tuple

import streamlit as st

from google.api_core.exceptions import ResourceExhausted
from openai import OpenAI, RateLimitError as OpenAIRateLimitError
from tenacity import (
    Retrying,
    retry,
    retry_if_exception_type,
    stop_after_attempt,
    wait_exponential,
)
import google.generativeai as genai

try:
    import anthropic
    from anthropic import RateLimitError as AnthropicRateLimitError
    ANTHROPIC_DISPONIBLE = True
except Exception:
    anthropic = None
    AnthropicRateLimitError = tuple()
    ANTHROPIC_DISPONIBLE = False

logger = logging.getLogger("core.ia")

# ═══════════════════════════════════════════════════════════════════════════
# CATÁLOGOS DE MODELOS POR PROVEEDOR
# ═══════════════════════════════════════════════════════════════════════════
MODELOS_GEMINI = [
    "gemini-3.5-flash", "gemini-2.5-flash", "gemini-2.5-pro", "gemini-2.0-flash",
    "gemini-1.5-pro", "gemini-1.5-flash",
]
MODELOS_OPENAI = ["gpt-4o-mini", "gpt-4o", "gpt-4.1-mini", "gpt-4.1"]
MODELOS_ANTHROPIC = [
    "claude-3-5-sonnet-latest", "claude-3-5-haiku-latest", "claude-3-opus-latest",
]


def catalogo_modelos(proveedor: str) -> list:
    prov = str(proveedor or "")
    if prov.startswith("OpenAI"):
        return list(MODELOS_OPENAI)
    if prov.startswith("Anthropic"):
        return list(MODELOS_ANTHROPIC)
    return list(MODELOS_GEMINI)


# ═══════════════════════════════════════════════════════════════════════════
# MARCADORES SEGUROS
# ═══════════════════════════════════════════════════════════════════════════
MARKER_NL = "<<NL>>"
MARKER_DQ = "<<DQ>>"
MARKER_TAB = "<<TAB>>"


def codificar_marcadores_texto(texto: str) -> str:
    texto = str(texto)
    texto = texto.replace("\r\n", MARKER_NL).replace("\n", MARKER_NL).replace("\r", MARKER_NL)
    texto = texto.replace("\t", MARKER_TAB)
    texto = texto.replace('"', MARKER_DQ)
    return texto


def decodificar_marcadores(obj: Any) -> Any:
    if isinstance(obj, str):
        return obj.replace(MARKER_NL, "\n").replace(MARKER_DQ, '"').replace(MARKER_TAB, "\t")
    if isinstance(obj, dict):
        return {decodificar_marcadores(k): decodificar_marcadores(v) for k, v in obj.items()}
    if isinstance(obj, list):
        return [decodificar_marcadores(item) for item in obj]
    return obj


# ═══════════════════════════════════════════════════════════════════════════
# JSON ROBUSTO
# ═══════════════════════════════════════════════════════════════════════════
def reparar_json_truncado(texto: str) -> str:
    in_string = False
    escape_next = False
    llaves = 0
    corchetes = 0
    last_safe_pos = 0
    for i, char in enumerate(texto):
        if escape_next:
            escape_next = False
            continue
        if in_string:
            if char == "\\":
                escape_next = True
            elif char == '"':
                in_string = False
                last_safe_pos = i + 1
            continue
        if char == '"':
            in_string = True
        elif char == "{":
            llaves += 1
            last_safe_pos = i + 1
        elif char == "}":
            llaves -= 1
            last_safe_pos = i + 1
        elif char == "[":
            corchetes += 1
            last_safe_pos = i + 1
        elif char == "]":
            corchetes -= 1
            last_safe_pos = i + 1
        elif char in (",", ":", " ", "\n", "\r", "\t"):
            last_safe_pos = i + 1
    reparado = texto[:last_safe_pos]
    if in_string:
        reparado += '"'
    reparado = reparado.rstrip()
    if reparado.endswith(","):
        reparado = reparado[:-1]
    reparado += "]" * max(corchetes, 0)
    reparado += "}" * max(llaves, 0)
    return reparado


def parsear_json_robusto(respuesta: str) -> Any:
    if not respuesta or not str(respuesta).strip():
        raise ValueError("La IA devolvió una respuesta vacía.")
    texto = str(respuesta).strip()
    # Quitar fences de markdown
    if texto.startswith("```json"):
        texto = texto[7:]
    elif texto.startswith("```"):
        texto = texto[3:]
    if texto.endswith("```"):
        texto = texto[:-3]
    texto = texto.strip()
    # Intento 1: parseo directo
    try:
        return json.loads(texto, strict=False)
    except json.JSONDecodeError:
        pass
    # Intento 2: extraer el objeto JSON más externo
    match = re.search(r"(\{[\s\S]*\})", texto)
    if match:
        try:
            return json.loads(match.group(1), strict=False)
        except json.JSONDecodeError:
            pass
    # Intento 3: reparar JSON truncado
    json_start = texto.find("{")
    if json_start >= 0:
        cuerpo = texto[json_start:]
        try:
            return json.loads(reparar_json_truncado(cuerpo), strict=False)
        except json.JSONDecodeError:
            pass
        for fin in range(len(cuerpo), max(len(cuerpo) - 8000, json_start), -1):
            if fin <= json_start:
                break
            if cuerpo[fin - 1] == "}":
                try:
                    return json.loads(reparar_json_truncado(cuerpo[:fin]), strict=False)
                except json.JSONDecodeError:
                    continue
    # Intento 4: limpiar caracteres de control
    try:
        return json.loads(re.sub(r"[\x00-\x1f]", " ", texto), strict=False)
    except json.JSONDecodeError:
        pass
    raise ValueError(f"JSON irrecuperable. Inicio de la respuesta: {texto[:400]}...")


# ═══════════════════════════════════════════════════════════════════════════
# CONFIGURACIÓN DESDE SESIÓN
# ═══════════════════════════════════════════════════════════════════════════
def config_ia() -> Dict[str, str]:
    proveedor = str(st.session_state.get("proveedor_ia_global") or "Google Gemini").strip()
    api_key = str(st.session_state.get("api_key_global") or "").strip()
    usar_custom = bool(st.session_state.get("usar_modelo_custom", False))
    custom = str(st.session_state.get("modelo_custom_text") or "").strip()
    modelo = custom if (usar_custom and custom) else str(
        st.session_state.get("modelo_global") or "gemini-2.5-flash"
    ).strip()
    return {"proveedor": proveedor, "api_key": api_key, "modelo": modelo}


# ═══════════════════════════════════════════════════════════════════════════
# LLAMADAS POR PROVEEDOR (devuelven (texto, truncado))
# ═══════════════════════════════════════════════════════════════════════════
@retry(retry=retry_if_exception_type(ResourceExhausted),
       wait=wait_exponential(multiplier=2, min=4, max=20), stop=stop_after_attempt(5), reraise=True)
def _gemini(api_key, modelo, prompt, modo, max_tokens, temperature, schema_json=None):
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel(modelo)
    kwargs = {"max_output_tokens": max_tokens, "temperature": temperature}
    if modo == "json":
        kwargs["response_mime_type"] = "application/json"
        if schema_json is not None:
            try:
                kwargs["response_schema"] = schema_json
            except Exception:
                pass
    try:
        respuesta = model.generate_content(
            prompt, generation_config=genai.types.GenerationConfig(**kwargs))
    except TypeError:
        kwargs.pop("response_mime_type", None)
        kwargs.pop("response_schema", None)
        respuesta = model.generate_content(
            prompt, generation_config=genai.types.GenerationConfig(**kwargs))
    truncado = False
    try:
        fr = respuesta.candidates[0].finish_reason
        truncado = "MAX_TOKENS" in str(fr).upper() or fr == 2
    except Exception:
        pass
    try:
        texto = respuesta.text
    except Exception:
        texto = ""
    return texto or "", truncado


@retry(retry=retry_if_exception_type(OpenAIRateLimitError),
       wait=wait_exponential(multiplier=2, min=4, max=20), stop=stop_after_attempt(5), reraise=True)
def _openai(api_key, modelo, prompt, modo, max_tokens, temperature, schema_json=None):
    client = OpenAI(api_key=api_key)
    payload = {"model": modelo, "messages": [{"role": "user", "content": prompt}],
               "temperature": temperature, "max_tokens": max_tokens}
    if modo == "json":
        payload["response_format"] = {"type": "json_object"}
    try:
        response = client.chat.completions.create(**payload)
    except Exception as exc:
        if "response_format" in str(exc).lower():
            payload.pop("response_format", None)
            response = client.chat.completions.create(**payload)
        else:
            raise
    return (response.choices[0].message.content or "",
            response.choices[0].finish_reason == "length")


def _anthropic(api_key, modelo, prompt, modo, max_tokens, temperature, schema_json=None):
    if not ANTHROPIC_DISPONIBLE:
        raise ValueError("Anthropic no está instalado. Ejecuta: pip install anthropic")
    if modo == "json":
        prompt += "\n\nDevuelve ÚNICAMENTE un objeto JSON válido, sin Markdown."

    def _llamada():
        cliente = anthropic.Anthropic(api_key=api_key)
        resp = cliente.messages.create(
            model=modelo, max_tokens=max_tokens, temperature=temperature,
            messages=[{"role": "user", "content": prompt}],
        )
        texto = "".join(b.text for b in resp.content if getattr(b, "type", "") == "text")
        return texto, resp.stop_reason == "max_tokens"

    for attempt in Retrying(
        stop=stop_after_attempt(5), wait=wait_exponential(multiplier=2, min=4, max=20),
        retry=retry_if_exception_type(AnthropicRateLimitError), reraise=True,
    ):
        with attempt:
            return _llamada()


def _llamar(proveedor, api_key, modelo, prompt, modo, max_tokens, temperature, schema_json=None):
    if str(proveedor).startswith("OpenAI"):
        return _openai(api_key, modelo, prompt, modo, max_tokens, temperature, schema_json)
    if str(proveedor).startswith("Anthropic"):
        return _anthropic(api_key, modelo, prompt, modo, max_tokens, temperature, schema_json)
    return _gemini(api_key, modelo, prompt, modo, max_tokens, temperature, schema_json)


# ═══════════════════════════════════════════════════════════════════════════
# AUDITORÍA
# ═══════════════════════════════════════════════════════════════════════════
def _auditar_ia(modulo, proveedor, modelo, max_tokens, truncado, reintento):
    try:
        from core import auth
        auth.registrar_evento(
            "Llamada IA", modulo,
            f"{proveedor} / {modelo} · tokens={max_tokens} · truncado={truncado} · reintento={reintento}",
        )
    except Exception:
        logger.warning("No se pudo auditar la llamada de IA.")


# ═══════════════════════════════════════════════════════════════════════════
# API PÚBLICA
# ═══════════════════════════════════════════════════════════════════════════
def solicitar_ia(prompt: str, modo: str = "json", max_tokens: int = 8192,
                 temperature: float = 0.2, tope: int = 32000,
                 modulo: str = "general", schema_json=None) -> Tuple[str, Dict[str, Any]]:
    """Llama al proveedor configurado. Si detecta corte por tokens, reintenta
    UNA vez con el doble de presupuesto. Devuelve (texto, flags)."""
    cfg = config_ia()
    if not cfg["api_key"]:
        raise ValueError("No hay API Key configurada. Ingrésala en Inicio → Configuración de IA.")
    texto, truncado = _llamar(cfg["proveedor"], cfg["api_key"], cfg["modelo"],
                              prompt, modo, max_tokens, temperature, schema_json)
    reintento = False
    if truncado and max_tokens < tope:
        texto2, truncado2 = _llamar(cfg["proveedor"], cfg["api_key"], cfg["modelo"],
                                    prompt, modo, min(max_tokens * 2, tope), temperature, schema_json)
        reintento = True
        if len(texto2 or "") >= len(texto or ""):
            texto, truncado = texto2, truncado2
    _auditar_ia(modulo, cfg["proveedor"], cfg["modelo"], max_tokens, truncado, reintento)
    return texto, {"truncado": truncado, "reintento": reintento,
                   "proveedor": cfg["proveedor"], "modelo": cfg["modelo"]}


def solicitar_json(prompt: str, max_tokens: int = 8192, temperature: float = 0.2,
                   tope: int = 32000, modulo: str = "general", schema_json=None) -> Tuple[Any, Dict[str, Any]]:
    """JSON listo para usar: llama, repara, decodifica marcadores."""
    texto, flags = solicitar_ia(prompt, modo="json", max_tokens=max_tokens,
                                temperature=temperature, tope=tope, modulo=modulo,
                                schema_json=schema_json)
    datos = parsear_json_robusto(texto)
    return decodificar_marcadores(datos), {**flags, "raw": texto}


def solicitar_texto(prompt: str, max_tokens: int = 8192, temperature: float = 0.3,
                    tope: int = 32000, modulo: str = "general", schema_json=None) -> Tuple[str, Dict[str, Any]]:
    """Modo texto/Markdown (redactor, simuladores, etc.)."""
    return solicitar_ia(prompt, modo="texto", max_tokens=max_tokens,
                        temperature=temperature, tope=tope, modulo=modulo,
                        schema_json=schema_json)


# ═══════════════════════════════════════════════════════════════════════════
# INTERFAZ ESTÁNDAR
# ═══════════════════════════════════════════════════════════════════════════
def render_error_ia(error: Exception, respuesta_cruda: Optional[str] = None) -> None:
    """Mensaje de error en español claro + diagnóstico opcional."""
    msg = str(error)
    bajo = msg.lower()
    if isinstance(error, ResourceExhausted) or "429" in bajo or "quota" in bajo or "rate" in bajo:
        st.error("❌ Límite de la API alcanzado. Espera unos momentos e inténtalo de nuevo.")
    elif isinstance(error, ValueError):
        st.error(f"⚠️ La IA devolvió una respuesta que no se pudo procesar: {msg}")
        st.info("💡 Sugerencia: sube el límite de tokens en ⚙️ Configuración avanzada o reduce la cantidad solicitada.")
    else:
        st.error(f"⚠️ Error al consultar la IA: {msg}")
    if respuesta_cruda:
        with st.expander("🔍 Ver respuesta cruda de la IA (diagnóstico)"):
            st.text(str(respuesta_cruda)[:4000])


def panel_sidebar_ia(titulo_modulo: str) -> None:
    """Bloque lateral estándar: estado del proveedor/modelo activo."""
    with st.sidebar:
        st.markdown(f"##### ⚡ {titulo_modulo}")
        cfg = config_ia()
        if not cfg["api_key"]:
            st.error("🔒 Configura tu API Key en la página de Inicio")
        else:
            st.success(f"✅ {cfg['proveedor']} · {cfg['modelo']}")


def control_avanzado(default_tokens: int = 16384, tope: int = 32000,
                     default_temp: float = 0.15) -> Tuple[int, float]:
    """Expander reutilizable para tokens y temperatura."""
    with st.expander("⚙️ Configuración avanzada de generación"):
        tokens = st.slider("🧠 Límite de tokens", 4096, tope, default_tokens, step=2048,
                           help="Si la respuesta sale cortada o el JSON falla, súbelo.")
        temp = st.slider("🎛️ Temperatura", 0.0, 1.0, default_temp, 0.05,
                         help="Valores bajos = respuestas más estables y fieles.")
    return int(tokens), float(temp)


def sanear_nombre_archivo(texto: str, por_defecto: str = "documento") -> str:
    """Genera un nombre de archivo seguro (ASCII, sin símbolos, máx. 60 caracteres).
    Nota: el guion '-' va al FINAL de la clase de caracteres para evitar
    el error 'bad character range' de Python."""
    try:
        texto = str(texto or "").strip() or por_defecto
        texto = unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode("ascii")
        texto = re.sub(r"[^A-Za-z0-9_\s-]", "", texto)
        texto = re.sub(r"\s+", "_", texto)
        return texto[:60].strip("_") or por_defecto
    except Exception:
        return por_defecto