"""
core/auth.py — Autenticación y gestión de accesos (v4 · reestructurado)
═══════════════════════════════════════════════════════════════════════════
Fuente única de verdad para usuarios, módulos, auditoría y directorio.

Correcciones de la reestructuración:
• eliminar_todos_usuarios / importar_usuario_con_modulos ahora son funciones
  de módulo (antes eran código muerto dentro de obtener_directorio → AttributeError).
• Sin contraseñas demo: importaciones sin contraseña quedan "Sin configurar".
• Esquema completo centralizado (usuarios, módulos, log, directorio_extras,
  config_coordinador) con migración automática e idempotente.
• Índices para acelerar consultas.
• Coordinador: secrets → env → SQLite → legado opcional (vacío = desactivado).
• Docentes: PBKDF2-HMAC-SHA256 (sal + 100k iteraciones), nunca texto plano.
"""
from __future__ import annotations

import hashlib
import logging
import os
import re
import secrets
import sqlite3
from datetime import datetime
from typing import Dict, List, Optional, Tuple

import streamlit as st

logger = logging.getLogger("core.auth")

# ═══════════════════════════════════════════════════════════════════════════
# CONSTANTES
# ═══════════════════════════════════════════════════════════════════════════
DB_NAME = "gestion_etp.db"
PBKDF2_ITER = 100_000
USUARIO_RE = re.compile(r"^[a-z0-9._-]{3,30}$")

# Legado opcional del coordinador: deja ambos vacíos para DESACTIVARLO.
COORD_LEGACY_USUARIO = ""
COORD_LEGACY_CLAVE = ""


# ═══════════════════════════════════════════════════════════════════════════
# HELPERS INTERNOS
# ═══════════════════════════════════════════════════════════════════════════
def _conn() -> sqlite3.Connection:
    return sqlite3.connect(DB_NAME, check_same_thread=False)


def _ahora() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def _columnas(cur: sqlite3.Cursor, tabla: str) -> List[str]:
    cur.execute(f"PRAGMA table_info({tabla})")
    return [r[1] for r in cur.fetchall()]


def _tabla_existe(cur: sqlite3.Cursor, tabla: str) -> bool:
    cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name=?", (tabla,))
    return cur.fetchone() is not None


# ═══════════════════════════════════════════════════════════════════════════
# HASH PBKDF2
# ═══════════════════════════════════════════════════════════════════════════
def _hash_clave(clave, salt: Optional[bytes] = None) -> str:
    if salt is None:
        salt = secrets.token_bytes(16)
    dk = hashlib.pbkdf2_hmac("sha256", str(clave).encode("utf-8"), salt, PBKDF2_ITER)
    return salt.hex() + "$" + dk.hex()


def _verificar_clave(clave, almacenado) -> bool:
    try:
        salt_hex, hash_hex = str(almacenado).split("$", 1)
        dk = hashlib.pbkdf2_hmac("sha256", str(clave).encode("utf-8"),
                                 bytes.fromhex(salt_hex), PBKDF2_ITER)
        return secrets.compare_digest(dk.hex(), hash_hex)
    except Exception:
        return False


def _normalizar_usuario(usuario) -> str:
    u = str(usuario or "").strip().lower()
    if not USUARIO_RE.fullmatch(u):
        raise ValueError("Usuario inválido: 3–30 caracteres (minúsculas, números, . _ -).")
    return u


def _usuario_base(nombre: str) -> str:
    return re.sub(r"[^a-z0-9._-]", "", str(nombre).strip().split()[0].lower()) or "docente"


# ═══════════════════════════════════════════════════════════════════════════
# ESQUEMA + MIGRACIÓN AUTOMÁTICA (idempotente, sin perder datos)
# ═══════════════════════════════════════════════════════════════════════════
def asegurar_esquema() -> None:
    conn = _conn()
    cur = conn.cursor()

    # ── 1) usuarios_docentes ──
    cur.execute("""CREATE TABLE IF NOT EXISTS usuarios_docentes (
        docente TEXT PRIMARY KEY,
        usuario TEXT UNIQUE NOT NULL,
        password_hash TEXT,
        creado_en TEXT
    )""")
    cols_u = _columnas(cur, "usuarios_docentes")
    if "password_hash" not in cols_u:
        cur.execute("ALTER TABLE usuarios_docentes ADD COLUMN password_hash TEXT")
    if "usuario" not in cols_u:
        cur.execute("ALTER TABLE usuarios_docentes ADD COLUMN usuario TEXT")
    if "creado_en" not in cols_u:
        cur.execute("ALTER TABLE usuarios_docentes ADD COLUMN creado_en TEXT")

    # Migración desde la tabla legacy `docentes` (texto plano → PBKDF2)
    if _tabla_existe(cur, "docentes"):
        cur.execute("SELECT docente, usuario, password FROM docentes")
        for docente, usuario, pwd in cur.fetchall():
            if not docente:
                continue
            cur.execute("SELECT 1 FROM usuarios_docentes WHERE docente=?", (docente,))
            if cur.fetchone():
                continue
            base = _usuario_base(usuario or docente)
            cur.execute("SELECT docente FROM usuarios_docentes WHERE usuario=?", (base,))
            if cur.fetchone():
                base = f"{base}{secrets.token_hex(2)}"
            cur.execute(
                "INSERT INTO usuarios_docentes (docente, usuario, password_hash, creado_en) "
                "VALUES (?,?,?,?)",
                (docente, base, _hash_clave(pwd or "1234"), _ahora()),
            )

    # ── 2) modulos_docentes (con migración de esquema antiguo) ──
    cur.execute("""CREATE TABLE IF NOT EXISTS modulos_docentes (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        docente TEXT NOT NULL,
        modulo TEXT NOT NULL,
        seccion TEXT NOT NULL
    )""")
    cols_m = _columnas(cur, "modulos_docentes")
    if "docente" not in cols_m:
        cur.execute("ALTER TABLE modulos_docentes ADD COLUMN docente TEXT")
        if "usuario" in cols_m:
            cur.execute("SELECT docente, usuario FROM usuarios_docentes")
            mapa = {u: d for d, u in cur.fetchall()}
            cur.execute("SELECT id, usuario FROM modulos_docentes")
            for mid, usu in cur.fetchall():
                cur.execute("UPDATE modulos_docentes SET docente=? WHERE id=?",
                            (mapa.get(usu, usu), mid))
        cur.execute("UPDATE modulos_docentes SET docente='' WHERE docente IS NULL")
    if "modulo" not in cols_m:
        cur.execute("ALTER TABLE modulos_docentes ADD COLUMN modulo TEXT")
        cur.execute("UPDATE modulos_docentes SET modulo='' WHERE modulo IS NULL")
    if "seccion" not in cols_m:
        cur.execute("ALTER TABLE modulos_docentes ADD COLUMN seccion TEXT")
        cur.execute("UPDATE modulos_docentes SET seccion='' WHERE seccion IS NULL")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_modulos_docente ON modulos_docentes(docente)")

    # ── 3) accesos_log ──
    cur.execute("""CREATE TABLE IF NOT EXISTS accesos_log (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        fecha TEXT, accion TEXT, usuario TEXT, detalle TEXT
    )""")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_log_fecha ON accesos_log(fecha)")

    # ── 4) directorio_extras (área técnica y horas) ──
    cur.execute("""CREATE TABLE IF NOT EXISTS directorio_extras (
        docente TEXT PRIMARY KEY,
        area_tecnica TEXT,
        horas_modulo TEXT
    )""")

    # ── 5) config_coordinador (credenciales SQLite del coordinador) ──
    cur.execute("""CREATE TABLE IF NOT EXISTS config_coordinador (
        id INTEGER PRIMARY KEY CHECK (id = 1),
        usuario TEXT NOT NULL,
        clave_hash TEXT NOT NULL,
        nombre TEXT,
        actualizado_en TEXT
    )""")

    conn.commit()
    conn.close()


# ═══════════════════════════════════════════════════════════════════════════
# AUDITORÍA
# ═══════════════════════════════════════════════════════════════════════════
def registrar_evento(accion: str, usuario: str, detalle: str = "") -> None:
    try:
        conn = _conn()
        conn.execute("INSERT INTO accesos_log (fecha, accion, usuario, detalle) VALUES (?,?,?,?)",
                     (_ahora(), accion, usuario, detalle))
        conn.commit()
        conn.close()
    except Exception:
        logger.warning("No se pudo registrar el evento: %s", accion)


def obtener_log() -> List[list]:
    conn = _conn()
    cur = conn.cursor()
    cur.execute("SELECT fecha, accion, usuario, detalle FROM accesos_log ORDER BY id DESC")
    rows = [list(r) for r in cur.fetchall()]
    conn.close()
    return rows


# ═══════════════════════════════════════════════════════════════════════════
# AUTENTICACIÓN DEL COORDINADOR (secrets → env → SQLite → legado)
# ═══════════════════════════════════════════════════════════════════════════
def obtener_credenciales_coordinador() -> Optional[Tuple[str, str, str]]:
    """Devuelve (usuario, clave_hash, nombre) desde SQLite o None."""
    try:
        conn = _conn()
        cur = conn.cursor()
        cur.execute("SELECT usuario, clave_hash, nombre FROM config_coordinador WHERE id = 1")
        row = cur.fetchone()
        conn.close()
        return row if row else None
    except Exception:
        return None


def guardar_credenciales_coordinador(usuario: str, clave: str, nombre: Optional[str] = None) -> None:
    conn = _conn()
    conn.execute('''
        INSERT INTO config_coordinador (id, usuario, clave_hash, nombre, actualizado_en)
        VALUES (1, ?, ?, ?, ?)
        ON CONFLICT(id) DO UPDATE SET usuario=excluded.usuario,
        clave_hash=excluded.clave_hash, nombre=excluded.nombre,
        actualizado_en=excluded.actualizado_en
    ''', (usuario, _hash_clave(clave), nombre, _ahora()))
    conn.commit()
    conn.close()


def autenticar_coordinador(usuario: str, clave: str) -> Tuple[bool, Optional[str]]:
    """Devuelve (ok, origen). Orden: secrets → env → SQLite → legado."""
    u = str(usuario or "").strip().lower()
    # 1) secrets.toml
    try:
        sec = st.secrets["coordinacion"]
        if u == str(sec["usuario"]).strip().lower() and str(clave) == str(sec["clave"]):
            return True, "secrets"
    except Exception:
        pass
    # 2) Variables de entorno
    eu, ec = os.environ.get("COORDINACION_USUARIO", ""), os.environ.get("COORDINACION_CLAVE", "")
    if eu and u == eu.strip().lower() and str(clave) == ec:
        return True, "env"
    # 3) SQLite (config_coordinador)
    creds = obtener_credenciales_coordinador()
    if creds and u == str(creds[0]).strip().lower() and _verificar_clave(clave, creds[1]):
        return True, "sqlite"
    # 4) Legado embebido (solo si está definido)
    if COORD_LEGACY_USUARIO and u == COORD_LEGACY_USUARIO and str(clave) == COORD_LEGACY_CLAVE:
        return True, "legacy"
    return False, None


# ═══════════════════════════════════════════════════════════════════════════
# AUTENTICACIÓN DE DOCENTES
# ═══════════════════════════════════════════════════════════════════════════
def autenticar_docente(usuario: str, clave: str) -> Optional[str]:
    """Devuelve el nombre del docente o None. Audita éxitos y fallos."""
    u = str(usuario or "").strip().lower()
    conn = _conn()
    cur = conn.cursor()
    cur.execute("SELECT docente, password_hash FROM usuarios_docentes WHERE usuario=?", (u,))
    row = cur.fetchone()
    conn.close()
    if not row or not row[1]:
        registrar_evento("Login fallido docente", u, "usuario inexistente o sin contraseña")
        return None
    if _verificar_clave(clave, row[1]):
        registrar_evento("Login docente", row[0], "autenticación correcta")
        return row[0]
    registrar_evento("Login fallido docente", u, "contraseña incorrecta")
    return None


# ═══════════════════════════════════════════════════════════════════════════
# LECTURAS
# ═══════════════════════════════════════════════════════════════════════════
def existe_docente(docente: str) -> bool:
    conn = _conn()
    cur = conn.cursor()
    cur.execute("SELECT 1 FROM usuarios_docentes WHERE docente=?", (docente,))
    ok = cur.fetchone() is not None
    conn.close()
    return ok


def listar_docentes() -> List[str]:
    conn = _conn()
    cur = conn.cursor()
    cur.execute("SELECT docente FROM usuarios_docentes ORDER BY docente")
    rows = [r[0] for r in cur.fetchall()]
    conn.close()
    return rows


def listar_asignaciones() -> List[Dict[str, str]]:
    conn = _conn()
    cur = conn.cursor()
    cur.execute("SELECT docente, modulo, seccion FROM modulos_docentes ORDER BY docente, modulo")
    rows = [{"docente": r[0], "modulo": r[1], "seccion": r[2]} for r in cur.fetchall()]
    conn.close()
    return rows


def obtener_modulos_usuario(docente: str) -> List[list]:
    conn = _conn()
    cur = conn.cursor()
    cur.execute("SELECT id, modulo, seccion FROM modulos_docentes WHERE docente=? ORDER BY modulo",
                (docente,))
    rows = [list(r) for r in cur.fetchall()]
    conn.close()
    return rows


def obtener_usuarios_resumen() -> List[list]:
    conn = _conn()
    cur = conn.cursor()
    cur.execute("""
        SELECT u.docente, u.usuario,
               CASE WHEN u.password_hash IS NULL OR u.password_hash = ''
                    THEN 'Sin configurar' ELSE 'Configurada' END,
               (SELECT COUNT(*) FROM modulos_docentes m WHERE m.docente = u.docente),
               (SELECT GROUP_CONCAT(m.modulo || ' | ' || m.seccion, '; ')
                FROM modulos_docentes m WHERE m.docente = u.docente)
        FROM usuarios_docentes u ORDER BY u.docente
    """)
    rows = [[r[0], r[1], r[2], r[3], r[4] or ""] for r in cur.fetchall()]
    conn.close()
    return rows


def obtener_directorio() -> List[Dict[str, str]]:
    """Directorio completo: una fila por asignación; docentes sin módulos
    aparecen una vez con modulo/sección vacíos. Incluye área técnica y horas."""
    asegurar_esquema()  # idempotente: garantiza directorio_extras
    conn = _conn()
    cur = conn.cursor()
    cur.execute("""
        SELECT u.docente, u.usuario,
               COALESCE(m.modulo, ''), COALESCE(m.seccion, ''),
               (SELECT COUNT(*) FROM modulos_docentes m2 WHERE m2.docente = u.docente),
               COALESCE(e.area_tecnica, ''), COALESCE(e.horas_modulo, '')
        FROM usuarios_docentes u
        LEFT JOIN modulos_docentes m ON m.docente = u.docente
        LEFT JOIN directorio_extras e ON e.docente = u.docente
        ORDER BY u.docente, m.modulo
    """)
    rows = [
        {
            "docente": r[0], "usuario": r[1], "modulo": r[2], "seccion": r[3],
            "total_modulos": r[4], "area_tecnica": r[5], "horas_modulo": r[6],
        }
        for r in cur.fetchall()
    ]
    conn.close()
    return rows


# ═══════════════════════════════════════════════════════════════════════════
# ESCRITURAS: USUARIOS
# ═══════════════════════════════════════════════════════════════════════════
def crear_usuario(nombre: str, usuario: Optional[str] = None) -> str:
    """Crea un usuario único SIN contraseña (queda 'Sin configurar')."""
    nombre = str(nombre or "").strip()
    if not nombre:
        raise ValueError("El nombre del docente es obligatorio.")
    conn = _conn()
    cur = conn.cursor()
    cur.execute("SELECT 1 FROM usuarios_docentes WHERE docente=?", (nombre,))
    if cur.fetchone():
        conn.close()
        raise ValueError("Ese docente ya existe.")
    u = _normalizar_usuario(usuario) if usuario else _usuario_base(nombre)
    cur.execute("SELECT docente FROM usuarios_docentes WHERE usuario=?", (u,))
    if cur.fetchone():
        u = f"{u}{secrets.token_hex(2)}"
    cur.execute(
        "INSERT INTO usuarios_docentes (docente, usuario, password_hash, creado_en) "
        "VALUES (?,?,NULL,?)",
        (nombre, u, _ahora()),
    )
    conn.commit()
    conn.close()
    return u


def eliminar_usuario(docente: str) -> None:
    conn = _conn()
    cur = conn.cursor()
    cur.execute("DELETE FROM usuarios_docentes WHERE docente=?", (docente,))
    cur.execute("DELETE FROM modulos_docentes WHERE docente=?", (docente,))
    cur.execute("DELETE FROM directorio_extras WHERE docente=?", (docente,))
    conn.commit()
    conn.close()


def actualizar_usuario(docente: str, nuevo_usuario: str) -> str:
    u = _normalizar_usuario(nuevo_usuario)
    conn = _conn()
    cur = conn.cursor()
    cur.execute("SELECT docente FROM usuarios_docentes WHERE usuario=?", (u,))
    row = cur.fetchone()
    if row and row[0] != docente:
        conn.close()
        raise ValueError("Ese usuario ya está en uso por otro docente.")
    cur.execute("UPDATE usuarios_docentes SET usuario=? WHERE docente=?", (u, docente))
    conn.commit()
    conn.close()
    return u


def restablecer_password(docente: str, nueva_clave: str) -> None:
    if not existe_docente(docente):
        raise ValueError("El docente no existe.")
    if len(str(nueva_clave)) < 6:
        raise ValueError("La contraseña debe tener al menos 6 caracteres.")
    conn = _conn()
    conn.execute("UPDATE usuarios_docentes SET password_hash=? WHERE docente=?",
                 (_hash_clave(nueva_clave), docente))
    conn.commit()
    conn.close()


# ═══════════════════════════════════════════════════════════════════════════
# ESCRITURAS: MÓDULOS
# ═══════════════════════════════════════════════════════════════════════════
def asignar_modulo(docente: str, modulo: str, seccion: str) -> None:
    modulo, seccion = str(modulo or "").strip(), str(seccion or "").strip()
    if not modulo or not seccion:
        raise ValueError("Módulo y sección son obligatorios.")
    conn = _conn()
    cur = conn.cursor()
    cur.execute("SELECT 1 FROM modulos_docentes WHERE docente=? AND modulo=? AND seccion=?",
                (docente, modulo, seccion))
    if cur.fetchone():
        conn.close()
        raise ValueError("Ese módulo ya está asignado a este docente.")
    cur.execute("INSERT INTO modulos_docentes (docente, modulo, seccion) VALUES (?,?,?)",
                (docente, modulo, seccion))
    conn.commit()
    conn.close()


def quitar_modulo(id_modulo: int) -> None:
    conn = _conn()
    conn.execute("DELETE FROM modulos_docentes WHERE id=?", (id_modulo,))
    conn.commit()
    conn.close()


# ═══════════════════════════════════════════════════════════════════════════
# OPERACIONES MASIVAS (importación desde Excel) — NIVEL DE MÓDULO
# ═══════════════════════════════════════════════════════════════════════════
def eliminar_todos_usuarios() -> None:
    """⚠️ Elimina TODOS los usuarios, módulos y extras.
    Se usa antes de importaciones masivas desde Excel para partir de cero."""
    conn = _conn()
    cur = conn.cursor()
    cur.execute("DELETE FROM modulos_docentes")
    cur.execute("DELETE FROM usuarios_docentes")
    cur.execute("DELETE FROM directorio_extras")
    conn.commit()
    conn.close()
    registrar_evento("Eliminar todos los usuarios", "coordinación",
                     "Vaciado previo a importación masiva desde Excel")


def importar_usuario_con_modulos(docente: str, usuario: Optional[str] = None,
                                 password_inicial: Optional[str] = None,
                                 modulos: Optional[List[Tuple[str, str]]] = None) -> str:
    """Crea (o reutiliza) un usuario y le asigna varios módulos de una vez.
    `modulos` es una lista de parejas: [(modulo, seccion), ...].
    Sin contraseña inicial → queda 'Sin configurar' (sin cuentas demo)."""
    docente = str(docente or "").strip()
    if not docente:
        raise ValueError("El nombre del docente es obligatorio.")
    conn = _conn()
    cur = conn.cursor()
    cur.execute("SELECT 1 FROM usuarios_docentes WHERE docente=?", (docente,))
    if not cur.fetchone():
        u = _normalizar_usuario(usuario) if usuario else _usuario_base(docente)
        cur.execute("SELECT docente FROM usuarios_docentes WHERE usuario=?", (u,))
        if cur.fetchone():
            u = f"{u}{secrets.token_hex(2)}"
        hash_inicial = _hash_clave(password_inicial) if password_inicial else None
        cur.execute(
            "INSERT INTO usuarios_docentes (docente, usuario, password_hash, creado_en) "
            "VALUES (?,?,?,?)",
            (docente, u, hash_inicial, _ahora()),
        )
    conn.commit()
    conn.close()
    for par in (modulos or []):
        try:
            asignar_modulo(docente, par[0], par[1])
        except ValueError:
            pass  # ya asignado: se ignora en importaciones masivas
    return docente