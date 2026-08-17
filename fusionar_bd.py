import sqlite3
import core.auth as auth

ORIGEN = r"C:\Users\User\OneDrive\Escritorio\Carpeta Operativa coordinacion Tecnica\Bernardo Programacion\Gestion-Modulos-Formativos-PSAC-main\gestion_etp.db"

con_o = sqlite3.connect(ORIGEN)
con_d = sqlite3.connect(auth.DB_PATH)

cols = {r[1] for r in con_o.execute("PRAGMA table_info(usuarios_docentes)")}
col_pass = "password" if "password" in cols else None
sel = "SELECT id, docente, usuario" + (f", {col_pass}" if col_pass else ", NULL") + " FROM usuarios_docentes"

nuevos = adoptados = modulos = 0
for uid_o, docente, usuario, plano in con_o.execute(sel):
    fila = con_d.execute("SELECT id FROM usuarios_docentes WHERE docente = ?", (docente,)).fetchone()
    if fila:
        uid_d = fila[0]
        ocupado = con_d.execute(
            "SELECT 1 FROM usuarios_docentes WHERE usuario = ? AND id <> ?", (usuario, uid_d)
        ).fetchone()
        if not ocupado:
            con_d.execute("UPDATE usuarios_docentes SET usuario = ? WHERE id = ?", (usuario, uid_d))
            adoptados += 1
    else:
        con_d.execute(
            "INSERT INTO usuarios_docentes (docente, usuario, password_hash, creado_el) "
            "VALUES (?, ?, ?, datetime('now'))",
            (docente, usuario, auth.hash_password(str(plano or "1234"))),
        )
        uid_d = con_d.lastrowid
        nuevos += 1
    for modulo, seccion in con_o.execute(
        "SELECT modulo, seccion FROM modulos_docentes WHERE usuario_id = ?", (uid_o,)
    ):
        con_d.execute(
            "INSERT OR IGNORE INTO modulos_docentes (usuario_id, modulo, seccion) VALUES (?, ?, ?)",
            (uid_d, modulo, seccion),
        )
        modulos += 1

con_d.commit()
con_o.close()
con_d.close()
print(f"✅ Fusión completada: {nuevos} usuarios nuevos, {adoptados} usuarios adoptados, {modulos} módulos sincronizados.")