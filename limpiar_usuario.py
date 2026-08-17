# limpiar_usuario.py
import sqlite3

conn = sqlite3.connect("gestion_etp.db")
cur = conn.cursor()

for tabla in ["usuarios_docentes", "modulos_docentes", "docentes"]:
    try:
        cur.execute(f"DELETE FROM {tabla} WHERE lower(docente) LIKE '%adarys%'")
        print(f"{tabla}: {cur.rowcount} fila(s) borrada(s)")
    except Exception as e:
        print(f"{tabla}: error -> {e}")

conn.commit()
conn.close()
print("✅ Listo. Recarga la app.")