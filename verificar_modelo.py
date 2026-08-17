import sqlite3

conn = sqlite3.connect("gestion_etp.db")
cur = conn.cursor()

n_usu = cur.execute("SELECT COUNT(*) FROM usuarios_docentes").fetchone()[0]
n_mod = cur.execute("SELECT COUNT(*) FROM modulos_docentes").fetchone()[0]
n_leg = cur.execute("SELECT COUNT(*) FROM docentes").fetchone()[0]
n_leg_uni = cur.execute("SELECT COUNT(DISTINCT TRIM(docente)) FROM docentes").fetchone()[0]

print(f"👤 Usuarios únicos (usuarios_docentes): {n_usu}")
print(f"📚 Asignaciones (modulos_docentes):     {n_mod}")
print(f"🗃️  Filas legacy 'docentes':            {n_leg}  ← asignaciones, NO usuarios")
print(f"👥 Docentes únicos en legacy:           {n_leg_uni}")

print("\nDocentes con más módulos asignados:")
for docente, total in cur.execute("""
    SELECT u.docente, COUNT(m.id)
    FROM usuarios_docentes u
    LEFT JOIN modulos_docentes m ON m.usuario_id = u.id
    GROUP BY u.id ORDER BY total DESC LIMIT 5
"""):
    print(f"  • {docente}: {total} módulos")
conn.close()