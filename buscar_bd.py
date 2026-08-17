import os, sqlite3, shutil
from pathlib import Path

RAICES = [
    r"C:\Users\User\OneDrive\Escritorio\Carpeta Operativa coordinacion Tecnica",
    r"C:\Users\User\OneDrive\Escritorio\Curriculum Bernardo",
]
DESTINO = Path(r"C:\Users\User\OneDrive\Escritorio\Carpeta Operativa coordinacion Tecnica\Bernardo Programacion\Proyecto Planificacion\gestion_etp.db")

print("Buscando copias de gestion_etp.db (omite carpetas inaccesibles)...\n")
resultados = []
for raiz in RAICES:
    # onerror=lambda: si OneDrive tiene carpetas rotas/en la nube, las salta y continúa
    for dirpath, dirnames, filenames in os.walk(raiz, onerror=lambda err: None):
        if "gestion_etp.db" not in filenames:
            continue
        ruta = Path(dirpath) / "gestion_etp.db"
        try:
            kb = ruta.stat().st_size // 1024
        except Exception:
            kb = 0
        det = {}
        try:
            con = sqlite3.connect(str(ruta))
            tablas = {r[0] for r in con.execute("SELECT name FROM sqlite_master WHERE type='table'")}
            for t in ("docentes", "usuarios_docentes", "calificaciones", "cronograma", "incidencias", "alertas", "acuerdos"):
                if t in tablas:
                    det[t] = con.execute("SELECT COUNT(*) FROM " + t).fetchone()[0]
            con.close()
        except Exception as e:
            det = {"error": str(e)}
        resultados.append((ruta, kb, det))
        print(f"{ruta}\n   {kb} KB | {det}\n")

def puntaje(item):
    d = item[2]
    return d.get("docentes", 0) * 10 + d.get("calificaciones", 0) + d.get("cronograma", 0) + d.get("incidencias", 0)

candidatas = [r for r in resultados if puntaje(r) > 0 and r[0].resolve() != DESTINO.resolve()]
if not candidatas:
    print("❌ No se encontró ninguna BD con datos en esas carpetas.")
else:
    mejor = max(candidatas, key=lambda i: (puntaje(i), i[1]))
    print("=" * 70)
    print("BD RECOMENDADA (la que tiene tus datos):")
    print(mejor[0])
    resp = input("\n¿Copiarla sobre la BD vacía de Proyecto Planificacion? (s/n): ").strip().lower()
    if resp == "s":
        shutil.copy2(mejor[0], DESTINO)
        print("✅ Copiada con éxito. Los originales quedan intactos como respaldo.")
        print("▶ Ahora ejecuta:  python prueba_auth.py")
    else:
        print("No se copió nada.")