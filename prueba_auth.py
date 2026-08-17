import logging
logging.basicConfig(level=logging.INFO)

import core.auth as auth

info = auth.diagnosticar()
for clave, valor in info.items():
    print(f"{clave}: {valor}")

auth.asegurar_esquema()
print("Docentes tras migración:", auth.listar_docentes())
print("Asignaciones:", auth.listar_asignaciones()[:10])