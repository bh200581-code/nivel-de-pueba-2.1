# 🏫 Sistema de Gestión Docente ETP - Fígital & IA

![Estado](https://img.shields.io/badge/Estado-En_Desarrollo-success)
![Versión](https://img.shields.io/badge/Versión-2026.08_Azul_Metálico-blue)
![Python](https://img.shields.io/badge/Python-3.11%2B-yellow)
![Streamlit](https://img.shields.io/badge/Streamlit-Framework-FF4B4B)

Plataforma integral diseñada para la coordinación académico-formativa y la gestión docente en la **Educación Técnico Profesional (ETP)**, estrictamente alineada al diseño curricular del MINERD. 

El sistema combina gestión administrativa (Fígital) con **Inteligencia Artificial generativa** para automatizar la creación de recursos didácticos, auditorías y evaluaciones, optimizando el tiempo del docente y del equipo de coordinación en el Politécnico Salesiano Arquides Calderón.

---

## ✨ Características Principales

### 👔 Módulo de Coordinación Pedagógica
* **Sala de Situación:** Panel gerencial con KPIs en tiempo real (alertas rojas, acuerdos, incidencias).
* **Auditoría de Planificaciones y Calificaciones:** Monitoreo y retroalimentación automatizada de los planes de unidad y diario.
* **Acompañamiento Docente Fígital:** Generación de rúbricas con IA, impresión, escaneo y corrección automática mediante visión artificial.
* **Gestión de Incidencias y Acuerdos:** Registro disciplinario y generación de informes ejecutivos.

### 🧑‍🏫 Módulo Docente ETP (Talleres y Módulos)
* **Planificación Modular y Diaria:** Generación de matrices basadas en Resultados de Aprendizaje (RA) y Criterios de Evaluación (CE).
* **Fábrica de Contenidos y Redactor Profundo:** Generación de material didáctico, analogías técnicas y libros anclados a PDFs curriculares.
* **Banco de Ítems PRO:** Creación asistida por IA de 12 tipos de ítems evaluativos basados en la taxonomía de Bloom.
* **Portal de Calificaciones:** Registro de notas por RA (Ev, R1, R2, R3) con cálculo automático de estatus de recuperación.

### 🤝 Herramientas Interactivas y Comunes
* **Pruebas Diagnósticas:** Aplicación digital vía enlace (sin login) o física (escaneo y corrección con IA).
* **Fábrica Visual IA:** Generador de presentaciones en PowerPoint e Infografías HTML con imágenes generadas automáticamente.
* **Juegos Interactivos:** Trivias, ruletas y juegos de memoria en archivos HTML autónomos.

---

## 🛠️ Stack Tecnológico

* **Frontend & Backend:** [Streamlit](https://streamlit.io/) (Python)
* **Base de Datos:** SQLite (`gestion_etp.db`)
* **Inteligencia Artificial:** Soporte multi-proveedor (Google Gemini, OpenAI ChatGPT, Anthropic Claude) vía API.
* **Procesamiento Documental:** `python-docx` (Word), `openpyxl` (Excel), `pypdf` (PDF).

---

## 🚀 Instalación y Ejecución Local

1. **Clonar el repositorio:**
   ```bash
   git clone [https://github.com/tu-usuario/gestion-etp.git](https://github.com/tu-usuario/gestion-etp.git)
   cd gestion-etp