"""
Configuración central del proyecto.
=====================================================================
>>> ÚNICO ARCHIVO QUE DEBES EDITAR ANTES DE ENTREGAR <<<
Cambia los nombres de los integrantes y listo.
=====================================================================
"""

# --- Datos de la entrega (aparecen en la app, el informe y el póster) ---
ASIGNATURA = "FITO9017 - Programación en Python"
ACTIVIDAD = "Solemne II - Proyecto Final DataViz Python"
UNIDAD_SEMANA = "Unidad 3 - Semana 13"
DOCENTE = "Docente instructor"

INTEGRANTES = [
    "Nombre Apellido 1",
    "Nombre Apellido 2",
    "Nombre Apellido 3",
]

# Se usa para nombrar el PDF: APELLIDO_NOMBRE_U3_SolemneII.pdf
NOMBRE_ARCHIVO_ENTREGA = "APELLIDO_NOMBRE_U3_SolemneII"


# --- Fuente de datos (API REST pública del Gobierno de Chile) ---
API_BASE = "https://datos.gob.cl/api/3/action"

# Recurso: "Generación Bruta Mensual SEN" (Comisión Nacional de Energía)
RESOURCE_ID = "389a1943-9c3d-4957-982a-58e3fb0c1bdb"

DATASET_TITULO = "Generación Bruta Mensual del Sistema Eléctrico Nacional (SEN)"
DATASET_ORGANISMO = "Comisión Nacional de Energía (CNE) - Ministerio de Energía"
DATASET_URL = "https://datos.gob.cl/dataset/generacion-bruta"
DATASET_LICENCIA = "Datos abiertos - Gobierno de Chile"

# Tamaño de página para la paginación de la API (máximo aceptado: 10.000)
PAGE_SIZE = 10_000
TIMEOUT = 60  # segundos por petición HTTP
