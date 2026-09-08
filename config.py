"""
Configuración del proyecto.
=====================================================================
Un solo lugar para cambiar la fuente de datos y los textos de la
aplicación, sin tocar el resto del código.
=====================================================================
"""

# --------------------------------------------------------------------------- #
# Identidad del proyecto
# --------------------------------------------------------------------------- #
TITULO = "La transición energética de Chile en datos"
SUBTITULO = "Análisis y presentación de datos utilizando APIs REST en Python"


# --------------------------------------------------------------------------- #
# Fuente de datos: API REST pública del Gobierno de Chile
# --------------------------------------------------------------------------- #
API_BASE = "https://datos.gob.cl/api/3/action"

# Recurso: "Generación Bruta Mensual SEN" (Comisión Nacional de Energía).
# Para usar otro conjunto de datos, reemplaza este identificador por el de
# cualquier recurso de datos.gob.cl publicado en el datastore de CKAN.
RESOURCE_ID = "389a1943-9c3d-4957-982a-58e3fb0c1bdb"

DATASET_TITULO = "Generación Bruta Mensual del Sistema Eléctrico Nacional (SEN)"
DATASET_ORGANISMO = "Comisión Nacional de Energía (CNE) - Ministerio de Energía"
DATASET_URL = "https://datos.gob.cl/dataset/generacion-bruta"
DATASET_LICENCIA = "Datos abiertos - Gobierno de Chile"

# Parámetros de las peticiones HTTP
PAGE_SIZE = 10_000   # máximo de filas que acepta la API por respuesta
TIMEOUT = 60         # segundos de espera por petición
