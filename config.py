"""
Configuración del proyecto.
=====================================================================
Un solo lugar para cambiar la fuente de datos o los metadatos del
informe, sin tocar el resto del código.
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


# --------------------------------------------------------------------------- #
# Metadatos del informe en Word
# --------------------------------------------------------------------------- #
# Sólo los usa `generar_informe.py` al construir el informe y el póster;
# la aplicación web no los muestra.

INFORME_AUTORES = [
    "Nombre Apellido 1",
    "Nombre Apellido 2",
    "Nombre Apellido 3",
]

INFORME_CONTEXTO = "FITO9017 - Programación en Python"
INFORME_REFERENCIA = "Unidad 3 - Semana 13"

# Nombre base de los archivos generados, sin extensión.
# Cámbialo si necesitas que los archivos sigan una convención concreta.
NOMBRE_ARCHIVO_INFORME = "informe-matriz-electrica"
