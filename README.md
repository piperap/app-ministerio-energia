# La matriz eléctrica de Chile en datos

Aplicación web que muestra cómo ha cambiado la generación eléctrica en Chile
entre 2014 y hoy. Los datos se obtienen en vivo desde la **API pública del
portal de datos abiertos del Gobierno de Chile** (datos.gob.cl) y se presentan
con filtros, indicadores y gráficos interactivos.

![La aplicación en funcionamiento](docs/captura-resumen.png)

---

## ▶️ Cómo ejecutarla en tu computador

### Lo que necesitas

- **Python 3.9 o superior** — para comprobar si lo tienes, abre una terminal y
  escribe `python3 --version`. Si no aparece una versión, descárgalo desde
  <https://www.python.org/downloads/>
- Conexión a Internet la primera vez (después funciona sin conexión).

### Tres comandos

Abre una terminal en la carpeta del proyecto y ejecuta:

```bash
pip install -r requirements.txt
streamlit run app.py
```

La aplicación se abre sola en tu navegador, en `http://localhost:8501`.
Para cerrarla, vuelve a la terminal y presiona `Ctrl + C`.

> **En macOS puedes evitar la terminal:** haz doble clic en el archivo
> `ejecutar.command`. Instala lo necesario y abre la aplicación por ti.

### ¿Se demora la primera vez?

Sí, entre 30 y 60 segundos: está descargando los 91.538 registros desde la
API. Queda guardado en `cache_datos.json.gz` y las siguientes veces abre al
instante.

---

## 🖥️ Cómo usarla

### Filtros (barra lateral izquierda)

| Filtro | Para qué sirve |
|---|---|
| **Período** | Elige el rango de años a analizar |
| **Tipo de energía** | Todas · ERNC (renovables no convencionales) · Convencional |
| **Tecnologías** | Déjalo vacío para ver todas, o elige las que te interesen |

Cada cambio recalcula al instante los indicadores, los gráficos y las tablas.

### Indicadores (arriba)

Generación total, participación de ERNC, crecimiento del período y cantidad de
centrales generadoras.

### Pestañas

| Pestaña | Qué encuentras |
|---|---|
| **Resumen** | Cómo se recompone la matriz año a año, con los hallazgos destacados |
| **Tecnologías** | Ranking por tecnología y evolución del mix energético |
| **Datos** | Los registros en tabla, descarga en CSV y el origen de los datos |

### Botones útiles

- **Actualizar datos desde la API** (barra lateral): descarga los datos más
  recientes desde datos.gob.cl.
- **Descargar la selección en CSV** (pestaña Datos): exporta lo que estés
  viendo. El archivo usa punto y coma y coma decimal, así que Excel en español
  lo abre correctamente.

---

## ⚙️ Configuración

Todo lo configurable está en **`config.py`**, en la parte de arriba:

```python
INTEGRANTES = [
    "Nombre Apellido 1",
    "Nombre Apellido 2",
    "Nombre Apellido 3",
]
```

Esos nombres aparecen en la barra lateral y al pie de la aplicación.

En el mismo archivo puedes cambiar el conjunto de datos: basta reemplazar
`RESOURCE_ID` por el identificador de otro recurso de datos.gob.cl que esté
publicado en el *datastore*. Si el nuevo recurso tiene otras columnas, habrá
que ajustar `api_datos.limpiar()`.

---

## 📡 De dónde vienen los datos

- **Recurso:** Generación Bruta Mensual del Sistema Eléctrico Nacional (SEN)
- **Organismo:** Comisión Nacional de Energía (CNE) — Ministerio de Energía
- **Portal:** <https://datos.gob.cl/dataset/generacion-bruta>
- **Licencia:** datos abiertos del Gobierno de Chile
- **Volumen:** 91.538 registros · 2014 a la fecha · 1.532 centrales ·
  13 tecnologías

La obtención es enteramente por API REST, sin descargar archivos a mano:

```
GET https://datos.gob.cl/api/3/action/datastore_search
    ?resource_id=389a1943-9c3d-4957-982a-58e3fb0c1bdb
    &limit=10000&offset=0
```

La API entrega como máximo 10.000 filas por respuesta, así que el programa
recorre la paginación con el parámetro `offset` hasta completar el total que
informa el servidor.

Para probar solamente la conexión con la API, sin abrir la aplicación:

```bash
python api_datos.py
```

---

## 🗂️ Estructura del proyecto

```
app.py                ← la aplicación web (Streamlit)
api_datos.py          ← consulta la API, pagina y limpia los datos
analisis.py           ← cálculos (pandas) y gráficos (matplotlib)
config.py             ← configuración editable
generar_entrega.py    ← genera un informe y un póster en Word
requirements.txt      ← librerías necesarias
ejecutar.command      ← doble clic para abrir la app (macOS)
.streamlit/           ← tema visual de la aplicación
cache_datos.json.gz   ← copia local de los datos (se crea automáticamente)
docs/                 ← capturas de pantalla
assets/               ← gráficos exportados en PNG
```

El código está separado en tres capas —obtención de datos, análisis y
presentación— para que la aplicación y los documentos generados usen
exactamente los mismos cálculos.

---

## 📄 Generar un informe y un póster

Además de la aplicación, el proyecto puede producir dos documentos de Word
editables con el análisis completo y los seis gráficos:

```bash
python generar_entrega.py
```

Crea `<nombre>.docx` (informe) y `<nombre>_Poster.docx` (póster de una página
horizontal), donde `<nombre>` es el valor de `NOMBRE_ARCHIVO_ENTREGA` en
`config.py`. Para obtener PDF, ábrelos en Word y usa
**Archivo → Guardar como… → PDF**.

---

## ☁️ Publicarla en línea (opcional)

Para tener una URL pública y gratuita, sin instalar nada:

1. Entra a <https://share.streamlit.io> e inicia sesión con GitHub.
2. **Create app → Deploy a public app from GitHub**.
3. Completa:
   - **Repository:** `piperap/app-ministerio-energia`
   - **Branch:** `main`
   - **Main file path:** `app.py`
4. **Deploy**.

---

## ❓ Problemas frecuentes

| Situación | Solución |
|---|---|
| `ModuleNotFoundError` | Faltan las librerías: `pip install -r requirements.txt` |
| `streamlit: command not found` | Usa `python -m streamlit run app.py` |
| `pip: command not found` | Usa `python3 -m pip install -r requirements.txt` |
| La app tarda al abrir por primera vez | Normal: está descargando los datos desde la API |
| Quiero los datos más recientes | Botón «Actualizar datos desde la API», o borra `cache_datos.json.gz` |
| No tengo Internet | Funciona igual, mientras exista `cache_datos.json.gz` |
| El puerto 8501 está ocupado | `streamlit run app.py --server.port 8600` |
| Los filtros no muestran nada | Esa combinación no tiene registros; amplía el período o cambia el tipo de energía |

---

## 🔒 Notas de seguridad

- La aplicación **no usa claves ni contraseñas**: la API de datos.gob.cl es
  pública y de solo lectura.
- Los datos son **datos abiertos de generación eléctrica** y no contienen
  información personal.
- Este repositorio es **público**: cualquier nombre que escribas en
  `config.py` queda visible en Internet.

---

## 🧰 Construido con

`streamlit` · `requests` · `pandas` · `matplotlib` · `python-docx`
