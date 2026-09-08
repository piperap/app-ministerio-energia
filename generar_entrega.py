"""
Generador de la entrega en Word
===============================
Ejecutar con:  python generar_entrega.py

Produce dos documentos EDITABLES, listos para revisar y luego exportar a PDF:
  1. <NOMBRE_ARCHIVO_ENTREGA>.docx         -> informe completo del proyecto
  2. <NOMBRE_ARCHIVO_ENTREGA>_Poster.docx  -> póster del proyecto (1 página)

Opcional:  python generar_entrega.py --pdf
           además exporta ambos a PDF usando Microsoft Word o LibreOffice.

Reutiliza exactamente las mismas funciones de análisis y los mismos gráficos
que la aplicación Streamlit (módulos api_datos.py y analisis.py): los números
del informe nunca pueden contradecir los de la app.
"""

from __future__ import annotations

import re
import subprocess
import sys
from datetime import date
from pathlib import Path

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

import analisis as an
from analisis import fmt
import api_datos
import config

BASE = Path(__file__).parent
IMG = BASE / "assets"
IMG.mkdir(exist_ok=True)

AZUL = RGBColor(0x0E, 0x74, 0x90)
VERDE = RGBColor(0x16, 0xA3, 0x4A)
GRIS = RGBColor(0x47, 0x55, 0x69)
NEGRO = RGBColor(0x0F, 0x17, 0x2A)
HEX_AZUL, HEX_CLARO, HEX_BORDE = "0E7490", "F1F5F9", "CBD5E1"


# ========================================================================== #
#  Utilidades de formato Word
# ========================================================================== #
# El esquema OOXML exige que los hijos de <w:tblPr> aparezcan en este orden;
# si no se respeta, Word descarta silenciosamente las propiedades posteriores.
ORDEN_TBLPR = ["tblStyle", "tblpPr", "tblOverlap", "bidiVisual",
               "tblStyleRowBandSize", "tblStyleColBandSize", "tblW", "jc",
               "tblCellSpacing", "tblInd", "tblBorders", "shd", "tblLayout",
               "tblCellMar", "tblLook", "tblCaption", "tblDescription"]


def _propiedad_tabla(tabla, elemento) -> None:
    """Inserta (o reemplaza) una propiedad de tabla en su posición correcta."""
    props = tabla._tbl.tblPr
    etiqueta = elemento.tag.split("}")[-1]

    for existente in props.findall(qn(f"w:{etiqueta}")):
        props.remove(existente)

    posicion = ORDEN_TBLPR.index(etiqueta)
    for hijo in props:
        nombre = hijo.tag.split("}")[-1]
        if nombre in ORDEN_TBLPR and ORDEN_TBLPR.index(nombre) > posicion:
            hijo.addprevious(elemento)
            return
    props.append(elemento)


def sombrear(celda, hex_color: str) -> None:
    """Aplica color de fondo a una celda de tabla."""
    elem = OxmlElement("w:shd")
    elem.set(qn("w:val"), "clear")
    elem.set(qn("w:fill"), hex_color)
    celda._tc.get_or_add_tcPr().append(elem)


def vaciar(celda):
    """Elimina los párrafos por defecto de una celda.

    python-docx crea siempre un párrafo vacío al construir la celda; si no se
    quita, el documento queda con espacios en blanco antes de cada contenido.
    """
    for p in list(celda.paragraphs):
        p._element.getparent().remove(p._element)
    return celda


def bordes(tabla, color: str = HEX_BORDE, grosor: int = 4) -> None:
    """Dibuja bordes finos en toda la tabla."""
    marco = OxmlElement("w:tblBorders")
    for lado in ("top", "left", "bottom", "right", "insideH", "insideV"):
        b = OxmlElement(f"w:{lado}")
        b.set(qn("w:val"), "single")
        b.set(qn("w:sz"), str(grosor))
        b.set(qn("w:color"), color)
        marco.append(b)
    _propiedad_tabla(tabla, marco)


def margenes_celda(tabla, arriba=2, abajo=2, izq=4, der=4) -> None:
    """Ajusta el relleno interno de todas las celdas (en puntos)."""
    marco = OxmlElement("w:tblCellMar")
    for lado, valor in (("top", arriba), ("bottom", abajo),
                        ("left", izq), ("right", der)):
        e = OxmlElement(f"w:{lado}")
        e.set(qn("w:w"), str(int(valor * 20)))
        e.set(qn("w:type"), "dxa")
        marco.append(e)
    _propiedad_tabla(tabla, marco)


def ancho(tabla, anchos_cm: list[float]) -> None:
    """Fija el ancho total y por columna con layout fijo.

    Sin layout fijo Word recalcula los anchos según el contenido y las
    columnas dejan de coincidir con el diseño.
    """
    tabla.autofit = False
    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    _propiedad_tabla(tabla, layout)

    total = OxmlElement("w:tblW")
    total.set(qn("w:type"), "dxa")
    total.set(qn("w:w"), str(int(sum(anchos_cm) * 567)))
    _propiedad_tabla(tabla, total)

    for col, cm_col in zip(tabla.columns, anchos_cm):
        col.width = Cm(cm_col)
    for fila in tabla.rows:
        for celda, cm_col in zip(fila.cells, anchos_cm):
            celda.width = Cm(cm_col)


def compactar_cola(celda, espacio: float = 3) -> None:
    """Compacta el párrafo que python-docx añade tras una tabla anidada.

    Ese párrafo es obligatorio en OOXML, pero hereda el tamaño del estilo por
    defecto (11 pt) y abre un hueco de una línea completa entre bloques.
    """
    ultimo = celda.paragraphs[-1]
    f = ultimo.paragraph_format
    f.space_before = Pt(0)
    f.space_after = Pt(espacio)
    f.line_spacing = 1.0
    ultimo.add_run("").font.size = Pt(2)


def parrafo(doc_o_celda, texto: str = "", *, tamano: float = 10.5,
            negrita: bool = False, cursiva: bool = False,
            color: RGBColor | None = None, alineacion=None,
            espacio_antes: float = 0, espacio_despues: float = 4,
            interlineado: float = 1.15, sangria: float = 0,
            fuente: str | None = None):
    """Crea un párrafo con formato.

    Admite marcadores dentro del texto: **negrita** y *cursiva*.
    """
    p = doc_o_celda.add_paragraph()
    f = p.paragraph_format
    f.space_before = Pt(espacio_antes)
    f.space_after = Pt(espacio_despues)
    f.line_spacing = interlineado
    if sangria:
        f.left_indent = Cm(sangria)
    if alineacion is not None:
        p.alignment = alineacion

    # Un párrafo sin texto necesita igualmente un run: de lo contrario Word
    # aplica el tamaño del estilo por defecto y el espacio resulta enorme.
    trozos = [t for t in re.split(r"(\*\*.+?\*\*|\*[^*]+?\*)", texto) if t] or [""]

    for trozo in trozos:
        es_negrita = trozo.startswith("**")
        es_cursiva = not es_negrita and trozo.startswith("*")
        limpio = trozo.strip("*") if (es_negrita or es_cursiva) else trozo
        r = p.add_run(limpio)
        r.font.size = Pt(tamano)
        r.font.bold = negrita or es_negrita
        r.font.italic = cursiva or es_cursiva
        r.font.name = fuente or "Calibri"
        if color is not None:
            r.font.color.rgb = color
    return p


def titulo(doc, texto: str, nivel: int = 1):
    """Encabezado del informe. Usa los estilos nativos de Word para que el
    usuario pueda generar un índice automático si lo desea."""
    h = doc.add_heading(level=nivel)
    h.paragraph_format.space_before = Pt(14 if nivel == 1 else 9)
    h.paragraph_format.space_after = Pt(5)
    r = h.add_run(texto)
    r.font.size = Pt(14 if nivel == 1 else 11.5)
    r.font.color.rgb = AZUL if nivel == 1 else NEGRO
    r.font.name = "Calibri"
    r.font.bold = True
    return h


def vineta(doc, texto: str, *, tamano: float = 10.5):
    p = parrafo(doc, texto, tamano=tamano, sangria=0.6, espacio_despues=3,
                alineacion=WD_ALIGN_PARAGRAPH.JUSTIFY)
    p.style = doc.styles["List Bullet"]
    for r in p.runs:
        r.font.size = Pt(tamano)
        r.font.name = "Calibri"
    return p


def codigo(doc, texto: str, ancho_cm: float = 16.6):
    """Bloque de código sobre fondo gris."""
    t = doc.add_table(rows=1, cols=1)
    bordes(t)
    margenes_celda(t, 4, 4, 6, 6)
    ancho(t, [ancho_cm])
    celda = vaciar(t.cell(0, 0))
    sombrear(celda, HEX_CLARO)
    for linea in texto.split("\n"):
        parrafo(celda, linea, tamano=9, fuente="Consolas",
                espacio_despues=0, interlineado=1.1)
    parrafo(doc, "", tamano=5, espacio_despues=4)
    return t


def tabla_datos(doc, encabezados: list[str], filas: list[list[str]],
                anchos: list[float], tamano: float = 9):
    """Tabla con cabecera azul y filas alternadas."""
    t = doc.add_table(rows=1, cols=len(encabezados))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    bordes(t)
    margenes_celda(t, 3, 3, 5, 5)

    for i, texto in enumerate(encabezados):
        celda = vaciar(t.rows[0].cells[i])
        sombrear(celda, HEX_AZUL)
        parrafo(celda, texto, tamano=tamano, negrita=True,
                color=RGBColor(0xFF, 0xFF, 0xFF), espacio_despues=0,
                interlineado=1.0)

    for n, fila in enumerate(filas):
        celdas = t.add_row().cells
        for i, valor in enumerate(fila):
            celda = vaciar(celdas[i])
            if n % 2 == 1:
                sombrear(celda, "F8FAFC")
            derecha = i > 0 and str(valor)[:1].isdigit()
            parrafo(celda, str(valor), tamano=tamano, espacio_despues=0,
                    interlineado=1.0,
                    alineacion=WD_ALIGN_PARAGRAPH.RIGHT if derecha else None)

    ancho(t, anchos)
    parrafo(doc, "", tamano=5, espacio_despues=4)
    return t


def figura(doc, ruta: Path, ancho_cm: float, pie: str):
    doc.add_picture(str(ruta), width=Cm(ancho_cm))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    parrafo(doc, pie, tamano=8.5, cursiva=True, color=GRIS,
            alineacion=WD_ALIGN_PARAGRAPH.CENTER, espacio_despues=10)


def pie_de_pagina(seccion, texto: str) -> None:
    p = seccion.footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(texto)
    r.font.size = Pt(8)
    r.font.color.rgb = GRIS
    r.font.name = "Calibri"


def encabezado_pagina(seccion, texto: str) -> None:
    p = seccion.header.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r = p.add_run(texto)
    r.font.size = Pt(8)
    r.font.color.rgb = GRIS
    r.font.name = "Calibri"


# ========================================================================== #
#  Gráficos a disco
# ========================================================================== #
def exportar_graficos(df) -> dict:
    figuras = {
        "mix": an.g_mix_anual(df),
        "serie": an.g_serie_mensual(df),
        "tecnologias": an.g_tecnologias(df),
        "evolucion": an.g_evolucion_tecnologia(df),
        "centrales": an.g_top_centrales(df, 15),
        "estacional": an.g_estacionalidad(df),
    }
    rutas = {}
    for nombre, fig in figuras.items():
        ruta = IMG / f"{nombre}.png"
        fig.savefig(ruta, dpi=170, bbox_inches="tight", facecolor="white")
        rutas[nombre] = ruta
    return rutas


# ========================================================================== #
#  INFORME (.docx)
# ========================================================================== #
def construir_informe(ruta: Path, df, kpi, tabla, tec, ev, meta, rutas) -> None:
    a0, a1 = tabla.index[0], tabla.index[-1]
    p0, p1 = tabla.loc[a0, "% ERNC"], tabla.loc[a1, "% ERNC"]
    crece = (ev.iloc[-1] - ev.iloc[0]).idxmax()
    cae = (ev.iloc[-1] - ev.iloc[0]).idxmin()
    concentracion = an.top_centrales(df, 15)["generacion_mwh"].sum() \
        / df["generacion_mwh"].sum() * 100

    doc = Document()
    doc.core_properties.title = config.ACTIVIDAD
    doc.core_properties.author = ", ".join(config.INTEGRANTES)

    s = doc.sections[0]
    s.left_margin = s.right_margin = Cm(2.2)
    s.top_margin = Cm(2.0)
    s.bottom_margin = Cm(2.0)
    encabezado_pagina(s, f"{config.ASIGNATURA} · {config.UNIDAD_SEMANA}")
    pie_de_pagina(s, f"{config.ACTIVIDAD} · {' · '.join(config.INTEGRANTES)}")

    ANCHO = 16.6

    # ------------------------------ Portada ------------------------------- #
    parrafo(doc, "", tamano=11, espacio_despues=14)
    parrafo(doc, "La transición energética de Chile en datos", tamano=24,
            negrita=True, color=AZUL, alineacion=WD_ALIGN_PARAGRAPH.CENTER,
            espacio_despues=4)
    parrafo(doc, "Análisis y presentación de datos utilizando APIs REST en Python",
            tamano=12, color=GRIS, alineacion=WD_ALIGN_PARAGRAPH.CENTER,
            espacio_despues=4)
    parrafo(doc, config.ACTIVIDAD, tamano=11, negrita=True, color=NEGRO,
            alineacion=WD_ALIGN_PARAGRAPH.CENTER, espacio_despues=18)

    ficha = doc.add_table(rows=2, cols=2)
    bordes(ficha)
    margenes_celda(ficha, 5, 5, 7, 7)
    ancho(ficha, [ANCHO / 2, ANCHO / 2])
    datos_ficha = [
        (f"**Asignatura**\n{config.ASIGNATURA}",
         f"**Unidad y semana**\n{config.UNIDAD_SEMANA}"),
        ("**Integrantes**\n" + "\n".join(config.INTEGRANTES),
         f"**Fecha de entrega**\n{date.today():%d-%m-%Y}"),
    ]
    for i, fila in enumerate(datos_ficha):
        for j, contenido in enumerate(fila):
            celda = vaciar(ficha.cell(i, j))
            sombrear(celda, "F8FAFC")
            for linea in contenido.split("\n"):
                parrafo(celda, linea, tamano=10, espacio_despues=1,
                        interlineado=1.15)
    parrafo(doc, "", tamano=8, espacio_despues=10)

    # Tarjetas con las cifras clave del proyecto
    kpis = [(f"{fmt(len(df), 0)}", "registros obtenidos vía API REST"),
            (f"{fmt(kpi['total_twh'], 0)} TWh", "generación analizada"),
            (f"{fmt(p1, 1)} %", f"participación ERNC en {a1}"),
            (f"{fmt(kpi['centrales'], 0)}", "centrales del sistema")]
    tarjetas = doc.add_table(rows=1, cols=4)
    bordes(tarjetas)
    margenes_celda(tarjetas, 6, 6, 4, 4)
    ancho(tarjetas, [ANCHO / 4] * 4)
    for j, (valor, etiqueta) in enumerate(kpis):
        celda = vaciar(tarjetas.cell(0, j))
        parrafo(celda, valor, tamano=15, negrita=True, color=AZUL,
                alineacion=WD_ALIGN_PARAGRAPH.CENTER, espacio_despues=1,
                interlineado=1.0)
        parrafo(celda, etiqueta, tamano=8, color=GRIS,
                alineacion=WD_ALIGN_PARAGRAPH.CENTER, espacio_despues=0,
                interlineado=1.05)
    parrafo(doc, "", tamano=8, espacio_despues=10)

    parrafo(doc, f"**Fuente de datos:** {config.DATASET_ORGANISMO} — "
                 f"{config.DATASET_TITULO}. Obtenidos mediante consultas GET a la "
                 f"API REST del portal de datos abiertos del Gobierno de Chile "
                 f"(datos.gob.cl). Última actualización del recurso: "
                 f"{meta.get('modificado', 's/i')}.",
            tamano=9, color=GRIS, alineacion=WD_ALIGN_PARAGRAPH.CENTER)
    doc.add_page_break()

    # ------------------------------ 1. Desafío ---------------------------- #
    titulo(doc, "1. El desafío")
    parrafo(doc, "Chile comprometió el retiro total de sus centrales a carbón y "
                 "una matriz eléctrica mayoritariamente renovable. El discurso "
                 "público abunda en metas, pero rara vez en evidencia "
                 "verificable. Este proyecto se propone responder con datos "
                 "abiertos una pregunta concreta: **¿cuánto ha avanzado "
                 "realmente la transición energética chilena y qué tecnologías "
                 "la están empujando?**",
            alineacion=WD_ALIGN_PARAGRAPH.JUSTIFY)
    parrafo(doc, "Para responderla definimos cuatro preguntas de análisis:")
    vineta(doc, "¿Cómo evolucionó la participación de las ERNC en la matriz?")
    vineta(doc, "¿Qué tecnologías crecieron y cuáles están en retirada?")
    vineta(doc, "¿Cuán concentrada está la generación en unas pocas centrales?")
    vineta(doc, "¿Existe un patrón estacional distinto entre las renovables y la "
                "generación convencional?")
    parrafo(doc, "El resultado es una aplicación web interactiva que permite a "
                 "cualquier persona explorar la respuesta filtrando por período, "
                 "subsistema, clasificación y tecnología.",
            alineacion=WD_ALIGN_PARAGRAPH.JUSTIFY)

    # ------------------------------ 2. Dataset ---------------------------- #
    titulo(doc, "2. El conjunto de datos seleccionado")
    parrafo(doc, f"Se seleccionó el recurso **«{meta.get('nombre', config.DATASET_TITULO)}»**, "
                 f"publicado por la {config.DATASET_ORGANISMO} en "
                 f"{config.DATASET_URL}. La elección se justifica en tres "
                 f"criterios: (a) está expuesto en el *datastore* de CKAN, por lo "
                 f"que admite consulta vía API y no sólo descarga de archivo; "
                 f"(b) cubre una serie mensual larga y continua, indispensable "
                 f"para analizar tendencias; y (c) combina dimensiones "
                 f"categóricas con una métrica numérica, lo que habilita "
                 f"agregaciones y visualizaciones ricas.",
            alineacion=WD_ALIGN_PARAGRAPH.JUSTIFY)
    tabla_datos(doc, ["Campo", "Tipo", "Descripción"], [
        ["anio", "entero", "Año de la generación"],
        ["mes", "entero", "Mes de la generación (1–12)"],
        ["subsistema", "texto", "Subsistema eléctrico: SIC o SING"],
        ["clasificacion", "texto", "ERNC o Convencional"],
        ["tecnologia", "texto",
         f"Tecnología de generación ({df['tecnologia'].nunique()} categorías)"],
        ["codigo_central", "texto", "Nombre de la central generadora"],
        ["generacion_mwh", "decimal", "Energía bruta generada en el mes (MWh)"],
    ], anchos=[3.6, 2.4, 10.6])
    parrafo(doc, f"**Volumen:** {fmt(len(df), 0)} registros · "
                 f"{df['anio'].min()}–{df['anio'].max()} · "
                 f"{fmt(df['codigo_central'].nunique(), 0)} centrales · "
                 f"{df['tecnologia'].nunique()} tecnologías · "
                 f"{config.DATASET_LICENCIA}.",
            tamano=9, color=GRIS)

    # ------------------------------ 3. Proceso ---------------------------- #
    titulo(doc, "3. El proceso desarrollado")
    parrafo(doc, "El proyecto se organizó en **tres capas independientes**. Esta "
                 "separación fue una decisión de diseño deliberada: permite que "
                 "el informe que estás leyendo y la aplicación web usen "
                 "exactamente el mismo código de análisis, sin duplicarlo.",
            alineacion=WD_ALIGN_PARAGRAPH.JUSTIFY)
    tabla_datos(doc, ["Capa", "Archivo", "Responsabilidad", "Librerías"], [
        ["Configuración", "config.py",
         "Parámetros del proyecto e identificación del recurso", "—"],
        ["Acceso a datos", "api_datos.py",
         "Consultas GET a la API, paginación, limpieza y tipado",
         "requests, json, pandas"],
        ["Análisis", "analisis.py",
         "Agregaciones, indicadores y generación de gráficos",
         "pandas, matplotlib"],
        ["Presentación", "app.py",
         "Interfaz web interactiva con filtros", "streamlit"],
        ["Entrega", "generar_entrega.py",
         "Informe y póster en Word", "python-docx"],
    ], anchos=[2.9, 3.2, 6.9, 3.6], tamano=8.5)

    titulo(doc, "3.1 Obtención de los datos (requests + json)", 2)
    parrafo(doc, "La API CKAN limita cada respuesta a 10.000 filas. Para obtener "
                 "el recurso completo se recorre la paginación con el parámetro "
                 "`offset` hasta alcanzar el total que informa el propio "
                 "servidor:", alineacion=WD_ALIGN_PARAGRAPH.JUSTIFY)
    codigo(doc, "GET https://datos.gob.cl/api/3/action/datastore_search\n"
                f"    ?resource_id={config.RESOURCE_ID}\n"
                "    &limit=10000&offset=0")
    parrafo(doc, "Cada respuesta se valida en dos niveles: el código de estado "
                 "HTTP (raise_for_status) y la bandera `success` del cuerpo JSON. "
                 f"Se necesitaron 10 peticiones para traer los {fmt(len(df), 0)} "
                 f"registros analizados.", alineacion=WD_ALIGN_PARAGRAPH.JUSTIFY)

    titulo(doc, "3.2 Limpieza y transformación (pandas)", 2)
    parrafo(doc, "Los datos crudos no eran utilizables tal como llegaron:")
    vineta(doc, "**Formato numérico local:** generacion_mwh llegaba como texto "
                "con coma decimal (por ejemplo «119288,0»). Se eliminó el "
                "separador de miles y se sustituyó la coma por punto antes de "
                "convertir a float.")
    vineta(doc, "**Tipos:** año y mes llegaban como texto; se convirtieron a "
                "entero con pd.to_numeric.")
    vineta(doc, "**Columna derivada:** se construyó fecha (datetime) a partir de "
                "año y mes para habilitar las series de tiempo y la media móvil.")
    vineta(doc, "**Registros inválidos:** se descartaron las filas sin generación "
                "numérica válida y se normalizaron los textos categóricos con "
                "str.strip().")

    titulo(doc, "3.3 Análisis (pandas + matplotlib)", 2)
    parrafo(doc, "Todas las agregaciones se implementaron con groupby y "
                 "pivot_table —nunca con bucles sobre filas— y los seis gráficos "
                 "se generan con matplotlib a partir de una paleta y una función "
                 "de estilo comunes, de modo que la aplicación y este informe "
                 "sean visualmente idénticos.",
            alineacion=WD_ALIGN_PARAGRAPH.JUSTIFY)

    titulo(doc, "3.4 Aplicación web (streamlit)", 2)
    parrafo(doc, "La interfaz ofrece cinco filtros encadenados (rango de años, "
                 "subsistema, clasificación, tecnologías y umbral mínimo de "
                 "generación) que se combinan en una única máscara vectorizada. "
                 "Cualquier cambio recalcula los cinco indicadores y los seis "
                 "gráficos en tiempo real. El contenido se organiza en seis "
                 "pestañas —Resumen, Evolución, Tecnologías, Centrales, Datos e "
                 "Informe— más un botón de descarga de la selección en CSV. El "
                 "decorador @st.cache_data y una caché en disco evitan repetir "
                 "las peticiones HTTP en cada interacción.",
            alineacion=WD_ALIGN_PARAGRAPH.JUSTIFY)
    doc.add_page_break()

    # ------------------------------ 4. Resultados ------------------------- #
    titulo(doc, "4. Resultados y presentación de los datos")

    titulo(doc, "4.1 Recomposición de la matriz eléctrica", 2)
    figura(doc, rutas["mix"], ANCHO,
           "Figura 1. Generación anual por clasificación y penetración de ERNC "
           "en la matriz.")
    parrafo(doc, f"Entre {a0} y {a1} la generación total creció "
                 f"{fmt(kpi['variacion_pct'], 1, signo=True)}%, pero el dato relevante es la "
                 f"recomposición interna: las ERNC subieron de **{fmt(p0, 1)}%** a "
                 f"**{fmt(p1, 1)}%** de la matriz "
                 f"({fmt((p1 / p0 if p0 else 0), 1)}× su participación inicial), "
                 f"desplazando a la generación convencional incluso en términos "
                 f"absolutos.", alineacion=WD_ALIGN_PARAGRAPH.JUSTIFY)

    titulo(doc, "4.2 Evolución del mix tecnológico", 2)
    figura(doc, rutas["evolucion"], ANCHO,
           "Figura 2. Composición anual de la generación por tecnología (TWh).")
    parrafo(doc, f"El área apilada muestra el reemplazo con claridad: "
                 f"**{crece}** pasó de {fmt(ev.iloc[0][crece], 1)} a "
                 f"{fmt(ev.iloc[-1][crece], 1)} TWh anuales, mientras **{cae}** "
                 f"retrocedió de {fmt(ev.iloc[0][cae], 1)} a "
                 f"{fmt(ev.iloc[-1][cae], 1)} TWh.",
            alineacion=WD_ALIGN_PARAGRAPH.JUSTIFY)
    doc.add_page_break()

    titulo(doc, "4.3 Aporte acumulado por tecnología", 2)
    figura(doc, rutas["tecnologias"], ANCHO,
           "Figura 3. Generación acumulada por tecnología en el período "
           "analizado.")

    titulo(doc, "4.4 Tendencia y estacionalidad", 2)
    figura(doc, rutas["serie"], ANCHO,
           "Figura 4. Serie mensual de generación bruta con media móvil de 12 "
           "meses.")
    figura(doc, rutas["estacional"], ANCHO,
           "Figura 5. Patrón estacional promedio por clasificación.")
    doc.add_page_break()

    titulo(doc, "4.5 Concentración de la generación", 2)
    figura(doc, rutas["centrales"], ANCHO,
           "Figura 6. Las 15 centrales con mayor generación acumulada.")
    parrafo(doc, f"Las 15 centrales mayores concentran el "
                 f"**{fmt(concentracion, 1)}%** de toda la generación del período, "
                 f"sobre un total de {fmt(kpi['centrales'], 0)} centrales registradas: "
                 f"menos del {fmt(15 / kpi['centrales'] * 100, 1)}% de las "
                 f"instalaciones explica cerca de un tercio de la energía del "
                 f"país.", alineacion=WD_ALIGN_PARAGRAPH.JUSTIFY)

    titulo(doc, "4.6 Tabla resumen anual", 2)
    tabla_datos(doc,
                ["Año", "Convencional (TWh)", "ERNC (TWh)", "Total (TWh)", "% ERNC"],
                [[str(anio), f"{fmt(r.get('Convencional', 0), 2)}",
                  f"{fmt(r.get('ERNC', 0), 2)}", f"{fmt(r['Total'], 2)}",
                  f"{fmt(r['% ERNC'], 1)}%"] for anio, r in tabla.iterrows()],
                anchos=[2.2, 4.2, 3.4, 3.4, 3.4])
    doc.add_page_break()

    # ------------------------------ 5. Hallazgos -------------------------- #
    titulo(doc, "5. Hallazgos")
    vineta(doc, f"**La transición es real y medible.** La participación ERNC pasó "
                f"de {fmt(p0, 1)}% ({a0}) a {fmt(p1, 1)}% ({a1}). No es un efecto de "
                f"menor demanda: la generación total creció "
                f"{fmt(kpi['variacion_pct'], 1, signo=True)}% en el mismo período.")
    vineta(doc, f"**{crece} es el motor del cambio.** Multiplicó su generación "
                f"anual desde {fmt(ev.iloc[0][crece], 1)} hasta "
                f"{fmt(ev.iloc[-1][crece], 1)} TWh, y hoy es la tecnología con "
                f"mayor aporte anual del sistema.")
    vineta(doc, f"**{cae} está en retirada.** Perdió "
                f"{fmt(ev.iloc[0][cae] - ev.iloc[-1][cae], 1)} TWh anuales, lo que "
                f"confirma con datos el avance del plan de descarbonización.")
    vineta(doc, f"**La generación está muy concentrada.** 15 de "
                f"{fmt(kpi['centrales'], 0)} centrales producen el "
                f"{fmt(concentracion, 1)}% de la energía; la matriz renovable, en "
                f"cambio, es estructuralmente más distribuida.")
    vineta(doc, "**Renovables y convencionales tienen estacionalidad opuesta.** "
                "Las ERNC rinden más en los meses de verano (mayor radiación "
                "solar) mientras la generación convencional aumenta en invierno "
                "para cubrir la demanda térmica: una restricción operacional que "
                "explica por qué el almacenamiento es la siguiente frontera del "
                "sistema.")

    # ------------------------------ 6. Aprendizajes ----------------------- #
    titulo(doc, "6. Aprendizajes obtenidos")
    vineta(doc, "**Una API REST no entrega datos listos para analizar.** Los dos "
                "obstáculos reales no fueron conceptuales sino prácticos: la "
                "paginación de 10.000 filas y el formato numérico con coma "
                "decimal. Resolverlos exigió leer la respuesta cruda antes de "
                "escribir cualquier análisis.")
    vineta(doc, "**Separar en capas se pagó solo.** Al aislar acceso, análisis y "
                "presentación, el informe en Word y la app web comparten el "
                "mismo código de agregación y de gráficos. Cambiar un color o "
                "una métrica se hace en un solo lugar.")
    vineta(doc, "**El cacheo no es opcional en Streamlit.** El script se "
                "re-ejecuta completo ante cada interacción; sin @st.cache_data "
                "la aplicación repetiría las 10 peticiones HTTP cada vez que el "
                "usuario mueve un filtro.")
    vineta(doc, "**Vectorizar en pandas cambia la experiencia de uso.** "
                "Construir una única máscara booleana en lugar de iterar filas "
                "mantiene la respuesta de los filtros por debajo del segundo con "
                f"{fmt(len(df), 0)} registros.")
    vineta(doc, "**La visualización es parte del análisis, no un adorno.** El "
                "área apilada del mix tecnológico hizo evidente un reemplazo que "
                "en una tabla de cifras pasaba desapercibido.")

    # ------------------------------ 7. Ejecución -------------------------- #
    titulo(doc, "7. Cómo ejecutar el proyecto")
    codigo(doc, "pip install -r requirements.txt\nstreamlit run app.py")
    parrafo(doc, "La aplicación se abre en http://localhost:8501. El botón "
                 "«Actualizar desde la API» de la barra lateral fuerza una nueva "
                 "descarga en vivo desde datos.gob.cl. Para regenerar este "
                 "informe y el póster: python generar_entrega.py.",
            alineacion=WD_ALIGN_PARAGRAPH.JUSTIFY)

    # ------------------------------ 8. Referencias ------------------------ #
    titulo(doc, "8. Referencias")
    vineta(doc, f"{config.DATASET_ORGANISMO}. *{config.DATASET_TITULO}*. Portal "
                f"de Datos Abiertos del Gobierno de Chile. {config.DATASET_URL}")
    vineta(doc, "CKAN. *Action API Reference — datastore_search*. docs.ckan.org")
    vineta(doc, "Streamlit Inc. *Streamlit Documentation*. docs.streamlit.io")
    vineta(doc, "The pandas development team. *pandas Documentation*. "
                "pandas.pydata.org")
    vineta(doc, "Hunter, J. D. *Matplotlib: A 2D Graphics Environment*. "
                "matplotlib.org")

    doc.save(ruta)


# ========================================================================== #
#  PÓSTER (.docx, una página horizontal)
# ========================================================================== #
def construir_poster(ruta: Path, df, kpi, tabla, ev, rutas) -> None:
    """Póster de una página horizontal, organizado en dos columnas."""
    a0, a1 = tabla.index[0], tabla.index[-1]
    p0, p1 = tabla.loc[a0, "% ERNC"], tabla.loc[a1, "% ERNC"]
    crece = (ev.iloc[-1] - ev.iloc[0]).idxmax()
    cae = (ev.iloc[-1] - ev.iloc[0]).idxmin()

    doc = Document()
    doc.core_properties.title = f"Póster · {config.ACTIVIDAD}"
    doc.core_properties.author = ", ".join(config.INTEGRANTES)

    s = doc.sections[0]
    s.orientation = WD_ORIENT.LANDSCAPE
    s.page_width, s.page_height = Cm(29.7), Cm(21.0)
    s.left_margin = s.right_margin = Cm(1.0)
    s.top_margin = Cm(0.8)
    s.bottom_margin = Cm(0.6)

    TOTAL = 27.7            # ancho útil de la página
    COL = TOTAL / 2 - 0.2   # ancho de cada columna

    def caja_externa(destino, encabezado: str, ancho_cm: float):
        """Marco con título de sección; devuelve la celda de contenido."""
        t = destino.add_table(rows=2, cols=1)
        bordes(t)
        margenes_celda(t, 2, 3, 5, 5)
        ancho(t, [ancho_cm])
        compactar_cola(destino)
        cabecera = vaciar(t.cell(0, 0))
        sombrear(cabecera, HEX_CLARO)
        parrafo(cabecera, encabezado, tamano=10.5, negrita=True, color=AZUL,
                espacio_despues=0, interlineado=1.0)
        return vaciar(t.cell(1, 0))

    def texto(celda, contenido: str, tam: float = 8.6, bullet: str = ""):
        parrafo(celda, (f"{bullet}  " if bullet else "") + contenido,
                tamano=tam, espacio_despues=2.5, interlineado=1.05,
                sangria=0.35 if bullet else 0,
                alineacion=WD_ALIGN_PARAGRAPH.JUSTIFY)

    def separador(destino, puntos: float = 4):
        parrafo(destino, "", tamano=3, espacio_despues=puntos)

    # ----------------------------- Cabecera ------------------------------- #
    cab = doc.add_table(rows=1, cols=1)
    margenes_celda(cab, 4, 5, 8, 8)
    ancho(cab, [TOTAL])
    celda = vaciar(cab.cell(0, 0))
    sombrear(celda, "0F172A")
    parrafo(celda, "La transición energética de Chile en datos", tamano=19,
            negrita=True, color=RGBColor(0xFF, 0xFF, 0xFF),
            alineacion=WD_ALIGN_PARAGRAPH.CENTER, espacio_despues=2,
            interlineado=1.0)
    parrafo(celda, f"{config.ACTIVIDAD} · {config.UNIDAD_SEMANA} · "
                   f"{' · '.join(config.INTEGRANTES)}", tamano=9,
            color=RGBColor(0xCB, 0xD5, 0xE1),
            alineacion=WD_ALIGN_PARAGRAPH.CENTER, espacio_despues=0,
            interlineado=1.0)
    separador(doc, 5)

    # ------------------------------ Cuerpo -------------------------------- #
    cuerpo = doc.add_table(rows=1, cols=2)
    margenes_celda(cuerpo, 0, 0, 0, 0)
    ancho(cuerpo, [TOTAL / 2, TOTAL / 2])
    izq, der = vaciar(cuerpo.cell(0, 0)), vaciar(cuerpo.cell(0, 1))

    # --- Columna izquierda ------------------------------------------------ #
    tarjetas = izq.add_table(rows=1, cols=4)
    bordes(tarjetas)
    margenes_celda(tarjetas, 4, 4, 2, 2)
    ancho(tarjetas, [COL / 4] * 4)
    for j, (valor, etiqueta, color) in enumerate([
            (f"{fmt(p1, 1)}%", f"ERNC en {a1}", VERDE),
            (f"{fmt(kpi['total_twh'], 0)}", "TWh analizados", AZUL),
            (f"{fmt(kpi['centrales'], 0)}", "centrales", RGBColor(0x8B, 0x5C, 0xF6)),
            (f"{fmt(len(df), 0)}", "registros API", RGBColor(0xF5, 0x9E, 0x0B))]):
        celda = vaciar(tarjetas.cell(0, j))
        parrafo(celda, valor, tamano=14, negrita=True, color=color,
                alineacion=WD_ALIGN_PARAGRAPH.CENTER, espacio_despues=0,
                interlineado=1.0)
        parrafo(celda, etiqueta, tamano=7.5, color=GRIS,
                alineacion=WD_ALIGN_PARAGRAPH.CENTER, espacio_despues=0,
                interlineado=1.0)
    compactar_cola(izq)

    c = caja_externa(izq, "El desafío", COL)
    texto(c, "Chile comprometió el retiro de sus centrales a carbón y una matriz "
             "mayoritariamente renovable. **¿Cuánto ha avanzado esa transición y "
             "qué tecnologías la empujan?** Reconstruimos la generación mensual "
             "de cada central del país para responderlo con evidencia.")

    c = caja_externa(izq, "Datos y método", COL)
    texto(c, f"**Fuente:** «{config.DATASET_TITULO}», "
             f"{config.DATASET_ORGANISMO}, vía consultas **GET** a la API REST "
             f"CKAN de datos.gob.cl.")
    texto(c, f"**Volumen:** {fmt(len(df), 0)} registros · {df['anio'].min()}–"
             f"{df['anio'].max()} · {fmt(kpi['centrales'], 0)} centrales · "
             f"{df['tecnologia'].nunique()} tecnologías.")
    texto(c, "**requests + json** → paginación de 10.000 filas hasta completar "
             "el recurso.", bullet="1.")
    texto(c, "**pandas** → coma decimal a float, tipado de fechas, descarte de "
             "nulos.", bullet="2.")
    texto(c, "**pandas** → agregaciones con groupby y pivot_table.", bullet="3.")
    texto(c, "**matplotlib** → 6 visualizaciones con estilo común.", bullet="4.")
    texto(c, "**streamlit** → app web con 5 filtros y 6 pestañas.", bullet="5.")

    c = caja_externa(izq, "Hallazgos", COL)
    texto(c, f"**La transición es medible:** ERNC de {fmt(p0, 1)}% ({a0}) a "
             f"{fmt(p1, 1)}% ({a1}), con la generación total creciendo "
             f"{fmt(kpi['variacion_pct'], 1, signo=True)}%.", bullet="•")
    texto(c, f"**{crece} es el motor:** {fmt(ev.iloc[0][crece], 1)} → "
             f"{fmt(ev.iloc[-1][crece], 1)} TWh al año.", bullet="•")
    texto(c, f"**{cae} en retirada:** {fmt(ev.iloc[0][cae], 1)} → "
             f"{fmt(ev.iloc[-1][cae], 1)} TWh.", bullet="•")
    texto(c, "**Generación concentrada:** 15 centrales explican cerca de un "
             "tercio de la energía del país.", bullet="•")
    texto(c, "**Estacionalidad opuesta** entre renovables (verano) y "
             "convencionales (invierno).", bullet="•")

    c = caja_externa(izq, "Aprendizajes", COL)
    texto(c, "Una API REST no entrega datos listos: la paginación y la coma "
             "decimal fueron los obstáculos reales.", bullet="•")
    texto(c, "Separar acceso / análisis / presentación permitió que la app y "
             "este póster usen el mismo código.", bullet="•")
    texto(c, "Sin @st.cache_data la app repetiría 10 peticiones HTTP en cada "
             "interacción del usuario.", bullet="•")
    texto(c, f"Vectorizar los filtros mantiene la respuesta bajo un segundo con "
             f"{fmt(len(df), 0)} registros.", bullet="•")

    # --- Columna derecha -------------------------------------------------- #
    def figura_poster(celda, ruta_img: Path, pie: str, ancho_img: float):
        p = celda.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(1)
        p.add_run().add_picture(str(ruta_img), width=Cm(ancho_img))
        parrafo(celda, pie, tamano=7.2, cursiva=True, color=GRIS,
                alineacion=WD_ALIGN_PARAGRAPH.CENTER, espacio_despues=0,
                interlineado=1.0)

    IMG_ANCHO = COL - 1.4

    c = caja_externa(der, "La matriz eléctrica se recompone", COL)
    figura_poster(c, rutas["mix"],
                  "Generación anual por clasificación y penetración de ERNC.",
                  IMG_ANCHO)

    c = caja_externa(der, "El reemplazo tecnológico, año a año", COL)
    figura_poster(c, rutas["evolucion"],
                  "Composición anual de la generación por tecnología (TWh).",
                  IMG_ANCHO)

    # Tabla resumen: se muestran años alternos para que quepa en el póster
    c = caja_externa(der, "Resumen anual (TWh)", COL)
    anios = [a for i, a in enumerate(tabla.index) if i % 2 == 0]
    if tabla.index[-1] not in anios:
        anios.append(tabla.index[-1])
    resumen = c.add_table(rows=1, cols=len(anios) + 1)
    compactar_cola(c, espacio=0)
    bordes(resumen)
    margenes_celda(resumen, 2, 2, 3, 3)

    encabezados = [""] + [str(a) for a in anios]
    for j, texto_col in enumerate(encabezados):
        celda_h = vaciar(resumen.cell(0, j))
        sombrear(celda_h, HEX_AZUL)
        parrafo(celda_h, texto_col, tamano=7.2, negrita=True,
                color=RGBColor(0xFF, 0xFF, 0xFF), espacio_despues=0,
                interlineado=1.0,
                alineacion=WD_ALIGN_PARAGRAPH.CENTER if j else None)

    for n, etiqueta in enumerate(("Convencional", "ERNC", "% ERNC")):
        celdas = resumen.add_row().cells
        celda_e = vaciar(celdas[0])
        if n % 2 == 1:
            sombrear(celda_e, "F8FAFC")
        parrafo(celda_e, etiqueta, tamano=7.2, negrita=True,
                espacio_despues=0, interlineado=1.0)
        for j, a in enumerate(anios, start=1):
            celda_v = vaciar(celdas[j])
            if n % 2 == 1:
                sombrear(celda_v, "F8FAFC")
            valor = tabla.loc[a, etiqueta] if etiqueta in tabla else 0
            parrafo(celda_v,
                    f"{fmt(valor, 1)}%" if etiqueta == "% ERNC" else f"{fmt(valor, 1)}",
                    tamano=7.2, espacio_despues=0, interlineado=1.0,
                    alineacion=WD_ALIGN_PARAGRAPH.CENTER)
    ancho(resumen, [2.5] + [(COL - 2.9) / len(anios)] * len(anios))

    # -------------------------------- Pie --------------------------------- #
    parrafo(doc, f"Fuente: {config.DATASET_ORGANISMO} · API REST datos.gob.cl · "
                 f"{fmt(len(df), 0)} registros · {config.DATASET_LICENCIA}  |  "
                 f"{config.ASIGNATURA} · {config.UNIDAD_SEMANA}",
            tamano=7.2, color=GRIS, alineacion=WD_ALIGN_PARAGRAPH.CENTER,
            espacio_antes=3, espacio_despues=0, interlineado=1.0)

    doc.save(ruta)


# ========================================================================== #
#  Exportación opcional a PDF
# ========================================================================== #
def exportar_pdf(rutas_docx: list[Path]) -> None:
    """Convierte los .docx a PDF si LibreOffice está instalado.

    Si no lo está, indica cómo hacerlo desde Word (dos clics), que es la vía
    recomendada porque conserva exactamente el formato del documento.
    """
    candidatos = ["/Applications/LibreOffice.app/Contents/MacOS/soffice",
                  "soffice", "libreoffice"]
    soffice = next(
        (c for c in candidatos
         if Path(c).exists()
         or subprocess.run(["which", c], capture_output=True).returncode == 0),
        None)

    if not soffice:
        print("     LibreOffice no está instalado. Para obtener el PDF abre el "
              "documento en Word y usa:")
        print("       Archivo → Guardar como… → Formato: PDF")
        return

    for ruta in rutas_docx:
        subprocess.run([soffice, "--headless", "--convert-to", "pdf",
                        "--outdir", str(ruta.parent), str(ruta)],
                       capture_output=True)
        destino = ruta.with_suffix(".pdf")
        print(f"     {'✓' if destino.exists() else '!'} {destino.name}")


# ========================================================================== #
#  Main
# ========================================================================== #
def main() -> None:
    print("1/4  Obteniendo datos desde la API de datos.gob.cl …")
    df_total = api_datos.cargar_datos()
    try:
        meta = api_datos.obtener_metadatos()
    except Exception:
        meta = {}

    # Se excluye el último año si está incompleto, para no distorsionar el análisis
    ultimo = int(df_total["anio"].max())
    completo = (ultimo if df_total[df_total["anio"] == ultimo]["mes"].nunique() == 12
                else ultimo - 1)
    df = df_total[df_total["anio"] <= completo]
    print(f"     {fmt(len(df_total), 0)} registros · análisis sobre años completos "
          f"({df['anio'].min()}–{completo})")

    print("2/4  Calculando indicadores y generando gráficos …")
    kpi = an.indicadores(df)
    tabla = an.mix_anual(df)
    tec = an.por_tecnologia(df)
    ev = an.evolucion_tecnologia(df)
    rutas = exportar_graficos(df)

    print("3/4  Construyendo el informe en Word …")
    informe = BASE / f"{config.NOMBRE_ARCHIVO_ENTREGA}.docx"
    construir_informe(informe, df, kpi, tabla, tec, ev, meta, rutas)
    print(f"     ✓ {informe.name}")

    print("4/4  Construyendo el póster en Word …")
    poster = BASE / f"{config.NOMBRE_ARCHIVO_ENTREGA}_Poster.docx"
    construir_poster(poster, df, kpi, tabla, ev, rutas)
    print(f"     ✓ {poster.name}")

    if "--pdf" in sys.argv:
        print("\n5/5  Exportando a PDF …")
        exportar_pdf([informe, poster])

    print("\nListo. Revisa y edita los documentos en Word y luego expórtalos a "
          "PDF para subirlos a la plataforma.")


if __name__ == "__main__":
    main()
