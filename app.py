"""
Aplicación web interactiva (Streamlit)
======================================
Ejecutar con:  streamlit run app.py

Diseño deliberadamente simple: tres filtros, cuatro indicadores y tres
pestañas. Toda la interfaz está en español, incluidos los encabezados de las
tablas y los controles que Streamlit rotula en inglés por defecto.
"""

from __future__ import annotations

import io

import pandas as pd
import streamlit as st

import analisis as an
import api_datos
import config
from analisis import fmt

# --------------------------------------------------------------------------- #
# 1. Configuración de la página y estilos
# --------------------------------------------------------------------------- #
st.set_page_config(
    page_title="Matriz eléctrica de Chile",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
  .block-container {padding-top: 1.8rem; padding-bottom: 3rem; max-width: 1150px;}

  /* Encabezado */
  .portada {
      background: linear-gradient(120deg, #0f172a 0%, #0e7490 60%, #16a34a 100%);
      color: #fff; padding: 1.4rem 1.7rem; border-radius: 12px; margin-bottom: 1.5rem;
  }
  .portada h1 {color:#fff; font-size: 1.7rem; margin: 0 0 .3rem 0;}
  .portada p  {color: rgba(255,255,255,.9); margin: 0; font-size: .88rem;}

  /* Tarjetas de indicadores */
  .dato {
      background:#fff; border:1px solid #e2e8f0; border-top: 3px solid #0e7490;
      border-radius: 10px; padding: .9rem 1rem; text-align:center; height:100%;
  }
  .dato .cifra {font-size:1.55rem; font-weight:700; color:#0f172a; line-height:1.2;}
  .dato .rotulo {font-size:.76rem; color:#64748b; margin-top:.2rem;}

  /* Cajas de hallazgos */
  .idea {
      background:#f8fafc; border-left:4px solid #16a34a; border-radius:8px;
      padding:.75rem 1rem; margin-bottom:.6rem; font-size:.9rem;
  }

  .stTabs [data-baseweb="tab"] {font-size:.95rem; font-weight:600;}
  section[data-testid="stSidebar"] {background:#f8fafc;}

  /* Se ocultan los controles de Streamlit que están en inglés */
  #MainMenu, footer, header [data-testid="stToolbar"],
  [data-testid="stElementToolbar"], [data-testid="stDecoration"] {
      visibility: hidden;
  }
</style>
""", unsafe_allow_html=True)


# --------------------------------------------------------------------------- #
# 2. Carga de datos desde la API (en caché para no repetir peticiones)
# --------------------------------------------------------------------------- #
@st.cache_data(show_spinner=False)
def cargar(_version: int = 0) -> pd.DataFrame:
    barra = st.progress(0.0, text="Consultando la API de datos.gob.cl …")

    def avance(descargados: int, total: int) -> None:
        barra.progress(min(descargados / max(total, 1), 1.0),
                       text=f"Descargando registros: {fmt(descargados)} "
                            f"de {fmt(total)}")

    datos = api_datos.cargar_datos(usar_cache=(_version == 0), progreso=avance)
    barra.empty()
    return datos


if "version" not in st.session_state:
    st.session_state.version = 0

try:
    df = cargar(st.session_state.version)
except Exception as error:
    st.error(f"No fue posible obtener los datos de la API: {error}")
    st.info("Revisa tu conexión a Internet y vuelve a intentarlo.")
    st.stop()


# --------------------------------------------------------------------------- #
# 3. Barra lateral: tres filtros
# --------------------------------------------------------------------------- #
ANIO_MIN, ANIO_MAX = int(df["anio"].min()), int(df["anio"].max())
# El último año de la fuente suele estar incompleto: no se incluye por defecto
ULTIMO_COMPLETO = (ANIO_MAX if df[df["anio"] == ANIO_MAX]["mes"].nunique() == 12
                   else ANIO_MAX - 1)
TECNOLOGIAS = sorted(df["tecnologia"].unique())

with st.sidebar:
    st.markdown("### Filtros")

    rango = st.slider("Período", ANIO_MIN, ANIO_MAX,
                      (ANIO_MIN, ULTIMO_COMPLETO), key="f_anios",
                      help="Años que se incluyen en el análisis.")

    tipo = st.radio("Tipo de energía", ["Todas", "ERNC", "Convencional"],
                    key="f_tipo",
                    help="ERNC: energías renovables no convencionales.")

    # Se deja vacío a propósito: así la barra lateral no se llena de etiquetas
    # y el campo vacío se interpreta como «todas las tecnologías».
    elegidas = st.multiselect(
        "Tecnologías", TECNOLOGIAS, default=[], key="f_tecnologias",
        placeholder="Todas las tecnologías",
        help="Deja el campo vacío para incluir todas las tecnologías.")
    tecnologias = elegidas or TECNOLOGIAS

    st.divider()
    if st.button("Actualizar datos desde la API", width="stretch",
                 key="f_recargar"):
        st.session_state.version += 1
        st.cache_data.clear()
        st.rerun()

    st.markdown("### Fuente")
    st.caption(config.DATASET_ORGANISMO)
    st.caption(f"[{config.DATASET_TITULO}]({config.DATASET_URL})")
    st.caption(f"{fmt(len(df))} registros · {ANIO_MIN}–{ANIO_MAX}")


# --- Aplicación de los filtros en una sola máscara vectorizada ------------- #
mascara = (df["anio"].between(*rango) & df["tecnologia"].isin(tecnologias))
if tipo != "Todas":
    mascara &= df["clasificacion"] == tipo
dff = df[mascara]


# --------------------------------------------------------------------------- #
# 4. Encabezado
# --------------------------------------------------------------------------- #
st.markdown(f"""
<div class="portada">
  <h1>⚡ {config.TITULO}</h1>
  <p>{config.DATASET_TITULO} · datos obtenidos por API REST desde datos.gob.cl</p>
</div>
""", unsafe_allow_html=True)

if dff.empty:
    st.warning("Los filtros seleccionados no devuelven registros. "
               "Ajústalos en la barra lateral.")
    st.stop()


# --------------------------------------------------------------------------- #
# 5. Indicadores
# --------------------------------------------------------------------------- #
kpi = an.indicadores(dff)


def tarjeta(cifra: str, rotulo: str) -> str:
    return (f'<div class="dato"><div class="cifra">{cifra}</div>'
            f'<div class="rotulo">{rotulo}</div></div>')


c1, c2, c3, c4 = st.columns(4)
c1.markdown(tarjeta(f"{fmt(kpi['total_twh'], 1)} TWh", "Generación total"),
            unsafe_allow_html=True)
c2.markdown(tarjeta(f"{fmt(kpi['pct_ernc'], 1)} %", "Participación de ERNC"),
            unsafe_allow_html=True)
c3.markdown(tarjeta(f"{fmt(kpi['variacion_pct'], 1, signo=True)} %",
                    "Crecimiento del período"), unsafe_allow_html=True)
c4.markdown(tarjeta(fmt(kpi["centrales"]), "Centrales generadoras"),
            unsafe_allow_html=True)

st.write("")


# --------------------------------------------------------------------------- #
# 6. Tres pestañas
# --------------------------------------------------------------------------- #
resumen, tecnologia, datos = st.tabs(
    ["Resumen", "Tecnologías", "Datos"])

# --- 6.1 Resumen ----------------------------------------------------------- #
with resumen:
    st.pyplot(an.g_mix_anual(dff), width="stretch")

    tabla = an.mix_anual(dff)
    ev = an.evolucion_tecnologia(dff)

    if "% ERNC" in tabla and len(tabla) >= 2:
        a0, a1 = tabla.index[0], tabla.index[-1]
        p0, p1 = tabla.loc[a0, "% ERNC"], tabla.loc[a1, "% ERNC"]
        st.markdown(
            f'<div class="idea">Las <b>ERNC</b> pasaron de '
            f'<b>{fmt(p0, 1)} %</b> en {a0} a <b>{fmt(p1, 1)} %</b> en {a1}: '
            f'<b>{fmt(p1 / p0 if p0 else 0, 1)} veces</b> su participación '
            f'inicial.</div>', unsafe_allow_html=True)

    if len(ev) >= 2:
        crece = (ev.iloc[-1] - ev.iloc[0]).idxmax()
        cae = (ev.iloc[-1] - ev.iloc[0]).idxmin()
        st.markdown(
            f'<div class="idea">La tecnología que más creció es '
            f'<b>{crece}</b>: de {fmt(ev.iloc[0][crece], 1)} a '
            f'{fmt(ev.iloc[-1][crece], 1)} TWh al año.</div>',
            unsafe_allow_html=True)
        st.markdown(
            f'<div class="idea">La que más retrocedió es <b>{cae}</b>: '
            f'de {fmt(ev.iloc[0][cae], 1)} a {fmt(ev.iloc[-1][cae], 1)} TWh '
            f'al año.</div>', unsafe_allow_html=True)

    st.markdown("##### Generación anual")
    vista = pd.DataFrame({"Año": [str(a) for a in tabla.index]})
    for columna in ("Convencional", "ERNC", "Total"):
        if columna in tabla:
            vista[f"{columna} (TWh)"] = [fmt(v, 2) for v in tabla[columna]]
    if "% ERNC" in tabla:
        vista["ERNC (%)"] = [fmt(v, 1) for v in tabla["% ERNC"]]
    st.dataframe(vista, width="stretch", hide_index=True,
                 height=min(38 * len(vista) + 40, 460))

# --- 6.2 Tecnologías ------------------------------------------------------- #
with tecnologia:
    st.pyplot(an.g_tecnologias(dff), width="stretch")
    st.pyplot(an.g_evolucion_tecnologia(dff), width="stretch")

    st.markdown("##### Detalle por tecnología")
    detalle = an.por_tecnologia(dff)
    st.dataframe(pd.DataFrame({
        "Tecnología": detalle["tecnologia"],
        "Clasificación": detalle["clasificacion"],
        "Generación (TWh)": [fmt(v, 2) for v in detalle["twh"]],
        "Participación (%)": [fmt(v, 1) for v in detalle["pct"]],
    }), width="stretch", hide_index=True,
        height=min(38 * len(detalle) + 40, 520))

# --- 6.3 Datos ------------------------------------------------------------- #
with datos:
    st.markdown(f"##### {fmt(len(dff))} registros seleccionados "
                f"de {fmt(len(df))}")
    # Se omite «fecha» porque Año y Mes ya la expresan de forma más legible
    columnas = [c for c in dff.columns if c != "fecha"]
    muestra = an.para_mostrar(dff[columnas].head(500)).copy()
    # Formato chileno también en la tabla cruda, para no mezclar notaciones
    muestra["Generación (MWh)"] = [fmt(v, 1)
                                   for v in muestra["Generación (MWh)"]]
    st.dataframe(muestra, width="stretch", hide_index=True)

    # Punto y coma con coma decimal: es el formato que Excel en español espera
    archivo = io.StringIO()
    an.para_mostrar(dff[columnas]).to_csv(archivo, index=False, sep=";",
                                          decimal=",")
    st.download_button("Descargar la selección en CSV", archivo.getvalue(),
                       file_name="generacion_sen.csv", mime="text/csv",
                       key="f_descargar")

    with st.expander("¿De dónde vienen estos datos?"):
        st.markdown(f"""
**Recurso:** {config.DATASET_TITULO}
**Organismo:** {config.DATASET_ORGANISMO}
**Portal:** {config.DATASET_URL}
**Licencia:** {config.DATASET_LICENCIA}

Los datos se obtienen con peticiones `GET` a la API REST del portal de datos
abiertos del Gobierno de Chile. La API entrega como máximo 10.000 filas por
respuesta, así que se recorre la paginación con el parámetro `offset` hasta
completar el total que informa el servidor:

```
GET {config.API_BASE}/datastore_search
    ?resource_id={config.RESOURCE_ID}
    &limit=10000&offset=0
```

**Campos disponibles**
""")
        st.dataframe(pd.DataFrame([
            ("Año", "Año de la generación"),
            ("Mes", "Mes de la generación (1 a 12)"),
            ("Subsistema", "Subsistema eléctrico: SIC o SING"),
            ("Clasificación", "ERNC o Convencional"),
            ("Tecnología", "Tecnología de generación (13 categorías)"),
            ("Central", "Nombre de la central generadora"),
            ("Generación (MWh)", "Energía bruta generada en el mes"),
        ], columns=["Campo", "Descripción"]),
            width="stretch", hide_index=True)


st.divider()
st.caption(
    f"Datos: [{config.DATASET_ORGANISMO}]({config.DATASET_URL}) · "
    f"obtenidos por API REST desde datos.gob.cl · {config.DATASET_LICENCIA}"
)
