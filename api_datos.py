"""
Capa de acceso a datos: consumo de la API REST pública de datos.gob.cl
(CKAN Action API) mediante peticiones GET.

Librerías: requests (HTTP) + json (parseo) + pandas (estructura tabular).

Endpoints utilizados
--------------------
GET /api/3/action/datastore_search  -> registros del recurso (paginado)
GET /api/3/action/resource_show     -> metadatos del recurso
"""

from __future__ import annotations

import gzip
import json
from pathlib import Path

import pandas as pd
import requests

import config

# La caché se guarda comprimida: 20 MB de JSON crudo bajan a ~2 MB.
CACHE = Path(__file__).parent / "cache_datos.json.gz"


# --------------------------------------------------------------------------- #
# 1. Llamada genérica a la API
# --------------------------------------------------------------------------- #
def _get(accion: str, **params) -> dict:
    """Ejecuta un GET contra la API CKAN y devuelve el bloque 'result'.

    Centralizar la petición aquí evita repetir código de red y de manejo
    de errores en el resto del proyecto.
    """
    url = f"{config.API_BASE}/{accion}"
    respuesta = requests.get(url, params=params, timeout=config.TIMEOUT)
    respuesta.raise_for_status()          # error HTTP -> excepción explícita

    cuerpo = json.loads(respuesta.text)   # uso explícito de la librería json
    if not cuerpo.get("success"):
        raise RuntimeError(f"La API respondió sin éxito: {cuerpo.get('error')}")
    return cuerpo["result"]


# --------------------------------------------------------------------------- #
# 2. Metadatos del recurso
# --------------------------------------------------------------------------- #
def obtener_metadatos() -> dict:
    """Devuelve los metadatos publicados del recurso consultado."""
    r = _get("resource_show", id=config.RESOURCE_ID)
    return {
        "nombre": r.get("name", ""),
        "formato": r.get("format", ""),
        "creado": (r.get("created") or "")[:10],
        "modificado": (r.get("last_modified") or r.get("created") or "")[:10],
        "url_descarga": r.get("url", ""),
    }


# --------------------------------------------------------------------------- #
# 3. Descarga completa del recurso (paginada)
# --------------------------------------------------------------------------- #
def descargar_registros(progreso=None) -> list[dict]:
    """Descarga TODOS los registros del recurso recorriendo la paginación.

    La API limita cada respuesta a 10.000 filas, por lo que se avanza con
    el parámetro `offset` hasta agotar el total informado por el servidor.

    Parameters
    ----------
    progreso : callable | None
        Función opcional `f(descargados, total)` para reportar avance
        (la usa Streamlit para dibujar la barra de progreso).
    """
    registros: list[dict] = []
    offset, total = 0, None

    while True:
        result = _get(
            "datastore_search",
            resource_id=config.RESOURCE_ID,
            limit=config.PAGE_SIZE,
            offset=offset,
        )
        total = result["total"]
        lote = result["records"]
        registros.extend(lote)

        if progreso:
            progreso(len(registros), total)

        offset += config.PAGE_SIZE
        if not lote or len(registros) >= total:
            break

    return registros


# --------------------------------------------------------------------------- #
# 4. Limpieza y tipado (pandas)
# --------------------------------------------------------------------------- #
def limpiar(registros: list[dict]) -> pd.DataFrame:
    """Convierte la respuesta JSON en un DataFrame limpio y tipado.

    Transformaciones aplicadas:
      * se descarta la columna técnica `_id`;
      * `generacion_mwh` viene como texto con coma decimal -> float;
      * `anio` y `mes` -> entero;
      * se construye una columna `fecha` (datetime) para las series de tiempo;
      * se eliminan filas sin generación válida y se normalizan textos.
    """
    df = pd.DataFrame(registros)

    df = df.drop(columns=[c for c in ("_id", "fecha_act") if c in df.columns])

    df["generacion_mwh"] = (
        df["generacion_mwh"].astype(str)
        .str.replace(".", "", regex=False)     # separador de miles
        .str.replace(",", ".", regex=False)    # coma decimal -> punto
        .pipe(pd.to_numeric, errors="coerce")
    )

    df["anio"] = pd.to_numeric(df["anio"], errors="coerce").astype("Int64")
    df["mes"] = pd.to_numeric(df["mes"], errors="coerce").astype("Int64")

    df = df.dropna(subset=["generacion_mwh", "anio", "mes"])
    df = df.astype({"anio": int, "mes": int})

    for col in ("tecnologia", "subsistema", "clasificacion", "codigo_central"):
        df[col] = df[col].astype(str).str.strip()

    df["fecha"] = pd.to_datetime(
        dict(year=df["anio"], month=df["mes"], day=1), errors="coerce"
    )
    orden = ["fecha", "anio", "mes", "subsistema", "clasificacion",
             "tecnologia", "codigo_central", "generacion_mwh"]
    return df[orden].sort_values("fecha").reset_index(drop=True)


# --------------------------------------------------------------------------- #
# 5. Punto de entrada único + caché local
# --------------------------------------------------------------------------- #
def cargar_datos(usar_cache: bool = True, progreso=None) -> pd.DataFrame:
    """Devuelve el DataFrame listo para analizar.

    Si existe una caché local (`cache_datos.json.gz`) la reutiliza; así la app
    no vuelve a golpear la API en cada recarga y funciona incluso sin
    conexión durante la presentación.
    """
    if usar_cache and CACHE.exists():
        with gzip.open(CACHE, "rt", encoding="utf-8") as f:
            registros = json.load(f)
    else:
        registros = descargar_registros(progreso=progreso)
        with gzip.open(CACHE, "wt", encoding="utf-8") as f:
            json.dump(registros, f, ensure_ascii=False)
    return limpiar(registros)


if __name__ == "__main__":
    # Ejecutar `python api_datos.py` para probar la conexión con la API.
    print("Consultando la API de datos.gob.cl ...")
    print("Metadatos:", obtener_metadatos())
    datos = cargar_datos(usar_cache=False)
    print(f"Registros descargados: {len(datos):,}")
    print(datos.head())
