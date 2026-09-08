"""
Capa de análisis: transformaciones con pandas y gráficos con matplotlib.

Todas las funciones reciben un DataFrame ya filtrado y devuelven
DataFrames agregados o figuras de matplotlib. No se hace ninguna llamada
de red aquí: así el análisis es reutilizable tanto por la app Streamlit
como por el generador del informe en PDF.
"""

from __future__ import annotations

import matplotlib
matplotlib.use("Agg")            # backend sin ventana (necesario en servidor)
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import pandas as pd

MWH_A_TWH = 1_000_000            # 1 TWh = 1.000.000 MWh

# Los nombres internos de las columnas replican los de la API; se traducen
# sólo al momento de mostrarlos, para no perder la trazabilidad con la fuente.
NOMBRES_COLUMNAS = {
    "fecha": "Fecha",
    "anio": "Año",
    "mes": "Mes",
    "subsistema": "Subsistema",
    "clasificacion": "Clasificación",
    "tecnologia": "Tecnología",
    "codigo_central": "Central",
    "generacion_mwh": "Generación (MWh)",
}


def para_mostrar(df: pd.DataFrame) -> pd.DataFrame:
    """Devuelve el DataFrame con los encabezados en español."""
    return df.rename(columns=NOMBRES_COLUMNAS)

# Paleta consistente en toda la aplicación y el informe
COLORES = {
    "ERNC": "#16a34a",
    "Convencional": "#64748b",
    "acento": "#0ea5e9",
    "alerta": "#dc2626",
}
PALETA = ["#0ea5e9", "#f59e0b", "#16a34a", "#dc2626", "#8b5cf6", "#0f766e",
          "#db2777", "#65a30d", "#f97316", "#2563eb", "#9333ea", "#059669",
          "#a16207"]


def fmt(valor, decimales: int = 0, signo: bool = False) -> str:
    """Formatea un número al estilo chileno: punto de miles y coma decimal.

    Mezclar «89,313» con «43.6» en un mismo informe se lee como descuido, así
    que todos los números del proyecto pasan por aquí.
    """
    texto = f"{valor:{'+' if signo else ''},.{decimales}f}"
    return texto.replace(",", "\x00").replace(".", ",").replace("\x00", ".")


def _estilo(ax, titulo: str, ylabel: str = "", xlabel: str = "") -> None:
    """Aplica un estilo visual uniforme a cualquier eje."""
    ax.set_title(titulo, fontsize=12, fontweight="bold", pad=12)
    ax.set_ylabel(ylabel, fontsize=9)
    ax.set_xlabel(xlabel, fontsize=9)
    ax.grid(axis="y", linestyle=":", alpha=.5)
    ax.set_axisbelow(True)
    for lado in ("top", "right"):
        ax.spines[lado].set_visible(False)
    ax.tick_params(labelsize=8)


# ========================================================================== #
#  A. AGREGACIONES (pandas)
# ========================================================================== #
def indicadores(df: pd.DataFrame) -> dict:
    """Indicadores globales (KPI) del subconjunto seleccionado."""
    total = df["generacion_mwh"].sum()
    ernc = df.loc[df["clasificacion"] == "ERNC", "generacion_mwh"].sum()

    por_anio = df.groupby("anio")["generacion_mwh"].sum()
    anios_completos = por_anio.index[:-1] if len(por_anio) > 1 else por_anio.index
    variacion = 0.0
    if len(anios_completos) >= 2:
        ini, fin = por_anio[anios_completos[0]], por_anio[anios_completos[-1]]
        variacion = (fin / ini - 1) * 100 if ini else 0.0

    lider = df.groupby("tecnologia")["generacion_mwh"].sum().idxmax() if len(df) else "-"

    return {
        "total_twh": total / MWH_A_TWH,
        "pct_ernc": (ernc / total * 100) if total else 0.0,
        "centrales": df["codigo_central"].nunique(),
        "tecnologias": df["tecnologia"].nunique(),
        "variacion_pct": variacion,
        "tecnologia_lider": lider,
        "periodo": (f"{df['fecha'].min():%m/%Y} – {df['fecha'].max():%m/%Y}"
                    if len(df) else "-"),
    }


def serie_mensual(df: pd.DataFrame) -> pd.DataFrame:
    """Generación total por mes (serie de tiempo)."""
    return (df.groupby("fecha", as_index=False)["generacion_mwh"].sum()
              .assign(twh=lambda d: d["generacion_mwh"] / MWH_A_TWH))


def mix_anual(df: pd.DataFrame) -> pd.DataFrame:
    """Participación anual ERNC vs Convencional, en TWh y en %."""
    tabla = (df.pivot_table(index="anio", columns="clasificacion",
                            values="generacion_mwh", aggfunc="sum")
               .fillna(0) / MWH_A_TWH)
    tabla["Total"] = tabla.sum(axis=1)
    if "ERNC" in tabla:
        tabla["% ERNC"] = (tabla["ERNC"] / tabla["Total"] * 100).round(1)
    return tabla.round(2)


def _categoria_dominante(df: pd.DataFrame, clave: str,
                         categoria: str) -> pd.DataFrame:
    """Asocia a cada `clave` la `categoria` que concentra más generación.

    Hace falta porque la fuente asigna dos categorías a unos pocos casos: hay
    4 registros de Gas Natural marcados como ERNC (de 5.472) y 10 centrales
    con dos tecnologías. Sin esta reducción la clave se duplica y los gráficos
    de barras superponen dos valores en la misma fila.
    """
    return (df.groupby([clave, categoria], as_index=False)["generacion_mwh"]
              .sum()
              .sort_values("generacion_mwh", ascending=False)
              .drop_duplicates(clave)[[clave, categoria]])


def por_tecnologia(df: pd.DataFrame) -> pd.DataFrame:
    """Ranking de tecnologías con su aporte porcentual."""
    total = df.groupby("tecnologia", as_index=False)["generacion_mwh"].sum()
    t = (total.merge(_categoria_dominante(df, "tecnologia", "clasificacion"),
                     on="tecnologia")
              .sort_values("generacion_mwh", ascending=False))
    t["twh"] = (t["generacion_mwh"] / MWH_A_TWH).round(2)
    t["pct"] = (t["generacion_mwh"] / t["generacion_mwh"].sum() * 100).round(1)
    return t.reset_index(drop=True)


def evolucion_tecnologia(df: pd.DataFrame) -> pd.DataFrame:
    """Matriz año × tecnología en TWh (para el área apilada)."""
    return (df.pivot_table(index="anio", columns="tecnologia",
                           values="generacion_mwh", aggfunc="sum")
              .fillna(0) / MWH_A_TWH)


def top_centrales(df: pd.DataFrame, n: int = 15) -> pd.DataFrame:
    """Las n centrales con mayor generación acumulada."""
    total = df.groupby("codigo_central", as_index=False)["generacion_mwh"].sum()
    t = (total.merge(_categoria_dominante(df, "codigo_central", "tecnologia"),
                     on="codigo_central")
              .nlargest(n, "generacion_mwh"))
    t["twh"] = (t["generacion_mwh"] / MWH_A_TWH).round(2)
    return t.reset_index(drop=True)


def estacionalidad(df: pd.DataFrame) -> pd.DataFrame:
    """Promedio mensual histórico por clasificación (patrón estacional)."""
    return (df.pivot_table(index="mes", columns="clasificacion",
                           values="generacion_mwh", aggfunc="mean")
              .fillna(0) / 1000)          # GWh promedio por central-mes


# ========================================================================== #
#  B. GRÁFICOS (matplotlib)
# ========================================================================== #
def g_serie_mensual(df: pd.DataFrame):
    datos = serie_mensual(df)
    fig, ax = plt.subplots(figsize=(9, 3.6))
    ax.plot(datos["fecha"], datos["twh"], color=COLORES["acento"], lw=1.8)
    ax.fill_between(datos["fecha"], datos["twh"], color=COLORES["acento"], alpha=.15)
    if len(datos) >= 12:
        ax.plot(datos["fecha"], datos["twh"].rolling(12).mean(),
                color=COLORES["alerta"], lw=1.6, ls="--",
                label="Media móvil 12 meses")
        ax.legend(fontsize=8, frameon=False)
    _estilo(ax, "Generación bruta mensual del SEN", "TWh")
    fig.tight_layout()
    return fig


def g_mix_anual(df: pd.DataFrame):
    tabla = mix_anual(df)
    cols = [c for c in ("Convencional", "ERNC") if c in tabla]
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(9, 3.6))

    base = 0
    for c in cols:
        ax1.bar(tabla.index.astype(str), tabla[c], bottom=base,
                color=COLORES[c], label=c, width=.68)
        base = base + tabla[c]
    _estilo(ax1, "Generación anual por clasificación", "TWh")
    ax1.legend(fontsize=8, frameon=False)
    ax1.tick_params(axis="x", rotation=90)

    if "% ERNC" in tabla:
        ax2.plot(tabla.index.astype(str), tabla["% ERNC"], marker="o", ms=4,
                 color=COLORES["ERNC"], lw=2)
        ax2.yaxis.set_major_formatter(mticker.PercentFormatter(decimals=0))
        _estilo(ax2, "Penetración de ERNC en la matriz", "% del total")
        ax2.tick_params(axis="x", rotation=90)
    fig.tight_layout()
    return fig


def g_tecnologias(df: pd.DataFrame):
    t = por_tecnologia(df)
    fig, ax = plt.subplots(figsize=(9, 4))

    # Se usan posiciones numéricas en lugar de los nombres: si dos filas
    # compartieran nombre, matplotlib las apilaría en la misma barra.
    posiciones = range(len(t))
    ax.barh(posiciones, t["twh"][::-1].values,
            color=[COLORES[c] for c in t["clasificacion"][::-1]])
    ax.set_yticks(list(posiciones))
    ax.set_yticklabels(t["tecnologia"][::-1].values)

    for y, valor, pct in zip(posiciones, t["twh"][::-1], t["pct"][::-1]):
        ax.text(valor + t["twh"].max() * .012, y,
                f"{fmt(valor, 1)} TWh ({fmt(pct, 1)}%)",
                va="center", fontsize=7.5)

    ax.set_xlim(0, t["twh"].max() * 1.3 if len(t) else 1)
    _estilo(ax, "Generación acumulada por tecnología", "", "TWh")
    fig.tight_layout()
    return fig


def g_evolucion_tecnologia(df: pd.DataFrame):
    tabla = evolucion_tecnologia(df)
    orden = tabla.sum().sort_values(ascending=False).index
    tabla = tabla[orden]
    fig, ax = plt.subplots(figsize=(9, 4))
    ax.stackplot(tabla.index, tabla.T.values, labels=tabla.columns,
                 colors=PALETA[:len(tabla.columns)], alpha=.9)
    ax.legend(fontsize=7, ncol=2, frameon=False, loc="upper left")
    ax.xaxis.set_major_locator(mticker.MaxNLocator(integer=True))
    _estilo(ax, "Evolución del mix tecnológico", "TWh", "Año")
    fig.tight_layout()
    return fig


def g_top_centrales(df: pd.DataFrame, n: int = 15):
    t = top_centrales(df, n)
    fig, ax = plt.subplots(figsize=(9, 4))
    posiciones = range(len(t))
    ax.barh(posiciones, t["twh"][::-1].values, color=COLORES["acento"])
    ax.set_yticks(list(posiciones))
    ax.set_yticklabels(t["codigo_central"][::-1].values)
    _estilo(ax, f"Top {n} centrales por generación acumulada", "", "TWh")
    fig.tight_layout()
    return fig


def g_estacionalidad(df: pd.DataFrame):
    tabla = estacionalidad(df)
    etiquetas = ["Ene", "Feb", "Mar", "Abr", "May", "Jun",
                 "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]
    fig, ax = plt.subplots(figsize=(9, 3.4))
    for c in tabla.columns:
        ax.plot([etiquetas[i - 1] for i in tabla.index], tabla[c],
                marker="o", ms=4, lw=2, color=COLORES.get(c, COLORES["acento"]),
                label=c)
    ax.legend(fontsize=8, frameon=False)
    _estilo(ax, "Patrón estacional (promedio por central-mes)", "GWh", "Mes")
    fig.tight_layout()
    return fig
