#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
==============================================================================
 ANÁLISIS DE ÍNDICES DE TEMPERATURA DE ESTACIÓN METEOROLÓGICA
==============================================================================
Este script lee el historial de la estación meteorológica (archivo Excel),
calcula índices de temperatura (medias, mínimas, máximas anuales, mensuales,
estacionales y de canícula), genera rankings de días/meses/años más
calurosos y otros índices climáticos relevantes, y produce un informe
completo en formato Word (.docx) con tablas y gráficos.

USO:
    python3 analisis_meteorologico.py [ruta_al_excel] [ruta_salida.docx]

Si no se indican rutas, se usan los valores por defecto definidos en
la sección CONFIGURACIÓN. El script está pensado para poder relanzarse
cada vez que el archivo Excel de la estación se actualice: basta con
sustituir el fichero de entrada (o apuntar a la nueva ruta) y volver a
ejecutar. Todo el contenido del informe (tablas, gráficos, rankings) se
recalcula automáticamente a partir de los datos disponibles en ese momento.
==============================================================================
"""

import os
import sys
import warnings
from datetime import datetime

import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker

warnings.filterwarnings("ignore")

# ==============================================================================
# CONFIGURACIÓN (editable)
# ==============================================================================

ARCHIVO_ENTRADA_DEFECTO = "weather_data_metric_es.xlsx"
ARCHIVO_SALIDA_DEFECTO = "Informe_Meteorologico.docx"
CARPETA_IMAGENES = "imgs_tmp"

# Periodo de "canícula" (los días más calurosos del verano, tradición española).
# Por defecto: 15 de julio - 15 de agosto (configurable).
CANICULA_INICIO = (7, 15)   # (mes, día)
CANICULA_FIN = (8, 15)      # (mes, día)

# Umbrales de índices térmicos (estándar climatológico, en °C)
UMBRAL_DIA_VERANO = 25.0     # "día de verano": Tmax >= 25°C
UMBRAL_DIA_CALUROSO = 30.0   # "día cálido": Tmax >= 30°C
UMBRAL_DIA_MUY_CALUROSO = 35.0  # "día muy cálido / ola de calor": Tmax >= 35°C
UMBRAL_NOCHE_TROPICAL = 20.0    # "noche tropical": Tmin >= 20°C
UMBRAL_NOCHE_TORRIDA = 25.0     # "noche tórrida": Tmin >= 25°C
UMBRAL_DIA_HELADA = 0.0         # "día de helada": Tmin <= 0°C

N_RANKING = 15  # nº de elementos en cada ranking (días/meses)

MESES_ES = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio",
    7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre",
    12: "Diciembre",
}
MESES_ES_ABR = {
    1: "Ene", 2: "Feb", 3: "Mar", 4: "Abr", 5: "May", 6: "Jun",
    7: "Jul", 8: "Ago", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dic",
}

# Estaciones meteorológicas (hemisferio norte)
def estacion_del_ano(mes, dia):
    md = (mes, dia)
    if (md >= (3, 21)) and (md < (6, 21)):
        return "Primavera"
    if (md >= (6, 21)) and (md < (9, 23)):
        return "Verano"
    if (md >= (9, 23)) and (md < (12, 21)):
        return "Otoño"
    return "Invierno"


# Paleta de colores coherente para todo el informe
COLOR_MAX = "#d64545"
COLOR_MEDIA = "#e8952e"
COLOR_MIN = "#3773b5"
COLOR_PRECIP = "#2f7a5c"
COLOR_NEUTRO = "#555555"
COLOR_FONDO_BARRA = "#4a6fa5"

plt.rcParams.update({
    "font.family": "DejaVu Sans",
    "font.size": 10.5,
    "axes.titlesize": 13,
    "axes.titleweight": "bold",
    "axes.labelsize": 10.5,
    "axes.edgecolor": "#888888",
    "axes.grid": True,
    "grid.color": "#dddddd",
    "grid.linewidth": 0.6,
    "figure.facecolor": "white",
    "axes.facecolor": "white",
    "legend.frameon": False,
})


# ==============================================================================
# CARGA Y PREPARACIÓN DE DATOS
# ==============================================================================

def cargar_datos(ruta_excel: str) -> pd.DataFrame:
    """Lee el Excel de la estación y devuelve un DataFrame diario unificado.

    El archivo de la estación ha ido cambiando de formato de exportación con
    el tiempo: los registros más antiguos usan columnas como
    'Temperatura Máxima (°C)' / 'Temperatura Mínima (°C)' / 'tempAvg',
    mientras que los más recientes usan 'Temperature_Max (℃)' /
    'Temperature_Min (℃)' / 'Temperature_Media (℃)'. Esta función unifica
    ambos formatos automáticamente (tomando el valor que esté disponible),
    de modo que el análisis cubra siempre todo el histórico, sin importar
    qué columnas use la exportación más reciente del archivo actualizado.
    """
    if not os.path.exists(ruta_excel):
        raise FileNotFoundError(f"No se encuentra el archivo de datos: {ruta_excel}")

    df = pd.read_excel(ruta_excel)

    if "Fecha" not in df.columns:
        raise ValueError("El archivo no contiene una columna 'Fecha'.")

    df["Fecha"] = pd.to_datetime(df["Fecha"])
    df = df.sort_values("Fecha").drop_duplicates(subset="Fecha").reset_index(drop=True)

    def coalesce(*cols):
        """Devuelve la primera columna no nula, fila a fila, entre las dadas."""
        cols_presentes = [c for c in cols if c in df.columns]
        if not cols_presentes:
            return pd.Series(np.nan, index=df.index)
        out = df[cols_presentes[0]].copy()
        for c in cols_presentes[1:]:
            out = out.fillna(df[c])
        return out

    salida = pd.DataFrame()
    salida["Fecha"] = df["Fecha"]
    salida["TMax"] = coalesce("Temperatura Máxima (°C)", "Temperature_Max (℃)")
    salida["TMin"] = coalesce("Temperatura Mínima (°C)", "Temperature_Min (℃)")
    salida["TMedia"] = coalesce("tempAvg", "Temperature_Media (℃)")
    salida["Precip"] = coalesce("Precipitación Total (mm)", "Daily_Max (mm)")

    # Índice de amplitud térmica diaria (DTR)
    salida["DTR"] = salida["TMax"] - salida["TMin"]

    # Descartar filas sin ningún dato de temperatura útil
    salida = salida[salida["TMax"].notna() | salida["TMedia"].notna()].copy()

    # Si falta TMedia pero hay TMax/TMin, se estima como su punto medio
    estim = (salida["TMax"] + salida["TMin"]) / 2.0
    salida["TMedia"] = salida["TMedia"].fillna(estim)

    salida["Año"] = salida["Fecha"].dt.year
    salida["Mes"] = salida["Fecha"].dt.month
    salida["Dia"] = salida["Fecha"].dt.day
    salida["MesNombre"] = salida["Mes"].map(MESES_ES)
    salida["Estacion"] = salida.apply(lambda r: estacion_del_ano(r["Mes"], r["Dia"]), axis=1)

    ini_m, ini_d = CANICULA_INICIO
    fin_m, fin_d = CANICULA_FIN
    salida["EnCanicula"] = salida.apply(
        lambda r: (r["Mes"], r["Dia"]) >= (ini_m, ini_d) and (r["Mes"], r["Dia"]) <= (fin_m, fin_d),
        axis=1,
    )

    salida = salida.sort_values("Fecha").reset_index(drop=True)
    return salida


def resumen_cobertura(df: pd.DataFrame) -> dict:
    fecha_ini = df["Fecha"].min()
    fecha_fin = df["Fecha"].max()
    dias_teoricos = (fecha_fin - fecha_ini).days + 1
    dias_con_dato = df["TMax"].notna().sum()
    return {
        "fecha_ini": fecha_ini,
        "fecha_fin": fecha_fin,
        "dias_teoricos": dias_teoricos,
        "dias_con_dato": int(dias_con_dato),
        "dias_faltantes": dias_teoricos - int(dias_con_dato),
        "pct_cobertura": 100.0 * dias_con_dato / dias_teoricos if dias_teoricos else 0.0,
    }


# ==============================================================================
# CÁLCULOS ESTADÍSTICOS
# ==============================================================================

def tabla_anual(df: pd.DataFrame) -> pd.DataFrame:
    g = df.groupby("Año").agg(
        Media=("TMedia", "mean"),
        Max_media=("TMax", "mean"),
        Min_media=("TMin", "mean"),
        Max_absoluta=("TMax", "max"),
        Min_absoluta=("TMin", "min"),
        Precip_total=("Precip", "sum"),
        Dias=("TMax", "count"),
    ).reset_index()
    return g


def tabla_mensual_climatologia(df: pd.DataFrame) -> pd.DataFrame:
    """Media histórica por mes del año (climatología), independiente del año."""
    g = df.groupby("Mes").agg(
        Media=("TMedia", "mean"),
        Max_media=("TMax", "mean"),
        Min_media=("TMin", "mean"),
        Max_absoluta=("TMax", "max"),
        Min_absoluta=("TMin", "min"),
        Precip_media=("Precip", "sum"),
    ).reset_index()
    g["Años_muestreados"] = df.groupby("Mes")["Año"].nunique().values
    g["Precip_media"] = g["Precip_media"] / g["Años_muestreados"].replace(0, np.nan)
    g["MesNombre"] = g["Mes"].map(MESES_ES)
    return g.sort_values("Mes")


def tabla_mensual_serie(df: pd.DataFrame) -> pd.DataFrame:
    """Media por combinación año-mes (serie temporal mensual)."""
    g = df.groupby(["Año", "Mes"]).agg(
        Media=("TMedia", "mean"),
        Max_media=("TMax", "mean"),
        Min_media=("TMin", "mean"),
        Precip_total=("Precip", "sum"),
        Dias=("TMax", "count"),
    ).reset_index()
    g["Etiqueta"] = g.apply(lambda r: f"{MESES_ES_ABR[int(r['Mes'])]} {int(r['Año'])}", axis=1)
    return g.sort_values(["Año", "Mes"])


def tabla_estacional(df: pd.DataFrame) -> pd.DataFrame:
    orden = ["Invierno", "Primavera", "Verano", "Otoño"]
    g = df.groupby("Estacion").agg(
        Media=("TMedia", "mean"),
        Max_media=("TMax", "mean"),
        Min_media=("TMin", "mean"),
        Max_absoluta=("TMax", "max"),
        Min_absoluta=("TMin", "min"),
        Precip_total=("Precip", "sum"),
    ).reindex(orden).reset_index()
    return g


def tabla_estacional_por_año(df: pd.DataFrame) -> pd.DataFrame:
    """Media estacional por año (para ver evolución estación a estación)."""

    def año_estacional(row):
        # El invierno "cruza" el año (Dic pertenece al invierno del año siguiente)
        if row["Estacion"] == "Invierno" and row["Mes"] == 12:
            return row["Año"] + 1
        return row["Año"]

    tmp = df.copy()
    tmp["AñoEstacional"] = tmp.apply(año_estacional, axis=1)
    g = tmp.groupby(["AñoEstacional", "Estacion"]).agg(
        Media=("TMedia", "mean"),
        Dias=("TMedia", "count"),
    ).reset_index()
    return g


def tabla_canicula(df: pd.DataFrame) -> pd.DataFrame:
    can = df[df["EnCanicula"]]
    g = can.groupby("Año").agg(
        Media=("TMedia", "mean"),
        Max_media=("TMax", "mean"),
        Max_absoluta=("TMax", "max"),
        Min_media=("TMin", "mean"),
        Dias=("TMax", "count"),
    ).reset_index()
    return g


def ranking_dias_calurosos(df: pd.DataFrame, n=N_RANKING) -> pd.DataFrame:
    r = df.nlargest(n, "TMax")[["Fecha", "TMax", "TMin", "TMedia"]].reset_index(drop=True)
    r.index = r.index + 1
    return r


def ranking_dias_frios(df: pd.DataFrame, n=N_RANKING) -> pd.DataFrame:
    r = df.nsmallest(n, "TMin")[["Fecha", "TMax", "TMin", "TMedia"]].reset_index(drop=True)
    r.index = r.index + 1
    return r


def ranking_meses_calurosos(df: pd.DataFrame, n=N_RANKING) -> pd.DataFrame:
    serie = tabla_mensual_serie(df)
    serie = serie[serie["Dias"] >= 20]  # exigir cobertura mínima del mes
    r = serie.nlargest(n, "Media")[["Etiqueta", "Media", "Max_media", "Min_media", "Dias"]].reset_index(drop=True)
    r.index = r.index + 1
    return r


def ranking_meses_frios(df: pd.DataFrame, n=N_RANKING) -> pd.DataFrame:
    serie = tabla_mensual_serie(df)
    serie = serie[serie["Dias"] >= 20]
    r = serie.nsmallest(n, "Media")[["Etiqueta", "Media", "Max_media", "Min_media", "Dias"]].reset_index(drop=True)
    r.index = r.index + 1
    return r


def ranking_años(df: pd.DataFrame) -> pd.DataFrame:
    an = tabla_anual(df)
    return an.sort_values("Media", ascending=False).reset_index(drop=True)


def indices_termicos_anuales(df: pd.DataFrame) -> pd.DataFrame:
    g = df.groupby("Año").apply(lambda x: pd.Series({
        "Dias_verano": (x["TMax"] >= UMBRAL_DIA_VERANO).sum(),
        "Dias_calurosos": (x["TMax"] >= UMBRAL_DIA_CALUROSO).sum(),
        "Dias_muy_calurosos": (x["TMax"] >= UMBRAL_DIA_MUY_CALUROSO).sum(),
        "Noches_tropicales": (x["TMin"] >= UMBRAL_NOCHE_TROPICAL).sum(),
        "Noches_torridas": (x["TMin"] >= UMBRAL_NOCHE_TORRIDA).sum(),
        "Dias_helada": (x["TMin"] <= UMBRAL_DIA_HELADA).sum(),
        "DTR_medio": x["DTR"].mean(),
    })).reset_index()
    return g


def racha_maxima_calor(df: pd.DataFrame, umbral=UMBRAL_DIA_MUY_CALUROSO):
    """Devuelve (longitud, fecha_inicio, fecha_fin) de la racha más larga de
    días consecutivos con TMax >= umbral."""
    d = df.sort_values("Fecha").reset_index(drop=True)
    cumple = (d["TMax"] >= umbral).values
    mejor_len, mejor_ini, mejor_fin = 0, None, None
    actual_len, actual_ini = 0, None
    for i, ok in enumerate(cumple):
        if ok:
            if actual_len == 0:
                actual_ini = d.loc[i, "Fecha"]
            actual_len += 1
            if actual_len > mejor_len:
                mejor_len = actual_len
                mejor_ini = actual_ini
                mejor_fin = d.loc[i, "Fecha"]
        else:
            actual_len = 0
    return mejor_len, mejor_ini, mejor_fin


def racha_maxima_sin_lluvia(df: pd.DataFrame, umbral_mm=0.1):
    d = df.sort_values("Fecha").reset_index(drop=True)
    seco = (d["Precip"].fillna(0) < umbral_mm).values
    mejor_len, mejor_ini, mejor_fin = 0, None, None
    actual_len, actual_ini = 0, None
    for i, ok in enumerate(seco):
        if ok:
            if actual_len == 0:
                actual_ini = d.loc[i, "Fecha"]
            actual_len += 1
            if actual_len > mejor_len:
                mejor_len = actual_len
                mejor_ini = actual_ini
                mejor_fin = d.loc[i, "Fecha"]
        else:
            actual_len = 0
    return mejor_len, mejor_ini, mejor_fin


def tendencia_lineal(anual: pd.DataFrame):
    """Regresión lineal simple de la temperatura media anual frente al año.
    Solo considera años con datos suficientes (>= 300 días) para no sesgar
    la tendencia con años parciales (primer y último año del histórico).
    Devuelve (pendiente_por_decada, r2, años_usados)."""
    completos = anual[anual["Dias"] >= 300]
    if len(completos) < 3:
        completos = anual  # si hay pocos años completos, usar todos igualmente
    x = completos["Año"].values.astype(float)
    y = completos["Media"].values.astype(float)
    if len(x) < 2:
        return None, None, completos["Año"].tolist()
    coef = np.polyfit(x, y, 1)
    pendiente_decada = coef[0] * 10
    y_pred = np.polyval(coef, x)
    ss_res = np.sum((y - y_pred) ** 2)
    ss_tot = np.sum((y - np.mean(y)) ** 2)
    r2 = 1 - ss_res / ss_tot if ss_tot > 0 else 0.0
    return pendiente_decada, r2, completos["Año"].tolist()


def precipitacion_anual(df: pd.DataFrame) -> pd.DataFrame:
    g = df.groupby("Año").agg(
        Precip_total=("Precip", "sum"),
        Dias_lluvia=("Precip", lambda s: (s.fillna(0) >= 1.0).sum()),
        Precip_max_dia=("Precip", "max"),
    ).reset_index()
    return g


# ==============================================================================
# GRÁFICOS
# ==============================================================================

def _guardar(fig, nombre, carpeta):
    ruta = os.path.join(carpeta, nombre)
    fig.savefig(ruta, dpi=170, bbox_inches="tight")
    plt.close(fig)
    return ruta


def grafico_evolucion_anual(anual: pd.DataFrame, pendiente, carpeta):
    fig, ax = plt.subplots(figsize=(9, 4.6))
    x = anual["Año"]
    ax.fill_between(x, anual["Min_media"], anual["Max_media"], color=COLOR_FONDO_BARRA, alpha=0.12,
                     label="Rango medio (min-max)")
    ax.plot(x, anual["Max_media"], color=COLOR_MAX, marker="o", markersize=4, linewidth=1.8, label="Máx. media")
    ax.plot(x, anual["Media"], color=COLOR_MEDIA, marker="o", markersize=4.5, linewidth=2.4, label="Media")
    ax.plot(x, anual["Min_media"], color=COLOR_MIN, marker="o", markersize=4, linewidth=1.8, label="Mín. media")

    if pendiente is not None:
        x_arr = x.values.astype(float)
        y_arr = anual["Media"].values.astype(float)
        coef = np.polyfit(x_arr, y_arr, 1)
        y_line = np.polyval(coef, x_arr)
        ax.plot(x, y_line, color="#333333", linestyle="--", linewidth=1.3,
                label=f"Tendencia ({pendiente:+.2f} °C/década)")

    ax.set_title("Evolución anual de la temperatura")
    ax.set_xlabel("Año")
    ax.set_ylabel("Temperatura (°C)")
    ax.xaxis.set_major_locator(mticker.MaxNLocator(integer=True))
    ax.legend(loc="upper left", fontsize=9, ncol=2)
    return _guardar(fig, "01_evolucion_anual.png", carpeta)


def grafico_climatologia_mensual(clima: pd.DataFrame, carpeta):
    fig, ax = plt.subplots(figsize=(9, 4.6))
    x = np.arange(len(clima))
    ancho = 0.27
    ax.bar(x - ancho, clima["Max_media"], width=ancho, color=COLOR_MAX, label="Máx. media")
    ax.bar(x, clima["Media"], width=ancho, color=COLOR_MEDIA, label="Media")
    ax.bar(x + ancho, clima["Min_media"], width=ancho, color=COLOR_MIN, label="Mín. media")
    ax.set_xticks(x)
    ax.set_xticklabels([MESES_ES_ABR[m] for m in clima["Mes"]])
    ax.set_title("Climatología mensual (media histórica por mes)")
    ax.set_ylabel("Temperatura (°C)")
    ax.legend(loc="upper left", fontsize=9, ncol=3)
    return _guardar(fig, "02_climatologia_mensual.png", carpeta)


def grafico_boxplot_mensual(df: pd.DataFrame, carpeta):
    fig, ax = plt.subplots(figsize=(9, 4.6))
    datos = [df.loc[df["Mes"] == m, "TMedia"].dropna().values for m in range(1, 13)]
    bp = ax.boxplot(datos, patch_artist=True, showfliers=False, widths=0.55)
    for patch in bp["boxes"]:
        patch.set_facecolor("#f0c179")
        patch.set_edgecolor("#8a5a12")
    for median in bp["medians"]:
        median.set_color("#8a2f1f")
        median.set_linewidth(1.8)
    ax.set_xticks(range(1, 13))
    ax.set_xticklabels([MESES_ES_ABR[m] for m in range(1, 13)])
    ax.set_title("Distribución de la temperatura media diaria por mes")
    ax.set_ylabel("Temperatura media diaria (°C)")
    return _guardar(fig, "03_distribucion_mensual.png", carpeta)


def grafico_estacional(estacional: pd.DataFrame, carpeta):
    fig, ax = plt.subplots(figsize=(8, 4.6))
    x = np.arange(len(estacional))
    ancho = 0.27
    ax.bar(x - ancho, estacional["Max_media"], width=ancho, color=COLOR_MAX, label="Máx. media")
    ax.bar(x, estacional["Media"], width=ancho, color=COLOR_MEDIA, label="Media")
    ax.bar(x + ancho, estacional["Min_media"], width=ancho, color=COLOR_MIN, label="Mín. media")
    ax.set_xticks(x)
    ax.set_xticklabels(estacional["Estacion"])
    ax.set_title("Temperatura media por estación del año")
    ax.set_ylabel("Temperatura (°C)")
    ax.legend(loc="upper left", fontsize=9, ncol=3)
    return _guardar(fig, "04_estacional.png", carpeta)


def grafico_estacional_evolucion(est_año: pd.DataFrame, carpeta):
    fig, ax = plt.subplots(figsize=(9, 4.6))
    colores = {"Invierno": COLOR_MIN, "Primavera": "#5aa469", "Verano": COLOR_MAX, "Otoño": "#c08a3e"}
    for est, color in colores.items():
        sub = est_año[est_año["Estacion"] == est].sort_values("AñoEstacional")
        sub = sub[sub["Dias"] >= 60]
        ax.plot(sub["AñoEstacional"], sub["Media"], marker="o", markersize=4, linewidth=1.8,
                color=color, label=est)
    ax.set_title("Evolución de la temperatura media por estación y año")
    ax.set_xlabel("Año")
    ax.set_ylabel("Temperatura media (°C)")
    ax.xaxis.set_major_locator(mticker.MaxNLocator(integer=True))
    ax.legend(loc="upper left", fontsize=9, ncol=4)
    return _guardar(fig, "05_estacional_evolucion.png", carpeta)


def grafico_canicula(canicula: pd.DataFrame, carpeta):
    fig, ax = plt.subplots(figsize=(8.5, 4.6))
    x = canicula["Año"].astype(int)
    ax.bar(x, canicula["Media"], color=COLOR_MEDIA, label="Media canícula", width=0.6)
    ax.plot(x, canicula["Max_absoluta"], color=COLOR_MAX, marker="o", markersize=5,
            linewidth=1.8, label="Máx. absoluta")
    ax.set_xticks(x)
    ini = f"{CANICULA_INICIO[1]:02d}/{CANICULA_INICIO[0]:02d}"
    fin = f"{CANICULA_FIN[1]:02d}/{CANICULA_FIN[0]:02d}"
    ax.set_title(f"Canícula ({ini} – {fin}): temperatura por año")
    ax.set_xlabel("Año")
    ax.set_ylabel("Temperatura (°C)")
    ax.legend(loc="lower right", fontsize=9)
    return _guardar(fig, "06_canicula.png", carpeta)


def grafico_ranking_dias(ranking: pd.DataFrame, carpeta, calurosos=True):
    fig, ax = plt.subplots(figsize=(8.5, 5.2))
    etiquetas = [f"{f.strftime('%d/%m/%Y')}" for f in ranking["Fecha"]]
    valores = ranking["TMax"] if calurosos else ranking["TMin"]
    color = COLOR_MAX if calurosos else COLOR_MIN
    y = np.arange(len(ranking))[::-1]
    ax.barh(y, valores, color=color)
    ax.set_yticks(y)
    ax.set_yticklabels(etiquetas, fontsize=9)
    for yi, v in zip(y, valores):
        ax.text(v, yi, f" {v:.1f}°C", va="center", fontsize=8.5)
    titulo = "Los días más calurosos" if calurosos else "Los días más fríos"
    ax.set_title(f"{titulo} del histórico (Top {len(ranking)})")
    ax.set_xlabel("Temperatura máxima (°C)" if calurosos else "Temperatura mínima (°C)")
    ax.margins(x=0.12)
    return _guardar(fig, "07_ranking_dias_calurosos.png" if calurosos else "08_ranking_dias_frios.png", carpeta)


def grafico_ranking_años(ranking_años_df: pd.DataFrame, carpeta):
    fig, ax = plt.subplots(figsize=(9, 5.4))
    r = ranking_años_df.sort_values("Media", ascending=True)
    y = np.arange(len(r))
    colores = plt.cm.RdYlBu_r(np.linspace(0.15, 0.95, len(r)))
    ax.barh(y, r["Media"], color=colores)
    ax.set_yticks(y)
    ax.set_yticklabels(r["Año"].astype(int))
    for yi, v, d in zip(y, r["Media"], r["Dias"]):
        marca = f" {v:.2f}°C" + ("  (parcial)" if d < 300 else "")
        ax.text(v, yi, marca, va="center", fontsize=8.5)
    ax.set_title("Ranking de años según temperatura media anual")
    ax.set_xlabel("Temperatura media (°C)")
    ax.margins(x=0.18)
    return _guardar(fig, "09_ranking_años.png", carpeta)


def grafico_indices_termicos(indices: pd.DataFrame, carpeta):
    fig, ax = plt.subplots(figsize=(9, 4.8))
    x = indices["Año"].astype(int)
    ax.bar(x, indices["Dias_verano"], color="#f2c14e", label="Días de verano (≥25°C)")
    ax.bar(x, indices["Dias_calurosos"], color="#e8873a", label="Días cálidos (≥30°C)")
    ax.bar(x, indices["Dias_muy_calurosos"], color="#c0392b", label="Días muy cálidos (≥35°C)")
    ax.plot(x, indices["Noches_tropicales"], color="#2c3e50", marker="o", markersize=4,
            linewidth=1.8, label="Noches tropicales (≥20°C)")
    ax.set_xticks(x)
    ax.set_title("Índices térmicos por año")
    ax.set_xlabel("Año")
    ax.set_ylabel("Nº de días")
    ax.legend(loc="upper left", fontsize=8.5, ncol=2)
    return _guardar(fig, "10_indices_termicos.png", carpeta)


def grafico_precipitacion_anual(precip: pd.DataFrame, carpeta):
    fig, ax = plt.subplots(figsize=(8.5, 4.6))
    x = precip["Año"].astype(int)
    ax.bar(x, precip["Precip_total"], color=COLOR_PRECIP)
    ax.set_xticks(x)
    ax.set_title("Precipitación total anual")
    ax.set_xlabel("Año")
    ax.set_ylabel("Precipitación (mm)")
    for xi, v in zip(x, precip["Precip_total"]):
        ax.text(xi, v, f"{v:.0f}", ha="center", va="bottom", fontsize=8.5)
    return _guardar(fig, "11_precipitacion_anual.png", carpeta)


# ==============================================================================
# GENERACIÓN DEL INFORME WORD
# ==============================================================================

from docx import Document
from docx.shared import Pt, Cm, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

FUENTE = "Arial"
COLOR_TITULO = RGBColor(0x1F, 0x3A, 0x5F)
COLOR_ACENTO = RGBColor(0xB0, 0x2E, 0x21)


def _set_estilos_base(doc):
    normal = doc.styles["Normal"]
    normal.font.name = FUENTE
    normal.font.size = Pt(10.5)
    rpr = normal.element.get_or_add_rPr()
    rFonts = rpr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rpr.append(rFonts)
    rFonts.set(qn('w:eastAsia'), FUENTE)

    for nombre, tam, color, negrita in [
        ("Title", 26, COLOR_TITULO, True),
        ("Heading 1", 17, COLOR_TITULO, True),
        ("Heading 2", 13.5, COLOR_ACENTO, True),
        ("Heading 3", 11.5, RGBColor(0x33, 0x33, 0x33), True),
    ]:
        try:
            st = doc.styles[nombre]
            st.font.name = FUENTE
            st.font.size = Pt(tam)
            st.font.color.rgb = color
            st.font.bold = negrita
        except KeyError:
            pass


def _add_footer_pagina(doc):
    section = doc.sections[0]
    footer = section.footer
    p = footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run()
    fld_char1 = OxmlElement('w:fldChar')
    fld_char1.set(qn('w:fldCharType'), 'begin')
    instr = OxmlElement('w:instrText')
    instr.set(qn('xml:space'), 'preserve')
    instr.text = 'PAGE'
    fld_char2 = OxmlElement('w:fldChar')
    fld_char2.set(qn('w:fldCharType'), 'end')
    run._r.append(fld_char1)
    run._r.append(instr)
    run._r.append(fld_char2)
    run.font.size = Pt(9)
    run.font.color.rgb = RGBColor(0x88, 0x88, 0x88)


def _tabla(doc, encabezados, filas, anchos_cm=None, alineaciones=None):
    tabla = doc.add_table(rows=1, cols=len(encabezados))
    tabla.alignment = WD_TABLE_ALIGNMENT.CENTER
    tabla.style = "Light Grid Accent 1"

    hdr = tabla.rows[0].cells
    for i, h in enumerate(encabezados):
        hdr[i].text = ""
        p = hdr[i].paragraphs[0]
        run = p.add_run(h)
        run.bold = True
        run.font.size = Pt(9.5)
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _sombrear_celda(hdr[i], "2F5B8A")

    for fila in filas:
        celdas = tabla.add_row().cells
        for i, val in enumerate(fila):
            celdas[i].text = ""
            p = celdas[i].paragraphs[0]
            run = p.add_run(str(val))
            run.font.size = Pt(9.5)
            if alineaciones and i < len(alineaciones):
                p.alignment = alineaciones[i]
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER if i > 0 else WD_ALIGN_PARAGRAPH.LEFT

    if anchos_cm:
        for row in tabla.rows:
            for i, w in enumerate(anchos_cm):
                row.cells[i].width = Cm(w)
    return tabla


def _sombrear_celda(celda, hex_color):
    tcPr = celda._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    tcPr.append(shd)


def _add_imagen(doc, ruta, ancho_cm=16):
    doc.add_picture(ruta, width=Cm(ancho_cm))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER


def _add_nota(doc, texto):
    p = doc.add_paragraph()
    run = p.add_run(texto)
    run.italic = True
    run.font.size = Pt(9)
    run.font.color.rgb = RGBColor(0x70, 0x70, 0x70)
    return p


def generar_informe(df, ruta_salida, carpeta_imgs, ruta_entrada):
    cobertura = resumen_cobertura(df)
    anual = tabla_anual(df)
    clima = tabla_mensual_climatologia(df)
    serie_mensual = tabla_mensual_serie(df)
    estacional = tabla_estacional(df)
    est_año = tabla_estacional_por_año(df)
    canicula = tabla_canicula(df)
    ranking_calurosos = ranking_dias_calurosos(df)
    ranking_frios = ranking_dias_frios(df)
    ranking_meses_c = ranking_meses_calurosos(df)
    ranking_meses_f = ranking_meses_frios(df)
    ranking_an = ranking_años(df)
    indices = indices_termicos_anuales(df)
    precip = precipitacion_anual(df)
    pendiente, r2, años_tendencia = tendencia_lineal(anual)
    racha_len, racha_ini, racha_fin = racha_maxima_calor(df)
    seca_len, seca_ini, seca_fin = racha_maxima_sin_lluvia(df)

    # --- Gráficos ---
    img1 = grafico_evolucion_anual(anual, pendiente, carpeta_imgs)
    img2 = grafico_climatologia_mensual(clima, carpeta_imgs)
    img3 = grafico_boxplot_mensual(df, carpeta_imgs)
    img4 = grafico_estacional(estacional, carpeta_imgs)
    img5 = grafico_estacional_evolucion(est_año, carpeta_imgs)
    img6 = grafico_canicula(canicula, carpeta_imgs) if len(canicula) else None
    img7 = grafico_ranking_dias(ranking_calurosos, carpeta_imgs, calurosos=True)
    img8 = grafico_ranking_dias(ranking_frios, carpeta_imgs, calurosos=False)
    img9 = grafico_ranking_años(ranking_an, carpeta_imgs)
    img10 = grafico_indices_termicos(indices, carpeta_imgs)
    img11 = grafico_precipitacion_anual(precip, carpeta_imgs)

    doc = Document()
    section = doc.sections[0]
    section.page_width = Cm(21)
    section.page_height = Cm(29.7)
    section.left_margin = Cm(2.2)
    section.right_margin = Cm(2.2)
    section.top_margin = Cm(2.0)
    section.bottom_margin = Cm(2.0)
    _set_estilos_base(doc)
    _add_footer_pagina(doc)

    # ---------------- Portada ----------------
    doc.add_paragraph().paragraph_format.space_before = Pt(40)
    titulo = doc.add_paragraph()
    titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = titulo.add_run("Informe de Índices de Temperatura")
    run.font.size = Pt(28)
    run.font.bold = True
    run.font.color.rgb = COLOR_TITULO

    subt = doc.add_paragraph()
    subt.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = subt.add_run("Historial de la Estación Meteorológica")
    run.font.size = Pt(16)
    run.font.color.rgb = COLOR_ACENTO

    doc.add_paragraph()
    linea = doc.add_paragraph()
    linea.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = linea.add_run(
        f"Periodo analizado: {cobertura['fecha_ini'].strftime('%d/%m/%Y')} — "
        f"{cobertura['fecha_fin'].strftime('%d/%m/%Y')}"
    )
    run.font.size = Pt(12)

    linea2 = doc.add_paragraph()
    linea2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = linea2.add_run(
        f"{cobertura['dias_con_dato']:,} días con datos "
        f"({cobertura['pct_cobertura']:.1f}% de cobertura del periodo)".replace(",", ".")
    )
    run.font.size = Pt(11)
    run.font.color.rgb = RGBColor(0x60, 0x60, 0x60)

    doc.add_paragraph()
    fecha_gen = doc.add_paragraph()
    fecha_gen.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = fecha_gen.add_run(f"Informe generado automáticamente el {datetime.now().strftime('%d/%m/%Y a las %H:%M')}")
    run.font.size = Pt(9.5)
    run.italic = True
    run.font.color.rgb = RGBColor(0x88, 0x88, 0x88)

    fuente_p = doc.add_paragraph()
    fuente_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = fuente_p.add_run(f"Fuente de datos: {os.path.basename(ruta_entrada)}")
    run.font.size = Pt(9.5)
    run.italic = True
    run.font.color.rgb = RGBColor(0x88, 0x88, 0x88)

    doc.add_page_break()

    # ---------------- 1. Resumen general ----------------
    doc.add_heading("1. Resumen general", level=1)
    doc.add_paragraph(
        "Este informe resume el comportamiento térmico y pluviométrico registrado por la "
        "estación meteorológica personal a lo largo de todo el histórico disponible. "
        "Incluye medias, máximas y mínimas a distintas escalas temporales (diaria, mensual, "
        "estacional y anual), un análisis específico del periodo de canícula, rankings de los "
        "días, meses y años más calurosos, y una serie de índices climáticos adicionales "
        "(días de verano, noches tropicales, olas de calor, tendencia de calentamiento, etc.)."
    )

    t_max_abs = df.loc[df["TMax"].idxmax()]
    t_min_abs = df.loc[df["TMin"].idxmin()]

    filas_resumen = [
        ("Temperatura media del periodo", f"{df['TMedia'].mean():.1f} °C"),
        ("Temperatura máxima absoluta", f"{t_max_abs['TMax']:.1f} °C  ({t_max_abs['Fecha'].strftime('%d/%m/%Y')})"),
        ("Temperatura mínima absoluta", f"{t_min_abs['TMin']:.1f} °C  ({t_min_abs['Fecha'].strftime('%d/%m/%Y')})"),
        ("Amplitud térmica diaria media", f"{df['DTR'].mean():.1f} °C"),
        ("Precipitación total registrada", f"{df['Precip'].sum():,.0f} mm".replace(",", ".")),
    ]
    ranking_an_completos = ranking_an[ranking_an["Dias"] >= 300]
    if len(ranking_an_completos):
        filas_resumen.append(("Año más cálido (años completos)",
                               f"{int(ranking_an_completos.iloc[0]['Año'])}  ({ranking_an_completos.iloc[0]['Media']:.2f} °C de media)"))
        filas_resumen.append(("Año más frío (años completos)",
                               f"{int(ranking_an_completos.iloc[-1]['Año'])}  ({ranking_an_completos.iloc[-1]['Media']:.2f} °C de media)"))
    if pendiente is not None:
        filas_resumen.append(("Tendencia de temperatura media", f"{pendiente:+.2f} °C / década  (R² = {r2:.2f})"))
    _tabla(doc, ["Indicador", "Valor"], filas_resumen, anchos_cm=[9, 8])

    if cobertura["dias_faltantes"] > 0:
        _add_nota(
            doc,
            f"Nota sobre cobertura de datos: existen {cobertura['dias_faltantes']} días sin registro dentro "
            f"del periodo analizado (fallos de conexión o de estación). Los cálculos se han realizado "
            f"únicamente sobre los días con datos disponibles."
        )
    if len(anual[anual["Dias"] < 300]):
        _add_nota(
            doc,
            "Los años marcados como parciales (menos de 300 días con datos, típicamente el primer y "
            "el último año del histórico) se muestran en las tablas a título informativo, pero se "
            "excluyen de los indicadores de \"año más cálido/frío\" y del cálculo de la tendencia, ya "
            "que al no cubrir todos los meses del año su media no es directamente comparable con la "
            "de un año completo."
        )

    # ---------------- 2. Evolución anual ----------------
    doc.add_heading("2. Evolución anual de la temperatura", level=1)
    doc.add_paragraph(
        "El siguiente gráfico muestra la evolución año a año de la temperatura media, así como "
        "la media de las máximas y de las mínimas diarias. La línea discontinua representa la "
        "tendencia lineal calculada sobre los años con cobertura de datos suficiente (≥300 días)."
    )
    _add_imagen(doc, img1)

    filas_anual = []
    for _, r in anual.iterrows():
        marca = " *" if r["Dias"] < 300 else ""
        filas_anual.append((
            f"{int(r['Año'])}{marca}", f"{r['Media']:.1f}", f"{r['Max_media']:.1f}", f"{r['Min_media']:.1f}",
            f"{r['Max_absoluta']:.1f}", f"{r['Min_absoluta']:.1f}", f"{r['Precip_total']:.0f}", f"{int(r['Dias'])}",
        ))
    doc.add_paragraph()
    _tabla(
        doc,
        ["Año", "T. media", "Máx. media", "Mín. media", "Máx. absol.", "Mín. absol.", "Precip. (mm)", "Días"],
        filas_anual,
        anchos_cm=[2.0, 2.0, 2.2, 2.2, 2.2, 2.2, 2.4, 1.6],
    )
    _add_nota(doc, "* Año con menos de 300 días de datos disponibles (año parcial, al inicio o al final del histórico).")

    # ---------------- 3. Climatología mensual ----------------
    doc.add_heading("3. Climatología mensual", level=1)
    doc.add_paragraph(
        "Media histórica de cada mes del año, calculada agregando todos los años disponibles. "
        "Permite identificar el patrón estacional típico del emplazamiento de la estación."
    )
    _add_imagen(doc, img2)
    doc.add_paragraph()
    _add_imagen(doc, img3)
    _add_nota(doc, "El diagrama de caja muestra la dispersión de las temperaturas medias diarias dentro de cada mes (caja = 25%-75%, línea = mediana; se excluyen valores atípicos extremos del trazado).")

    filas_clima = []
    for _, r in clima.iterrows():
        filas_clima.append((
            r["MesNombre"], f"{r['Media']:.1f}", f"{r['Max_media']:.1f}", f"{r['Min_media']:.1f}",
            f"{r['Max_absoluta']:.1f}", f"{r['Min_absoluta']:.1f}", f"{r['Precip_media']:.0f}",
        ))
    doc.add_paragraph()
    _tabla(
        doc,
        ["Mes", "T. media", "Máx. media", "Mín. media", "Máx. absol.", "Mín. absol.", "Precip. media (mm)"],
        filas_clima,
        anchos_cm=[3.0, 2.2, 2.2, 2.2, 2.2, 2.2, 2.6],
    )

    # ---------------- 4. Análisis estacional ----------------
    doc.add_heading("4. Análisis por estación del año", level=1)
    doc.add_paragraph(
        "Las estaciones se han delimitado de forma astronómica (equinoccios y solsticios). "
        "La tabla y el gráfico muestran las temperaturas típicas de cada estación en el conjunto "
        "del histórico, y el gráfico posterior muestra su evolución año a año."
    )
    _add_imagen(doc, img4)

    filas_est = []
    for _, r in estacional.iterrows():
        filas_est.append((
            r["Estacion"], f"{r['Media']:.1f}", f"{r['Max_media']:.1f}", f"{r['Min_media']:.1f}",
            f"{r['Max_absoluta']:.1f}", f"{r['Min_absoluta']:.1f}", f"{r['Precip_total']:.0f}",
        ))
    doc.add_paragraph()
    _tabla(
        doc,
        ["Estación", "T. media", "Máx. media", "Mín. media", "Máx. absol.", "Mín. absol.", "Precip. total (mm)"],
        filas_est,
        anchos_cm=[2.6, 2.2, 2.2, 2.2, 2.2, 2.2, 2.8],
    )

    doc.add_paragraph()
    doc.add_paragraph("Evolución de la temperatura media por estación a lo largo de los años:")
    _add_imagen(doc, img5)

    # ---------------- 5. Canícula ----------------
    doc.add_heading("5. Canícula", level=1)
    ini_txt = f"{CANICULA_INICIO[1]:02d}/{CANICULA_INICIO[0]:02d}"
    fin_txt = f"{CANICULA_FIN[1]:02d}/{CANICULA_FIN[0]:02d}"
    doc.add_paragraph(
        f"Se ha definido la canícula como el periodo comprendido entre el {ini_txt} y el {fin_txt} "
        f"de cada año (parámetro configurable en el script), correspondiente tradicionalmente a los "
        f"días más calurosos del verano."
    )
    if img6:
        _add_imagen(doc, img6)
        filas_can = []
        for _, r in canicula.iterrows():
            filas_can.append((
                int(r["Año"]), f"{r['Media']:.1f}", f"{r['Max_media']:.1f}", f"{r['Max_absoluta']:.1f}",
                f"{r['Min_media']:.1f}", int(r["Dias"]),
            ))
        doc.add_paragraph()
        _tabla(
            doc,
            ["Año", "T. media", "Máx. media", "Máx. absol.", "Mín. media", "Días con datos"],
            filas_can,
            anchos_cm=[2.2, 2.4, 2.6, 2.6, 2.6, 3.0],
        )
        if len(canicula):
            fila_top = canicula.loc[canicula["Media"].idxmax()]
            doc.add_paragraph()
            doc.add_paragraph(
                f"La canícula más calurosa registrada fue la de {int(fila_top['Año'])}, con una "
                f"temperatura media de {fila_top['Media']:.1f} °C y una máxima absoluta de "
                f"{fila_top['Max_absoluta']:.1f} °C."
            )
    else:
        _add_nota(doc, "No hay datos suficientes en el histórico para analizar el periodo de canícula.")

    # ---------------- 6. Rankings ----------------
    doc.add_heading("6. Rankings de temperatura", level=1)

    doc.add_heading("6.1 Días más calurosos", level=2)
    _add_imagen(doc, img7)
    filas_rc = [(i, f["Fecha"].strftime("%d/%m/%Y"), f"{f['TMax']:.1f}", f"{f['TMin']:.1f}", f"{f['TMedia']:.1f}")
                for i, f in ranking_calurosos.iterrows()]
    doc.add_paragraph()
    _tabla(doc, ["#", "Fecha", "T. máx.", "T. mín.", "T. media"], filas_rc, anchos_cm=[1.2, 3.5, 2.5, 2.5, 2.5])

    doc.add_heading("6.2 Días más fríos", level=2)
    _add_imagen(doc, img8)
    filas_rf = [(i, f["Fecha"].strftime("%d/%m/%Y"), f"{f['TMax']:.1f}", f"{f['TMin']:.1f}", f"{f['TMedia']:.1f}")
                for i, f in ranking_frios.iterrows()]
    doc.add_paragraph()
    _tabla(doc, ["#", "Fecha", "T. máx.", "T. mín.", "T. media"], filas_rf, anchos_cm=[1.2, 3.5, 2.5, 2.5, 2.5])

    doc.add_heading("6.3 Meses más calurosos y más fríos", level=2)
    doc.add_paragraph("Se consideran únicamente los meses con al menos 20 días de datos registrados.")
    filas_mc = [(i, r["Etiqueta"], f"{r['Media']:.1f}", f"{r['Max_media']:.1f}", f"{r['Min_media']:.1f}")
                for i, r in ranking_meses_c.iterrows()]
    doc.add_paragraph("Meses más calurosos:")
    _tabla(doc, ["#", "Mes", "T. media", "Máx. media", "Mín. media"], filas_mc, anchos_cm=[1.2, 4.0, 2.6, 2.6, 2.6])

    doc.add_paragraph()
    filas_mf = [(i, r["Etiqueta"], f"{r['Media']:.1f}", f"{r['Max_media']:.1f}", f"{r['Min_media']:.1f}")
                for i, r in ranking_meses_f.iterrows()]
    doc.add_paragraph("Meses más fríos:")
    _tabla(doc, ["#", "Mes", "T. media", "Máx. media", "Mín. media"], filas_mf, anchos_cm=[1.2, 4.0, 2.6, 2.6, 2.6])

    doc.add_heading("6.4 Ranking de años", level=2)
    _add_imagen(doc, img9)

    # ---------------- 7. Índices térmicos adicionales ----------------
    doc.add_heading("7. Índices térmicos adicionales", level=1)
    doc.add_paragraph(
        "Además de las medias y extremos habituales, se han calculado los siguientes índices "
        "climáticos estándar, de utilidad para caracterizar la severidad térmica de cada año:"
    )
    bullets = [
        f"Día de verano: máxima ≥ {UMBRAL_DIA_VERANO:.0f} °C.",
        f"Día cálido: máxima ≥ {UMBRAL_DIA_CALUROSO:.0f} °C.",
        f"Día muy cálido / ola de calor: máxima ≥ {UMBRAL_DIA_MUY_CALUROSO:.0f} °C.",
        f"Noche tropical: mínima ≥ {UMBRAL_NOCHE_TROPICAL:.0f} °C.",
        f"Noche tórrida: mínima ≥ {UMBRAL_NOCHE_TORRIDA:.0f} °C.",
        f"Día de helada: mínima ≤ {UMBRAL_DIA_HELADA:.0f} °C.",
    ]
    for b in bullets:
        doc.add_paragraph(b, style="List Bullet")

    _add_imagen(doc, img10)

    filas_idx = []
    for _, r in indices.iterrows():
        filas_idx.append((
            int(r["Año"]), int(r["Dias_verano"]), int(r["Dias_calurosos"]), int(r["Dias_muy_calurosos"]),
            int(r["Noches_tropicales"]), int(r["Noches_torridas"]), int(r["Dias_helada"]), f"{r['DTR_medio']:.1f}",
        ))
    doc.add_paragraph()
    _tabla(
        doc,
        ["Año", "Días verano", "Días cálidos", "Días muy cálidos", "Noches trop.", "Noches tórridas", "Días helada", "DTR medio"],
        filas_idx,
        anchos_cm=[1.6, 2.0, 2.0, 2.2, 2.0, 2.2, 2.0, 2.0],
    )

    doc.add_paragraph()
    if racha_ini is not None:
        doc.add_paragraph(
            f"Racha de calor más larga: {racha_len} días consecutivos con máxima ≥ "
            f"{UMBRAL_DIA_MUY_CALUROSO:.0f} °C, del {racha_ini.strftime('%d/%m/%Y')} al "
            f"{racha_fin.strftime('%d/%m/%Y')}."
        )
    if pendiente is not None:
        doc.add_paragraph(
            f"Tendencia de calentamiento/enfriamiento: {pendiente:+.2f} °C por década "
            f"(regresión lineal sobre la temperatura media anual, R² = {r2:.2f}), calculada con "
            f"los años de cobertura suficiente."
        )

    # ---------------- 8. Precipitación ----------------
    doc.add_heading("8. Precipitación", level=1)
    doc.add_paragraph(
        "Aunque el foco del informe son los índices de temperatura, se incluye un resumen "
        "pluviométrico anual como contexto climático adicional."
    )
    _add_imagen(doc, img11)

    filas_precip = []
    for _, r in precip.iterrows():
        filas_precip.append((int(r["Año"]), f"{r['Precip_total']:.0f}", int(r["Dias_lluvia"]), f"{r['Precip_max_dia']:.1f}"))
    doc.add_paragraph()
    _tabla(doc, ["Año", "Precipitación total (mm)", "Días de lluvia (≥1 mm)", "Máx. en un día (mm)"], filas_precip,
           anchos_cm=[2.5, 4.5, 4.5, 4.5])

    if seca_ini is not None:
        doc.add_paragraph()
        doc.add_paragraph(
            f"Racha más larga sin lluvia apreciable: {seca_len} días consecutivos, del "
            f"{seca_ini.strftime('%d/%m/%Y')} al {seca_fin.strftime('%d/%m/%Y')}."
        )

    # ---------------- Cierre ----------------
    doc.add_paragraph()
    _add_nota(
        doc,
        "Informe generado automáticamente a partir del histórico exportado por la estación "
        "meteorológica. Para actualizarlo, sustituya el archivo Excel de origen por su versión "
        "más reciente y vuelva a ejecutar el script analisis_meteorologico.py."
    )

    doc.save(ruta_salida)
    return ruta_salida


# ==============================================================================
# PROGRAMA PRINCIPAL
# ==============================================================================

def main():
    ruta_entrada = sys.argv[1] if len(sys.argv) > 1 else ARCHIVO_ENTRADA_DEFECTO
    ruta_salida = sys.argv[2] if len(sys.argv) > 2 else ARCHIVO_SALIDA_DEFECTO

    print(f"→ Leyendo datos de: {ruta_entrada}")
    df = cargar_datos(ruta_entrada)
    cobertura = resumen_cobertura(df)
    print(f"  Periodo: {cobertura['fecha_ini'].date()} a {cobertura['fecha_fin'].date()}  "
          f"({cobertura['dias_con_dato']} días con datos, {cobertura['pct_cobertura']:.1f}% cobertura)")

    os.makedirs(CARPETA_IMAGENES, exist_ok=True)

    print("→ Calculando índices y generando gráficos...")
    ruta_final = generar_informe(df, ruta_salida, CARPETA_IMAGENES, ruta_entrada)

    print(f"✔ Informe generado correctamente en: {os.path.abspath(ruta_final)}")


if __name__ == "__main__":
    main()
