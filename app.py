# -*- coding: utf-8 -*-
"""
==============================================================================
 app.py — Estación Los Tomillares (ITORRE77): panel en vivo
==============================================================================
App de Streamlit que muestra, sin depender de que nadie la genere a mano,
los indicadores de temperatura y de peligro de incendio de la estación,
actualizados hasta el día anterior a la consulta.

No duplica lógica de cálculo: reutiliza tal cual las funciones de
`analisis_meteorologico.py` y `informe_incendios.py` (los mismos scripts
que generan los informes en Word), así que cualquier corrección o mejora
que se haga allí se refleja aquí automáticamente.

DESPLIEGUE: ver README.md para los pasos en Streamlit Community Cloud.
==============================================================================
"""

import os
import tempfile
import datetime as dt

import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt

import data_fetch
import analisis_meteorologico as am
import informe_incendios as fi

st.set_page_config(
    page_title="Estación Los Tomillares — Panel en vivo",
    page_icon="🌡️",
    layout="wide",
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
HISTORICO_BASE = os.path.join(BASE_DIR, "historico_base.xlsx")

# Cambia esto a False antes de desplegar si NO quieres que la dirección
# exacta de la estación aparezca en una página pública (ver nota en el README).
MOSTRAR_DIRECCION = os.getenv("MOSTRAR_DIRECCION", "1") not in ("0", "false", "False")


# ------------------------------------------------------------------
# Carga y cálculo de datos (con caché de 24h: no golpea la API de
# Ecowitt en cada visita, solo una vez al día como máximo)
# ------------------------------------------------------------------
@st.cache_data(ttl=24 * 3600, show_spinner="Actualizando datos de la estación…")
def cargar_todo():
    with tempfile.TemporaryDirectory() as tmp:
        ruta_combinada = os.path.join(tmp, "combinado.xlsx")
        resumen = data_fetch.construir_dataset_actualizado(HISTORICO_BASE, ruta_combinada)

        df_temp = am.cargar_datos(ruta_combinada)
        df_fuego = fi.load_data(ruta_combinada)

    return df_temp, df_fuego, resumen


def metric_card(col, label, value, help_text=None, delta=None):
    col.metric(label, value, delta=delta, help=help_text)


# ------------------------------------------------------------------
# Cabecera
# ------------------------------------------------------------------
st.title("🌡️ Estación Meteorológica Los Tomillares (ITORRE77)")
if MOSTRAR_DIRECCION:
    st.caption("c/Jaral, 23, Los Tomillares, Torremocha de Jarama · "
               "wunderground.com/dashboard/pws/ITORRE77")
else:
    st.caption("Torremocha de Jarama · wunderground.com/dashboard/pws/ITORRE77")

try:
    df_temp, df_fuego, resumen = cargar_todo()
except Exception as e:
    st.error(f"No se han podido cargar los datos: {e}")
    st.stop()

# Aviso de estado de la actualización diaria
ultimo_dato = df_fuego["Fecha"].max().date()
if resumen["dias_nuevos"] > 0:
    st.success(f"✅ Datos actualizados hasta el {ultimo_dato:%d/%m/%Y} "
               f"({resumen['dias_nuevos']} día(s) incorporado(s) desde la API de Ecowitt).")
elif resumen["credenciales"]:
    st.info(f"ℹ️ Datos al día: el histórico ya llegaba hasta el {ultimo_dato:%d/%m/%Y}.")
else:
    st.warning(f"⚠️ No hay credenciales de Ecowitt configuradas todavía — mostrando el "
               f"histórico base, actualizado hasta el {ultimo_dato:%d/%m/%Y}. "
               f"Añade ECOWITT_APPLICATION_KEY / ECOWITT_API_KEY / ECOWITT_MAC en Secrets "
               f"para que se actualice sola cada día.")

st.caption(f"Última comprobación: {dt.datetime.now():%d/%m/%Y %H:%M}")

tab_temp, tab_fuego = st.tabs(["🌡️ Temperatura", "🔥 Peligro de incendio y sequía"])


# ==============================================================================
# PESTAÑA 1 — TEMPERATURA
# ==============================================================================
with tab_temp:
    dias = st.slider("Días a analizar", 30, min(2200, len(df_temp)), 365, step=5, key="dias_temp")
    d = df_temp.tail(dias).copy()

    anual = am.tabla_anual(d)
    ranking_an = am.ranking_años(d)
    # La tendencia a largo plazo se calcula SIEMPRE sobre el histórico completo,
    # no sobre la ventana de días seleccionada: con una ventana corta (p.ej. 365
    # días arbitrarios que no coinciden con años naturales completos) la
    # regresión pierde sentido — el ajuste puede dar un R² artificialmente
    # perfecto sobre 1-2 años parciales.
    pendiente, r2, _ = am.tendencia_lineal(am.tabla_anual(df_temp))

    c1, c2, c3, c4 = st.columns(4)
    metric_card(c1, "Temperatura media del periodo", f"{d['TMedia'].mean():.1f} °C")
    fila_max = d.loc[d["TMax"].idxmax()]
    metric_card(c2, "Máxima absoluta", f"{fila_max['TMax']:.1f} °C", f"el {fila_max['Fecha']:%d/%m/%Y}")
    fila_min = d.loc[d["TMin"].idxmin()]
    metric_card(c3, "Mínima absoluta", f"{fila_min['TMin']:.1f} °C", f"el {fila_min['Fecha']:%d/%m/%Y}")
    if pendiente is not None:
        metric_card(c4, "Tendencia (histórico completo)", f"{pendiente:+.2f} °C/década", f"R² = {r2:.2f}")

    st.subheader("Evolución de la temperatura")
    fig, ax = plt.subplots(figsize=(11, 4))
    ax.plot(d["Fecha"], d["TMax"], color="#d64545", linewidth=0.8, alpha=0.7, label="Máxima diaria")
    ax.plot(d["Fecha"], d["TMedia"], color="#e8952e", linewidth=1.1, label="Media diaria")
    ax.plot(d["Fecha"], d["TMin"], color="#3773b5", linewidth=0.8, alpha=0.7, label="Mínima diaria")
    ax.legend(loc="upper left", fontsize=9)
    ax.set_ylabel("°C")
    ax.grid(alpha=0.3)
    st.pyplot(fig, use_container_width=True)

    col_a, col_b = st.columns(2)
    with col_a:
        st.subheader("Ranking de días más calurosos")
        st.dataframe(am.ranking_dias_calurosos(d, 10), use_container_width=True, hide_index=True)
    with col_b:
        st.subheader("Ranking de años")
        st.dataframe(
            ranking_an[["Año", "Media", "Max_absoluta", "Min_absoluta", "Dias"]]
            .rename(columns={"Media": "T. media", "Max_absoluta": "Máx. absol.",
                              "Min_absoluta": "Mín. absol.", "Dias": "Días"}),
            use_container_width=True, hide_index=True,
        )

    st.subheader("Climatología mensual")
    clima = am.tabla_mensual_climatologia(d)
    fig2, ax2 = plt.subplots(figsize=(11, 4))
    ax2.bar(clima["MesNombre"], clima["Max_media"], color="#d64545", alpha=0.85, label="Máx. media")
    ax2.bar(clima["MesNombre"], clima["Media"], color="#e8952e", alpha=0.85, label="Media")
    ax2.bar(clima["MesNombre"], clima["Min_media"], color="#3773b5", alpha=0.85, label="Mín. media")
    ax2.legend(fontsize=9)
    plt.xticks(rotation=30)
    st.pyplot(fig2, use_container_width=True)


# ==============================================================================
# PESTAÑA 2 — INCENDIOS Y SEQUÍA
# ==============================================================================
with tab_fuego:
    dias_f = st.slider("Días a analizar", 30, min(2200, len(df_fuego)), 90, step=5, key="dias_fuego")

    df_fuego = df_fuego.copy()
    df_fuego["ET0_mm"] = fi.calc_et0(df_fuego)
    df_fuego["KBDI"] = fi.calculate_kbdi(df_fuego, fi.media_anios_completos(df_fuego)[0] or 478.275)
    df_fuego["FWI"] = fi.run_fwi_system(df_fuego)
    df_fuego["Angstrom"] = [fi.calc_angstrom(r["Temperatura Máxima (°C)"], r["Humidity_Min (%)"])
                             for _, r in df_fuego.iterrows()]
    df_fuego["Regla_30"] = df_fuego.apply(fi.eval_regla_30, axis=1)

    dfa = df_fuego.tail(dias_f)
    ultimo = dfa.iloc[-1]

    c1, c2, c3, c4, c5 = st.columns(5)
    metric_card(c1, "KBDI (sequía)", f"{ultimo['KBDI']:.0f}", fi.kbdi_label(ultimo["KBDI"]))
    metric_card(c2, "FWI (propagación)", f"{ultimo['FWI']:.0f}", fi.fwi_label(ultimo["FWI"]))
    metric_card(c3, "Ångström", f"{ultimo['Angstrom']:.2f}", fi.angstrom_label(ultimo["Angstrom"]))
    metric_card(c4, "T. máxima ayer", f"{ultimo['Temperatura Máxima (°C)']:.1f} °C")
    metric_card(c5, "Racha máx. ayer", f"{ultimo['wind_gust']:.0f} km/h")

    nivel = ultimo["Regla_30"]
    color_alerta = {"Extremo (3/3)": "🔴", "Alerta (2/3)": "🟠"}.get(nivel, "🟡")
    st.markdown(f"### {color_alerta} Regla 30-30-30: **{nivel}**")

    st.subheader("Evolución del KBDI (sequía) y el FWI (peligro de propagación)")
    fig3, (axk, axf) = plt.subplots(2, 1, figsize=(11, 6), sharex=True)
    axk.plot(dfa["Fecha"], dfa["KBDI"], color="#c0392b", linewidth=1.5)
    axk.axhline(400, color="#888", linestyle="--", linewidth=0.8)
    axk.set_ylabel("KBDI")
    axk.grid(alpha=0.3)
    axf.plot(dfa["Fecha"], dfa["FWI"], color="#8e44ad", linewidth=1.5)
    axf.axhline(50, color="#888", linestyle="--", linewidth=0.8)
    axf.set_ylabel("FWI")
    axf.grid(alpha=0.3)
    plt.xticks(rotation=30)
    st.pyplot(fig3, use_container_width=True)

    st.subheader("Últimos 15 días con racha máxima ≥ 60 km/h")
    rachas = dfa[dfa["wind_gust"] >= 60][["Fecha", "wind_gust", "Temperatura Máxima (°C)", "Humidity_Min (%)"]]
    rachas = rachas.rename(columns={"wind_gust": "Racha máx. (km/h)"}).sort_values("Fecha", ascending=False)
    if len(rachas):
        st.dataframe(rachas.head(15), use_container_width=True, hide_index=True)
    else:
        st.caption("Sin rachas ≥ 60 km/h en el periodo seleccionado.")

st.divider()
st.caption(
    "Este panel se genera automáticamente a partir de una estación meteorológica personal; "
    "no sustituye a las fuentes oficiales (AEMET, protección civil) para decisiones de seguridad. "
    "Código y metodología: proyecto abierto de la estación ITORRE77."
)
