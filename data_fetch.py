# -*- coding: utf-8 -*-
"""
==============================================================================
 data_fetch.py — Cliente de la API de Ecowitt + fusión con el histórico
==============================================================================
Obtiene los días que falten entre el histórico base (el Excel que ya
manejas manualmente, empaquetado en el repositorio como `historico_base.xlsx`)
y "ayer", llamando a la API en la nube de Ecowitt, y devuelve un DataFrame
con el mismo esquema de columnas que ya usan `analisis_meteorologico.py` e
`informe_incendios.py` — así los dos scripts se reutilizan sin tocarlos.

CREDENCIALES (nunca van en el código ni en el repositorio):
Se leen de `st.secrets` en producción (Streamlit Community Cloud) o de
variables de entorno en local. Hacen falta tres valores, todos disponibles
en https://www.ecowitt.net/user/index (apartado "API"):
    ECOWITT_APPLICATION_KEY
    ECOWITT_API_KEY
    ECOWITT_MAC              (dirección MAC o IMEI del panel GW1000/GW1100)

NOTA IMPORTANTE: la API v3 de Ecowitt (doc.ecowitt.net) es estable desde
hace años y esta integración sigue su estructura de siempre (endpoints
`device/real_time` y `device/history`, parámetros application_key/api_key/
mac/call_back...). No he podido abrir la documentación interactiva desde
aquí (bloquea el acceso automatizado), así que en cuanto tengas tus
credenciales, haz una primera llamada de prueba y compara los nombres de
campo del JSON de respuesta con los que se usan más abajo en
`_parse_history_response` — pueden variar ligeramente según el modelo de
estación/sensores conectados.
==============================================================================
"""

import os
import math
import datetime as dt
from typing import Optional

import numpy as np
import pandas as pd
import requests

ECOWITT_HISTORY_URL = "https://api.ecowitt.net/api/v3/device/history"

# Escribimos TANTO el nombre de columna "antiguo" (el que leen directamente
# analisis_meteorologico.py, informe_incendios.py, informe_tecnico.py y
# riego_jardin.py) COMO el "nuevo" (por si algún cálculo lo prefiere). No
# conviene depender de que cada script haga bien el *coalesce* entre ambos
# formatos: ya hemos visto que no todos lo hacen de forma consistente.
COLUMNAS_SALIDA = [
    "Fecha",
    "Temperatura Máxima (°C)", "Temperatura Mínima (°C)", "tempAvg",
    "Temperature_Max (℃)", "Temperature_Min (℃)", "Temperature_Media (℃)",
    "Humidity_Max (%)", "Humidity_Min (%)",
    "Precipitación Total (mm)",
    "Viento_Media (km/h)", "Viento_Racha_Max (km/h)", "Viento_Dir_Predominante (º)",
    "Solar_Media (W/m²)", "Vpd_Max (hPa)", "dewptAvg",
]


def _secret(nombre: str) -> Optional[str]:
    """Lee una credencial de st.secrets (si hay Streamlit disponible y
    configurado) o, si no, de una variable de entorno. Nunca lanza error
    solo por no encontrarla: quien llame decide qué hacer si falta."""
    try:
        import streamlit as st
        if nombre in st.secrets:
            return st.secrets[nombre]
    except Exception:
        pass
    return os.getenv(nombre)


def credenciales_disponibles() -> bool:
    return all(_secret(k) for k in ("ECOWITT_APPLICATION_KEY", "ECOWITT_API_KEY", "ECOWITT_MAC"))


def _es_kpa(t_c):
    """Presión de vapor de saturación (kPa) — FAO-56."""
    return 0.6108 * math.exp(17.27 * t_c / (t_c + 237.3))


def _vpd_hpa(t_c, rh):
    """Déficit de presión de vapor en hPa a partir de temperatura y humedad."""
    if pd.isna(t_c) or pd.isna(rh):
        return np.nan
    es = _es_kpa(t_c)
    ea = es * (rh / 100.0)
    return max(0.0, (es - ea) * 10.0)  # kPa -> hPa


def _dewpoint_c(t_c, rh):
    """Punto de rocío (°C) — fórmula de Magnus, suficiente para este uso."""
    if pd.isna(t_c) or pd.isna(rh) or rh <= 0:
        return np.nan
    a, b = 17.27, 237.3
    alpha = (a * t_c) / (b + t_c) + math.log(rh / 100.0)
    return (b * alpha) / (a - alpha)


def _f_to_c(f):
    return (f - 32.0) * 5.0 / 9.0 if f is not None else np.nan


def _mph_to_kmh(mph):
    return mph * 1.60934 if mph is not None else np.nan


def _in_to_mm(inch):
    return inch * 25.4 if inch is not None else np.nan


def fetch_history(fecha_inicio: dt.date, fecha_fin: dt.date, timeout: int = 20) -> pd.DataFrame:
    """Descarga el histórico horario/diario de Ecowitt entre dos fechas
    (ambas inclusive) y lo agrega a resolución DIARIA con el esquema de
    columnas que ya esperan los scripts de análisis.

    Devuelve un DataFrame vacío (con las columnas correctas) si faltan
    credenciales o si la llamada falla — así la app puede seguir mostrando
    el histórico base aunque la API no responda ese día, en vez de romperse.
    """
    if fecha_inicio > fecha_fin:
        return pd.DataFrame(columns=COLUMNAS_SALIDA)

    app_key = _secret("ECOWITT_APPLICATION_KEY")
    api_key = _secret("ECOWITT_API_KEY")
    mac = _secret("ECOWITT_MAC")
    if not (app_key and api_key and mac):
        return pd.DataFrame(columns=COLUMNAS_SALIDA)

    params = {
        "application_key": app_key,
        "api_key": api_key,
        "mac": mac,
        "start_date": f"{fecha_inicio:%Y-%m-%d} 00:00:00",
        "end_date": f"{fecha_fin:%Y-%m-%d} 23:59:59",
        "cycle_type": "auto",
        "call_back": "outdoor,wind,rainfall,solar_and_uvi,indoor",
        "temp_unitid": 1,       # °C
        "wind_speed_unitid": 7,  # km/h
        "rainfall_unitid": 12,   # mm
        "solar_irradiance_unitid": 1,  # W/m²
    }
    try:
        r = requests.get(ECOWITT_HISTORY_URL, params=params, timeout=timeout)
        r.raise_for_status()
        payload = r.json()
    except Exception as e:
        print(f"⚠️  Ecowitt API: no se pudo obtener el histórico ({e}). "
              f"Se sigue solo con el histórico base.")
        return pd.DataFrame(columns=COLUMNAS_SALIDA)

    return _parse_history_response(payload)


def _serie_horaria_a_diaria(serie_horaria: dict) -> pd.Series:
    """Convierte el dict {timestamp_unix: valor} que devuelve Ecowitt en
    una Serie diaria (usa la fecha local del propio timestamp)."""
    if not serie_horaria:
        return pd.Series(dtype=float)
    idx = pd.to_datetime([int(k) for k in serie_horaria.keys()], unit="s")
    vals = pd.to_numeric(pd.Series(list(serie_horaria.values())), errors="coerce")
    s = pd.Series(vals.values, index=idx)
    return s


def _parse_history_response(payload: dict) -> pd.DataFrame:
    """Adapta el JSON de `device/history` al esquema COLUMNAS_SALIDA.

    Estructura esperada (API v3, `call_back` con varios grupos separados
    por comas): payload['data'][grupo][campo]['list'] = {timestamp: valor}.
    Si tu estación devuelve nombres de campo distintos (varía según el
    modelo/sensores), ajusta aquí las claves marcadas con "# campo:".
    """
    data = (payload or {}).get("data", {})
    if not data:
        return pd.DataFrame(columns=COLUMNAS_SALIDA)

    def lista(grupo, campo):
        try:
            return data[grupo][campo]["list"]
        except (KeyError, TypeError):
            return {}

    outdoor_temp = _serie_horaria_a_diaria(lista("outdoor", "temperature"))       # campo: outdoor.temperature
    outdoor_hum = _serie_horaria_a_diaria(lista("outdoor", "humidity"))           # campo: outdoor.humidity
    wind_speed = _serie_horaria_a_diaria(lista("wind", "wind_speed"))             # campo: wind.wind_speed
    wind_gust = _serie_horaria_a_diaria(lista("wind", "wind_gust"))               # campo: wind.wind_gust
    wind_dir = _serie_horaria_a_diaria(lista("wind", "wind_direction"))           # campo: wind.wind_direction
    rain_event = _serie_horaria_a_diaria(lista("rainfall", "daily"))              # campo: rainfall.daily
    solar = _serie_horaria_a_diaria(lista("solar_and_uvi", "solar"))              # campo: solar_and_uvi.solar

    if outdoor_temp.empty:
        return pd.DataFrame(columns=COLUMNAS_SALIDA)

    dias = sorted(set(outdoor_temp.index.date))
    filas = []
    for d in dias:
        t_dia = outdoor_temp[outdoor_temp.index.date == d]
        h_dia = outdoor_hum[outdoor_hum.index.date == d] if not outdoor_hum.empty else pd.Series(dtype=float)
        ws_dia = wind_speed[wind_speed.index.date == d] if not wind_speed.empty else pd.Series(dtype=float)
        wg_dia = wind_gust[wind_gust.index.date == d] if not wind_gust.empty else pd.Series(dtype=float)
        wd_dia = wind_dir[wind_dir.index.date == d] if not wind_dir.empty else pd.Series(dtype=float)
        r_dia = rain_event[rain_event.index.date == d] if not rain_event.empty else pd.Series(dtype=float)
        s_dia = solar[solar.index.date == d] if not solar.empty else pd.Series(dtype=float)

        tmax = float(t_dia.max()) if len(t_dia) else np.nan
        tmin = float(t_dia.min()) if len(t_dia) else np.nan
        tmed = float(t_dia.mean()) if len(t_dia) else np.nan
        hmax = float(h_dia.max()) if len(h_dia) else np.nan
        hmin = float(h_dia.min()) if len(h_dia) else np.nan

        filas.append({
            "Fecha": pd.Timestamp(d),
            "Temperatura Máxima (°C)": tmax,
            "Temperatura Mínima (°C)": tmin,
            "tempAvg": tmed,
            "Temperature_Max (℃)": tmax,
            "Temperature_Min (℃)": tmin,
            "Temperature_Media (℃)": tmed,
            "Humidity_Max (%)": hmax,
            "Humidity_Min (%)": hmin,
            "Precipitación Total (mm)": float(r_dia.max()) if len(r_dia) else 0.0,
            "Viento_Media (km/h)": float(ws_dia.mean()) if len(ws_dia) else np.nan,
            "Viento_Racha_Max (km/h)": float(wg_dia.max()) if len(wg_dia) else np.nan,
            "Viento_Dir_Predominante (º)": float(wd_dia.mode().iloc[0]) if len(wd_dia) else np.nan,
            "Solar_Media (W/m²)": float(s_dia.mean()) if len(s_dia) else np.nan,
            "Vpd_Max (hPa)": _vpd_hpa(tmax, hmin) if not (pd.isna(tmax) or pd.isna(hmin)) else np.nan,
            "dewptAvg": _dewpoint_c(tmed, (h_dia.mean() if len(h_dia) else np.nan)),
        })

    return pd.DataFrame(filas, columns=COLUMNAS_SALIDA)


def construir_dataset_actualizado(ruta_base_xlsx: str, ruta_salida_xlsx: str) -> dict:
    """Función principal: lee el histórico base, calcula qué días faltan
    hasta ayer, los descarga de Ecowitt si hay credenciales, fusiona todo y
    escribe el resultado en `ruta_salida_xlsx` (mismo formato que ya leen
    `analisis_meteorologico.cargar_datos` e `informe_incendios.load_data`).

    Devuelve un pequeño resumen (para mostrar en la app) con cuántos días
    nuevos se han incorporado y si la API respondió o no.
    """
    base = pd.read_excel(ruta_base_xlsx)
    base["Fecha"] = pd.to_datetime(base["Fecha"])
    ultima_fecha_base = base["Fecha"].max().date()
    ayer = (dt.datetime.utcnow() - dt.timedelta(days=1)).date()

    resumen = {"ultima_fecha_base": ultima_fecha_base, "ayer": ayer,
               "dias_nuevos": 0, "api_ok": False, "credenciales": credenciales_disponibles()}

    if ultima_fecha_base >= ayer:
        base.to_excel(ruta_salida_xlsx, index=False)
        return resumen

    nuevos = fetch_history(ultima_fecha_base + dt.timedelta(days=1), ayer)
    resumen["dias_nuevos"] = len(nuevos)
    resumen["api_ok"] = len(nuevos) > 0

    if len(nuevos):
        combinado = pd.concat([base, nuevos], ignore_index=True, sort=False)
        combinado = combinado.sort_values("Fecha").drop_duplicates(subset="Fecha", keep="last")
    else:
        combinado = base

    combinado.to_excel(ruta_salida_xlsx, index=False)
    return resumen
