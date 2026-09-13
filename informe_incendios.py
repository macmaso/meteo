#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
=============================================================================
INFORME DE PELIGRO DE INCENDIOS E INDICADORES DE SEQUÍA
Estación ITORRE77 - Los Tomillares
=============================================================================
Genera Word + PNG (permanentes) + CSV con:
· Cabecera con procedencia de datos y "Resumen de valores a día …"
· Pie de página con numeración automática (pág. X de Y)
· Tabla-calor 30-30-30 (fila "Día" separada) + Histórico de días extremos
  (con columna "Total año" y media anual 2021-2025)
· KBDI, FWI, CBI, Angström, VPD, Precipitación, Histórico de precipitación
  acumulada mensual, ET₀, Rosa de vientos
· Tabla diaria con unidades y flechas de tendencia
Uso:
python informe_incendios.py
python informe_incendios.py --days 60
=============================================================================
"""
import os, re, math, argparse
from datetime import datetime
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as patches
import seaborn as sns
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass
# ============================ CONFIGURACIÓN (editable) ============================
STATION        = os.getenv("STATION_NAME", "Los Tomillares (ITORRE77)")
STATION_DIR    = os.getenv("STATION_ADDRESS", "c/Jaral, 23, Los Tomillares, Torremocha de Jarama")
STATION_URL    = os.getenv("STATION_URL", "wunderground.com/dashboard/pws/ITORRE77")
AUTOR          = os.getenv("REPORT_AUTHOR", "Manuel Martínez Sotillos")
LATITUD_GRADOS = float(os.getenv("STATION_LAT", "40.75"))
ALTITUD_M      = float(os.getenv("STATION_ALT_M", "850"))
DPI            = int(os.getenv("REPORT_DPI", "150"))
INPUT_XLSX     = os.getenv("ECOWITT_OUTPUT_FILE", "weather_data_metric_es.xlsx")
BASE_DIR       = os.path.dirname(os.path.abspath(__file__))
INPUT_XLSX     = os.path.join(BASE_DIR, INPUT_XLSX)
KBDI_RAIN_MM_ENV = os.getenv("KBDI_ANNUAL_RAIN_MM")
DEFAULT_DAYS = int(os.getenv("REPORT_DEFAULT_DAYS", "90"))
REINICIO_FWI_ANUAL = os.getenv("REINICIO_FWI_ANUAL", "1") not in ("0", "false", "False")
sns.set_theme(style="whitegrid")
MES_LARGO = {1:"Enero",2:"Febrero",3:"Marzo",4:"Abril",5:"Mayo",6:"Junio",7:"Julio",
             8:"Agosto",9:"Septiembre",10:"Octubre",11:"Noviembre",12:"Diciembre"}
MES_ES = {1:"Ene",2:"Feb",3:"Mar",4:"Abr",5:"May",6:"Jun",
          7:"Jul",8:"Ago",9:"Sep",10:"Oct",11:"Nov",12:"Dic"}
GRIS = RGBColor(0x4B,0x55,0x63); AZUL = RGBColor(0x1F,0x29,0x37)
ROJO = RGBColor(0x99,0x1B,0x1B)
# =============================================================================
# ETIQUETAS
# =============================================================================
def kbdi_label(v):
    if v <= 200: return "Bajo Riesgo"
    if v <= 400: return "Sequía Moderada"
    if v <= 600: return "Sequía Severa"
    return "Sequía Extrema"
def fwi_label(v):
    if v < 11.2: return "Bajo"
    if v < 21.3: return "Moderado"
    if v < 38.0: return "Alto"
    if v < 50.0: return "Muy Alto"
    return "Peligro Extremo"
def cbi_label(v):
    if v < 50:   return "Bajo"
    if v < 75:   return "Moderado"
    if v < 90:   return "Alto"
    if v < 97.5: return "Muy Alto"
    return "Extremo"
def angstrom_label(v):   # valores MÁS BAJOS = MAYOR peligro (Ångström, 1949)
    if v < 2.0: return "Peligro Extremo"
    if v < 2.5: return "Peligro Alto"
    if v < 3.0: return "Peligro Moderado"
    return "Peligro Bajo"
def _txt_percentil(pct):
    if pct is None:
        return None
    return f"Percentil {pct:.0f} del histórico propio de la estación (más alto que el {pct:.0f}% de los días registrados)."
def vpd_label(v):
    if v < 15: return "Bajo estrés"
    if v < 25: return "Moderado"
    if v < 35: return "Alto"
    return "Extremo"
DESCRIPCIONES = {
    'regla3030': ("La regla 30-30-30 marca el peligro extremo cuando coinciden temperatura "
                  "superior a 30 °C, humedad relativa inferior al 30 % y rachas de viento "
                  "superiores a 30 km/h. La tabla muestra, día a día, qué variables cumplen "
                  "su condición (celda coloreada con el nivel del día) y cuáles NO (gris); "
                  "la fila superior resume el nivel de alerta de cada jornada."),
    'heatmap': ("El mapa de calor muestra el número de días con peligro extremo (3/3 en Regla "
                "30-30-30) por mes y año en todo el histórico, con la columna 'Total año' a la "
                "derecha. Permite apreciar la estacionalidad del riesgo y si los días extremos "
                "aumentan o disminuyen. Para la media anual se excluyen 2020 (la estación empieza "
                "en mayo) y el año en curso."),
    'precipmensual': ("El mapa de calor muestra la precipitación total acumulada en cada mes "
                      "(en mm) para todo el histórico de la estación, con la columna 'Total año' "
                      "con el acumulado anual. Permite identificar de un vistazo los meses secos y "
                      "húmedos de cada año y comparar temporadas entre sí: una sucesión de meses con "
                      "valores muy por debajo de lo habitual anticipa el aumento del riesgo de "
                      "incendio y la necesidad de riego. El año en curso se muestra con acumulado "
                      "parcial hasta la última fecha disponible."),
    'kbdi': ("El índice de sequía Keetch-Byram (KBDI) estima el déficit acumulado de humedad "
             "en el suelo y el mantillo orgánico (combustible profundo), escala 0-800: <200 "
             "bajo; 200-400 moderada; 400-600 severa; >600 extrema. Junto a la escala genérica "
             "se indica el percentil que ocupa el valor de hoy dentro del histórico propio de "
             "esta estación, como referencia local adicional."),
    'fwi': ("El Fire Weather Index (FWI) es el índice final del sistema canadiense CFFDRS. "
            "Integra la humedad de combustibles finos, medios y profundos con el viento. "
            "Escala: <11,2 bajo; 11,2-21,3 moderado; 21,3-38 alto; 38-50 muy alto; >50 extremo. "
            "Esta escala está calibrada para un sistema que se reinicia cada temporada de "
            "incendios; este informe reinicia igualmente los códigos cada 1 de enero (ver "
            "configuración del script), pero en veranos mediterráneos muy secos el código de "
            "sequía profunda (DC) puede seguir empujando el FWI por encima de los valores "
            "habituales en otros climas. Por eso se acompaña del percentil histórico de la "
            "propia estación, más fiable para comparar un día con la experiencia local."),
    'cbi': ("El índice de combustión de Chandler (CBI) estima la inflamabilidad instantánea a "
            "partir de temperatura, humedad relativa y racha máxima. Escala: <50 bajo; 50-75 "
            "moderado; 75-90 alto; 90-97,5 muy alto; >97,5 extremo."),
    'angstrom': ("El índice de Angström (Ångström, 1949) combina humedad relativa y temperatura "
                 "mediante I = RH/20 + (27−T)/10. Los valores BAJOS indican MAYOR peligro "
                 "(aire cálido y seco): <2,0 extremo; 2,0-2,5 alto; 2,5-3,0 moderado; >3,0 bajo. "
                 "Por eso la zona roja queda en la parte baja de la gráfica."),
    'vpd': ("El déficit de presión de vapor (VPD) mide la demanda evaporativa de la atmósfera. "
            "Un VPD alto seca rápidamente el combustible fino. Escala (hPa): <15 bajo estrés; "
            "15-25 moderado; 25-35 alto; >35 extremo."),
    'rain': ("La precipitación es el factor que más reduce el peligro. Las lluvias efectivas "
             "(>= 5 mm) recargan el suelo y reinician los acumuladores de sequía. El título "
             "indica el total acumulado del periodo."),
    'et0': ("La evapotranspiración de referencia (ET₀) es la cantidad de agua que la "
            "atmósfera 'roba' cada día al suelo y a la vegetación, calculada con la ecuación "
            "FAO-56 Penman-Monteith a partir de temperatura, humedad, viento y radiación "
            "solar (1 mm = 1 L/m²). El valor acumulado del periodo refleja la demanda "
            "evaporativa total: cuanto mayor, más rápido se seca el combustible fino."),
    'windrose': ("La rosa de vientos muestra la distribución de velocidad y dirección del "
                 "viento. Permite identificar las direcciones predominantes y las más veloces, "
                 "que marcarían la trayectoria de propagación en caso de incendio."),
}
# =============================================================================
# DERIVACIONES
# =============================================================================
def rh_from_t_td(t_c, td_c):
    if pd.isna(t_c) or pd.isna(td_c): return np.nan
    es = 6.112*math.exp(17.67*t_c/(t_c+243.5))
    ea = 6.112*math.exp(17.67*td_c/(td_c+243.5))
    return max(1.0, min(100.0, 100.0*ea/es))
def vpd_from_t_rh(t_c, rh):
    if pd.isna(t_c) or pd.isna(rh): return np.nan
    es = 6.112*math.exp(17.67*t_c/(t_c+243.5))
    return max(0.0, es*(1.0-rh/100.0))
# =============================================================================
# KBDI
# =============================================================================
def calculate_kbdi(df, r_anual_mm=478.275):
    out=[]; q=0.0; in_rain=False; cum=0.0; net_prev=0.0
    for _, r in df.iterrows():
        t = r['Temperatura Máxima (°C)'] if not pd.isna(r['Temperatura Máxima (°C)']) else 20.0
        p = r['Precipitación Total (mm)'] if not pd.isna(r['Precipitación Total (mm)']) else 0.0
        t_f=t*1.8+32.0; p_in=p/25.4; r_in=r_anual_mm/25.4
        if p_in > 0.0:
            if not in_rain: in_rain, cum, net_prev = True, p_in, 0.0
            else: cum += p_in
            if cum > 0.20:
                nt=cum-0.20; nd=nt-net_prev; net_prev=nt
            else: nd=0.0
        else:
            in_rain, cum, net_prev, nd = False, 0.0, 0.0, 0.0
        q = max(0.0, q - nd*100.0)
        if t_f > 50.0:
            num=(800.0-q)*(0.968*math.exp(0.0486*t_f)-8.30)*0.001
            den=1.0+10.88*math.exp(-0.0441*r_in)
            q=min(800.0, q+num/den)
        out.append(q)
    return out
# =============================================================================
# FWI
# =============================================================================
def calc_ffmc(t, rh, w, rain, prev):
    if rain > 0.5:
        rf=rain-0.5
        mo=147.2*(101.0-prev)/(59.5+prev)
        mr=mo+42.5*rf*math.exp(-100.0/(251.0-mo))*(1.0-math.exp(-6.93/rf))
        if mo>150.0: mr+=0.0015*((mo-150.0)**2)*(rf**0.5)
        if mr>250.0: mr=250.0
        prev=59.5*(250.0-mr)/(147.2+mr)
    mo=147.2*(101.0-prev)/(59.5+prev)
    ed=0.942*(rh**0.679)+11.0*math.exp((rh-100.0)/10.0)+0.18*(21.1-t)*(1.0-math.exp(-0.115*rh))
    if mo>ed:
        ew=0.618*(rh**0.753)+10.0*math.exp((rh-100.0)/10.0)+0.18*(21.1-t)*(1.0-math.exp(-0.115*rh))
        if mo>ew:
            kl=0.424*(1.0-(rh/100.0)**1.7)+0.0694*(w**0.5)*(1.0-(rh/100.0)**8.0)
            kw=kl*0.581*math.exp(0.0365*t)
            m=ew+(mo-ew)*(10.0**(-kw))
        else: m=mo
    else:
        ko=0.424*(1.0-((100.0-rh)/100.0)**1.7)+0.0694*(w**0.5)*(1.0-((100.0-rh)/100.0)**8.0)
        kd=ko*0.581*math.exp(0.0365*t)
        m=ed-(ed-mo)*(10.0**(-kd))
    return min(99.0, max(0.0, 59.5*(250.0-m)/(147.2+m)))
def calc_dmc(t, rh, rain, month, prev):
    el=[6.5,7.5,9.0,12.8,13.9,13.9,12.4,10.9,9.4,8.0,7.0,6.0]; t=max(-1.1,t)
    prev=max(0.0, prev)
    if rain>1.5:
        rw=0.92*rain-1.27; wmi=20.0+280.0/math.exp(0.023*prev)
        if prev<=33.0: b=100.0/(0.5+0.3*prev)
        elif prev<=65.0: b=14.0-1.3*math.log(prev)
        else: b=6.2*math.log(prev)-17.2
        wmr=wmi+1000.0*rw/(48.77+max(b,1e-6)*rw)
        arg=max(wmr-20.0, 1e-6)
        prev=max(0.0, 43.43*(5.634-math.log(arg)))
    rk=1.894*(t+1.1)*(100.0-rh)*el[month-1]*0.0001 if t>-1.1 else 0.0
    return max(0.0, prev+rk)
def calc_dc(t, rain, month, prev):
    fl=[-1.6,-1.6,-1.6,0.9,3.8,5.8,6.4,5.0,2.4,0.4,-1.6,-1.6]; t=max(-2.8,t)
    prev=max(0.0, prev)
    if rain>2.8:
        rw=0.83*rain-1.27; smi=800.0*math.exp(-prev/400.0)
        arg=max(smi+3.937*rw, 1e-6)
        prev=max(0.0, 400.0*math.log(800.0/arg))
    v=0.36*(t+2.8)+fl[month-1] if t>-2.8 else 0.0
    return max(0.0, prev+0.5*max(0.0,v))
def calc_isi(w, ffmc):
    fm=147.2*(101.0-ffmc)/(59.5+ffmc)
    return 0.208*math.exp(0.05039*w)*91.9*math.exp(-0.1386*fm)*(1.0+(fm**5.31)/4.93e7)
def calc_bui(dmc, dc):
    if dmc<=0 and dc<=0: return 0.0
    if dmc<=0.4*dc: bui = 0.8*dmc*dc/(dmc+0.4*dc)
    else: bui = dmc-(1.0-0.8*dc/(dmc+0.4*dc))*(0.92+(0.0114*dmc)**1.7)
    return max(0.0, bui)
def calc_fwi(isi, bui):
    bb=0.1*isi*(0.626*(bui**0.809)+2.0) if bui<=80.0 else 0.1*isi*(1000.0/(25.0+108.64*math.exp(-0.023*bui)))
    return bb if bb<=1.0 else math.exp(2.72+0.434*math.log(bb))
def run_fwi_system(df, reinicio_anual=REINICIO_FWI_ANUAL):
    out=[]; ff,dm,dc = 85.0,6.0,15.0; año_prev=None
    for _, r in df.iterrows():
        t  = r['Temperatura Máxima (°C)'] if not pd.isna(r['Temperatura Máxima (°C)']) else 25.0
        rh = r['Humidity_Min (%)'] if not pd.isna(r['Humidity_Min (%)']) else 30.0
        w  = r['wind_avg'] if not pd.isna(r['wind_avg']) else 10.0
        rain = r['Precipitación Total (mm)'] if not pd.isna(r['Precipitación Total (mm)']) else 0.0
        m = r['Fecha'].month
        año = r['Fecha'].year
        if reinicio_anual and año_prev is not None and año != año_prev:
            ff, dm, dc = 85.0, 6.0, 15.0
        año_prev = año
        ff=calc_ffmc(t,rh,w,rain,ff); dm=calc_dmc(t,rh,rain,m,dm); dc=calc_dc(t,rain,m,dc)
        out.append(calc_fwi(calc_isi(w,ff), calc_bui(dm,dc)))
    return out
def calc_cbi(t, rh, g):
    tf=t*1.8+32.0; mph=g/1.60934
    return ((110.0-1.373*rh)-0.54*(10.20-tf)+124.0*(10.0**(-0.0142*rh)))*(10.0**(0.02*mph))/10.0
def calc_angstrom(t, rh):   # Fórmula real de Ångström (1949): I = RH/20 + (27−T)/10
    return rh/20.0 + (27.0 - t)/10.0
def percentil_historico(serie, valor):
    s = pd.Series(serie).dropna()
    if len(s) == 0 or pd.isna(valor):
        return None
    return float((s <= valor).mean() * 100.0)
def eval_regla_30(r):
    c = sum([r['Temperatura Máxima (°C)']>30,
             r['Humidity_Min (%)']<30 if not pd.isna(r['Humidity_Min (%)']) else False,
             r['wind_gust']>30 if not pd.isna(r['wind_gust']) else False])
    return {3:'Extremo (3/3)',2:'Alerta (2/3)'}.get(c,'Base (1/3)')
# =============================================================================
# ET0 (FAO-56)
# =============================================================================
def e_sat_k(t): return 0.6108*math.exp(17.27*t/(t+237.3))
def ra_dia(J, lat_grados=LATITUD_GRADOS):
    phi=math.radians(lat_grados)
    dr=1+0.033*math.cos(2*math.pi*J/365)
    dec=0.409*math.sin(2*math.pi*J/365-1.39)
    ws=math.acos(max(-1.0,min(1.0,-math.tan(phi)*math.tan(dec))))
    return (24*60/math.pi)*0.0820*dr*(ws*math.sin(phi)*math.sin(dec)+math.cos(phi)*math.cos(dec)*math.sin(ws))
def calc_et0(df, lat_grados=LATITUD_GRADOS, altitud_m=ALTITUD_M):
    out=[]
    for _, r in df.iterrows():
        t  = r['Temperatura Máxima (°C)'] if not pd.isna(r['Temperatura Máxima (°C)']) else 20.0
        tn = r['Temperatura Mínima (°C)'] if 'Temperatura Mínima (°C)' in r and not pd.isna(r['Temperatura Mínima (°C)']) else t-8.0
        rh = r['Humidity_Min (%)'] if not pd.isna(r['Humidity_Min (%)']) else 40.0
        w  = (r['wind_avg']/3.6) if not pd.isna(r['wind_avg']) else 2.2
        rs = r['Solar_Media (W/m²)'] if 'Solar_Media (W/m²)' in r and not pd.isna(r['Solar_Media (W/m²)']) else np.nan
        J  = r['Fecha'].timetuple().tm_yday
        es=(e_sat_k(t)+e_sat_k(tn))/2.0
        ea=es*rh/100.0
        Rs = rs*0.0864 if not pd.isna(rs) else 0.16*math.sqrt(max(t-tn,0.5))*ra_dia(J, lat_grados)
        Rso=(0.75+2e-5*altitud_m)*ra_dia(J, lat_grados)
        ratio=min(Rs/Rso,1.0) if Rso>0 else 1.0
        Rns=0.77*Rs
        Rnl=4.903e-9*((t+273.16)**4+(tn+273.16)**4)/2*(0.34-0.14*math.sqrt(max(ea,0.05)))*(1.35*ratio-0.35)
        Rn=Rns-Rnl
        gamma=0.000665*(101.3*((293-0.0065*altitud_m)/293)**5.26)
        delta=4098*e_sat_k((t+tn)/2)/((t+tn)/2+237.3)**2
        out.append(max(0.0,(0.408*delta*Rn+gamma*(900/((t+tn)/2+273))*w*(es-ea))/(delta+gamma*(1+0.34*w))))
    return out
# =============================================================================
# CARGA
# =============================================================================
def rellenar_huecos_diarios(df):
    df = df.set_index('Fecha').sort_index()
    calendario_completo = pd.date_range(df.index.min(), df.index.max(), freq='D')
    huecos = calendario_completo.difference(df.index)
    df = df.reindex(calendario_completo)
    df['Interpolado'] = df.index.isin(huecos)
    cols_continuas = [c for c in [
        'Temperatura Máxima (°C)', 'Temperatura Mínima (°C)', 'Humidity_Min (%)',
        'wind_avg', 'wind_gust', 'dewptAvg', 'Solar_Media (W/m²)', 'Vpd_Max (hPa)',
        'Viento_Dir_Predominante (º)',
    ] if c in df.columns]
    if cols_continuas:
        df[cols_continuas] = df[cols_continuas].interpolate(method='linear', limit=5, limit_area='inside')
    if 'Precipitación Total (mm)' in df.columns:
        df['Precip_estimada'] = df['Precipitación Total (mm)'].isna()
        df['Precipitación Total (mm)'] = df['Precipitación Total (mm)'].fillna(0.0)
    else:
        df['Precip_estimada'] = False
    df.index.name = 'Fecha'
    df = df.reset_index()
    return df
def load_data(path):
    if not os.path.exists(path):
        raise SystemExit(f"❌ No se encontró '{path}'. Colócalo junto al script.")
    df = pd.read_excel(path)
    df['Fecha'] = pd.to_datetime(df['Fecha'], errors='coerce')
    df = df.dropna(subset=['Fecha']).sort_values('Fecha').drop_duplicates(subset='Fecha').reset_index(drop=True)
    if 'Temperature_Max (℃)' in df.columns:
        df['Temperatura Máxima (°C)'] = df['Temperatura Máxima (°C)'].fillna(df['Temperature_Max (℃)'])
    if 'Temperatura Mínima (°C)' in df.columns and 'Temperature_Min (℃)' in df.columns:
        df['Temperatura Mínima (°C)'] = df['Temperatura Mínima (°C)'].fillna(df['Temperature_Min (℃)'])
    df['Temperatura Máxima (°C)'] = df['Temperatura Máxima (°C)'].ffill()
    if 'Humidity_Min (%)' not in df.columns: df['Humidity_Min (%)']=np.nan
    m = df['Humidity_Min (%)'].isna()
    if m.any():
        td = df['dewptAvg'] if 'dewptAvg' in df.columns else pd.Series(np.nan, index=df.index)
        df.loc[m,'Humidity_Min (%)']=[rh_from_t_td(t,d) for t,d in zip(df.loc[m,'Temperatura Máxima (°C)'], td[m])]
    if 'Vpd_Max (hPa)' not in df.columns: df['Vpd_Max (hPa)']=np.nan
    mv = df['Vpd_Max (hPa)'].isna()
    if mv.any():
        df.loc[mv,'Vpd_Max (hPa)']=[vpd_from_t_rh(t,h) for t,h in zip(df.loc[mv,'Temperatura Máxima (°C)'], df.loc[mv,'Humidity_Min (%)'])]
    def wa(r):
        for c in ('windspeedAvg','Viento_Media (km/h)'):
            if c in r and not pd.isna(r[c]) and r[c]>=0: return float(r[c])
        return np.nan
    def wg(r):
        for c in ('windgustHigh','Viento_Racha_Max (km/h)'):
            if c in r and not pd.isna(r[c]): return float(r[c])
        return np.nan
    df['wind_avg']=df.apply(wa,axis=1); df['wind_gust']=df.apply(wg,axis=1)
    df = rellenar_huecos_diarios(df)
    df['wind_avg']=df['wind_avg'].ffill().fillna(df['wind_avg'].median())
    df['wind_gust']=df['wind_gust'].ffill().fillna(df['wind_avg']*1.5)
    df['Temperatura Máxima (°C)']=df['Temperatura Máxima (°C)'].ffill()
    if 'Temperatura Mínima (°C)' in df.columns:
        df['Temperatura Mínima (°C)']=df['Temperatura Mínima (°C)'].ffill()
    if 'Humidity_Min (%)' in df.columns:
        df['Humidity_Min (%)']=df['Humidity_Min (%)'].ffill()
    df['Precipitación Total (mm)']=df['Precipitación Total (mm)'].fillna(0.0)
    return df
# =============================================================================
# PIVOTS Y MEDIAS ANUALES
# =============================================================================
def pivot_extremos(df):
    d = df.copy()
    d['y'], d['m'] = d['Fecha'].dt.year, d['Fecha'].dt.month
    piv = (d[d['Regla_30']=='Extremo (3/3)'].groupby(['y','m']).size()
           .unstack(fill_value=0).reindex(columns=range(1,13), fill_value=0))
    piv = piv.reindex(sorted(d['y'].unique()), fill_value=0)
    return piv
def media_anios_validos(df):
    d = df.copy()
    d['y'] = d['Fecha'].dt.year
    anio_actual = datetime.now().year
    anios_validos = [y for y in sorted(d['y'].unique()) if y != 2020 and y != anio_actual]
    ext = d[d['Regla_30'] == 'Extremo (3/3)']
    totales = (ext.groupby(ext['Fecha'].dt.year).size()
               .reindex(anios_validos, fill_value=0))
    media = totales.mean() if len(totales) else 0.0
    rango = f"{min(anios_validos)}–{max(anios_validos)}" if anios_validos else ""
    return media, rango
def pivot_mensual(df):
    d = df.copy()
    d['y'], d['m'] = d['Fecha'].dt.year, d['Fecha'].dt.month
    piv = (d.groupby(['y','m'])['Precipitación Total (mm)'].sum()
           .unstack(fill_value=0)
           .reindex(columns=range(1,13), fill_value=0))
    return piv
def media_anios_completos(df):
    d = df.copy()
    d['y'] = d['Fecha'].dt.year
    d['m'] = d['Fecha'].dt.month
    sums, anios = [], []
    for y, g in d.groupby('y'):
        if g['m'].nunique() == 12:
            sums.append(g['Precipitación Total (mm)'].sum())
            anios.append(y)
    if not sums:
        return None, ""
    return sum(sums)/len(sums), f"{min(anios)}–{max(anios)}"
# =============================================================================
# GRÁFICAS (PNG permanentes)
# =============================================================================
def _save(fig, name):
    p = os.path.join(BASE_DIR, f"incendios_{name}.png")
    fig.savefig(p, dpi=DPI, bbox_inches='tight'); plt.close(fig)
    return p
def _add_trend(ax, fechas, vals):
    msk=~np.isnan(vals)
    if msk.sum()>2:
        x=np.arange(len(vals)); m,b=np.polyfit(x[msk], vals[msk], 1)
        ax.plot(fechas, m*x+b, color='#4a148c', ls='--', lw=1.8, label='Tendencia')
def annotate_with_arrow(ax, ld, lv, txt, edge, yo):
    ax.annotate(txt, xy=(ld,lv), xytext=(ld-pd.Timedelta(days=3), lv+yo),
                arrowprops=dict(facecolor='black', shrink=0.05, width=1.2, headwidth=5),
                fontweight='bold', fontsize=8.5,
                bbox=dict(boxstyle='round,pad=0.2', fc='white', ec=edge, lw=1.2))
def plot_index(dfa, col, ylab, title, bands, labf, name, ylim=None, color='#d95f02'):
    fig, ax = plt.subplots(figsize=(10,4.0))
    ax.plot(dfa['Fecha'], dfa[col], marker='o', ms=3, lw=2, color=color, label=col)
    _add_trend(ax, dfa['Fecha'], dfa[col].values.astype(float))
    for lo,hi,c,a,lab in bands: ax.axhspan(lo,hi,color=c,alpha=a,label=lab)
    st=max(1,len(dfa)//12)
    ax.set_xticks(dfa['Fecha'][::st])
    ax.set_xticklabels([d.strftime('%d/%m') for d in dfa['Fecha'][::st]], rotation=45, fontsize=7.5)
    if ylim: ax.set_ylim(*ylim)
    ax.set_ylabel(ylab, fontsize=10); ax.set_title(title, fontsize=11, fontweight='bold', pad=10)
    ax.grid(True, ls='--', alpha=0.5); ax.legend(loc='upper left', fontsize=8)
    lv=dfa[col].iloc[-1]; rng=dfa[col].max()-dfa[col].min()
    yo=-rng*0.25 if lv>dfa[col].mean() else rng*0.25
    annotate_with_arrow(ax, dfa['Fecha'].iloc[-1], lv, f"{lv:.1f}\n({labf(lv)})", '#d62728', yo)
    plt.tight_layout(); return _save(fig, name)
def plot_regla3030_tabla(dfa, name="regla3030", por_bloque=30):
    filas=[('Temp. Máxima (°C)','Temperatura Máxima (°C)','T30'),
           ('Humedad mínima (%)','Humidity_Min (%)','H30'),
           ('Vel. Máx. Viento (km/h)','wind_gust','V30')]
    col_n={3:'#d62728', 2:'#f8b4ae', 1:'#f2c94c', 0:'#f7f3e9'}
    gris='#f1f1f1'
    bloques=[dfa.iloc[i:i+por_bloque].reset_index(drop=True) for i in range(0,len(dfa),por_bloque)]
    cw=ch=1.0; hh=0.75; sep=0.45; gap=1.2; x_etiq=4.4
    n_cols=max(len(b) for b in bloques)
    alto_bloque=hh+sep+3*ch+gap
    u=0.40
    x_units=x_etiq+n_cols*cw+0.2
    y_units=len(bloques)*alto_bloque+0.6
    mL,mR,mT,mB = 0.2,0.2,0.7,1.45
    fig=plt.figure(figsize=(x_units*u+mL+mR, y_units*u+mT+mB))
    ax=fig.add_axes([mL/fig.get_figwidth(), mB/fig.get_figheight(),
                     x_units*u/fig.get_figwidth(), y_units*u/fig.get_figheight()])
    for b, blk in enumerate(bloques):
        y_top=-b*alto_bloque
        ax.text(x_etiq-0.18, y_top-hh/2, 'Día', ha='right', va='center', fontsize=8, fontweight='bold')
        ax.text(0.1, y_top-hh/2, f"{blk['Fecha'].iloc[0]:%d/%m} – {blk['Fecha'].iloc[-1]:%d/%m}",
                ha='left', va='center', fontsize=7, color='#444444', fontweight='bold')
        for r,(lab,col,cond) in enumerate(filas):
            ax.text(x_etiq-0.18, y_top-hh-sep-(r+0.5)*ch, lab, ha='right', va='center',
                    fontsize=8, fontweight='bold')
        for j in range(len(blk)):
            row=blk.iloc[j]; x=x_etiq+j*cw
            nc=int(row['T30'])+int(row['H30'])+int(row['V30'])
            dc=col_n[nc]
            ax.add_patch(patches.Rectangle((x, y_top-hh), cw, hh, facecolor=dc,
                         edgecolor='white', linewidth=0.5, zorder=1))
            ax.text(x+cw/2, y_top-hh/2, row['Fecha'].strftime('%d'), ha='center', va='center',
                    fontsize=6.5, fontweight='bold',
                    color='white' if nc==3 else '#333333', zorder=2)
            for r,(lab,col,cond) in enumerate(filas):
                y=y_top-hh-sep-(r+1)*ch
                val=row[col]; cumple=bool(row[cond])
                face=dc if cumple else gris
                tc='white' if (cumple and nc==3) else ('#333333' if cumple else '#999999')
                ax.add_patch(patches.Rectangle((x,y), cw, ch, facecolor=face,
                             edgecolor='white', linewidth=0.5, zorder=1))
                txt='–' if pd.isna(val) else f"{val:.0f}"
                ax.text(x+cw/2, y+ch/2, txt, ha='center', va='center', fontsize=7,
                        color=tc, zorder=2)
    ax.set_xlim(0, x_units)
    ax.set_ylim(-(len(bloques)-1)*alto_bloque-hh-sep-3*ch-0.3, 0.3)
    ax.axis('off')
    ax.set_title(f"Regla 30-30-30 — detalle diario ({dfa['Fecha'].iloc[0]:%d/%m} – {dfa['Fecha'].iloc[-1]:%d/%m})",
                 fontsize=10.5, fontweight='bold')
    leyenda=[patches.Patch(facecolor=col_n[3], label='Día 3/3 — peligro extremo'),
             patches.Patch(facecolor=col_n[2], label='Día 2/3 — alerta sequedad y calor'),
             patches.Patch(facecolor=col_n[1], label='Día 1/3 — riesgo base'),
             patches.Patch(facecolor=col_n[0], edgecolor='#cccccc', label='Día 0/3 — NO cumple Regla'),
             patches.Patch(facecolor=gris, edgecolor='#cccccc', label='Variable que NO cumple la condición de la Regla')]
    fig.subplots_adjust(bottom=0.12)
    fig.legend(handles=leyenda, loc='lower center', ncol=2, fontsize=8, frameon=False)
    return _save(fig, name)
def plot_heatmap(df, name="heatmap"):
    piv = pivot_extremos(df)
    totales = piv.sum(axis=1)
    n_rows, n_cols = piv.shape
    fig, ax = plt.subplots(figsize=(10, max(3, n_rows*0.5)))
    sns.heatmap(piv, annot=True, fmt='d', cmap='YlOrRd', ax=ax,
                cbar_kws={'label':'Días Extremo (3/3)'}, annot_kws={'size':8})
    ax.text(n_cols + 0.5, -0.6, "Total año", ha='center', va='center',
            fontsize=8, fontweight='bold')
    for i, total in enumerate(totales):
        ax.text(n_cols + 0.5, i + 0.5, f"{total:.0f}", ha='center', va='center',
                fontsize=8, fontweight='bold')
    ax.set_xlim(0, n_cols + 1.2)
    ax.set_xticklabels(list(MES_ES.values()))
    ax.tick_params(axis='y', labelrotation=0)
    ax.set_title("Días con peligro extremo (3/3 en Regla 30-30-30) por año/mes",
                 fontsize=10, fontweight='bold')
    ax.set_ylabel("Año"); ax.set_xlabel("Mes")
    plt.tight_layout(); return _save(fig, name)
def plot_precip_mensual(piv, name="precipmensual"):
    totales = piv.sum(axis=1)
    n_rows, n_cols = piv.shape
    fig, ax = plt.subplots(figsize=(11, max(3, n_rows*0.45)))
    sns.heatmap(piv.round(1), annot=True, fmt='.1f', cmap='Blues', ax=ax,
                linewidths=0.5, linecolor='white',
                cbar_kws={'label':'Precipitación (mm)'}, annot_kws={'size':6.5})
    ax.set_xticklabels(list(MES_ES.values()), rotation=0)
    ax.axvline(n_cols, color='#888888', lw=1.0)
    ax.text(n_cols + 0.6, -0.6, "Total año", ha='center', va='center',
            fontsize=8, fontweight='bold')
    for i, total in enumerate(totales):
        ax.text(n_cols + 0.6, i + 0.5, f"{total:.0f}", ha='center', va='center',
                fontsize=7.5, fontweight='bold')
    ax.set_xlim(0, n_cols + 1.2)
    ax.set_ylim(n_rows + 0.8, -0.6)
    ax.text(0, n_rows + 0.55, "* Año en curso con acumulado parcial",
            ha='left', va='center', fontsize=7, style='italic', color='#666666')
    ax.tick_params(axis='y', labelrotation=0)
    ax.set_title("Histórico de precipitación acumulada mensual (mm)",
                 fontsize=10.5, fontweight='bold')
    plt.tight_layout()
    return _save(fig, name)
def plot_et0(dfa, name="et0"):
    fig, ax1 = plt.subplots(figsize=(10,3.5))
    ax1.bar(dfa['Fecha'], dfa['ET0_mm'], color='#74a9cf', label='ET₀ diaria')
    ax2=ax1.twinx()
    ax2.plot(dfa['Fecha'], dfa['ET0_mm'].cumsum(), color='#d73027', lw=2, label='Acumulada')
    st=max(1,len(dfa)//12)
    ax1.set_xticks(dfa['Fecha'][::st])
    ax1.set_xticklabels([d.strftime('%d/%m') for d in dfa['Fecha'][::st]], rotation=45, fontsize=7.5)
    ax1.set_ylabel('mm/día (L/m²)'); ax2.set_ylabel('Acumulada (mm)')
    ax1.set_title('Evapotranspiración de referencia (ET0) diaria y acumulada', fontsize=11, fontweight='bold')
    ax1.legend(loc='upper left', fontsize=8); ax2.legend(loc='upper right', fontsize=8)
    plt.tight_layout(); return _save(fig, name)
def plot_rain(dfa, name="rain"):
    fig, ax1 = plt.subplots(figsize=(10,3.5))
    ax1.bar(dfa['Fecha'], dfa['Precipitación Total (mm)'], color='#4575b4', label='Lluvia diaria')
    ax2=ax1.twinx()
    ax2.plot(dfa['Fecha'], dfa['Precipitación Total (mm)'].cumsum(), color='#d73027', lw=2, label='Acumulada')
    st=max(1,len(dfa)//12)
    ax1.set_xticks(dfa['Fecha'][::st])
    ax1.set_xticklabels([d.strftime('%d/%m') for d in dfa['Fecha'][::st]], rotation=45, fontsize=7.5)
    ax1.set_ylabel('mm/día'); ax2.set_ylabel('Acumulada (mm)')
    ax1.set_title('Precipitación diaria y acumulada', fontsize=10, fontweight='bold')
    ax1.legend(loc='upper left', fontsize=8); ax2.legend(loc='upper right', fontsize=8)
    plt.tight_layout(); return _save(fig, name)
def plot_wind_rose(dfa, name="windrose"):
    col='Viento_Dir_Predominante (º)'
    if col not in dfa.columns: return None
    validos = dfa[[col,'wind_avg']].dropna()
    if len(validos)==0: return None
    dirs = validos[col].values
    spd = validos['wind_avg'].values
    sector_bins = np.arange(-11.25, 360, 22.5)
    sector_idx = np.digitize((dirs + 11.25) % 360, sector_bins) - 1
    sector_idx = np.clip(sector_idx, 0, 15)
    sector_centros = np.deg2rad(np.arange(0, 360, 22.5))
    tramos = [(0,10,'#a6cee3','0-10 km/h'), (10,20,'#41ab5d','10-20 km/h'),
              (20,35,'#fdae61','20-35 km/h'), (35,999,'#d73027','>35 km/h')]
    fig, ax = plt.subplots(figsize=(7,7), subplot_kw={'projection':'polar'})
    acumulado = np.zeros(16)
    for lo, hi, color, etiqueta in tramos:
        conteo = np.zeros(16)
        for s in range(16):
            mask = (sector_idx == s) & (spd >= lo) & (spd < hi)
            conteo[s] = mask.sum()
        ax.bar(sector_centros, conteo, width=np.deg2rad(20), bottom=acumulado,
               color=color, label=etiqueta, edgecolor='white', linewidth=0.4)
        acumulado += conteo
    ax.set_theta_zero_location('N'); ax.set_theta_direction(-1)
    ax.set_xticks(sector_centros[::2])
    ax.set_xticklabels(['N','NNE','NE','ENE','E','ESE','SE','SSE',
                        'S','SSO','SO','OSO','O','ONO','NO','NNO'][::2], fontsize=8)
    ax.set_title("Rosa de vientos (frecuencia y velocidad)", fontsize=11, fontweight='bold', pad=20)
    ax.legend(loc='upper right', bbox_to_anchor=(1.3, 1.1), fontsize=8, title='Velocidad media diaria')
    return _save(fig, name)
# =============================================================================
# WORD
# =============================================================================
def set_cell_bg(cell, hexcolor):
    tcPr=cell._element.get_or_add_tcPr()
    tcPr.append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hexcolor}"/>'))
def get_arrow(c,p):
    if p is None or pd.isna(p) or pd.isna(c): return ""
    d=c-p
    if abs(d)<0.1: return " ="
    return " ▲" if d>0 else " ▼"
def _add_field(par, code):
    r=par.add_run(); fc=OxmlElement('w:fldChar'); fc.set(qn('w:fldCharType'),'begin'); r._r.append(fc)
    r2=par.add_run(); it=OxmlElement('w:instrText'); it.set(qn('xml:space'),'preserve'); it.text=code; r2._r.append(it)
    r3=par.add_run(); fc2=OxmlElement('w:fldChar'); fc2.set(qn('w:fldCharType'),'end'); r3._r.append(fc2)
def add_footer(doc):
    hoy=datetime.now()
    f=doc.sections[0].footer
    p=f.paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=p.add_run(f"Informe de Peligro de Incendios e Indicadores de Sequía – {STATION}\n")
    r.font.size=Pt(8); r.font.color.rgb=GRIS
    r2=p.add_run(f"{AUTOR} – {hoy.day} {MES_LARGO[hoy.month]} {hoy.year} – pág. ")
    r2.font.size=Pt(8); r2.font.color.rgb=GRIS
    _add_field(p,'PAGE')
    r3=p.add_run(" de "); r3.font.size=Pt(8); r3.font.color.rgb=GRIS
    _add_field(p,'NUMPAGES')
def parr(doc, text, size=10.5, bold=False, italic=False, color=None, align=None, after=6):
    p = doc.add_paragraph()
    r = p.add_run(text)
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.italic = italic
    if color: r.font.color.rgb = color
    if align: p.alignment = align
    p.paragraph_format.space_after = Pt(after)
    return p
def add_index_section(doc, img, title, desc, num, subtitulo=None):
    p=doc.add_paragraph(); p.paragraph_format.space_before=Pt(12); p.paragraph_format.space_after=Pt(4)
    r=p.add_run(f"{num}. {title}"); r.bold=True; r.font.size=Pt(12); r.font.color.rgb=AZUL
    if subtitulo:
        ps=doc.add_paragraph(); ps.paragraph_format.space_after=Pt(2)
        rs=ps.add_run(subtitulo); rs.font.size=Pt(9.5); rs.font.color.rgb=RGBColor(0x37,0x41,0x51)
    if desc:
        pd_=doc.add_paragraph(); pd_.paragraph_format.space_after=Pt(6)
        rd=pd_.add_run(desc); rd.font.size=Pt(9.5); rd.font.color.rgb=RGBColor(0x37,0x41,0x51)
    if img and os.path.exists(img):
        doc.add_picture(img, width=Inches(6.6))
        doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
def _set_fuente_base(doc, fuente="Arial"):
    normal = doc.styles["Normal"]
    normal.font.name = fuente
    rpr = normal.element.get_or_add_rPr()
    rFonts = rpr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rpr.append(rFonts)
    rFonts.set(qn('w:ascii'), fuente)
    rFonts.set(qn('w:hAnsi'), fuente)
    rFonts.set(qn('w:eastAsia'), fuente)
def build_word(dfa, imgs, meta):
    doc=docx.Document()
    _set_fuente_base(doc)
    for s in doc.sections:
        s.top_margin=s.bottom_margin=Inches(0.6); s.left_margin=s.right_margin=Inches(0.6)
    add_footer(doc)
    parr(doc, "Informe de Peligro de Incendios e Indicadores de Sequía", 18, True, color=RGBColor(0x11,0x18,0x27))
    parr(doc, f"Realizado con los datos extraídos de la Estación Meteorológica situada en "
              f"{STATION_DIR} ({STATION_URL}) — Últimos {meta['n']} días ({meta['rango']})",
         9, italic=True, color=GRIS)
    if meta.get('dias_interpolados', 0) > 0:
        parr(doc, f"{meta['dias_interpolados']} de los {meta['n']} días del periodo no tenían registro y se han "
                  f"estimado por interpolación para no distorsionar los índices acumulativos (KBDI, FWI, ET₀).",
             8.5, italic=True, color=GRIS, after=4)
    parr(doc, f"Resumen de valores a día {meta['hoy']:%d/%m/%Y}", 12, True, color=AZUL, after=4)
    kt=doc.add_table(rows=3, cols=7); kt.alignment=WD_TABLE_ALIGNMENT.CENTER
    headers=['DÍAS SIN LLUVIA\n(≥5 mm)','PRECIP. ACUM.\n(PERIODO)','ET₀ AYER\n(mm = L/m²)','KBDI','FWI','CBI','REGLA 30\n(3/3)']
    vals=[meta['dias_sin'], f"{meta['prec_acum']:.1f} mm", f"{meta['et0']:.1f}",
          f"{meta['kbdi']:.1f}", f"{meta['fwi']:.1f}", f"{meta['cbi']:.1f}", f"{meta['n33']} de {meta['n']}"]
    subs=[meta['sub_lluvia'], f"en {meta['n']} días", f"Acum.: {meta['et0_acum']:.0f} mm",
          kbdi_label(meta['kbdi']), fwi_label(meta['fwi']), cbi_label(meta['cbi']), 'Días extremos']
    for i in range(7):
        for row_i,(txt,size,bold,ital,col) in enumerate([
            (headers[i],6.5,True,False,(0x4B,0x55,0x63)),
            (vals[i],11,True,False,(0x99,0x1B,0x1B)),
            (subs[i],6.5,False,True,(0x6B,0x72,0x80))]):
            c=kt.cell(row_i,i); c.text=txt
            c.paragraphs[0].alignment=WD_ALIGN_PARAGRAPH.CENTER
            for rr in c.paragraphs[0].runs:
                rr.font.size=Pt(size); rr.font.bold=bold; rr.font.italic=ital; rr.font.color.rgb=RGBColor(*col)
            set_cell_bg(kt.cell(row_i,i),'FEF2F2')
    doc.add_paragraph()
    parr(doc, f"1. Evaluación y Evolución Gráfica (Últimos {meta['n']} días)", 12, True, color=AZUL)
    sub_ext = (f"Media anual de días extremos en los años completos ({meta['rango_ext']}): "
               f"{meta['media_ext']:.1f} días/año.") if meta['rango_ext'] else None
    sub_prec = (f"La media anual de precipitación en los años completos ({meta['rango_prec']}) "
                f"es de {meta['media_prec']:.0f} mm/año.") if meta['media_prec'] is not None else None
    sections=[
        ('regla3030', f"Regla 30-30-30 — {meta['n33']} días extremos de {meta['n']}", None),
        ('heatmap',   "Histórico de días extremos", sub_ext),
        ('kbdi',      f"KBDI — Valor actual: {meta['kbdi']:.1f} ({kbdi_label(meta['kbdi'])})", _txt_percentil(meta.get('pct_kbdi'))),
        ('fwi',       f"FWI — Valor actual: {meta['fwi']:.1f} ({fwi_label(meta['fwi'])})", _txt_percentil(meta.get('pct_fwi'))),
        ('cbi',       f"CBI — Valor actual: {meta['cbi']:.1f} ({cbi_label(meta['cbi'])})", _txt_percentil(meta.get('pct_cbi'))),
        ('angstrom',  f"Angström — Valor actual: {meta['angstrom']:.2f} ({angstrom_label(meta['angstrom'])})", _txt_percentil(meta.get('pct_angstrom'))),
        ('vpd',       f"VPD — Valor actual: {meta['vpd']:.1f} hPa ({vpd_label(meta['vpd'])})", _txt_percentil(meta.get('pct_vpd'))),
        ('rain',      f"Precipitación del periodo — {meta['prec_acum']:.1f} mm acumulados", None),
        ('precipmensual', "Histórico de precipitación acumulada mensual", sub_prec),
        ('et0',       f"Evapotranspiración de referencia (ET0) — ayer {meta['et0']:.1f} mm · acumulada {meta['et0_acum']:.0f} mm", None),
        ('windrose',  "Distribución del viento", None),
    ]
    for idx,(key,tit,sub) in enumerate(sections, start=1):
        add_index_section(doc, imgs.get(key), tit, DESCRIPCIONES.get(key,""), idx, subtitulo=sub)
    doc.add_page_break()
    parr(doc, f"2. Tabla de Valores Diarios (Últimos {meta['n']} días)", 12, True, color=AZUL)
    cols=['Fecha','T máx (°C)','RH mín (%)','Racha (km/h)','Lluvia (mm)','Regla 30','KBDI (0-800)','FWI (índice)','CBI (índice)']
    tbl=doc.add_table(rows=len(dfa)+1, cols=len(cols)); tbl.alignment=WD_TABLE_ALIGNMENT.CENTER
    for i,h in enumerate(cols):
        c=tbl.cell(0,i); c.text=h; c.paragraphs[0].alignment=WD_ALIGN_PARAGRAPH.CENTER
        set_cell_bg(c,'1E293B')
        for rr in c.paragraphs[0].runs:
            rr.font.size=Pt(7); rr.font.bold=True; rr.font.color.rgb=RGBColor(0xFF,0xFF,0xFF)
    for ri in range(len(dfa)):
        row=dfa.iloc[ri]; prev=dfa.iloc[ri-1] if ri>0 else None
        h_val=row['Humidity_Min (%)']
        h_str=(f"{h_val:.0f}%{get_arrow(h_val, prev['Humidity_Min (%)'] if prev is not None else None)}"
               if not pd.isna(h_val) else "N/A")
        vals=[row['Fecha'].strftime('%d/%m/%Y'),
              f"{row['Temperatura Máxima (°C)']:.1f}{get_arrow(row['Temperatura Máxima (°C)'], prev['Temperatura Máxima (°C)'] if prev is not None else None)}",
              h_str,
              f"{row['wind_gust']:.1f}{get_arrow(row['wind_gust'], prev['wind_gust'] if prev is not None else None)}",
              f"{row['Precipitación Total (mm)']:.1f}", row['Regla_30'],
              f"{row['KBDI']:.0f}{get_arrow(row['KBDI'], prev['KBDI'] if prev is not None else None)}",
              f"{row['FWI']:.1f}{get_arrow(row['FWI'], prev['FWI'] if prev is not None else None)}",
              f"{row['CBI']:.1f}{get_arrow(row['CBI'], prev['CBI'] if prev is not None else None)}"]
        bg='F8FAFC' if (ri+1)%2==0 else 'FFFFFF'
        for ci,v in enumerate(vals):
            c=tbl.cell(ri+1,ci); c.text=str(v); c.paragraphs[0].alignment=WD_ALIGN_PARAGRAPH.CENTER
            set_cell_bg(c,bg)
            for rr in c.paragraphs[0].runs:
                rr.font.size=Pt(7)
                if ci==5 and 'Extremo' in str(v):
                    rr.font.bold=True; rr.font.color.rgb=ROJO
    return doc
# =============================================================================
# MAIN
# =============================================================================
def main():
    ap=argparse.ArgumentParser(description="Genera informe de peligro de incendios.")
    ap.add_argument("--days", type=int, default=None,
                    help=f"Días a incluir en el informe (por defecto {DEFAULT_DAYS} si no se indica y no se usa --interactive).")
    ap.add_argument("--input", default=INPUT_XLSX)
    ap.add_argument("--output", default=None, help="Ruta del .docx de salida (por defecto se genera junto al script).")
    ap.add_argument("--interactive", action="store_true",
                    help="Si no se indica --days, pregunta por terminal en vez de usar el valor por defecto.")
    args=ap.parse_args()
    df=load_data(args.input)
    if KBDI_RAIN_MM_ENV is not None:
        kbdi_rain_mm = float(KBDI_RAIN_MM_ENV)
        origen_lluvia = "fijado por variable de entorno KBDI_ANNUAL_RAIN_MM"
    else:
        media_prec_kbdi, rango_prec_kbdi = media_anios_completos(df)
        if media_prec_kbdi is not None:
            kbdi_rain_mm = media_prec_kbdi
            origen_lluvia = f"media de los años completos del histórico ({rango_prec_kbdi})"
        else:
            kbdi_rain_mm = 478.275
            origen_lluvia = "valor por defecto (sin años completos aún en el histórico)"
    print(f"   Lluvia anual de referencia para KBDI: {kbdi_rain_mm:.0f} mm ({origen_lluvia})")
    print("🧮 Calculando índices...")
    df['ET0_mm']=calc_et0(df)
    df['KBDI']=calculate_kbdi(df, kbdi_rain_mm)
    df['FWI']=run_fwi_system(df)
    df['CBI']=[calc_cbi(r['Temperatura Máxima (°C)'],
                        r['Humidity_Min (%)'] if not pd.isna(r['Humidity_Min (%)']) else 30.0,
                        r['wind_gust']) for _,r in df.iterrows()]
    df['Angstrom']=[calc_angstrom(r['Temperatura Máxima (°C)'],
                        r['Humidity_Min (%)'] if not pd.isna(r['Humidity_Min (%)']) else 30.0)
                    for _,r in df.iterrows()]
    df['Regla_30']=df.apply(eval_regla_30, axis=1)
    dfc=df.iloc[:-1].copy().reset_index(drop=True)
    if args.days is not None:
        n=args.days
    elif args.interactive:
        while True:
            try:
                n=int(input(f"📊 Días a analizar (máx {len(dfc)}): "))
                if n>0: break
            except ValueError: print("   ⚠️ Introduce un entero.")
    else:
        n=DEFAULT_DAYS
        print(f"   (sin --days ni --interactive: se usan los últimos {n} días por defecto)")
    n=min(n, len(dfc))
    dfa=dfc.tail(n).copy().reset_index(drop=True)
    dfa['T30']=dfa['Temperatura Máxima (°C)']>30
    dfa['H30']=dfa['Humidity_Min (%)']<30
    dfa['V30']=dfa['wind_gust']>30
    last=dfa['Fecha'].iloc[-1]; first=dfa['Fecha'].iloc[0]
    ev=dfc[(dfc['Fecha']<=last)&(dfc['Precipitación Total (mm)']>=5.0)]
    if not ev.empty:
        lr=ev['Fecha'].max(); dias_sin=f"{(last-lr).days} días"; sub=f"Desde el {lr.strftime('%d/%m')}"
    else: dias_sin="N/A"; sub="Sin lluvias efectivas"
    media_ext, rango_ext = media_anios_validos(df)
    media_prec, rango_prec = media_anios_completos(df)
    meta={'n':n,'rango':f"{first:%d/%m/%Y} - {last:%d/%m/%Y}",'hoy':datetime.now(),
          'dias_sin':dias_sin,'sub_lluvia':sub,
          'prec_acum':float(dfa['Precipitación Total (mm)'].sum()),
          'et0':float(dfa['ET0_mm'].iloc[-1]),'et0_acum':float(dfa['ET0_mm'].sum()),
          'kbdi':dfa['KBDI'].iloc[-1],'fwi':dfa['FWI'].iloc[-1],'cbi':dfa['CBI'].iloc[-1],
          'angstrom':dfa['Angstrom'].iloc[-1],
          'vpd':dfa['Vpd_Max (hPa)'].iloc[-1] if 'Vpd_Max (hPa)' in dfa else 0.0,
          'n33':int((dfa['Regla_30']=='Extremo (3/3)').sum()),
          'media_ext':media_ext,'rango_ext':rango_ext,
          'media_prec':media_prec,'rango_prec':rango_prec,
          'pct_kbdi':percentil_historico(df['KBDI'], dfa['KBDI'].iloc[-1]),
          'pct_fwi':percentil_historico(df['FWI'], dfa['FWI'].iloc[-1]),
          'pct_cbi':percentil_historico(df['CBI'], dfa['CBI'].iloc[-1]),
          'pct_angstrom':percentil_historico(df['Angstrom'], dfa['Angstrom'].iloc[-1]),
          'pct_vpd':percentil_historico(df['Vpd_Max (hPa)'], dfa['Vpd_Max (hPa)'].iloc[-1]) if 'Vpd_Max (hPa)' in dfa else None,
          'kbdi_rain_mm':kbdi_rain_mm, 'origen_lluvia':origen_lluvia,
          'dias_interpolados':int(dfa['Interpolado'].sum()) if 'Interpolado' in dfa else 0}
    print("🎨 Generando gráficas (PNG permanentes)...")
    imgs={
        'regla3030': plot_regla3030_tabla(dfa),
        'heatmap': plot_heatmap(df),
        'precipmensual': plot_precip_mensual(pivot_mensual(df)),
        'kbdi': plot_index(dfa,'KBDI','KBDI','Índice Keetch-Byram (KBDI)',
                 [(0,200,'#2ca02c',0.12,'Bajo'),(200,400,'#ff7f0e',0.12,'Moderada'),
                  (400,600,'#d62728',0.15,'Severa'),(600,800,'#9467bd',0.15,'Extrema')],
                 kbdi_label,'kbdi',ylim=(0,850),color='#d95f02'),
        'fwi': plot_index(dfa,'FWI','FWI','Fire Weather Index (FWI - CFFDRS)',
                 [(0,11.2,'#2ca02c',0.12,'Bajo'),(11.2,21.3,'#bcbd22',0.12,'Moderado'),
                  (21.3,38.0,'#ff7f0e',0.12,'Alto'),(38.0,50.0,'#e377c2',0.15,'Muy Alto'),
                  (50.0,150,'#d62728',0.18,'Extremo')], fwi_label,'fwi',color='#ce1256'),
        'cbi': plot_index(dfa,'CBI','CBI','Índice de Combustión de Chandler (CBI)',
                 [(0,50,'#2ca02c',0.12,'Bajo'),(50,75,'#ff7f0e',0.15,'Moderado'),
                  (75,90,'#e377c2',0.15,'Alto'),(90,125,'#d62728',0.15,'Extremo')],
                 cbi_label,'cbi',ylim=(0,125),color='#02818a'),
        'angstrom': plot_index(dfa,'Angstrom','Angström','Índice de Angström',
                 [(-3,2.0,'#d62728',0.15,'Peligro extremo'),(2.0,2.5,'#ff7f0e',0.12,'Peligro alto'),
                  (2.5,3.0,'#bcbd22',0.12,'Peligro moderado'),(3.0,8.0,'#2ca02c',0.12,'Peligro bajo')],
                 angstrom_label,'angstrom',color='#756bb1'),
        'vpd': plot_index(dfa,'Vpd_Max (hPa)','VPD (hPa)','Déficit de presión de vapor (VPD)',
                 [(0,15,'#2ca02c',0.10,'Bajo'),(15,25,'#ff7f0e',0.12,'Moderado'),
                  (25,35,'#d62728',0.12,'Alto'),(35,60,'#8b0000',0.14,'Extremo')],
                 vpd_label,'vpd',color='#e31a1c') if 'Vpd_Max (hPa)' in dfa and dfa['Vpd_Max (hPa)'].notna().any() else None,
        'rain': plot_rain(dfa),
        'et0': plot_et0(dfa),
        'windrose': plot_wind_rose(dfa),
    }
    print("📝 Generando Word...")
    doc=build_word(dfa, imgs, meta)
    if args.output:
        fname=args.output
    else:
        slug=re.sub(r'[^A-Za-z0-9]+','_',STATION).strip('_')
        fname=os.path.join(BASE_DIR, f"Informe_Incendios_{slug}_{n}d_{last:%Y%m%d}.docx")
    doc.save(fname); print(f"   ✅ Word generado: {fname}")
    csv_path=os.path.join(BASE_DIR, "incendios_calculo.csv")
    cols_csv=['Fecha','Temperatura Máxima (°C)','Temperatura Mínima (°C)','Humidity_Min (%)',
              'wind_gust','Precipitación Total (mm)','ET0_mm','KBDI','FWI','CBI','Angstrom',
              'Vpd_Max (hPa)','Regla_30','Interpolado','Precip_estimada']
    cols_csv=[c for c in cols_csv if c in df.columns]
    df[cols_csv].to_csv(csv_path, index=False)
    print(f"   ✅ CSV generado: {csv_path}")
    print("🎉 Informe completado.")
if __name__ == "__main__":
    main()