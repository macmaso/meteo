# Estación Los Tomillares — Panel en vivo

App de Streamlit con los datos e índices de la estación (temperatura y
peligro de incendio/sequía) actualizados hasta el día anterior a la
consulta. Reutiliza directamente `analisis_meteorologico.py` e
`informe_incendios.py` — la misma lógica de cálculo que los informes en
Word, sin duplicar nada.

## Cómo funciona la actualización diaria

1. `historico_base.xlsx` es tu histórico actual, empaquetado en el repositorio
   como punto de partida (no hace falta que esté al día: la app rellena
   automáticamente los días que falten).
2. Al abrir la app (o cada 24h, gracias a la caché), `data_fetch.py` mira
   qué día es el último del histórico base y pide a la API de Ecowitt los
   días que falten hasta ayer.
3. Si no hay credenciales configuradas, la app sigue funcionando y muestra
   el histórico base tal cual (con un aviso), en vez de romperse.

Esto significa que **no necesitas volver a subir el Excel a mano**: solo
tienes que configurar las credenciales una vez (paso 4 más abajo).

## Desplegar en Streamlit Community Cloud (gratis)

### 1. Crear el repositorio en GitHub
- Crea un repositorio nuevo (puede ser privado; Streamlit Community Cloud
  puede desplegar desde repos privados si conectas tu cuenta de GitHub).
- Sube todos los archivos de esta carpeta **excepto** `.streamlit/secrets.toml`
  si llegas a crearlo en local (el `.gitignore` ya lo excluye).

```bash
cd estacion-tomillares-web
git init
git add .
git commit -m "Primera versión del panel en vivo"
git branch -M main
git remote add origin https://github.com/TU_USUARIO/estacion-tomillares-web.git
git push -u origin main
```

### 2. Conseguir tus claves de Ecowitt
Entra en <https://www.ecowitt.net/user/index> (con la misma cuenta que usa
tu estación), busca el apartado de API y genera:
- `Application Key`
- `API Key`
- Apunta también la **MAC o IMEI** de tu panel (GW1000/GW1100), visible en
  los ajustes del dispositivo.

### 3. Crear la app en Streamlit Community Cloud
- Entra en <https://share.streamlit.io> con tu cuenta de GitHub.
- "New app" → selecciona el repositorio, la rama `main` y el archivo
  principal `app.py`.
- Despliega. La primera vez tardará un minuto en instalar dependencias.

### 4. Configurar las credenciales (Secrets)
En el panel de tu app, dentro de Streamlit Community Cloud:
`Settings → Secrets`, y pega:

```toml
ECOWITT_APPLICATION_KEY = "tu_application_key"
ECOWITT_API_KEY = "tu_api_key"
ECOWITT_MAC = "AA:BB:CC:DD:EE:FF"
```

Guarda — la app se reinicia sola y a partir de ahí ya puede completar el
histórico por su cuenta cada día.

### 5. Compartir el enlace
Streamlit te da una URL pública del tipo
`https://TU-APP.streamlit.app`. Cualquiera con el enlace puede consultarla;
no hace falta que tengan cuenta de nada.

## Probar en local antes de desplegar

```bash
pip install -r requirements.txt
cp .streamlit/secrets.toml.example .streamlit/secrets.toml   # y rellena tus claves
streamlit run app.py
```

Si no rellenas `secrets.toml`, la app arranca igualmente y muestra el
histórico base con un aviso de que faltan credenciales.

## Nota de privacidad

Por defecto la app muestra la dirección exacta de la estación (la misma
que ya aparece en los informes en Word). Al ser ahora una página pública en
vez de un documento que tú compartes, es buen momento para decidir
conscientemente si quieres que siga así. Para ocultarla, en Secrets añade:

```toml
MOSTRAR_DIRECCION = "0"
```

## Actualizar el histórico base

Cuando quieras "fijar" el histórico base con datos más recientes (por
ejemplo, una vez al año), simplemente sustituye `historico_base.xlsx` por
tu Excel actualizado y vuelve a subirlo al repositorio — la app seguirá
completando automáticamente lo que falte desde ahí hasta ayer.

## Limitaciones conocidas

- Streamlit Community Cloud (plan gratuito) "duerme" la app tras un tiempo
  sin visitas; el primer visitante tras el sueño esperará unos segundos
  mientras arranca.
- La API de Ecowitt conserva el detalle horario solo un tiempo limitado
  (meses recientes a alta resolución, más atrás a resolución menor) — por
  eso el histórico base en el repositorio es importante: es la memoria
  larga que la API por sí sola no garantiza.
- El mapeo de campos de la API (`data_fetch.py`, función
  `_parse_history_response`) sigue la estructura estándar de la API v3 de
  Ecowitt; si tu estación tiene sensores distintos, puede que algún nombre
  de campo cambie — la primera vez que actives las credenciales, revisa el
  aviso de "datos actualizados" en la app para confirmar que ha
  incorporado los días nuevos correctamente.
