# 🏪 Simulador de canibalización de tiendas

App Streamlit que cruza la capa DANE (personas y viviendas por manzana), las tiendas operando y el histórico de ventas para estimar cuánto canibaliza una tienda nueva y generar un informe PDF.

## Qué hace
1. Lee `Book.xlsx` (tiendas) y `LibroVentas.xlsx` (ventas) desde OneDrive/SharePoint (Microsoft Graph) o por carga manual.
2. Capa DANE: personas y viviendas por manzana (ponderadas por área dentro de cada isócrona).
3. Tiendas nuevas: se digitan con coordenada (tabla editable).
4. Isócrona de **300 m** (editable): caminando real con OpenRouteService, o círculo simulado.
5. Muestra distancia a cada tienda vecina, % de traslape de población y **% de canibalización sugerido** (traslape × tasa de migración), **editable** por tienda.
6. Impacto en **ventas y contribución de los últimos 3 meses** (en pesos y %), y gráfico de barras de **ventas de 6 meses**.
7. Caja de comentarios + **PDF** con mapa estilo OpenStreetMap, tablas, gráfico y comentarios.
8. (Opcional) Cargar isócronas **reales** de encuestas origen-destino (GeoJSON con columna `id`).

## Instalación local
```bash
python -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
cp .streamlit/secrets.toml.example .streamlit/secrets.toml   # editar
streamlit run app.py
```

## Subir a GitHub
```bash
git init && git add . && git commit -m "App canibalización"
git branch -M main
git remote add origin https://github.com/TU_USUARIO/canibalizacion-app.git
git push -u origin main
```
Luego despliega en https://share.streamlit.io (Community Cloud) y pega el contenido de `secrets.toml` en *Settings → Secrets*. Si el repo es público, **nunca** subas secretos ni la capa DANE si es sensible.

## Conectar OneDrive/SharePoint
En *Settings → Secrets* (o `.streamlit/secrets.toml` en local) agrega:
```toml
ONEDRIVE_BOOK_URL = "https://...Book.xlsx link compartido..."
ONEDRIVE_SALES_URL = "https://...LibroVentas.xlsx link compartido..."
ONEDRIVE_LAYER_URL = "https://...VIV&HAB.zip link compartido..."
```
La app descarga cada archivo con el link y conserva la capa en caché durante una hora. La capa nacional no se convierte completa en memoria: después de ingresar la coordenada, se lee solo una ventana cercana a la tienda nueva. **Requisito:** el link debe estar compartido como *"Cualquier persona con el vínculo puede ver"*. Si el link es solo de tu organización, la descarga directa falla y hay que usar un enlace público o el respaldo opcional con Microsoft Graph (`[azure]` tenant_id, client_id, client_secret). La capa local en `data/` tiene prioridad sobre OneDrive.

## Capa DANE
Descarga del Geoportal DANE el **Marco Geoestadístico / CNPV 2018 a nivel de manzana** de tu departamento o ciudad (evita el archivo nacional, es muy pesado). Guárdalo en `data/` (`.gpkg`, `.geojson` o `.zip` con shapefile) o súbelo desde la app. Elige en el mapeo las columnas de personas y viviendas; para la capa `VIVIENDAS_Y_HABITANTES` la app reconoce automáticamente `TP27_PERSO` y `TVIVIENDA`.

El archivo `.lyrx` recibido (`VIVIENDAS_Y_HABITANTES_COLOMBIA.lyrx`) es una **definición de capa de ArcGIS**: contiene simbología, nombres de campos y una referencia a una geodatabase local, pero no contiene los polígonos ni los valores. Para usarlo en la app hay que exportar desde ArcGIS Pro la capa como **GeoPackage**, **GeoJSON** o **ZIP de shapefile**, incluyendo geometría y atributos. La app acepta el `.lyrx` para mostrar una advertencia explicativa, pero no puede calcular isócronas con ese archivo solo.

## Isócronas caminando reales
Crea una API key gratuita en openrouteservice.org y ponla en `[ors] api_key`. Sin key usa círculo de 300 m.

## Supuestos
- Canibalización sugerida = % de la población de la isócrona de la tienda existente que queda compartida con la nueva × tasa de migración (slider, por defecto 30 %). Calíbrala con aperturas históricas.
- "Impacto en precio" se entrega en **pesos (COP) y %** para ventas y contribución.
- Los últimos 3/6 meses se calculan hacia atrás desde el último mes con datos en la base de ventas.

## Si no carga el Excel
Abre el panel **🔎 Diagnóstico de secrets** en la barra lateral:
- ❌ junto a una clave: la clave no se detectó. En TOML las claves `ONEDRIVE_*` deben ir **antes** de cualquier `[tabla]` (por ejemplo `[ors]`), si no quedan dentro de esa tabla.
- Error "pide inicio de sesión": el link no está como *Cualquier persona con el vínculo*. Streamlit Cloud está fuera de tu organización y no puede iniciar sesión; la única vía automática en ese caso es el registro de app en Azure (`[azure]`).
