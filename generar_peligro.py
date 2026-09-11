import json
import requests

from io import BytesIO
from datetime import datetime, timezone

from openpyxl import load_workbook

URL_XLSX = (
"https://gencat.cat/medinatural/incendis/mapes/"
"taula_muni_perill_avui.xlsx"
)

ARCHIVO_SALIDA = "peligro.json"

def descargar_xlsx():
print("Descargando XLSX oficial...")

```
marca_tiempo = int(
    datetime.now(timezone.utc).timestamp()
)

url = (
    URL_XLSX
    + "?nocache="
    + str(marca_tiempo)
)

print("URL de descarga:")
print(url)

respuesta = requests.get(
    url,
    headers={
        "Cache-Control": "no-cache, no-store, must-revalidate",
        "Pragma": "no-cache",
        "Expires": "0",
        "User-Agent": "Mozilla/5.0 PlaAlfaESP32"
    },
    timeout=30
)

print(
    "HTTP:",
    respuesta.status_code
)

print(
    "Bytes recibidos:",
    len(respuesta.content)
)

if respuesta.status_code != 200:
    raise RuntimeError(
        "No se ha podido descargar el XLSX"
    )

return respuesta.content
```

def normalizar_codigo(codigo):
if codigo is None:
return ""

```
texto = str(codigo).strip()

if texto.isdigit():
    texto = texto.zfill(5)

return texto
```

def generar_json():
contenido = descargar_xlsx()

```
libro = load_workbook(
    filename=BytesIO(contenido),
    read_only=True,
    data_only=True
)

hoja = libro.active

municipios = {}

fecha_actualizacion = ""

for fila in hoja.iter_rows(
    min_row=4,
    max_row=950,
    min_col=2,
    max_col=6,
    values_only=True
):
    codigo = fila[0]
    nombre = fila[1]
    comarca = fila[2]
    peligro = fila[3]
    fecha = fila[4]

    if nombre is None:
        continue

    codigo = normalizar_codigo(codigo)

    nombre = str(nombre).strip()

    comarca = (
        str(comarca).strip()
        if comarca is not None
        else ""
    )

    peligro = (
        str(peligro).strip()
        if peligro is not None
        else ""
    )

    if isinstance(fecha, datetime):
        fecha_texto = fecha.strftime("%d/%m/%y")

    elif fecha is not None:
        fecha_texto = str(fecha).strip()

    else:
        fecha_texto = ""

    if fecha_texto:
        fecha_actualizacion = fecha_texto

    if not codigo:
        continue

    municipios[codigo] = {
        "municipio": nombre,
        "comarca": comarca,
        "peligro": peligro,
        "fecha": fecha_texto
    }

libro.close()

resultado = {
    "fecha": fecha_actualizacion,
    "municipios": municipios
}

with open(
    ARCHIVO_SALIDA,
    "w",
    encoding="utf-8"
) as archivo:
    json.dump(
        resultado,
        archivo,
        ensure_ascii=False,
        indent=2
    )

print()
print("======================================")
print("RESULTADO ACTUALIZACION")
print("======================================")

print(
    "Fecha detectada:",
    fecha_actualizacion
)

print(
    "Municipios generados:",
    len(municipios)
)

for codigo in (
    "08148",
    "08270",
    "08305"
):
    if codigo in municipios:
        datos = municipios[codigo]

        print(
            datos["municipio"],
            "|",
            datos["peligro"],
            "|",
            datos["fecha"]
        )

print(
    "Archivo creado:",
    ARCHIVO_SALIDA
)

print("======================================")
```

if **name** == "**main**":
generar_json()
