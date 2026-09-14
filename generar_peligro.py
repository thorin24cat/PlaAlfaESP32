import json
import re
import requests

from html.parser import HTMLParser
from io import BytesIO
from datetime import datetime, timezone

from openpyxl import load_workbook


URL_HTML = (
    "https://gencat.cat/medinatural/incendis/mapes/"
    "taula_muni_perill_avui.html"
)

# Se mantiene únicamente para obtener la correspondencia
# municipio -> código oficial.
# Los datos de peligro YA NO se obtienen del XLSX.
URL_XLSX_CODIGOS = (
    "https://gencat.cat/medinatural/incendis/mapes/"
    "taula_muni_perill_avui.xlsx"
)

ARCHIVO_SALIDA = "peligro.json"


# ============================================================
# UTILIDADES
# ============================================================

def limpiar_texto(texto):
    if texto is None:
        return ""

    texto = str(texto)

    texto = texto.replace("\xa0", " ")
    texto = re.sub(r"\s+", " ", texto)

    return texto.strip()


def normalizar_nombre(nombre):
    texto = limpiar_texto(nombre).lower()

    reemplazos = {
        "à": "a",
        "á": "a",
        "ä": "a",
        "â": "a",
        "è": "e",
        "é": "e",
        "ë": "e",
        "ê": "e",
        "ì": "i",
        "í": "i",
        "ï": "i",
        "î": "i",
        "ò": "o",
        "ó": "o",
        "ö": "o",
        "ô": "o",
        "ù": "u",
        "ú": "u",
        "ü": "u",
        "û": "u",
        "ç": "c"
    }

    for origen, destino in reemplazos.items():
        texto = texto.replace(origen, destino)

    return texto


def normalizar_codigo(codigo):
    if codigo is None:
        return ""

    texto = str(codigo).strip()

    if texto.isdigit():
        texto = texto.zfill(5)

    return texto


# ============================================================
# PARSER HTML
# ============================================================

class TablaHTMLParser(HTMLParser):

    def __init__(self):
        super().__init__()

        self.filas = []

        self.en_fila = False
        self.en_celda = False

        self.fila_actual = []
        self.celda_actual = ""

    def handle_starttag(self, tag, attrs):

        tag = tag.lower()

        if tag == "tr":

            self.en_fila = True
            self.fila_actual = []

        elif tag in ("td", "th") and self.en_fila:

            self.en_celda = True
            self.celda_actual = ""

    def handle_endtag(self, tag):

        tag = tag.lower()

        if tag in ("td", "th"):

            if self.en_celda:

                texto = limpiar_texto(
                    self.celda_actual
                )

                self.fila_actual.append(
                    texto
                )

                self.celda_actual = ""
                self.en_celda = False

        elif tag == "tr":

            if self.en_fila:

                if self.fila_actual:

                    self.filas.append(
                        self.fila_actual
                    )

                self.fila_actual = []
                self.en_fila = False

    def handle_data(self, data):

        if self.en_celda:

            self.celda_actual += data


# ============================================================
# DESCARGAR HTML OFICIAL
# ============================================================

def descargar_html():

    print()
    print("======================================")
    print("DESCARGA HTML OFICIAL")
    print("======================================")

    marca_tiempo = int(
        datetime.now(timezone.utc).timestamp()
    )

    url = (
        URL_HTML
        + "?nocache="
        + str(marca_tiempo)
    )

    print("URL:")
    print(url)

    respuesta = requests.get(
        url,
        headers={
            "Cache-Control":
                "no-cache, no-store, must-revalidate",
            "Pragma": "no-cache",
            "Expires": "0",
            "User-Agent":
                "Mozilla/5.0 PlaAlfaESP32"
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
            "No se ha podido descargar "
            "la tabla HTML oficial"
        )

    return respuesta.content


# ============================================================
# LEER TABLA HTML
# ============================================================

def leer_tabla_html(contenido):

    parser = TablaHTMLParser()

    parser.feed(
        contenido.decode(
            "utf-8",
            errors="replace"
        )
    )

    print()
    print(
        "Filas HTML detectadas:",
        len(parser.filas)
    )

    if not parser.filas:

        raise RuntimeError(
            "No se han encontrado filas "
            "en la tabla HTML"
        )

    # --------------------------------------------------------
    # Buscar la fila de cabecera.
    # --------------------------------------------------------

    indice_cabecera = -1

    for i, fila in enumerate(
        parser.filas
    ):

        texto_fila = " ".join(
            fila
        ).upper()

        if (
            "MUNICIPI" in texto_fila
            and "COMARCA" in texto_fila
            and "PERILL" in texto_fila
        ):

            indice_cabecera = i
            break

    if indice_cabecera < 0:

        raise RuntimeError(
            "No se ha encontrado la cabecera "
            "MUNICIPI / COMARCA / PERILL"
        )

    cabecera = parser.filas[
        indice_cabecera
    ]

    print()
    print(
        "Cabecera detectada:"
    )
    print(
        cabecera
    )

    # --------------------------------------------------------
    # Detectar posiciones de columnas.
    # --------------------------------------------------------

    columnas = {}

    for i, nombre in enumerate(
        cabecera
    ):

        nombre_normalizado = (
            limpiar_texto(nombre)
            .upper()
        )

        if "MUNICIPI" in nombre_normalizado:
            columnas["municipio"] = i

        elif "COMARCA" in nombre_normalizado:
            columnas["comarca"] = i

        elif "PERILL" in nombre_normalizado:
            columnas["peligro"] = i

        elif "DATA" in nombre_normalizado:
            columnas["fecha"] = i

    print()
    print(
        "Columnas detectadas:"
    )
    print(
        columnas
    )

    necesarias = (
        "municipio",
        "comarca",
        "peligro"
    )

    for columna in necesarias:

        if columna not in columnas:

            raise RuntimeError(
                "No se ha encontrado la columna "
                + columna
                + " en la tabla HTML"
            )

    # --------------------------------------------------------
    # Convertir filas.
    # --------------------------------------------------------

    registros = []

    for fila in parser.filas[
        indice_cabecera + 1:
    ]:

        if len(fila) <= columnas["municipio"]:
            continue

        municipio = limpiar_texto(
            fila[
                columnas["municipio"]
            ]
        )

        if not municipio:
            continue

        comarca = ""

        if (
            "comarca" in columnas
            and len(fila) > columnas["comarca"]
        ):

            comarca = limpiar_texto(
                fila[
                    columnas["comarca"]
                ]
            )

        peligro = ""

        if (
            "peligro" in columnas
            and len(fila) > columnas["peligro"]
        ):

            peligro = limpiar_texto(
                fila[
                    columnas["peligro"]
                ]
            )

        fecha = ""

        if (
            "fecha" in columnas
            and len(fila) > columnas["fecha"]
        ):

            fecha = limpiar_texto(
                fila[
                    columnas["fecha"]
                ]
            )

        # Evitar filas que no sean municipios.
        if municipio.upper() in (
            "MUNICIPI",
            "TOTAL",
            "TOTAL MUNICIPIS"
        ):
            continue

        registros.append(
            {
                "municipio": municipio,
                "comarca": comarca,
                "peligro": peligro,
                "fecha": fecha
            }
        )

    print()
    print(
        "Municipios encontrados en HTML:",
        len(registros)
    )

    return registros


# ============================================================
# DESCARGAR CÓDIGOS MUNICIPALES
# ============================================================

def descargar_mapa_codigos():

    print()
    print("======================================")
    print("DESCARGA TABLA DE CÓDIGOS")
    print("======================================")

    marca_tiempo = int(
        datetime.now(timezone.utc).timestamp()
    )

    url = (
        URL_XLSX_CODIGOS
        + "?nocache="
        + str(marca_tiempo)
    )

    respuesta = requests.get(
        url,
        headers={
            "Cache-Control":
                "no-cache, no-store, must-revalidate",
            "Pragma": "no-cache",
            "Expires": "0",
            "User-Agent":
                "Mozilla/5.0 PlaAlfaESP32"
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
            "No se ha podido descargar "
            "la tabla de códigos municipales"
        )

    return respuesta.content


def obtener_codigos_municipios():

    contenido = descargar_mapa_codigos()

    libro = load_workbook(
        filename=BytesIO(contenido),
        read_only=True,
        data_only=True
    )

    hoja = libro.active

    codigos = {}

    for fila in hoja.iter_rows(
        min_row=4,
        max_row=950,
        min_col=2,
        max_col=3,
        values_only=True
    ):

        codigo = fila[0]
        nombre = fila[1]

        if nombre is None:
            continue

        codigo = normalizar_codigo(
            codigo
        )

        nombre = limpiar_texto(
            nombre
        )

        if not codigo or not nombre:
            continue

        codigos[
            normalizar_nombre(nombre)
        ] = codigo

    libro.close()

    print()
    print(
        "Municipios con código:",
        len(codigos)
    )

    return codigos


# ============================================================
# GENERAR JSON
# ============================================================

def generar_json():

    # --------------------------------------------------------
    # 1. HTML = fuente REAL del peligro actual.
    # --------------------------------------------------------

    contenido_html = descargar_html()

    registros = leer_tabla_html(
        contenido_html
    )

    # --------------------------------------------------------
    # 2. XLSX = solamente correspondencia nombre -> código.
    # --------------------------------------------------------

    codigos = obtener_codigos_municipios()

    municipios = {}

    fecha_actualizacion = ""

    for registro in registros:

        nombre = registro[
            "municipio"
        ]

        nombre_normalizado = (
            normalizar_nombre(
                nombre
            )
        )

        codigo = codigos.get(
            nombre_normalizado,
            ""
        )

        if not codigo:

            print(
                "AVISO: municipio sin código:",
                nombre
            )

            continue

        peligro = registro[
            "peligro"
        ]

        comarca = registro[
            "comarca"
        ]

        fecha = registro[
            "fecha"
        ]

        if fecha:
            fecha_actualizacion = fecha

        municipios[codigo] = {

            "municipio": nombre,

            "comarca": comarca,

            "peligro": peligro,

            "fecha": fecha
        }

    # --------------------------------------------------------
    # Comprobar que tenemos los municipios que utiliza
    # actualmente el proyecto.
    # --------------------------------------------------------

    print()
    print(
        "======================================"
    )
    print(
        "COMPROBACION MUNICIPIOS"
    )
    print(
        "======================================"
    )

    for codigo in (
        "08148",
        "08270",
        "08305"
    ):

        if codigo in municipios:

            datos = municipios[
                codigo
            ]

            print(
                datos["municipio"],
                "|",
                datos["peligro"],
                "|",
                datos["fecha"],
                "|",
                codigo
            )

        else:

            print(
                "ERROR: no encontrado:",
                codigo
            )

    # --------------------------------------------------------
    # Comprobar que no hemos perdido demasiados municipios.
    # --------------------------------------------------------

    print()
    print(
        "Municipios HTML:",
        len(registros)
    )

    print(
        "Municipios con código:",
        len(municipios)
    )

    if len(municipios) < 900:

        raise RuntimeError(
            "Se han generado menos de 900 "
            "municipios con código. "
            "Se detiene la actualización "
            "para no publicar un archivo "
            "incompleto."
        )

    # --------------------------------------------------------
    # Guardar resultado.
    # --------------------------------------------------------

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
    print(
        "======================================"
    )
    print(
        "RESULTADO ACTUALIZACION"
    )
    print(
        "======================================"
    )

    print(
        "Fecha detectada:",
        fecha_actualizacion
    )

    print(
        "Municipios HTML:",
        len(registros)
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

            datos = municipios[
                codigo
            ]

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

    print(
        "======================================"
    )


# ============================================================
# PROGRAMA PRINCIPAL
# ============================================================

if __name__ == "__main__":
    generar_json()
