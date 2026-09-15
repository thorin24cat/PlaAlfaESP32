# ============================================================
# PLA ALFA ESP32
# GENERADOR DE PELIGRO
#
# VERSION: 4.0.0
#
# Fuente del peligro:
#   Gencat - tabla municipal AVUI
#
# La página "Avui" utiliza un htmlwidget/DataTables.
# Los datos de los 947 municipios están dentro de un bloque:
#
# <script type="application/json" data-for="...">
#
# y concretamente en:
#
# x -> data
#
# Las columnas son:
#
#   0 = MUNICIPI
#   1 = COMARCA
#   2 = PERILL
#   3 = DATA
#
# El XLSX se utiliza SOLO para obtener:
#
#   municipio -> código municipal
#
# ============================================================

import json
import re
import requests

from datetime import datetime, timezone
from io import BytesIO

from openpyxl import load_workbook


VERSION = "4.0.0"


# ============================================================
# FUENTES
# ============================================================

URL_HTML = (
    "https://gencat.cat/medinatural/incendis/mapes/"
    "taula_muni_perill_avui.html"
)

URL_XLSX_CODIGOS = (
    "https://gencat.cat/medinatural/incendis/mapes/"
    "taula_muni_perill_avui.xlsx"
)

ARCHIVO_SALIDA = "peligro.json"


# ============================================================
# MUNICIPIOS DEL PROYECTO
# ============================================================

MUNICIPIOS_PROYECTO = {
    "08148": "Olivella",
    "08270": "Sitges",
    "08305": "Vilafranca del Penedès"
}


# ============================================================
# UTILIDADES
# ============================================================

def limpiar_texto(texto):

    if texto is None:
        return ""

    texto = str(texto)

    texto = texto.replace(
        "\xa0",
        " "
    )

    texto = re.sub(
        r"\s+",
        " ",
        texto
    )

    return texto.strip()


def normalizar_nombre(nombre):

    texto = limpiar_texto(
        nombre
    ).lower()

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

        texto = texto.replace(
            origen,
            destino
        )

    return texto


def normalizar_codigo(codigo):

    if codigo is None:
        return ""

    texto = str(
        codigo
    ).strip()

    if texto.isdigit():

        texto = texto.zfill(5)

    return texto


# ============================================================
# DESCARGAR URL
# ============================================================

def descargar_url(url):

    marca_tiempo = int(
        datetime.now(
            timezone.utc
        ).timestamp()
    )

    separador = (
        "&"
        if "?" in url
        else "?"
    )

    url_final = (
        url
        + separador
        + "nocache="
        + str(marca_tiempo)
    )

    print()
    print(
        "URL:"
    )
    print(
        url_final
    )

    respuesta = requests.get(
        url_final,
        headers={
            "Cache-Control":
                "no-cache, no-store, must-revalidate",
            "Pragma": "no-cache",
            "Expires": "0",
            "User-Agent":
                "Mozilla/5.0 PlaAlfaESP32"
        },
        timeout=60
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
            "HTTP "
            + str(
                respuesta.status_code
            )
            + " al descargar "
            + url
        )

    return respuesta.content


# ============================================================
# OBTENER CODIGOS MUNICIPALES
# ============================================================

def obtener_codigos_municipios():

    print()
    print(
        "======================================"
    )
    print(
        "OBTENER CODIGOS MUNICIPALES"
    )
    print(
        "======================================"
    )

    contenido = descargar_url(
        URL_XLSX_CODIGOS
    )

    libro = load_workbook(
        filename=BytesIO(
            contenido
        ),
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

        if not codigo:
            continue

        if not nombre:
            continue

        codigos[
            normalizar_nombre(
                nombre
            )
        ] = codigo

    libro.close()

    print(
        "Municipios con código:",
        len(codigos)
    )

    if len(codigos) < 900:

        raise RuntimeError(
            "La tabla de códigos contiene "
            "menos de 900 municipios."
        )

    return codigos


# ============================================================
# EXTRAER EL JSON EMBEBIDO EN EL HTML
# ============================================================

def extraer_json_embebido(contenido):

    print()
    print(
        "======================================"
    )
    print(
        "EXTRAER DATOS JSON DE AVUI"
    )
    print(
        "======================================"
    )

    texto = contenido.decode(
        "utf-8",
        errors="replace"
    )

    print(
        "Caracteres HTML:",
        len(texto)
    )

    # --------------------------------------------------------
    # La página contiene:
    #
    # <script type="application/json" data-for="...">
    # { ... }
    # </script>
    #
    # Buscamos TODOS los bloques de este tipo.
    # --------------------------------------------------------

    patron = re.compile(
        r'<script\s+'
        r'type=["\']application/json["\']\s+'
        r'data-for=["\'][^"\']+["\']\s*>'
        r'(.*?)'
        r'</script>',
        re.IGNORECASE |
        re.DOTALL
    )

    bloques = patron.findall(
        texto
    )

    print(
        "Bloques JSON encontrados:",
        len(bloques)
    )

    if not bloques:

        raise RuntimeError(
            "No se ha encontrado el bloque "
            "JSON de la tabla DataTables."
        )

    # --------------------------------------------------------
    # Buscar el bloque que contiene:
    #
    # x -> data
    #
    # y que tenga las listas de municipios.
    # --------------------------------------------------------

    datos_tabla = None

    for bloque in bloques:

        bloque = bloque.strip()

        if not bloque:
            continue

        try:

            objeto = json.loads(
                bloque
            )

        except json.JSONDecodeError:

            continue

        if not isinstance(
            objeto,
            dict
        ):
            continue

        x = objeto.get(
            "x"
        )

        if not isinstance(
            x,
            dict
        ):
            continue

        data = x.get(
            "data"
        )

        if not isinstance(
            data,
            list
        ):
            continue

        if len(data) < 3:
            continue

        # ----------------------------------------------------
        # La primera columna debe contener los municipios.
        # ----------------------------------------------------

        primera_columna = data[0]

        if not isinstance(
            primera_columna,
            list
        ):
            continue

        if len(
            primera_columna
        ) < 900:

            continue

        # ----------------------------------------------------
        # Comprobar que realmente parece una lista municipal.
        # ----------------------------------------------------

        nombres = [
            normalizar_nombre(
                valor
            )
            for valor in primera_columna[:100]
        ]

        if (
            "abella de la conca"
            in nombres
            or
            "abrera"
            in nombres
            or
            "olivella"
            in [
                normalizar_nombre(
                    valor
                )
                for valor in primera_columna
            ]
        ):

            datos_tabla = data
            break

    if datos_tabla is None:

        raise RuntimeError(
            "No se ha encontrado el bloque "
            "de datos municipal con 947 registros."
        )

    print(
        "Columnas encontradas:",
        len(datos_tabla)
    )

    for i, columna in enumerate(
        datos_tabla
    ):

        if isinstance(
            columna,
            list
        ):

            print(
                "Columna",
                i,
                ":",
                len(columna),
                "registros"
            )

        else:

            print(
                "Columna",
                i,
                ": tipo no esperado"
            )

    return datos_tabla


# ============================================================
# CONVERTIR DATOS DE LA TABLA
# ============================================================

def convertir_datos_tabla(
    datos_tabla
):

    print()
    print(
        "======================================"
    )
    print(
        "CONVERTIR TABLA"
    )
    print(
        "======================================"
    )

    if len(datos_tabla) < 3:

        raise RuntimeError(
            "La tabla no contiene las "
            "columnas necesarias."
        )

    municipios = datos_tabla[0]
    comarcas = datos_tabla[1]
    peligros = datos_tabla[2]

    if len(datos_tabla) >= 4:

        fechas = datos_tabla[3]

    else:

        fechas = []

    cantidad = len(
        municipios
    )

    print(
        "Municipios:",
        cantidad
    )

    print(
        "Comarcas:",
        len(comarcas)
    )

    print(
        "Peligros:",
        len(peligros)
    )

    print(
        "Fechas:",
        len(fechas)
    )

    if cantidad < 900:

        raise RuntimeError(
            "La columna MUNICIPI contiene "
            "menos de 900 registros."
        )

    if len(comarcas) < cantidad:

        raise RuntimeError(
            "La columna COMARCA está incompleta."
        )

    if len(peligros) < cantidad:

        raise RuntimeError(
            "La columna PERILL está incompleta."
        )

    registros = []

    for i in range(
        cantidad
    ):

        municipio = limpiar_texto(
            municipios[i]
        )

        comarca = limpiar_texto(
            comarcas[i]
        )

        peligro = limpiar_texto(
            peligros[i]
        )

        fecha = ""

        if i < len(fechas):

            fecha = limpiar_texto(
                fechas[i]
            )

        if not municipio:
            continue

        registros.append(
            {
                "municipio":
                    municipio,

                "comarca":
                    comarca,

                "peligro":
                    peligro,

                "fecha":
                    fecha
            }
        )

    print(
        "Registros convertidos:",
        len(registros)
    )

    if len(registros) < 900:

        raise RuntimeError(
            "No se han podido convertir "
            "al menos 900 municipios."
        )

    return registros


# ============================================================
# CREAR PELIGRO.JSON
# ============================================================

def construir_municipios(
    registros,
    codigos
):

    print()
    print(
        "======================================"
    )
    print(
        "ASIGNAR CODIGOS MUNICIPALES"
    )
    print(
        "======================================"
    )

    municipios = {}

    sin_codigo = []

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

            sin_codigo.append(
                nombre
            )

            continue

        municipios[codigo] = {

            "municipio":
                nombre,

            "comarca":
                registro[
                    "comarca"
                ],

            "peligro":
                registro[
                    "peligro"
                ],

            "fecha":
                registro[
                    "fecha"
                ]
        }

    print(
        "Municipios con código:",
        len(municipios)
    )

    print(
        "Municipios sin código:",
        len(sin_codigo)
    )

    if sin_codigo:

        print()
        print(
            "Primeros municipios sin código:"
        )

        for nombre in sin_codigo[:20]:

            print(
                " -",
                nombre
            )

    if len(municipios) < 900:

        raise RuntimeError(
            "Se han podido asociar códigos "
            "a menos de 900 municipios."
        )

    return municipios


# ============================================================
# COMPROBAR LOS TRES MUNICIPIOS
# ============================================================

def comprobar_municipios(
    municipios
):

    print()
    print(
        "======================================"
    )
    print(
        "COMPROBACION MUNICIPIOS DEL PROYECTO"
    )
    print(
        "======================================"
    )

    for codigo, nombre_esperado in (
        MUNICIPIOS_PROYECTO.items()
    ):

        if codigo not in municipios:

            raise RuntimeError(
                "No se ha encontrado "
                + nombre_esperado
                + " ("
                + codigo
                + ")"
            )

        datos = municipios[
            codigo
        ]

        print(
            codigo,
            "|",
            datos[
                "municipio"
            ],
            "|",
            datos[
                "comarca"
            ],
            "|",
            datos[
                "peligro"
            ],
            "|",
            datos[
                "fecha"
            ]
        )


# ============================================================
# GUARDAR JSON
# ============================================================

def guardar_json(
    municipios
):

    fechas = []

    for datos in municipios.values():

        fecha = limpiar_texto(
            datos.get(
                "fecha",
                ""
            )
        )

        if fecha:

            fechas.append(
                fecha
            )

    fecha_actualizacion = ""

    if fechas:

        fecha_actualizacion = max(
            fechas
        )

    resultado = {

        "fecha":
            fecha_actualizacion,

        "municipios":
            municipios
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
        "ARCHIVO GENERADO"
    )
    print(
        "======================================"
    )

    print(
        "Fecha:",
        fecha_actualizacion
    )

    print(
        "Municipios:",
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
                datos[
                    "municipio"
                ],
                "|",
                datos[
                    "peligro"
                ],
                "|",
                datos[
                    "fecha"
                ]
            )

    print(
        "Archivo:",
        ARCHIVO_SALIDA
    )

    print(
        "======================================"
    )


# ============================================================
# PRINCIPAL
# ============================================================

def generar_json():

    print()
    print(
        "======================================"
    )
    print(
        "PLA ALFA PELIGRO v"
        + VERSION
    )
    print(
        "======================================"
    )

    # --------------------------------------------------------
    # 1. Obtener códigos municipales.
    # --------------------------------------------------------

    codigos = obtener_codigos_municipios()

    # --------------------------------------------------------
    # 2. Descargar página AVUI.
    # --------------------------------------------------------

    print()
    print(
        "======================================"
    )
    print(
        "DESCARGA HTML OFICIAL AVUI"
    )
    print(
        "======================================"
    )

    contenido = descargar_url(
        URL_HTML
    )

    # --------------------------------------------------------
    # 3. Extraer JSON que contiene los 947 municipios.
    # --------------------------------------------------------

    datos_tabla = extraer_json_embebido(
        contenido
    )

    # --------------------------------------------------------
    # 4. Convertir columnas en registros.
    # --------------------------------------------------------

    registros = convertir_datos_tabla(
        datos_tabla
    )

    # --------------------------------------------------------
    # 5. Asignar códigos.
    # --------------------------------------------------------

    municipios = construir_municipios(
        registros,
        codigos
    )

    # --------------------------------------------------------
    # 6. Comprobar Olivella, Sitges y Vilafranca.
    # --------------------------------------------------------

    comprobar_municipios(
        municipios
    )

    # --------------------------------------------------------
    # 7. Guardar peligro.json.
    # --------------------------------------------------------

    guardar_json(
        municipios
    )


# ============================================================
# INICIO
# ============================================================

if __name__ == "__main__":

    generar_json()
