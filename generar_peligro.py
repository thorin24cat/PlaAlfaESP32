# ============================================================
# PLA ALFA ESP32
# GENERADOR DE PELIGRO
#
# VERSION: 3.0.0
#
# Fuente del peligro:
#   Gencat - tabla municipal AVUI
#
# El XLSX se utiliza únicamente para obtener:
#   municipio -> código municipal
#
# ============================================================

import json
import re
import requests

from datetime import datetime, timezone
from io import BytesIO

from openpyxl import load_workbook


VERSION = "3.0.0"


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
# DESCARGAR
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
# CODIGOS MUNICIPALES
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
# EXTRAER DATOS DEL HTML
# ============================================================

def extraer_datos_html(contenido):

    print()
    print(
        "======================================"
    )
    print(
        "PROCESAR HTML AVUI"
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
    # Normalizar HTML.
    # --------------------------------------------------------

    texto = texto.replace(
        "\r",
        " "
    )

    texto = texto.replace(
        "\n",
        " "
    )

    texto = texto.replace(
        "\t",
        " "
    )

    # --------------------------------------------------------
    # Primero intentamos localizar directamente los datos
    # como texto HTML.
    #
    # La página oficial utiliza:
    #
    # MUNICIPI | COMARCA | PERILL | DATA DEMÀ
    # --------------------------------------------------------

    patron_fila = re.compile(
        r"""
        (?:
            <td[^>]*>
            \s*
            (?P<municipio>[^<]+)
            \s*
            </td>
        )
        \s*
        (?:
            <td[^>]*>
            \s*
            (?P<comarca>[^<]+)
            \s*
            </td>
        )
        \s*
        (?:
            <td[^>]*>
            \s*
            (?P<peligro>[^<]+)
            \s*
            </td>
        )
        \s*
        (?:
            <td[^>]*>
            \s*
            (?P<fecha>[^<]+)
            \s*
            </td>
        )
        """,
        re.IGNORECASE |
        re.VERBOSE
    )

    coincidencias = list(
        patron_fila.finditer(
            texto
        )
    )

    print(
        "Filas <td> encontradas:",
        len(coincidencias)
    )

    registros = []

    for coincidencia in coincidencias:

        municipio = limpiar_texto(
            coincidencia.group(
                "municipio"
            )
        )

        comarca = limpiar_texto(
            coincidencia.group(
                "comarca"
            )
        )

        peligro = limpiar_texto(
            coincidencia.group(
                "peligro"
            )
        )

        fecha = limpiar_texto(
            coincidencia.group(
                "fecha"
            )
        )

        if not municipio:
            continue

        registros.append(
            {
                "municipio": municipio,
                "comarca": comarca,
                "peligro": peligro,
                "fecha": fecha
            }
        )

    # --------------------------------------------------------
    # Si no hay filas <td>, la página está usando una
    # estructura JavaScript.
    #
    # En ese caso buscamos los registros en el contenido
    # JavaScript embebido.
    # --------------------------------------------------------

    if len(registros) < 900:

        print()
        print(
            "No se han encontrado 900 filas HTML."
        )

        print(
            "Buscando datos embebidos..."
        )

        # ----------------------------------------------------
        # Buscar específicamente los municipios conocidos.
        # Esto permite identificar rápidamente la estructura.
        # ----------------------------------------------------

        municipios_buscar = [
            "Olivella",
            "Sitges",
            "Vilafranca del Penedès",
            "Abella de la Conca",
            "Abrera"
        ]

        for nombre in municipios_buscar:

            posicion = texto.lower().find(
                nombre.lower()
            )

            if posicion >= 0:

                inicio = max(
                    0,
                    posicion - 500
                )

                final = min(
                    len(texto),
                    posicion + 1000
                )

                fragmento = texto[
                    inicio:final
                ]

                print()
                print(
                    "Municipio localizado:",
                    nombre
                )

                print(
                    fragmento
                )

        # ----------------------------------------------------
        # Buscar arrays JavaScript con cadenas.
        # ----------------------------------------------------

        patrones = [

            re.compile(
                r"""
                \[
                \s*
                ["']([^"']+)["']
                \s*,\s*
                ["']([^"']+)["']
                \s*,\s*
                ["']([^"']+)["']
                \s*,\s*
                ["']([^"']+)["']
                \s*
                \]
                """,
                re.IGNORECASE |
                re.VERBOSE
            ),

            re.compile(
                r"""
                \{
                [^{}]{0,500}?
                (?:municipi|municipio)
                [^{}]{0,500}?
                \}
                """,
                re.IGNORECASE |
                re.VERBOSE
            )
        ]

        for patron in patrones:

            encontrados = list(
                patron.finditer(
                    texto
                )
            )

            print(
                "Coincidencias patrón:",
                len(encontrados)
            )

            for coincidencia in encontrados:

                grupos = (
                    coincidencia.groups()
                )

                if len(grupos) == 4:

                    municipio = limpiar_texto(
                        grupos[0]
                    )

                    comarca = limpiar_texto(
                        grupos[1]
                    )

                    peligro = limpiar_texto(
                        grupos[2]
                    )

                    fecha = limpiar_texto(
                        grupos[3]
                    )

                    if municipio:

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

    # --------------------------------------------------------
    # Eliminar duplicados.
    # --------------------------------------------------------

    unicos = []

    vistos = set()

    for registro in registros:

        clave = (
            normalizar_nombre(
                registro[
                    "municipio"
                ]
            ),
            registro[
                "comarca"
            ],
            registro[
                "peligro"
            ],
            registro[
                "fecha"
            ]
        )

        if clave in vistos:
            continue

        vistos.add(
            clave
        )

        unicos.append(
            registro
        )

    registros = unicos

    print()
    print(
        "Registros finales:",
        len(registros)
    )

    return registros


# ============================================================
# CONSTRUIR PELIGRO.JSON
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
        "CONSTRUIR MUNICIPIOS"
    )
    print(
        "======================================"
    )

    municipios = {}

    for registro in registros:

        nombre = registro[
            "municipio"
        ]

        codigo = codigos.get(
            normalizar_nombre(
                nombre
            ),
            ""
        )

        if not codigo:

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

    return municipios


# ============================================================
# COMPROBAR MUNICIPIOS
# ============================================================

def comprobar_municipios(
    municipios
):

    print()
    print(
        "======================================"
    )
    print(
        "COMPROBACION"
    )
    print(
        "======================================"
    )

    errores = 0

    for codigo, nombre_esperado in (
        MUNICIPIOS_PROYECTO.items()
    ):

        if codigo not in municipios:

            print(
                "ERROR:",
                codigo,
                nombre_esperado,
                "NO ENCONTRADO"
            )

            errores += 1

            continue

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

    if errores:

        raise RuntimeError(
            "No se han encontrado "
            "todos los municipios "
            "del proyecto."
        )


# ============================================================
# GUARDAR
# ============================================================

def guardar_json(
    municipios
):

    fechas = []

    for datos in municipios.values():

        fecha = datos.get(
            "fecha",
            ""
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
    # Códigos municipales
    # --------------------------------------------------------

    codigos = obtener_codigos_municipios()

    # --------------------------------------------------------
    # HTML AVUI
    # --------------------------------------------------------

    print()
    print(
        "DESCARGA HTML OFICIAL AVUI"
    )

    contenido = descargar_url(
        URL_HTML
    )

    # --------------------------------------------------------
    # Extraer datos
    # --------------------------------------------------------

    registros = extraer_datos_html(
        contenido
    )

    if len(registros) < 900:

        raise RuntimeError(
            "No se han podido obtener "
            "los 947 municipios de la tabla "
            "oficial."
        )

    # --------------------------------------------------------
    # Crear JSON
    # --------------------------------------------------------

    municipios = construir_municipios(
        registros,
        codigos
    )

    if len(municipios) < 900:

        raise RuntimeError(
            "No se han podido asociar "
            "los códigos de al menos "
            "900 municipios."
        )

    # --------------------------------------------------------
    # Comprobar municipios usados
    # --------------------------------------------------------

    comprobar_municipios(
        municipios
    )

    # --------------------------------------------------------
    # Guardar
    # --------------------------------------------------------

    guardar_json(
        municipios
    )


# ============================================================
# INICIO
# ============================================================

if __name__ == "__main__":

    generar_json()
