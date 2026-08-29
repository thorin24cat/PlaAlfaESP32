import json
import os


# ==================================================
# ARCHIVOS
# ==================================================

ARCHIVO_ENTRADA = "peligro.json"
CARPETA_SALIDA = "peligro"


# ==================================================
# CARGAR PELIGRO.JSON
# ==================================================

def cargar_datos():

    print()
    print("======================================")
    print("     CARGANDO PELIGRO.JSON")
    print("======================================")

    with open(
        ARCHIVO_ENTRADA,
        "r",
        encoding="utf-8"
    ) as archivo:

        datos = json.load(archivo)

    print(
        "Municipios encontrados:",
        len(datos["municipios"])
    )

    return datos


# ==================================================
# CREAR ARCHIVOS
# ==================================================

def generar_archivos(datos):

    municipios = datos["municipios"]

    os.makedirs(
        CARPETA_SALIDA,
        exist_ok=True
    )

    contador = 0

    for codigo, municipio in municipios.items():

        codigo = str(codigo).strip()

        if not codigo:
            continue

        archivo_salida = os.path.join(
            CARPETA_SALIDA,
            codigo + ".json"
        )

        resultado = {
            "codigo": codigo,
            "municipio": municipio.get(
                "municipio",
                ""
            ),
            "comarca": municipio.get(
                "comarca",
                ""
            ),
            "peligro": municipio.get(
                "peligro",
                ""
            ),
            "fecha": municipio.get(
                "fecha",
                datos.get("fecha", "")
            )
        }

        with open(
            archivo_salida,
            "w",
            encoding="utf-8"
        ) as archivo:

            json.dump(
                resultado,
                archivo,
                ensure_ascii=False,
                separators=(",", ":")
            )

        contador += 1

    print()
    print(
        "Archivos generados:",
        contador
    )


# ==================================================
# MAIN
# ==================================================

def main():

    datos = cargar_datos()

    generar_archivos(
        datos
    )

    print()
    print("======================================")
    print("     GENERACION FINALIZADA")
    print("======================================")


if __name__ == "__main__":

    main()
