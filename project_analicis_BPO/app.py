
# ==========================================
# PROYECTO: ANALISIS DE AGENTES BPO
# APLICACION WEB CON PYTHON Y HTML
# ==========================================

from tkinter import Tk, filedialog
from openpyxl import load_workbook
from flask import Flask, render_template
from threading import Timer
import webbrowser


# 1. Crear la aplicacion Flask
app = Flask(__name__)


# 2. Definir las columnas necesarias
columnas_requeridas = [
    "ID del agente",
    "Llamadas",
    "Satisfacción (%)",
    "Cumplimiento (%)",
    "Ausencias",
    "Tiempo promedio (min)"
]


# 3. Abrir el explorador de Windows
def seleccionar_excel():

    ventana = Tk()
    ventana.withdraw()
    ventana.attributes("-topmost", True)

    archivo = filedialog.askopenfilename(
        title="Selecciona el archivo de agentes",
        filetypes=[
            ("Archivos Excel", "*.xlsx")
        ]
    )

    ventana.destroy()

    return archivo


# 4. Leer y validar el Excel
def analizar_excel(archivo):

    agentes = []
    errores = []

    libro = None

    try:

        # Abrir el archivo
        libro = load_workbook(
            archivo,
            data_only=True
        )

        hoja = libro.active

        # Leer los encabezados de la fila 4
        encabezados = [
            celda.value.strip()
            if isinstance(celda.value, str)
            else celda.value
            for celda in hoja[4]
        ]

        # Guardar la posicion de cada columna
        posiciones = {}

        for posicion, encabezado in enumerate(encabezados):

            if encabezado is not None:
                posiciones[encabezado] = posicion + 1

        # Comprobar las columnas necesarias
        columnas_faltantes = []

        for columna in columnas_requeridas:

            if columna not in posiciones:
                columnas_faltantes.append(columna)

        if columnas_faltantes:

            errores.append("Faltan las siguientes columnas:")

            for columna in columnas_faltantes:
                errores.append(columna)

            return agentes, errores

        # 5. Leer los agentes desde la fila 5
        for numero_fila in range(5, hoja.max_row + 1):

            agente = {}

            for columna in columnas_requeridas:

                posicion = posiciones[columna]

                valor = hoja.cell(
                    row=numero_fila,
                    column=posicion
                ).value

                agente[columna] = valor

            # Ignorar filas completamente vacias
            if all(
                valor is None or
                (isinstance(valor, str) and not valor.strip())
                for valor in agente.values()
            ):
                continue

            agentes.append(agente)

        # 6. Validar los datos
        columnas_numericas = [
            "Llamadas",
            "Satisfacción (%)",
            "Cumplimiento (%)",
            "Ausencias",
            "Tiempo promedio (min)"
        ]

        for numero, agente in enumerate(agentes, start=5):

            # Comprobar el ID
            identificacion = agente["ID del agente"]

            if (
                identificacion is None
                or str(identificacion).strip() == ""
            ):

                errores.append(
                    f"Fila {numero}: falta el ID del agente."
                )

            # Comprobar las metricas numericas
            for columna in columnas_numericas:

                valor = agente[columna]

                if valor is None or isinstance(valor, bool):

                    errores.append(
                        f"Fila {numero}: falta el dato {columna}."
                    )

                elif not isinstance(valor, (int, float)):

                    errores.append(
                        f"Fila {numero}: {columna} debe ser numerico."
                    )

            # Validar porcentajes
            for columna in [
                "Satisfacción (%)",
                "Cumplimiento (%)"
            ]:

                valor = agente[columna]

                if (
                    isinstance(valor, (int, float))
                    and not isinstance(valor, bool)
                ):

                    if valor < 0 or valor > 100:

                        errores.append(
                            f"Fila {numero}: {columna} debe estar entre 0 y 100."
                        )

            # Validar llamadas y ausencias
            for columna in ["Llamadas", "Ausencias"]:

                valor = agente[columna]

                if (
                    isinstance(valor, (int, float))
                    and not isinstance(valor, bool)
                ):

                    if valor < 0:

                        errores.append(
                            f"Fila {numero}: {columna} no puede ser negativo."
                        )

                    elif not float(valor).is_integer():

                        errores.append(
                            f"Fila {numero}: {columna} debe ser un numero entero."
                        )

            # Validar tiempo promedio
            tiempo = agente["Tiempo promedio (min)"]

            if (
                isinstance(tiempo, (int, float))
                and not isinstance(tiempo, bool)
            ):

                if tiempo <= 0:

                    errores.append(
                        f"Fila {numero}: el tiempo promedio debe ser mayor que cero."
                    )

        # Comprobar que existan agentes
        if not agentes:

            errores.append(
                "El archivo no contiene agentes para analizar."
            )

        # 7. Realizar el analisis
        if not errores:

            for agente in agentes:

                satisfaccion = agente["Satisfacción (%)"]
                cumplimiento = agente["Cumplimiento (%)"]
                ausencias = agente["Ausencias"]
                tiempo = agente["Tiempo promedio (min)"]

                # Clasificar los agentes
                if (
                    satisfaccion < 70
                    or cumplimiento < 80
                    or ausencias >= 4
                    or tiempo > 7
                ):

                    agente["estado"] = "Crítico"
                    agente["decision"] = "Requiere capacitación"
                    agente["color"] = "critico"

                else:

                    agente["estado"] = "Bueno"
                    agente["decision"] = "Rendimiento adecuado"
                    agente["color"] = "bueno"

    except Exception as error:

        errores.append(
            f"No se pudo procesar el archivo: {error}"
        )

    finally:

        if libro is not None:
            libro.close()

    return agentes, errores


# 8. Preparar los resultados
agentes = []
errores = []
nombre_archivo = ""

total_agentes = 0
total_criticos = 0
total_buenos = 0


# 9. Crear la pagina de resultados
@app.route("/")
def inicio():

    return render_template(
        "index.html",
        agentes=agentes,
        errores=errores,
        nombre_archivo=nombre_archivo,
        total_agentes=total_agentes,
        total_criticos=total_criticos,
        total_buenos=total_buenos
    )


# 10. Ejecutar el programa
if __name__ == "__main__":

    print("\n====================================")
    print(" SISTEMA DE ANALISIS DE AGENTES BPO")
    print("====================================")

    # Seleccionar el archivo
    archivo = seleccionar_excel()

    if not archivo:

        print("\nNo seleccionaste ningun archivo.")
        print("El programa ha finalizado.")

    else:

        import os

        nombre_archivo = os.path.basename(archivo)

        print("\nArchivo seleccionado:")
        print(nombre_archivo)

        print("\nProcesando informacion...")

        # Analizar el Excel
        agentes, errores = analizar_excel(archivo)

        # Calcular resultados si no hay errores
        if not errores:

            total_agentes = len(agentes)

            total_criticos = sum(
                1 for agente in agentes
                if agente["estado"] == "Crítico"
            )

            total_buenos = total_agentes - total_criticos

            print("\nAnalisis completado correctamente.")

        else:

            print("\nSe encontraron problemas en el archivo.")

            for error in errores:
                print("-", error)

        # Abrir Chrome cuando Flask este listo
        Timer(
            1.5,
            lambda: webbrowser.open("http://127.0.0.1:5000")
        ).start()

        print("\nAbriendo pagina de resultados...")

        # Iniciar la pagina web
        app.run(
            host="127.0.0.1",
            port=5000,
            debug=False,
            use_reloader=False
        )
