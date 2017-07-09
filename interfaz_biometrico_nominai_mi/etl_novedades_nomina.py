
# coding: utf-8

import pandas as pd
import os
import configparser
import datetime as dt
import sys
import csv
import ntpath


def ConfigSectionMap(section):
    """
    Funcion para la captura de parámetros del archivo de configuración
    :param section: Sección del archivo de parámetros de los que capturará el parámetro
    :return:
    """

    dict1 = {}
    options = config.options(section)
    for option in options:
        try:
            dict1[option] = config.get(section, option)
            if dict1[option] == -1:
                DebugPrint("skip: %s" % option)
        except:
            print("exception on %s!" % option)
            dict1[option] = None
    return dict1


def get_config_file():
    for loc in os.curdir, os.path.expanduser("~"), "/etc/interfaz_biometrico", os.environ.get("CONFPATH"):
        try:
            open(os.path.join(loc, "config.ini"))
            return os.path.join(loc, "config.ini")
        except IOError:
            pass


def convert_file(file):
    """
    Convertir archivo con la ruta del parámetro de entrada, deja el archivo convertido en la ruta de salida y mueve el
    archivo origen a la carpeta de procesados

    :param file: Ruta del archivo a convertir
    :return:
    """

    in_file = pd.read_excel(file)
    #     Borrar columnas del dataframe que sobran
    del in_file["Codigo"]
    del in_file["Empleado"]
    del in_file["VTOTAL"]
    del in_file["Empresa"]

    concept_dict = ConfigSectionMap("conceptos")
    column_names = [concept_dict["rn"],
                    concept_dict["he"],
                    concept_dict["hen"],
                    concept_dict["hfd"],
                    concept_dict["hefn"],
                    "CC"]
    in_file.columns = column_names + ['Cedula']
    in_file.set_index(['Cedula', 'CC'], inplace=True)
    in_file_stacked = in_file.stack()

    stacked_df = pd.DataFrame(in_file_stacked)
    stacked_df.reset_index(level=[0, 1, 2], inplace=True)

    stacked_df = stacked_df[stacked_df[0] != 0]
    out_frame = pd.DataFrame(columns=["Cedula", "Codigo Concepto", "Codigo Centro de costos",
                                      "Codigo Concepto Referencia", "Horas", "Valor", "Periodo",
                                      "Fecha", "Salario", "Unidades Producidas", "Es Prestacion",
                                      "Numero Prestamo", "Dias mes 1", "Dias mes 2",
                                      "Fecha Inicio", "Base"])

    out_frame[["Cedula", "Codigo Centro de costos", "Codigo Concepto", "Horas"]] = stacked_df
    out_frame["Fecha"] = dt.datetime.today().strftime("%m/%d/%Y")
    out_frame["Periodo"] = dt.datetime.today().strftime("%m")
    out_frame["Valor"] = 0
    out_frame["Codigo Concepto Referencia"] = ""
    out_frame["Salario"] = 0
    out_frame["Unidades Producidas"] = 0
    out_frame["Es Prestacion"] = ""
    out_frame["Numero Prestamo"] = ""
    out_frame["Dias mes 1"] = 0
    out_frame["Dias mes 2"] = 0
    out_frame["Fecha Inicio"] = ""
    out_frame["Base"] = 0
    out_frame["Horas"] = out_frame["Horas"]

    out_frame.set_index("Cedula", inplace=True)
    path,filename = ntpath.split(file)
    out_frame.to_csv(out_path + "/"+filename+".txt", sep="\t", quoting=csv.QUOTE_NONE)
    os.rename(file, "./Procesado/"+filename)
    print("\""+filename+"\" ... OK. Archivo generado en \""+out_path+"/"+filename+".txt\"")


def main():
    """
    Función principal para la ejecución de la transformación del archivo
    :return:
    """
    global config
    global out_path
    config = configparser.ConfigParser()
    config.read(get_config_file())
    in_path = config.get("rutas", "in")
    out_path = config.get("rutas", "out")

    # Para capturar todos los archivs .xlsx que estén en la carpeta de entrada para procesarlos.

    if not os.listdir(in_path):
        sys.exit('Directorio de entrada vacío')
    else:
        for file in os.listdir(in_path):
            if file.endswith(".xlsx"):
                file=os.path.join(in_path, file)
                print("Convirtiendo archivos origen...")
                convert_file(file)



if __name__ == "__main__":
    main()
