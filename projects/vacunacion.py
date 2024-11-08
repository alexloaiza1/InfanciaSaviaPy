import pandas as pd

ruta_archivo = "C:\\Users\\alex2\\Music\\InfanciaProjectPy\\Data\\Guía ACA Creación de empresas 3 (2).xlsm"

def leer_xlsm(ruta_archivo):

    """
    Lee un archivo XLSM y devuelve un DataFrame de pandas.

    :param ruta_archivo: Ruta del archivo XLSM.
    :return: DataFrame de pandas con los datos del XLSM.
    """
    try:
        df = pd.read_excel(ruta_archivo, engine='openpyxl')
        return df
    except FileNotFoundError:
        print(f"El archivo en la ruta {ruta_archivo} no fue encontrado.")
    except pd.errors.EmptyDataError:
        print("El archivo XLSM está vacío.")
    except pd.errors.ParserError:
        print("Error al parsear el archivo XLSM.")
    except Exception as e:
        print(f"Ocurrió un error: {e}")
        
        
        numero = 0
        numero = numero + 1
