# Programa para buscar componentes en una carpeta de archivos pdf
# ----------------------------------------------------------------------------------------------------------------------------------------
# 
# 
# 
# Dseñado por el Ingeniero en Mecatrónica y desarrollador de software Nestor A. Ruiz a petición y 
# con requerimientos necesarios de la Ingeniera en mecatrónica Pamela Mejía Duran, Hardware engineer de 
# BCS-Automative Interface Solutions
#
#
#
# ----------------------------------------------------------------------------------------------------------------------------------------

# librerias necesarias
import fitz
import os
import pandas as pd
from tkinter import Tk
from tkinter.filedialog import askopenfilename
import BOMandPP
# Crear una instancia de Tk
root = Tk()
# Ocultar la ventana de Tk
root.withdraw()

def main():
    # ----------------------------------------------------------------------------------------------------------------------------------------
    print('Selecciona el archivo de la BOM')
    Ruta_Excel = askopenfilename()
    # ----------------------------------------------------------------------------------------------------------------------------------------
    nombre_hoja = 'Sheet1'
    nombre_columna = 'referenceDesignator'
    dataframe1 = pd.read_excel(Ruta_Excel, sheet_name=nombre_hoja)

    lista_datos = dataframe1[nombre_columna].tolist()
    Checklist = []

    # Separacion de componentes en la lista de excel
    for componentes in range (len(lista_datos)-2):
        Firstlist = lista_datos[componentes+1]
        Secondlist = Firstlist.split(',')
        Checklist = Checklist + Secondlist
    # print(Checklist)

    # ----------------------------------------------------------------------------------------------------------------------------------------
    # print(os.listdir()) # comando para obtener la LISTA de documentos de la carpeta
    # print(f'Nombre de los archivos leidos de la carpeta son: {archivos}')
    # print(f'nombreArchivosaBuscar = {nombreArchivosaBuscar}')
    # for NeedToBeDeleted in archivos:
    #     for archbuscados in range(len(nombreArchivosaBuscar)):
    #         if NeedToBeDeleted != nombreArchivosaBuscar[archbuscados]:
    #             print(NeedToBeDeleted)
    #             print(archbuscados)
    #             archivos.remove(NeedToBeDeleted)
    # print(f'Nombre de los archivos leidos de la carpeta con sobrantes borrados son: {archivos}')

    # nombre_archivoBOM = os.path.basename(Ruta_Excel)
    # archivos = []
    # CuantosArchivos = input(int('En cuantos archivos necesitas buscar?\n'))
    # for CArch in range(CuantosArchivos):
    #     print('Selecciona el archivo en donde se van a buscar')
    #     Ruta_Archivos = askopenfilename()

    archivos = os.listdir()
    archivos.remove("venv")
    archivos.remove("Projects")
    archivos.remove("__pycache__")

    archivos.remove("BOMDetails 1811198-01-C_04_Audio_Amplifier_Full.xlsx") #BOM
    archivos.remove("1811198-01-C_04(Pick and Place).xlsx") #P&P.xlsx

    archivos.remove("BuscadorV2.py")
    archivos.remove("BuscadorPDF.py")
    archivos.remove("BOMandPP.py")
    # ----------------------------------------------------------------------------------------------------------------------------------------
    print(f'\n------> Lista de archivos: <------\n{archivos}\n')
    ElementosNoEncontrados = []
    PASSresults = []
    FAILresults = []
    countPASS = 0
    countFAIL = 0
    text = ""
    Checklist.remove("PCB1")
    for documentosCarpeta in archivos:
        documento = fitz.open(documentosCarpeta)
        PaginasTotales = documento.page_count
        countFAIL = 0
        countPASS = 0
        for NumeroDepagina in range(PaginasTotales):
                Condicion = False
                pagina = documento.load_page(NumeroDepagina)
                # textpagina = pagina.get_text("text")
                textpagina = pagina.get_text("raw")
                text = text + textpagina
        for i in range(len(Checklist)):            
            Condicion = Checklist[i] in text
            # pagina.add_highlight_annot(Checklist[i])
            if Condicion == True:
                print(f'Electronic component {Checklist[i]} in document {documentosCarpeta} == TRUE')
                countPASS = countPASS+1
            elif (Condicion == False):
                if (NumeroDepagina+1) == PaginasTotales:
                    ElementosNoEncontrados.append(Checklist[i])
                    print(f'Electronic component {Checklist[i]} in document {documentosCarpeta} == FALSE')
                    countFAIL = countFAIL+1
        PASSresults.append(countPASS)
        FAILresults.append(countFAIL)

    print(f'Elementos no encontrados : {ElementosNoEncontrados}')
    for arch in range (len(archivos)): 
        print(f'En el archivo {archivos[arch]} se tuvo PASS: {PASSresults[arch]}, FAIL: {FAILresults[arch]}')
    
    # Se llama a la función para comparar la BOM con el P&P
    BOMandPP.BOM_PP(Ruta_Excel)

if __name__ == "__main__":
    main()