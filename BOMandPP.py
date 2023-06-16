import pandas as pd
from tkinter import Tk
from tkinter.filedialog import askopenfilename

def BOM_PP(Ruta_BOM):
    # Crear una instancia de Tk
    root = Tk()
    # Ocultar la ventana de Tk
    root.withdraw()
    # ---------------------------Inicio del programa---------------------------
    nombreHojaBOM = 'Sheet1'
    nombreColumnaBOM = 'referenceDesignator'
    # nombreHojaBOM = input(str('Escribe el nombre de la hoja donde se encuentran los elementos de la BOM:\n'))
    # nombreColumnaBOM = input(str('Escribe el nombre de la columna donde se encuentran los elementos de la BOM (lo que está escrito en la primera celda de la columna):\n'))
    
    print('\nSelecciona el archivo del Pick & Place en formato ".xlsx"')
    Ruta_PP = askopenfilename()
    # Ruta_PP = 'C:\PruebaBCSPame\PruebaBCSPame\PruebaBCSPame\\1755080-01-G_01(Pick and place).xls'
    nombreHojaPP = '1811198-01-C_04'
    nombreColumnaPP = 'Altium Designer Pick and Place Locations'
    # nombreHojaPP = input(str('Escribe el nombre de la hoja donde se encuentran los elementos del Pick & Place:\n'))
    # nombreColumnaPP = input(str('Escribe el nombre de la Columna donde se encuentran los elementos del Pick & Place (Lo que está escrito en la primera celda de la columna):\n'))

    dataframeBOM = pd.read_excel(Ruta_BOM, sheet_name=nombreHojaBOM)
    dataframePP = pd.read_excel(Ruta_PP, sheet_name=nombreHojaPP)

    # ---------------------------------------Creación de listas definitivas---------------------------------------
    Lista_elementos_BOM = dataframeBOM[nombreColumnaBOM].tolist()
    ChecklistBOM = []
    # Separacion de componentes en la lista de excel
    for componentesBOM in range(len(Lista_elementos_BOM)-1):
        FirstlistBOM = Lista_elementos_BOM[componentesBOM+1]
        SecondlistBOM = str(FirstlistBOM).split(',')
        ChecklistBOM = ChecklistBOM + SecondlistBOM

    Lista_elementos_PP = dataframePP[nombreColumnaPP].tolist()
    ChecklistPP = []
    # Separacion de componentes en la lista de excel
    for componentesPP in range(len(Lista_elementos_PP)-12):
        FirstlistPP = Lista_elementos_PP[componentesPP+12]
        SecondlistPP = FirstlistPP.split(',')
        ChecklistPP = ChecklistPP + SecondlistPP

    # ---------------------------------------Conteo de elementos para buscar dentro del P&P ---------------------------------------
    failbom = 0
    passbom = 0
    elementosfailBOM = []
    for elementosBOM in ChecklistBOM:
        if elementosBOM in ChecklistPP:
            passbom += 1
        else:
            failbom += 1
            elementosfailBOM.append(elementosBOM)
    print(f'\nDe la busqueda de los elementos de la BOM en el P&P\n PASS = {passbom}, Fails en la BOM = {failbom}')
    print(f'Elementos no encontrados: {elementosfailBOM}')

    # ---------------------------------------Conteo de elementos para buscar dentro de la BOM ---------------------------------------
    failpp = 0
    passpp = 0
    elementosfailPP = []
    for elementospp in ChecklistPP:
        if elementospp in ChecklistBOM:
            passpp += 1
        else:
            failpp += 1
            elementosfailPP.append(elementospp)
    print(f'\nDe la busqueda de los elementos deL P&P en la BOM\n PASS = {passpp}, Fails en la BOM = {failpp}')
    print(f'Elementos no encontrados: {elementosfailPP}')
        