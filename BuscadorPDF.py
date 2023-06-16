# Programa para buscar componentes en una carpeta de archivos pdf
# librerias necesarias
import fitz
import os
# import pandas as pd

# print(os.listdir()) # comando para obtener la LISTA de documentos de la carpeta
archivos = os.listdir() 
archivos.remove("BuscadorPDF.py")
archivos.remove("venv")
archivos.remove("BOM.xlsx")
archivos.remove("BuscadorV2.py")
archivos.remove("~$BOM.xlsx")
print(f'\n------> Lista de archivos: <------\n{archivos}\n')

palabraAbuscar = str(input('Que palabra quieres buscar?: '))

for documentosCarpeta in archivos:
    print("\nEn el archivo:", documentosCarpeta + " lo que busca")
    documento = fitz.open(documentosCarpeta)
    # Numero total de páginas en el documento abierto
    PaginasTotales = documento.page_count
    text = ""
    Condicionstatus = False
    # print("Lo que busca")
    for NumeroDepagina in range(PaginasTotales): 
        pagina = documento.load_page(NumeroDepagina)
        # text = text + pagina.get_text("text")
        text = pagina.get_text("text")
        Condicion = palabraAbuscar in text
        if Condicion == True:
            print("Se encuentra en la página: ", NumeroDepagina + 1)
            Condicionstatus = True
        elif (Condicion == False):
            if NumeroDepagina + 1 == PaginasTotales:
                if Condicionstatus == False:
                    print("NO se encuentra")
                else: Condicionstatus == True
