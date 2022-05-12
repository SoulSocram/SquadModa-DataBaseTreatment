import string
import numpy as np

class functions:

    def loadData(xlsxFileRead,fileJson):

        matrizBi = functions.createMatriz(fileJson, xlsxFileRead)

        rowNumber = 2
        itensCount = 0

        while xlsxFileRead[f'A{rowNumber}'].value != None:
            itensCount += 1
            rowNumber +=1

        rowNumber = 2

        while xlsxFileRead[f'A{rowNumber}'].value != None:
            matrizBi = functions.toCheck(fileJson, xlsxFileRead, rowNumber, rowNumber-2, matrizBi)
            print("\033c")
            rowNumber += 1
            print(f'Itens a serem carregados: {rowNumber - 2}/{itensCount}')
 
        np.savetxt('BaseBinária.csv', matrizBi, fmt="%i", delimiter=",")

        return "Base carregada com sucesso!"

    def toCheck(fileJson, xlsxFileRead,cellNumber, rowNumber, matriz):

        alpha = list(string.ascii_uppercase[1:20])
        columnNumber = 0
        itemNumber = 0

        for category in range(len(fileJson)):
            for group in range(len(fileJson[category]['grupos'])):
                if fileJson[category]['categoria'] == "profissões":
                    columnNumber = 7
                for item in range(len(fileJson[category]['grupos'][group]['itens'])):
                    itemName = fileJson[category]['grupos'][group]['itens'][item]['name']
                    cellValue = xlsxFileRead[f'{alpha[columnNumber]}{cellNumber}'].value
                    if cellValue != None:
                        if f'{itemName}' in cellValue:
                            matriz[rowNumber,itemNumber] = 1
                    itemNumber += 1
                columnNumber += 1

        return matriz

    def createMatriz(jsonFile, xlsxFileRead):

        countItens = 0
        countSearch = 0

        for category in range(len(jsonFile)):
            for group in range(len(jsonFile[category]['grupos'])):
                for item in range(len(jsonFile[category]['grupos'][group]['itens'])):
                    countItens += 1

        while xlsxFileRead[f'A{countSearch+2}'].value != None:
            countSearch += 1

        matrizBi = np.zeros((countSearch,countItens), int)

        return matrizBi

        
