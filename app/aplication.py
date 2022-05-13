import string
import numpy as np

class functions:

    def loadData(xlsxFileRead,fileJson):

        matrizBi = functions.createMatriz(fileJson, xlsxFileRead)

        rowNumber = 2
        itensCount = 0
        styleBase = ''
        styleComplement = ''

        while xlsxFileRead[f'A{rowNumber}'].value != None:
            itensCount += 1
            rowNumber +=1

        rowNumber = 2

        matrizResult = np.zeros((1,itensCount), str)

        while xlsxFileRead[f'A{rowNumber}'].value != None:

            countStyleBase = np.zeros((1,3), int)
            countStyleComplement = np.zeros((1,4), int)

            matrizBi = functions.toCheck(fileJson, xlsxFileRead, rowNumber, matrizBi, countStyleBase, countStyleComplement)

            maxBase = np.argmax(countStyleBase, axis=1)
            maxComplement = np.argmax(countStyleComplement, axis=1)

            match maxBase:
                case 0:
                    styleBase = 'Natural'
                case 1:
                    styleBase = 'Clássica'
                case 2:
                    styleBase = 'Elegante'

            match maxComplement:
                case 0:
                    styleComplement = 'Romântica'
                case 1:
                    styleComplement = 'Sexy'
                case 2:
                    styleComplement = 'Criativa'
                case 3:
                    styleComplement = 'Dramático'

            print(f'{styleBase}, {styleComplement}')

            matrizResult[0, rowNumber-2] = f'{styleBase}, {styleComplement}'

            rowNumber += 1
            print(f'Itens a serem carregados: {rowNumber - 2}/{itensCount}')
 
        np.savetxt('BaseBinária.csv', matrizBi, fmt="%i", delimiter=",")

        print(matrizResult)

        return "Base carregada com sucesso!"

    def toCheck(fileJson, xlsxFileRead, rowNumber, matriz, countStyleBase, countStyleComplement):

        alpha = list(string.ascii_uppercase[1:20])
        columnNumber = 0
        itemNumber = 0

        for category in range(len(fileJson)):
            for group in range(len(fileJson[category]['grupos'])):
                if fileJson[category]['categoria'] == "profissões":
                    columnNumber = 7
                for item in range(len(fileJson[category]['grupos'][group]['itens'])):
                    itemName = fileJson[category]['grupos'][group]['itens'][item]['name']
                    cellValue = xlsxFileRead[f'{alpha[columnNumber]}{rowNumber}'].value
                    if cellValue != None:
                        if f'{itemName}' in cellValue:
                            matriz[(rowNumber-2),itemNumber] = 1
                            if fileJson[category]['grupos'][group]['id'] <= 2:
                                countStyleBase[0,fileJson[category]['grupos'][group]['id']] += 1
                            else:
                                countStyleComplement[0,(fileJson[category]['grupos'][group]['id']-3)] += 1
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

    def teste():

        countStyleComplement = np.array([[1, 2, 1]])
        maxIndex = np.argmax(countStyleComplement, axis=1)
        print(maxIndex)

        
