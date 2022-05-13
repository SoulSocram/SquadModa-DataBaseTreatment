from xlsxManegement import xlsx
from jsonManegement import jsonArq
from aplication import functions

pathXlsx = './dataBase/pesquisa.xlsx'
pathJson = './dataBase/groups.json'

fileXlsx = xlsx.openFile(pathXlsx)
fileJson = jsonArq.openFile(pathJson)

worksheetRead = xlsx.openWorkbook(fileXlsx, 'Pesquisa')
worksheetWrite = xlsx.openWorkbook(fileXlsx, 'Tratamento')

result = functions.loadData(worksheetRead, fileJson)

#functions.teste()

print(result)

xlsx.close(fileXlsx)