from xlsxManegement import xlsx
from jsonManegement import jsonArq
from aplication import functions

'''

Relação dos ID's:

0 - natural 
1 - clássica
2 - elegante 
3 - romântica 
4 - sexy 
5 - criativa 
6 - dramático

'''

pathXlsx = './dataBase/pesquisa.xlsx'
pathJson = './dataBase/groups.json'

fileXlsx = xlsx.openFile(pathXlsx)
fileJson = jsonArq.openFile(pathJson)

worksheetRead = xlsx.openWorkbook(fileXlsx, 'Respostas ao formulário 1')
#worksheetWrite = xlsx.openWorkbook(fileXlsx, 'Tratamento')

result = functions.loadData(worksheetRead, fileJson)

print(result)

xlsx.close(fileXlsx)