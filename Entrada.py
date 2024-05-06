import EstacaZero

listaSolos = [23, 23, 23, 23, 33, 33, 33, 33, 33, 12, 12, 12, 12, 12, 12]

listaNspt =  [ 0,  1,  1,  3,  6, 10, 11, 16, 30, 50, 50, 50, 50, 50, 50]


#listaSolos = [14, 14, 14, 14, 14, 14, 14, 14, 14, 14, 14, 14, 14, 14, 14, 14, 14, 14, 14, 14, 14, 14, 14, 23, 23, 23, 23]
#listaNspt =  [ 5, 2,   3,  2,  4,  4,  7,  9,  9,  7,  7,  9, 14, 16, 15, 13, 14, 16, 21, 28, 25, 30, 35, 37, 56, 80, 75]

estaca = "Hélice Contínua" #Tipo da estaca: Franki, Metálica, Pré-Moldada, Escavada, Raiz, Hélice Contínua, Ômega

diametro = 0.3 #Diâmetro em Metros

cargaAdmissivel = 100 #kN

niveldAgua = 5


print(EstacaZero.excelExport(listaSolos, estaca, diametro, listaNspt, cargaAdmissivel, niveldAgua, "Resultados.xlsx"))

print(EstacaZero.wordExport(listaSolos, estaca, diametro, listaNspt, cargaAdmissivel, niveldAgua, "Resultados.docx"))