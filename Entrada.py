import EstacaZero
from time import sleep
from colorama import init, Fore

#listaSolos = [23, 23, 23, 23, 33, 33, 33, 33, 33, 12, 12, 12, 12, 12, 12]
#listaNspt =  [ 0,  1,  1,  3,  6, 10, 11, 16, 30, 50, 50, 50, 50, 50, 50]

listaSolos = []
listaNspt = []

cor_azul = Fore.BLUE

texto = "Bem Vindo ao Estaca Zero"

print("")
print(cor_azul + texto.center(80))
print("")

while True:
    print("""Escolha o tipo de solo: 
           
    Areia = 1                        Silte = 2                       Argila = 3
    Areia siltosa = 12               Silte arenoso = 21              Argila arenosa = 31
    Areia siltoargilosa = 13         Silte arenoargiloso = 22        Argila arenossiltosa = 32
    Areia argilosa = 14              Silte argiloso = 23             Argila siltosa = 33
    Areia argilossiltosa = 15        Silte argiloarenoso = 24        Argila siltoarenosa = 34
    """)

    tipo_solo = int(input("Digite o tipo de solo: "))

    while tipo_solo not in [1, 12, 13, 14, 15, 2, 21, 22, 23, 24, 3, 31, 32, 33, 34]:

        print("Opção inválida! Por favor, escolha um tipo de solo válido.")
        
        tipo_solo = int(input("Digite o tipo de solo: "))

    listaSolos.append(tipo_solo)

    nspt_valor = int(input("Digite o valor NSPT correspondente: "))

    listaNspt.append(nspt_valor)

    print("")

    finalizar = str(input("Digite 'x' para finalizar ou 'c' para continuar preenchendo:  ")).upper()
    print("")

    if finalizar == 'X':
        
        break

print("""Escolha o tipo de estaca: 
      
    [0] Hélice Contínua    [1] Escavada    [2] Raiz    [3] Pré-Moldada    [4] Franki   [5] Ômega   [6] Metálica
      """)

estaca_opcoes = {
    '0': "Hélice Contínua",
    '1': "Escavada",
    '2': "Raiz",
    '3': "Pré-Moldada",
    '4': "Franki",
    '5': "Ômega",
    '6': "Metálica"
}

estaca = input("Tipo de estaca: ")

while estaca not in estaca_opcoes:

    print("Dado inválido. Por favor, escolha uma opção válida.")

    estaca = input("Tipo de estaca: ")

estaca = estaca_opcoes[estaca]

diametro = float(input("Diâmetro da estaca (em centimetros): ")) / 100

cargaAdmissivel = float(input("Carga admissível esperada (em kN): "))

niveldAgua = float(input("Nível da água (em metros): "))

print(EstacaZero.excelExport(listaSolos, estaca, diametro, listaNspt, cargaAdmissivel, niveldAgua, "Resultados.xlsx"))

print(EstacaZero.wordExport(listaSolos, estaca, diametro, listaNspt, cargaAdmissivel, niveldAgua, "Resultados.docx"))

sleep(10)
