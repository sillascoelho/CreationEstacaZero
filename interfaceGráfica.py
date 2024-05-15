from EstacaZero import wordExport, excelExport
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from tkinter import filedialog

#ÍNICIANDO A INTERFACE

estacaZero = tk.Tk()
estacaZero.title("Calculadora Estaca Zero")
estacaZero.rowconfigure(0, weight=1)
estacaZero.columnconfigure(0, weight=1)

#DADOS QUE SERÃO UTILIZADOS E PREENCHIDOS.

fonte_negrito = ('Helvetica', 9, 'bold')
cota = -1
listaSolos = []
listaNspt = []
diametro = tk.IntVar()
cargaAdmissivel = tk.IntVar()
niveldAgua = tk.IntVar()

""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""

listaEstacas_dict = {
    "Hélice Contínua",
    "Escavada",
    "Raiz",
    "Pré-Moldada",
    "Franki",
    "Ômega",
    "Metálica"
            }

listaEstacas = list(listaEstacas_dict)
tipoEstaca_msg = tk.Label(text="Tipo de Estaca: ", fg="black", bg="white", font=fonte_negrito)
tipoEstaca_msg.grid(row=0, column=0, sticky='NSEW')
tipoEstaca = ttk.Combobox(estacaZero, values=listaEstacas, state="readonly")
tipoEstaca.grid(row=0, column=1, sticky='NSEW')

diametro_msg = tk.Label(text="Diâmetro da Estaca (cm): ", fg="black", bg="white", font=fonte_negrito)
diametro_msg.grid(row=1, column=0, sticky='NSEW')
diametroEstaca = tk.Entry()
diametroEstaca.grid(row=1, column=1, sticky='NSEW')

cargaAdmissivel_msg = tk.Label(text="Carga Admissível Esperada (kN): ", fg="black", bg="white", font=fonte_negrito)
cargaAdmissivel_msg.grid(row=2, column=0, sticky='NSEW')
cargaAdmissivelEsperada = tk.Entry()
cargaAdmissivelEsperada.grid(row=2, column=1, sticky='NSEW')

niveldAgua_msg = tk.Label(text="Nível de Água (m): ", fg="black", bg="white", font=fonte_negrito)
niveldAgua_msg.grid(row=3, column=0, sticky='NSEW')
niveldeAgua = tk.Entry()
niveldeAgua.grid(row=3, column=1, sticky='NSEW')

listadeSolos = [
    'Areia',
    'Areia Siltosa',
    'Areia Siltoargilosa',
    'Areia Argilosa',
    'Areia Argilossiltosa',
    'Silte',
    'Silte Arenoso',
    'Silte Arenoargiloso',
    'Silte Argiloso',
    'Silte Argiloarenoso',
    'Argila',
    'Argila Arenosa',
    'Argila Arenossiltosa',
    'Argila Siltosa',
    'Argila Siltoarenosa'
                    ]

camadas_msg = tk.Label(text="Indique o Solo e Nspt correspondente", fg="black", bg="white", font=fonte_negrito)
espaco_msg = tk.Label(text="", fg="black", bg="white")
camadas_msg.grid(row=4, column=0, sticky='NSEW', columnspan=1)
espaco_msg.grid(row=4, column=1, sticky='NSEW', columnspan=1)

listadeSolos = list(listadeSolos)
listaSolos_msg = tk.Label(text="Solo da Camada: ", fg="black", bg="white", font=fonte_negrito)
listaSolos_msg.grid(row=5, column=0, sticky='NSEW')
escolhaSolo = ttk.Combobox(estacaZero, values=listadeSolos, state="readonly")
escolhaSolo.grid(row=5, column=1, sticky='NSEW')

valorNspt_msg = tk.Label(text="Nspt da Camada: ", fg="black", bg="white", font=fonte_negrito)
valorNspt_msg.grid(row=6, column=0, sticky='NSEW')
valorNspt = tk.Entry()
valorNspt.grid(row=6, column=1, sticky='NSEW')
    
def salvar_camada():
    global cota
    global listaSolos
    global listaNspt

    listaSolos_dict = {
    'Areia': 1,
    'Areia Siltosa': 12,
    'Areia Siltoargilosa': 13,
    'Areia Argilosa': 14,
    'Areia Argilossiltosa': 15,
    'Silte': 2,
    'Silte Arenoso': 21,
    'Silte Arenoargiloso': 22,
    'Silte Argiloso': 23,
    'Silte Argiloarenoso': 24,
    'Argila': 3,
    'Argila Arenosa': 31,
    'Argila Arenossiltosa': 32,
    'Argila Siltosa': 33,
    'Argila Siltoarenosa': 34 
                    }

    solo_selecionado = escolhaSolo.get()
    valorSolo = listaSolos_dict[escolhaSolo.get()]
    nspt_digitado = valorNspt.get()
    dados.insert(tk.END, f"Cota {cota} m, Solo: {solo_selecionado}, Nspt: {nspt_digitado}\n")
    if not nspt_digitado.isdigit():
        tk.messagebox.showerror("Atenção", "Por favor, insira apenas números para Nspt.")
        return
    listaSolos.append(valorSolo)
    listaNspt.append(int(nspt_digitado))

    cota = cota - 1

salvarCamada = tk.Button(text="Adicionar Camada", fg='white', bg='brown', font=fonte_negrito, command=salvar_camada)
salvarCamada.grid(row=7, column=0, columnspan=2, sticky='NSEW')

def limpartudo():
    global cota
    listaSolos.clear()
    listaNspt.clear()
    dados.delete('1.0', tk.END)
    cota = -1

limparCamadas = tk.Button(text="Limpar Camadas", fg='white', bg='orange', font=fonte_negrito, command=limpartudo)
limparCamadas.grid(row=8, column=0, columnspan=2, sticky='NSEW')

dados = tk.Text(width=48, height=20)
dados.grid(row=9, columnspan=2)

def gerarMemorial():

    if (tipoEstaca.get() == "" or
        diametroEstaca.get() == "" or
        cargaAdmissivelEsperada.get() == "" or
        niveldeAgua.get() == "" or
        len(listaSolos) == 0 or
        len(listaNspt) == 0):

        messagebox.showerror("ATENÇÃO", "Por favor, preencha todos os campos obrigatórios.")

    else:

        estaca = tipoEstaca.get()
        diametro = float(diametroEstaca.get()) / 100
        cargaAdmissivel = float(cargaAdmissivelEsperada.get())
        niveldAgua = float(niveldeAgua.get())
        fileName = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word Document", "*.docx")], initialfile="Memorial Estaca Zero")
        if not fileName:
            return

        wordGerado = wordExport(listaSolos, estaca, diametro, listaNspt, cargaAdmissivel, niveldAgua, fileName)

        messagebox.showinfo("Memorial Gerado", f"O memorial foi exportado com sucesso!")

    return wordGerado

exportarWord = tk.Button(text="Gerar Memorial em Word", fg='white', bg='blue', font=fonte_negrito, command=gerarMemorial)
exportarWord.grid(row=10, column=0, columnspan=2, sticky='NSEW')

def gerarExcel(): 

    if (tipoEstaca.get() == "" or
        diametroEstaca.get() == "" or
        cargaAdmissivelEsperada.get() == "" or
        niveldeAgua.get() == "" or
        len(listaSolos) == 0 or
        len(listaNspt) == 0):
        messagebox.showerror("ATENÇÃO", "Por favor, preencha todos os campos obrigatórios.")

    else:

        estaca = tipoEstaca.get()
        diametro = float(diametroEstaca.get()) / 100
        cargaAdmissivel = float(cargaAdmissivelEsperada.get())
        niveldAgua = float(niveldeAgua.get())
        fileName = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Workbook", "*.xlsx")], initialfile="Resultados Estaca Zero")
        if not fileName:
            return
        
        excelGerado = excelExport(listaSolos, estaca, diametro, listaNspt, cargaAdmissivel, niveldAgua, fileName)

        messagebox.showinfo("Resultados Gerados", f"Os resultados foram exportados com sucesso!")

    return excelGerado

exportarExcel = tk.Button(text="Exportar Resultados para Excel", fg='white', bg='green', font=fonte_negrito, command=gerarExcel)
exportarExcel.grid(row=11, column=0, columnspan=2, sticky='NSEW')

estacaZero.mainloop()