from tkinter import *
from tkinter import ttk
import openpyxl
from openpyxl import Workbook

def on_entry_change(entry, color):
    entry.configure(background=color)

def salvar_valores():
    wb = openpyxl.load_workbook("lancamento.xlsx")
    ws = wb.active

    corrida_atual = corrida.get()
    data_atual = data.get()

    for item, entry in zip(itens, entries):
        valor_inserido = entry.get()
        processo_selecionado = processo.get()

        if valor_inserido and processo_selecionado:
   
            ws.append([corrida_atual, data_atual, '', item, valor_inserido, processo_selecionado])

    
    wb.save("lancamento.xlsx")


    for entry in entries:
        entry.delete(0, END)
        entry.configure(background='white')
    processo.set("Processo")

janela = Tk()
janela.title('Inserir Valores')
janela.geometry('1200x700') 

corrida_label = Label(janela, text='Corrida:',background='#BCFFC1',anchor='e', font=('Arial', 15, 'bold'))
corrida_label.grid(row=0, column=0, pady=5, padx=5)
corrida = Entry(janela, width=15,background='#BCFFC1', font=('Arial', 10))
corrida.grid(row=0, column=1, pady=5, padx=5)

data_label = Label(janela, text='Data:',background='#BCFFC1',anchor='e', font=('Arial', 15, 'bold'))
data_label.grid(row=0, column=2, pady=5, padx=5)
data = Entry(janela, width=15,background='#BCFFC1', font=('Arial', 10))
data.grid(row=0, column=3, pady=5, padx=5)

itens = [
    "Sucata aço inox 310",
    "Retorno Inox 310/HK",
    "Sucata Alto Cromo III A >24%", 
    "Retorno Alto Cromo III A",
    "Sucata Alto Cromo II D - 18% a 24%",
    "Retorno Alto Cromo IID",
    "Sucata Alto Cromo II B - 14% a 17%",
    "Sucata Alto Cromo IIA - 11% a 14%",
    "Retorno Alto Cromo II A",
    "Sucata de Ferro Nodular",
    "Retorno Ferro nodular",
    "Sucata de Ferro Cinzento",
    "Retorno Ferro cinzento",
    "Sucata Ferro Branco",
    "Retorno Ferro Branco",
    "Sucata Aço Manganês",
    "Retorno Aço Manganês",
    "Sucata Aço Mn XT 710",
    "Retorno Aço Mn XT 710",
    "Sucata de Aço Fardinho A",
    "Sucata de Aço Fardinho B",
    "Sucata de Aço Fardinho C",
    "Sucata de Aço Oxicorte A",
    "Sucata de Aço Oxicorte B",
    "Retorno Sucata Mista",
    "Retorno Baixa Liga A",
    "Retorno Baixa Liga B",
    "Retorno Aço Carbono (1020/1030/1045)",
    "Gusa de Acearia",
    "Gusa Nodular",
    "Sucata aço inox 304",
    "Retorno Inox 304/HF",
    "Sucata aço Inox 316",
    "Sucata Inox 430",
    "Fe Cromo A/C",
    "Fe Cromo B/C", 
    "Fe Mn A/C",
    "Fe Mn M/C",
    "Fe Molibidenio",
    "Fe Si 75% Pedra",
    "Grafite",
    "Mn Eletrolitico",
    "Minério de Ferro",
    "Mn Eletrolitico - Granulado/Cano",
    "Fe Si Mg Granulado",
    "Aluminio",
    "CaSi",
    "FeSi 75% Inoculante (FeSiCaAlBa)",
    "FERRO SILÍCIO MAGNÉSIO",
    "Fe Silicio Zirconio",
    "Fe Titanio",
    "Escorificante",
    "Ponta de Pirometro STM",
    "Po exotermico ",
    "Espectrometria"

]

lista_processo = ['Alimentação', 'Correção', 'Desoxidação', 'Outros']


num_colunas = 3


entries = []
for i, item in enumerate(itens):
    row_number = i // num_colunas + 2  
    col_number = i % num_colunas * 2  

    label = Label(janela, text=item, anchor='w', font=('Arial', 10))
    label.grid(row=row_number, column=col_number, pady=5, padx=5)

    entry = Entry(janela, width=15, font=('Arial', 10))
    entry.grid(row=row_number, column=col_number + 1, pady=5, padx=5)
    entry.bind('<KeyRelease>', lambda event, entry=entry: on_entry_change(entry, 'yellow'))
    entries.append(entry)


processo = ttk.Combobox(janela, values=lista_processo, width=15, font=('Arial', 10))
processo.set("Processo")
processo.grid(row=0, column=4, columnspan=1, pady=10, padx=5)


row_number += 3  
botao_salvar = Button(janela, text='Salvar Valores', command=salvar_valores, font=('Arial', 12, 'bold'), bg='#4CAF50', fg='white')
botao_salvar.grid(row=0, column=5, columnspan=num_colunas * 2, pady=10)


janela.mainloop()
