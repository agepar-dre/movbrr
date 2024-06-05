"""
Created on Wed Jan 31 10:40:36 2024

@author: cecil.skaleski, est.angelo

Interface para a aplicação de movimentação da BRR confome disposto na Nota Técnica n° XX/2024.

Requeriments
conda install conda-forge::matplotlib
conda install anaconda::openpyxl
conda install conda-forge::natsort
conda install -c conda-forge numpy-financial
conda install anaconda::xlrd

Gerar executável
pyinstaller main_mov_BRR.py --hidden-import openpyxl.cell._writer --hidden-import matplotlib.backends.backend_pdf --add-data "1_ENTRADA_CONVERTE;1_ENTRADA_CONVERTE" --add-data "2_SAIDA_CONVERTE;2_SAIDA_CONVERTE" --add-data "3_ENTRADA_CONSOLIDA;3_ENTRADA_CONSOLIDA" --add-data "4_SAIDA_CONSOLIDA;4_SAIDA_CONSOLIDA" --add-data "5_ENTRADA_MOVIMENTA;5_ENTRADA_MOVIMENTA" --add-data "6_SAIDA_MOVIMENTA;6_SAIDA_MOVIMENTA"
"""



import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
import converte_BRR_v7
import consolida_BRR_v5
import movimenta_BRR_v8


def show_tab(event):
    tab_name = tab_control.select()
    if tab_name == tab1:
        print("Você está na aba 'Converte BRR'")
    elif tab_name == tab2:
        print("Você está na aba 'Consolida BRR'")
    elif tab_name == tab3:
        print("Você está na aba 'Movimenta BRR'")

# Criando a janela principal
root = tk.Tk()
root.wm_attributes('-topmost', 1)
root.title("Ferramenta de movimentação da BRR")
root.geometry("700x350")
root.configure(bg="white")

# Definindo estilo
style = ttk.Style()
style.theme_create("MyStyle", parent="alt", settings={
    "TNotebook": {"configure": {"background": "navy"}},
    "TNotebook.Tab": {
        "configure": {"padding": [20, 5], "background": "blue", "foreground": "white"},
        "map": {"background": [("selected", "royalblue")]}
    },
    "TFrame": {"configure": {"background": "white"}}
})
style.theme_use("MyStyle")

# Criando o controle de abas
tab_control = ttk.Notebook(root)

# Cria as guias
#Guia 1: Converte BRR
tab1 = ttk.Frame(tab_control)
converte_BRR_v7.make_frame(tab1)

#Guia 1: Consolida BRR
tab2 = ttk.Frame(tab_control)
consolida_BRR_v5.make_frame(tab2)

#Guia 3: Movimenta BRR
tab3 = ttk.Frame(tab_control)
movimenta_BRR_v8.make_frame(tab3)

tab_control.add(tab1, text='Converte BRR')
tab_control.add(tab2, text='Consolida BRR')
tab_control.add(tab3, text='Movimenta BRR')

tab_control.pack(expand=1, fill="both")

# Evento para mostrar a aba selecionada
tab_control.bind("<<NotebookTabChanged>>", show_tab)

root.mainloop()
