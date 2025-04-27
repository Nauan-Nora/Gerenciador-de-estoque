from kivy.uix.treeview import TreeView
import openpyxl
import customtkinter as ctk
import tkinter as tk
from tkinter import ttk

arquivo = openpyxl.load_workbook("dados.xlsx")
ativa= arquivo['Plan1']


def cadastrar_produto():
    v_nome = str(nome.get())
    v_codigo = int(codigo.get())
    v_preco = float(preco.get())
    v_quantidade = int(quantidade.get())
    v_venda = float(venda.get())

    ultima_linha = ativa.max_row + 1
    ativa.cell(row=ultima_linha, column=1, value=v_nome)
    ativa.cell(row=ultima_linha, column=2, value=v_codigo)
    ativa.cell(row=ultima_linha, column=3, value=v_preco)
    ativa.cell(row=ultima_linha, column=4, value=v_quantidade)
    ativa.cell(row=ultima_linha, column=5, value=v_venda)

    arquivo.save("dados.xlsx")
    limpar_campos()
    mostra_estoque()

def limpar_campos():
    v_nome = nome.delete(0, "end")
    v_codigo = codigo.delete(0, "end")
    v_preco = preco.delete(0, "end")
    v_quantidade = quantidade.delete(0, "end")
    v_venda = venda.delete(0, "end")

def mostra_estoque():
    # Limpa a TreeView antes de inserir os dados novamente
    mostra.delete(*mostra.get_children())

    # Obtém os cabeçalhos da planilha (primeira linha)
    cabecalhos = [cell.value for cell in ativa[1]]
    mostra["columns"] = cabecalhos
    mostra["show"] = "headings"
    for col in cabecalhos:
        mostra.heading(col, text=col)

    # Itera pelas linhas da planilha (a partir da segunda linha, para pular o cabeçalho)
    for row in ativa.iter_rows(min_row=2, values_only=True):
        mostra.insert('', tk.END, values=row)


ctk.set_appearance_mode("white")
ctk.set_default_color_theme("green")

app = ctk.CTk()
app.geometry("500x500")
app.title("Gerenciador de estoque")

app.grid_rowconfigure(0, weight=1)
app.grid_rowconfigure(1, weight=1)
app.grid_columnconfigure(0, weight=1)

area_cadastro = ctk.CTkFrame(app)
area_cadastro.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
area_cadastro.grid_columnconfigure(0, weight=1)

titulo = ctk.CTkLabel(area_cadastro, text="Formulario de cadastro")
titulo.grid(row=0, column=0, pady=5, sticky="ew")

nome = ctk.CTkEntry(area_cadastro, placeholder_text="Nome")
nome.grid(row=1, column=0, pady=5, sticky="ew")

codigo = ctk.CTkEntry(area_cadastro, placeholder_text="Código")
codigo.grid(row=2, column=0, pady=5, sticky="ew")

preco = ctk.CTkEntry(area_cadastro, placeholder_text="Custo")
preco.grid(row=3, column=0, pady=5, sticky="ew")

quantidade = ctk.CTkEntry(area_cadastro, placeholder_text="Quantidade")
quantidade.grid(row=4, column=0, pady=5, sticky="ew")

venda = ctk.CTkEntry(area_cadastro, placeholder_text="Valor de venda")
venda.grid(row=5, column=0, pady=5, sticky="ew")

buton_cadastrar = ctk.CTkButton(area_cadastro, text="Cadastrar", command=cadastrar_produto)
buton_cadastrar.grid(row=6, column=0, pady=5, sticky="ew")

buton_apagar = ctk.CTkButton(area_cadastro, text="Limpar formulario", command=limpar_campos)
buton_apagar.grid(row=7, column=0, pady=5, sticky="ew")

# Área de visualização centralizada
area_visualizacao = ctk.CTkFrame(app)
area_visualizacao.grid(row=1, column=0, padx=20, pady=20, sticky="nsew")
area_visualizacao.grid_columnconfigure(0, weight=1)
area_visualizacao.grid_rowconfigure(0, weight=1)

mostra = ttk.Treeview(
    area_visualizacao,
    columns=("Nome", "Código", "Custo", "Quantidade", "Valor de venda"),
    show="headings"
    
)
tabelaScorll = ttk.Scrollbar(area_visualizacao)
tabelaScorll.grid(sticky="nse")
mostra.heading("Nome", text="Nome")
mostra.heading("Código", text="Código")
mostra.heading("Custo", text="Custo")
mostra.heading("Quantidade", text="Quantidade")
mostra.heading("Valor de venda", text="Valor de venda")
mostra.grid(row=0, column=0, sticky="nsew")

mostra_estoque()
app.mainloop()