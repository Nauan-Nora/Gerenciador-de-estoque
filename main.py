from kivy.uix.treeview import TreeView
from tkinter import messagebox
import openpyxl
import customtkinter as ctk
import tkinter as tk
from tkinter import ttk

arquivo = openpyxl.load_workbook("dados.xlsx")
ativa= arquivo['Plan1']
tema_atual = "white"


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
    mostra_treeview.delete(*mostra_treeview.get_children())

    cabecalhos = [cell.value for cell in ativa[1]]
    mostra_treeview["columns"] = cabecalhos
    mostra_treeview["show"] = "headings"
    for col in cabecalhos:
        mostra_treeview.heading(col, text=col)

    for row in ativa.iter_rows(min_row=2, values_only=True):
        mostra_treeview.insert('', tk.END, values=row)

def verifica_preenchimento():
    v_nome = nome.get()
    v_codigo = codigo.get()
    v_preco = preco.get()
    v_quantidade = quantidade.get()
    v_venda = venda.get()

    try:
        if v_nome == "":
            messagebox.showerror("ERRO DE PREENCIMENTO", "O campo de nome esta vazio")
        elif v_codigo == "":
            messagebox.showerror("ERRO DE PREENCIMENTO", "O compo de código esta vazio")
        elif v_preco == "":
            messagebox.showerror("ERRO DE PREENCIMENTO", " O campo de preço esta vazio")
        elif v_quantidade == "":
            messagebox.showerror("ERRO DE PREENCIMENTO", "O campo de quantidade esta vazio")
        elif v_venda == "":
            messagebox.showerror("ERRO DE PREENCIMENTO", "O campo de valor de venda esta vazio")
            pass
        else:
            cadastrar_produto()

    except:
        messagebox.showerror("ERRO", " Houve um erra não esperado, entre em contato com o suporte técnico")


ctk.set_appearance_mode(tema_atual)
ctk.set_default_color_theme("green")

app = ctk.CTk()
app.geometry("830x710")
app.title("Gerenciador de estoque")

app.grid_rowconfigure(0, weight=0)
app.grid_rowconfigure(1, weight=0) 
app.grid_rowconfigure(2, weight=1) 
app.grid_columnconfigure(0, weight=1)

style = ttk.Style(app)
style.theme_use("default")
style.configure("Light.Treeview", background="white", foreground="black")
style.configure("Light.Treeview.Heading", background="#f0f0f0", foreground="black")
style.map("Light.Treeview", background=[("selected", "#aed1fc")])

style.configure("Dark.Treeview", background="#333333", foreground="white")
style.configure("Dark.Treeview.Heading", background="#555555", foreground="white")
style.map("Dark.Treeview", background=[("selected", "#5699bc")])

frame_superior = ctk.CTkFrame(app)
frame_superior.grid(row=0, column=0, padx=20, pady=20, sticky="ew")
frame_superior.grid_columnconfigure(0, weight=1)
frame_superior.grid_columnconfigure(1, weight=0)

frame_cadastro = ctk.CTkFrame(frame_superior)
frame_cadastro.grid(row=0, column=0, padx=(0, 10), pady=10, sticky="nsew")
frame_cadastro.grid_columnconfigure(0, weight=1)

titulo = ctk.CTkLabel(frame_cadastro, text="Formulario de cadastro")
titulo.grid(row=0, column=0, pady=5, sticky="ew")

nome = ctk.CTkEntry(frame_cadastro, placeholder_text="Nome")
nome.grid(row=1, column=0, pady=5, sticky="ew")

codigo = ctk.CTkEntry(frame_cadastro, placeholder_text="Código")
codigo.grid(row=2, column=0, pady=5, sticky="ew")

preco = ctk.CTkEntry(frame_cadastro, placeholder_text="Custo (usar ponto ao envez de virgula)")
preco.grid(row=3, column=0, pady=5, sticky="ew")

quantidade = ctk.CTkEntry(frame_cadastro, placeholder_text="Quantidade")
quantidade.grid(row=4, column=0, pady=5, sticky="ew")

venda = ctk.CTkEntry(frame_cadastro, placeholder_text="Valor de venda (usar ponto ao envez de virgula)")
venda.grid(row=5, column=0, pady=5, sticky="ew")

buton = ctk.CTkButton(frame_cadastro, text="Cadastrar", command=verifica_preenchimento)
buton.grid(row=6, column=0, pady=5, sticky="ew")

buton_limpar = ctk.CTkButton(frame_cadastro, text="Limpar formulario", command=limpar_campos)
buton_limpar.grid(row=7, column=0, pady=5, sticky="ew")

frame_opcoes = ctk.CTkFrame(frame_superior)
frame_opcoes.grid(row=0, column=1, padx=(10, 0), pady=10, sticky="nsew")
frame_opcoes.grid_columnconfigure(0, weight=1)
frame_opcoes.grid_rowconfigure(1, weight=0)

funcionalidades_label = ctk.CTkLabel(frame_opcoes, text="Funcionalidades")
funcionalidades_label.grid(row=0, column=0, pady=(5, 10), padx=10, sticky="ew")

pesquisa_entry = ctk.CTkEntry(frame_opcoes, placeholder_text="Insira um nome")
pesquisa_entry.grid(row=1, column=0, pady=5, padx=10, sticky="ew")

pesquisa_button = ctk.CTkButton(frame_opcoes, text="Pesquisar")
pesquisa_button.grid(row=2, column=0, pady=5, padx=10, sticky="ew")

apaga_button = ctk.CTkButton(frame_opcoes, text="Apagar Item Selecionado")
apaga_button.grid(row=3, column=0, pady=(5, 10), padx=10, sticky="ew")

area_visualizacao = ctk.CTkFrame(app)
area_visualizacao.grid(row=2, column=0, padx=20, pady=20, sticky="nsew")
area_visualizacao.grid_columnconfigure(0, weight=1)
area_visualizacao.grid_rowconfigure(0, weight=1)

mostra_treeview = ttk.Treeview(
    area_visualizacao,
    columns=("Código", "Nome", "Custo", "Quantidade", "Valor de venda"),
    show="headings",
    style=f"{'Dark' if tema_atual == 'dark' else 'Light'}.Treeview"
)
tabelaScorll = ttk.Scrollbar(area_visualizacao, orient="vertical", command=mostra_treeview.yview)
mostra_treeview.configure(yscrollcommand=tabelaScorll.set)
mostra_treeview.grid(row=0, column=0, sticky="nsew")
tabelaScorll.grid(row=0, column=1, sticky="ns")

mostra_treeview.heading("Código", text="Código")
mostra_treeview.heading("Nome", text="Nome")
mostra_treeview.heading("Custo", text="Custo")
mostra_treeview.heading("Quantidade", text="Quantidade")
mostra_treeview.heading("Valor de venda", text="Valor de venda")
mostra_treeview.column("#0", width=0, stretch=tk.NO)
mostra_treeview.column("Código", anchor=tk.CENTER, width=80)
mostra_treeview.column("Nome", anchor=tk.W, width=150)
mostra_treeview.column("Custo", anchor=tk.E, width=80)
mostra_treeview.column("Quantidade", anchor=tk.CENTER, width=100)
mostra_treeview.column("Valor de venda", anchor=tk.E, width=120)

mostra_estoque()
app.mainloop()