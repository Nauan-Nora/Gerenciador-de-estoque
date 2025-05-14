from kivy.uix.treeview import TreeView
import openpyxl
import customtkinter as ctk
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox

arquivo = openpyxl.load_workbook("dados.xlsx")
ativa = arquivo['Plan1']
tema_atual = "white"

app = ctk.CTk()
app.geometry("300x400")
app.title("Gerenciador de estoque")
app.iconbitmap('assets/icon.ico')

def cadastro_produto():
    main_frame.place_forget()

    def cadastrar_produto():

        v_nome = str(nome_entry.get())
        v_codigo = int(codigo_entry.get())
        v_preco = float(preco_entry.get())
        v_quantidade = int(quantidade_entry.get())
        v_venda = float(venda_entry.get())

        ultima_linha = ativa.max_row + 1
        ativa.cell(row=ultima_linha, column=1, value=v_codigo)
        ativa.cell(row=ultima_linha, column=2, value=v_nome)
        ativa.cell(row=ultima_linha, column=3, value=v_preco)
        ativa.cell(row=ultima_linha, column=4, value=v_quantidade)
        ativa.cell(row=ultima_linha, column=5, value=v_venda)

        arquivo.save("dados.xlsx")
        limpar_campos()
        mostra_estoque()
        nome_entry.focus_set()

    def limpar_campos():
        nome_entry.delete(0, "end")
        codigo_entry.delete(0, "end")
        preco_entry.delete(0, "end")
        quantidade_entry.delete(0, "end")
        venda_entry.delete(0, "end")

    def mostra_estoque():
        mostra_treeview.delete(*mostra_treeview.get_children())
        try:
            cabecalhos = [cell.value for cell in ativa[1]]
            mostra_treeview["columns"] = cabecalhos
            mostra_treeview["show"] = "headings"
            for col_index, col in enumerate(cabecalhos):
                mostra_treeview.heading(col, text=col)
                mostra_treeview.column(col, width=100)
            for row in ativa.iter_rows(min_row=2, values_only=True):
                mostra_treeview.insert('', tk.END, values=row)
        except IndexError:
            print("A planilha está vazia ou a primeira linha não contém cabeçalhos.")

    def pesquisar_item():
        termo_pesquisa = pesquisa_entry.get().lower()
        for item in mostra_treeview.get_children():
            valores = mostra_treeview.item(item, 'values')
            if valores:
                for valor in valores:
                    if termo_pesquisa in str(valor).lower():
                        mostra_treeview.see(item)
                        return

    def apagar_item():
        item_selecionado = mostra_treeview.selection()
        print(f"Itens selecionados na TreeView: {item_selecionado}")
        if not item_selecionado:
            messagebox.showerror("ATENÇÂO", "Não a nenhum item selecionado")
            return 

        item_id = item_selecionado[0]
        valores = mostra_treeview.item(item_id, 'values')
        print(f"Valores do item selecionado na TreeView: {valores}")
        if valores:
            nome_apagar = str(valores[1]).lower()
            print(f"Nome a apagar (convertido para minúsculo): {nome_apagar}")

            linhas_para_remover = []
            for row_index, row in enumerate(ativa.iter_rows(min_row=2), start=2):
                try:
                    nome_excel = str(row[1].value).lower() if row[1].value else ""
                    print(f"Nome na linha {row_index} da planilha (convertido para minúsculo): {nome_excel}")
                    if nome_apagar == nome_excel:
                        linhas_para_remover.append(row_index)
                except IndexError:
                    print(f"Erro ao acessar a coluna 'Nome' na linha {row_index} da planilha.")

            print(f"Linhas a serem removidas da planilha: {linhas_para_remover}")
            for linha in reversed(linhas_para_remover):
                try:
                    ativa.delete_rows(linha)
                except IndexError:
                    print(f"Erro ao tentar deletar a linha {linha} da planilha.")

            arquivo.save("dados.xlsx")
            mostra_estoque()
        else:
            messagebox.showerror("ERRO", "Não foi possivel obter os valores da tabela. Verifique se o dados.xlsx esta salvo no diretorio correto")
        arquivo.save("dados.xlsx")
        mostra_estoque()

    def verifica_preenchimento():
        v_nome = nome_entry.get()
        v_codigo = codigo_entry.get()
        v_preco = preco_entry.get()
        v_quantidade = quantidade_entry.get()
        v_venda = venda_entry.get()

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
        
    def menu_principal():
        frame_cadastro.grid_forget()
        frame_opcoes.grid_forget()
        area_visualizacao.grid_forget()
        frame_superior.grid_forget()
        main_frame.place(relx=0.5, rely=0.5, anchor=tk.CENTER)
        app.geometry("300x400")
         

    ctk.set_appearance_mode(tema_atual)
    ctk.set_default_color_theme("green")

    app.geometry("830x710")

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

    frame_superior = ctk.CTkFrame(app, fg_color="transparent")
    frame_superior.grid(row=0, column=0, padx=20, pady=20, sticky="ew")
    frame_superior.grid_columnconfigure(0, weight=1)
    frame_superior.grid_columnconfigure(1, weight=0)

    frame_cadastro = ctk.CTkFrame(frame_superior)
    frame_cadastro.grid(row=0, column=0, padx=(0, 10), pady=10, sticky="nsew")
    frame_cadastro.grid_columnconfigure(0, weight=1)

    titulo = ctk.CTkLabel(frame_cadastro, text="Formulario de cadastro")
    titulo.grid(row=0, column=0, pady=5, sticky="ew")

    nome_entry = ctk.CTkEntry(frame_cadastro, placeholder_text="Nome")
    nome_entry.grid(row=1, column=0, pady=5, sticky="ew")

    codigo_entry = ctk.CTkEntry(frame_cadastro, placeholder_text="Código")
    codigo_entry.grid(row=2, column=0, pady=5, sticky="ew")

    preco_entry = ctk.CTkEntry(frame_cadastro, placeholder_text="Custo (usar ponto ao envez de virgula)")
    preco_entry.grid(row=3, column=0, pady=5, sticky="ew")

    quantidade_entry = ctk.CTkEntry(frame_cadastro, placeholder_text="Quantidade")
    quantidade_entry.grid(row=4, column=0, pady=5, sticky="ew")

    venda_entry = ctk.CTkEntry(frame_cadastro, placeholder_text="Valor de venda (usar ponto ao envez de virgula)")
    venda_entry.grid(row=5, column=0, pady=5, sticky="ew")

    buton_cadastrar = ctk.CTkButton(frame_cadastro, text="Cadastrar", command=verifica_preenchimento)
    buton_cadastrar.grid(row=6, column=0, pady=5, sticky="ew")

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

    pesquisa_button = ctk.CTkButton(frame_opcoes, text="Pesquisar", command=pesquisar_item)
    pesquisa_button.grid(row=2, column=0, pady=5, padx=10, sticky="ew")

    apaga_button = ctk.CTkButton(frame_opcoes, text="Apagar Item Selecionado", command=apagar_item)
    apaga_button.grid(row=3, column=0, pady=(5, 10), padx=10, sticky="ew")

    volta_button = ctk.CTkButton(frame_opcoes, text="Menu Principal", command=menu_principal)
    volta_button.grid(row=4, column=0, pady=(5, 10), padx=10, sticky="ew")
 
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

main_frame = ctk.CTkFrame(app, fg_color="transparent")
main_frame.place(relx=0.5, rely=0.5, anchor=tk.CENTER)


botao_cadas = ctk.CTkButton(main_frame, text="Sistema de cadastro", width=200, height=100, command=cadastro_produto)
botao_cadas.pack( pady=10)

botao_baixa = ctk.CTkButton(main_frame, text="Baixa de estoque", width=200, height=100)
botao_baixa.pack( pady=10)

botao_edicao = ctk.CTkButton(main_frame, text="Edição de produtos", width=200, height=100)
botao_edicao.pack(pady=10)

app.mainloop()