import tkinter as tk
from tkinter import ttk, messagebox
import openpyxl
from openpyxl import Workbook
import os

# Nome do arquivo Excel
ARQUIVO_EXCEL = "produtos.xlsx"

# Criar planilha se não existir
if not os.path.exists(ARQUIVO_EXCEL):
    wb = Workbook()
    ws = wb.active
    ws.title = "Produtos"
    ws.append(["Categoria", "Subcategoria", "Nome", "Quantidade", "Preço"])
    wb.save(ARQUIVO_EXCEL)

# Função para salvar no Excel
def salvar():
    categoria = combo_categoria.get()
    subcategoria = combo_subcategoria.get()
    nome = entry_nome.get()
    quantidade = entry_qtd.get()
    preco = entry_preco.get()

    if not (categoria and subcategoria and nome and quantidade and preco):
        messagebox.showwarning("Aviso", "Preencha todos os campos!")
        return

    wb = openpyxl.load_workbook(ARQUIVO_EXCEL)
    ws = wb["Produtos"]
    ws.append([categoria, subcategoria, nome, quantidade, preco])
    wb.save(ARQUIVO_EXCEL)

    messagebox.showinfo("Sucesso", f"{nome} salvo com sucesso!")
    entry_nome.delete(0, tk.END)
    entry_qtd.delete(0, tk.END)
    entry_preco.delete(0, tk.END)

# Atualizar subcategorias conforme a categoria
def atualizar_subcategorias(event):
    categoria = combo_categoria.get()
    if categoria == "Bebidas":
        combo_subcategoria["values"] = ["Alcoólicas", "Refrigerantes"]
    elif categoria == "Comidas":
        combo_subcategoria["values"] = ["Salgados", "Petiscos"]

# Interface Tkinter
root = tk.Tk()
root.title("Cadastro de Produtos")
root.geometry("400x350")

tk.Label(root, text="Categoria:").pack(pady=5)
combo_categoria = ttk.Combobox(root, values=["Bebidas", "Comidas"])
combo_categoria.pack()
combo_categoria.bind("<<ComboboxSelected>>", atualizar_subcategorias)

tk.Label(root, text="Subcategoria:").pack(pady=5)
combo_subcategoria = ttk.Combobox(root, values=[])
combo_subcategoria.pack()

tk.Label(root, text="Nome do Produto:").pack(pady=5)
entry_nome = tk.Entry(root)
entry_nome.pack()

tk.Label(root, text="Quantidade:").pack(pady=5)
entry_qtd = tk.Entry(root)
entry_qtd.pack()

tk.Label(root, text="Preço:").pack(pady=5)
entry_preco = tk.Entry(root)
entry_preco.pack()

tk.Button(root, text="Salvar no Excel", command=salvar).pack(pady=20)

root.mainloop()
