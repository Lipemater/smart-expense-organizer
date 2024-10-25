import os
import tkinter as tk
from tkinter import filedialog, messagebox
import subprocess
from processamento import fechar_arquivo



def carregar_arquivo(entry_planilha, tabela_entrada, tabela_saida):
    caminho_csv = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])

    caminho_xlsx = filedialog.askopenfilename(filetypes=[("XLSX files", "*.xlsx")])


    if not caminho_csv or not caminho_xlsx:
        messagebox.showerror("Erro", "Por favor, selecione ambos os arquivos CSV e XLSX.")
        return

    fechar_arquivo(caminho_xlsx, caminho_csv, entry_planilha, tabela_entrada, tabela_saida)
    return


# Cria a janela principal
def criar_interface():

    root = tk.Tk()
    root.title("Carregar Extrato Bancário")

    # Dimensões da janela e centralização
    root.geometry("500x400")  # Tamanho da janela
    root.configure(bg="#2E3440")  # Fundo da janela com uma cor moderna (cinza escuro)
    root.resizable(False, False)  # Janela fixa, sem redimensionamento

    # Centralizar a janela na tela
    window_width = 500
    window_height = 400
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    position_top = int(screen_height/2 - window_height/2)
    position_right = int(screen_width/2 - window_width/2)
    root.geometry(f'{window_width}x{window_height}+{position_right}+{position_top}')
    
    # Fonte e cores modernas
    font_label = ("Segoe UI", 12)
    font_entry = ("Segoe UI", 10)
    label_color = "#D8DEE9"  # Cor de texto clara
    entry_bg = "#4C566A"  # Cor de fundo das entradas
    entry_fg = "#ECEFF4"  # Cor do texto das entradas
    button_bg = "#5E81AC"  # Azul suave para botões
    button_fg = "#ECEFF4"  # Cor clara para o texto dos botões

    # Frame para centralizar widgets
    frame = tk.Frame(root, bg="#2E3440")
    frame.pack(expand=True)

    # Labels e Entradas com estilo moderno
    label_planilha = tk.Label(frame, text="Nome da Planilha:", bg="#2E3440", fg=label_color, font=font_label)
    label_planilha.grid(row=0, column=0, pady=10, sticky="e")
    
    entry_planilha = tk.Entry(frame, width=30, font=font_entry, bg=entry_bg, fg=entry_fg, bd=0, insertbackground=entry_fg)
    entry_planilha.grid(row=0, column=1, padx=10, pady=10)

    label_tabela_entrada = tk.Label(frame, text="Tabela de Entradas:", bg="#2E3440", fg=label_color, font=font_label)
    label_tabela_entrada.grid(row=1, column=0, pady=10, sticky="e")

    tabela_entrada = tk.Entry(frame, width=30, font=font_entry, bg=entry_bg, fg=entry_fg, bd=0, insertbackground=entry_fg)
    tabela_entrada.grid(row=1, column=1, padx=10, pady=10)

    label_tabela_saida = tk.Label(frame, text="Tabela de Saídas:", bg="#2E3440", fg=label_color, font=font_label)
    label_tabela_saida.grid(row=2, column=0, pady=10, sticky="e")

    tabela_saida = tk.Entry(frame, width=30, font=font_entry, bg=entry_bg, fg=entry_fg, bd=0, insertbackground=entry_fg)
    tabela_saida.grid(row=2, column=1, padx=10, pady=10)

    # Botão moderno
    botao_carregar = tk.Button(frame, text="Carregar Arquivo", 
                               command=lambda: carregar_arquivo(entry_planilha, tabela_entrada, tabela_saida),
                               bg=button_bg, fg=button_fg, font=font_label, activebackground="#4C566A", bd=0, padx=20, pady=5)
    botao_carregar.grid(row=3, columnspan=2, pady=30)

    # Iniciar a interface
    root.mainloop()