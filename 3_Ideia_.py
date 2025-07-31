import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog # Importe simpledialog
import pandas as pd

# Variáveis globais para armazenar os DataFrames
df_descricao = None
df_medidas = None
coluna_chave_descricao = None # Nova variável para a coluna chave da planilha de descrição
coluna_chave_medidas = None  # Nova variável para a coluna chave da planilha de medidas

def carregar_planilha_descricao():
    global df_descricao, coluna_chave_descricao
    filepath = filedialog.askopenfilename(
        title="Selecione a Planilha de Descrição",
        filetypes=[("Arquivos Excel", "*.xlsx *.xls")]
    )
    if filepath:
        try:
            df_descricao = pd.read_excel(filepath)
            messagebox.showinfo("Sucesso", "Planilha de descrição carregada com sucesso!")
            # Após carregar, perguntar qual coluna usar como chave
            coluna_chave_descricao = simpledialog.askstring(
                "Chave da Planilha de Descrição",
                "Digite o nome da coluna que identifica o item na Planilha de Descrição (ex: 'ID_Produto', 'Nome Material'):",
                parent=root # Define a janela principal como pai do diálogo
            )
            if not coluna_chave_descricao or coluna_chave_descricao not in df_descricao.columns:
                messagebox.showwarning("Aviso", "Coluna chave da descrição inválida ou não informada. Por favor, recarregue.")
                df_descricao = None # Reinicia o df para forçar o usuário a carregar novamente
                coluna_chave_descricao = None
            else:
                messagebox.showinfo("Chave Definida", f"Chave da descrição definida como: '{coluna_chave_descricao}'")
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível carregar a planilha de descrição: {e}")

def carregar_planilha_medidas():
    global df_medidas, coluna_chave_medidas
    filepath = filedialog.askopenfilename(
        title="Selecione a Planilha de Medidas",
        filetypes=[("Arquivos Excel", "*.xlsx *.xls")]
    )
    if filepath:
        try:
            df_medidas = pd.read_excel(filepath)
            messagebox.showinfo("Sucesso", "Planilha de medidas carregada com sucesso!")
            # Após carregar, perguntar qual coluna usar como chave
            coluna_chave_medidas = simpledialog.askstring(
                "Chave da Planilha de Medidas",
                "Digite o nome da coluna que identifica o item na Planilha de Medidas (ex: 'Codigo', 'Descricao Produto'):",
                parent=root # Define a janela principal como pai do diálogo
            )
            if not coluna_chave_medidas or coluna_chave_medidas not in df_medidas.columns:
                messagebox.showwarning("Aviso", "Coluna chave das medidas inválida ou não informada. Por favor, recarregue.")
                df_medidas = None # Reinicia o df para forçar o usuário a carregar novamente
                coluna_chave_medidas = None
            else:
                messagebox.showinfo("Chave Definida", f"Chave das medidas definida como: '{coluna_chave_medidas}'")
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível carregar a planilha de medidas: {e}")

def realizar_juncao():
    if df_descricao is None or df_medidas is None:
        messagebox.showwarning("Aviso", "Por favor, carregue ambas as planilhas primeiro.")
        return
    if coluna_chave_descricao is None or coluna_chave_medidas is None:
        messagebox.showwarning("Aviso", "Por favor, defina as colunas chave para ambas as planilhas.")
        return

    try:
        # AQUI USAMOS left_on e right_on com as colunas que o usuário informou
        df_junto = pd.merge(df_descricao, df_medidas, left_on=coluna_chave_descricao, right_on=coluna_chave_medidas, how='inner')
        messagebox.showinfo("Sucesso", "Junção realizada com sucesso! Verifique o console para pré-visualização.")
        print("--- DataFrame Consolidado ---")
        print(df_junto.head()) # Exibe as primeiras linhas no console

        # Opcional: Salvar o DataFrame resultante em um novo arquivo
        output_filepath = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Arquivos Excel", "*.xlsx")],
            title="Salvar Planilha Consolidada"
        )
        if output_filepath:
            df_junto.to_excel(output_filepath, index=False)
            messagebox.showinfo("Sucesso", f"Planilha consolidada salva em: {output_filepath}")

    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao realizar a junção: {e}")


# --- Configuração da Janela Principal ---
root = tk.Tk()
root.title("Ferramenta de Junção de Planilhas Dinâmica")

# --- Botão para Carregar Planilha de Descrição ---
btn_carregar_descricao = tk.Button(
    root,
    text="1. Carregar Planilha de Descrição",
    command=carregar_planilha_descricao
)
btn_carregar_descricao.pack(pady=10)

# --- Botão para Carregar Planilha de Medidas ---
btn_carregar_medidas = tk.Button(
    root,
    text="2. Carregar Planilha de Medidas",
    command=carregar_planilha_medidas
)
btn_carregar_medidas.pack(pady=10)

# --- Botão para Realizar a Junção ---
btn_realizar_juncao = tk.Button(
    root,
    text="3. Realizar Junção e Salvar",
    command=realizar_juncao
)
btn_realizar_juncao.pack(pady=20)

# --- Iniciar o loop principal da interface ---
root.mainloop()

