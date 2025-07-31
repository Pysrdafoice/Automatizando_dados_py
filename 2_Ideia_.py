import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from difflib import get_close_matches

class SistemaMateriais:
    def __init__(self, root):
        self.root = root
        self.root.title("Sistema de Cálculo de Materiais")
        self.df_descricoes = None
        self.df_medidas = None
        self.criar_interface()

    def criar_interface(self):
        # Frame principal
        self.frame = tk.Frame(self.root, padx=20, pady=20)
        self.frame.pack()

        # Botão para carregar arquivos
        self.btn_carregar = tk.Button(
            self.frame,
            text="Carregar Planilhas",
            command=self.carregar_planilhas,
            width=25
        )
        self.btn_carregar.grid(row=0, column=0, pady=10)

        # Campo de busca
        self.label_busca = tk.Label(self.frame, text="Digite a descrição do material:")
        self.entry_busca = tk.Entry(self.frame, width=40, state='disabled')
        self.btn_buscar = tk.Button(
            self.frame,
            text="Buscar",
            state='disabled',
            command=self.buscar_material
        )

        # Listbox com scrollbar
        self.listbox = tk.Listbox(self.frame, width=50, height=10)
        self.scrollbar = tk.Scrollbar(self.frame, orient="vertical")
        self.listbox.config(yscrollcommand=self.scrollbar.set)
        self.scrollbar.config(command=self.listbox.yview)

        # Posicionamento dos elementos
        self.label_busca.grid(row=1, column=0, sticky='w', pady=(10, 0))
        self.entry_busca.grid(row=2, column=0, pady=(0, 10))
        self.btn_buscar.grid(row=2, column=1, padx=(5, 0))
        self.listbox.grid(row=3, column=0, columnspan=2, pady=10)
        self.scrollbar.grid(row=3, column=2, sticky='ns')

        # Bind para seleção
        self.listbox.bind('<<ListboxSelect>>', self.selecionar_material)

    def carregar_planilhas(self):
        arquivo_desc = filedialog.askopenfilename(
            title="Selecione a planilha de DESCRIÇÕES",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )

        arquivo_medidas = filedialog.askopenfilename(
            title="Selecione a planilha de MEDIDAS",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )

        try:
            self.df_descricoes = pd.read_excel(arquivo_desc)
            self.df_medidas = pd.read_excel(arquivo_medidas)

            self.entry_busca.config(state='normal')
            self.btn_buscar.config(state='normal')

            messagebox.showinfo("Sucesso", "Planilhas carregadas com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao carregar planilhas:\n{str(e)}")

    def buscar_material(self):
        termo_busca = self.entry_busca.get().strip().lower()

        if not termo_busca:
            messagebox.showwarning("Aviso", "Digite uma descrição para buscar")
            return

        self.listbox.delete(0, tk.END)

        descricoes = self.df_descricoes['Descrição'].str.lower().tolist()
        resultados = get_close_matches(termo_busca, descricoes, n=10, cutoff=0.6)

        if not resultados:
            self.listbox.insert(tk.END, "Nenhum material encontrado")
            return

        for resultado in resultados:
            self.listbox.insert(tk.END, resultado)

    def selecionar_material(self, event):
        try:
            selecionado = self.listbox.get(self.listbox.curselection())
        except tk.TclError:
            return  # Nenhuma seleção válida

        self.janela_calculo = tk.Toplevel(self.root)
        self.janela_calculo.title(f"Cálculo para: {selecionado}")

        medidas = self.df_medidas[self.df_medidas['Descrição'].str.lower() == selecionado.lower()]

        if medidas.empty:
            tk.Label(self.janela_calculo, text="Medidas não disponíveis para este material").pack()
            return

        tk.Label(self.janela_calculo, text=f"Material: {selecionado}").pack(pady=5)

        tk.Label(self.janela_calculo, text="Selecione a medida:").pack()
        self.var_medida = tk.StringVar()
        self.combo_medidas = ttk.Combobox(
            self.janela_calculo,
            textvariable=self.var_medida,
            values=medidas['Medida'].tolist(),
            state="readonly"
        )
        self.combo_medidas.pack(pady=5)

        tk.Label(self.janela_calculo, text="Quantidade necessária:").pack()
        self.entry_quantidade = tk.Entry(self.janela_calculo)
        self.entry_quantidade.pack(pady=5)

        tk.Button(
            self.janela_calculo,
            text="Calcular",
            command=lambda: self.calcular_quantidade(selecionado)
        ).pack(pady=10)

    def calcular_quantidade(self, descricao):
        try:
            medida_selecionada = self.var_medida.get()
            quantidade_necessaria = float(self.entry_quantidade.get())

            material = self.df_medidas[
                (self.df_medidas['Descrição'].str.lower() == descricao.lower()) &
                (self.df_medidas['Medida'] == medida_selecionada)
            ]

            if material.empty:
                raise ValueError("Medida não encontrada")

            quantidade_por_unidade = material['Quantidade'].values[0]
            materiais_necessarios = quantidade_necessaria / quantidade_por_unidade
            materiais_necessarios = int(materiais_necessarios) + 1 if materiais_necessarios % 1 > 0 else int(materiais_necessarios)

            idx = self.df_descricoes[
                self.df_descricoes['Descrição'].str.lower() == descricao.lower()
            ].index

            if not idx.empty:
                self.df_descricoes.at[idx[0], 'Medida Utilizada'] = medida_selecionada
                self.df_descricoes.at[idx[0], 'Quantidade Calculada'] = materiais_necessarios

            messagebox.showinfo(
                "Resultado",
                f"Serão necessários {materiais_necessarios} unidades de {descricao} ({medida_selecionada})"
            )

            self.janela_calculo.destroy()

        except ValueError as e:
            messagebox.showerror("Erro", f"Entrada inválida: {str(e)}")

# Execução do programa
if __name__ == "__main__":
    root = tk.Tk()
    app = SistemaMateriais(root)
    root.mainloop()

