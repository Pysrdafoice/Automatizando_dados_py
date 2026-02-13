import tkinter as tk
from tkinter import ttk
from ParametrosProcessamento import ParametrosProcessamento

class TelaSelecaoReferencia:
    def __init__(self, root, item_orcamento, parametros: ParametrosProcessamento):
        self.root = root
        self.root.title("Seleção de Referência - Automação de Orçamento")
        self.root.geometry("900x400")
        self.root.configure(bg="#f5f5f5")

        # --- Cabeçalho ---
        frame_topo = tk.Frame(self.root, bg="#004c99")
        frame_topo.pack(fill="x")

        lbl_titulo = tk.Label(
            frame_topo, text="Seleção de Referência para o Item do Orçamento",
            fg="white", bg="#004c99", font=("Segoe UI", 14, "bold"), pady=10
        )
        lbl_titulo.pack()

        # --- Descrição do item do orçamento ---
        frame_item = tk.Frame(self.root, bg="#f5f5f5", pady=10)
        frame_item.pack(fill="x")

        lbl_item = tk.Label(
            frame_item, text=f"Item do Orçamento: {item_orcamento}",
            font=("Segoe UI", 12, "bold"), bg="#f5f5f5"
        )
        lbl_item.pack(padx=20, anchor="w")

        # --- Área de seleção de referências ---
        frame_refs = tk.Frame(self.root, bg="#ffffff", bd=1, relief="solid")
        frame_refs.pack(padx=20, pady=10, fill="both", expand=True)

        self.var_selecao = tk.StringVar(value=None)

        # Cabeçalho da tabela
        colunas = ["", "Descrição da Referência", "Unidade", "Mão de Obra", "Material", "Similaridade"]
        for i, texto in enumerate(colunas):
            lbl = tk.Label(frame_refs, text=texto, font=("Segoe UI", 10, "bold"),
                           bg="#d9d9d9", borderwidth=1, relief="solid", width=20, anchor="center")
            lbl.grid(row=0, column=i, sticky="nsew")

        # Linhas de referências
        for idx, ref in enumerate(referencias, start=1):
            tk.Radiobutton(frame_refs, variable=self.var_selecao, value=ref["id"],
                           bg="#ffffff").grid(row=idx, column=0, sticky="nsew", padx=5)
            tk.Label(frame_refs, text=ref["descricao"], bg="#ffffff", anchor="w").grid(row=idx, column=1, sticky="nsew")
            tk.Label(frame_refs, text=ref["unidade"], bg="#ffffff").grid(row=idx, column=2, sticky="nsew")
            tk.Label(frame_refs, text=f'R$ {ref["mao_obra"]:.2f}', bg="#ffffff").grid(row=idx, column=3, sticky="nsew")
            tk.Label(frame_refs, text=f'R$ {ref["material"]:.2f}', bg="#ffffff").grid(row=idx, column=4, sticky="nsew")
            tk.Label(frame_refs, text=f'{ref["similaridade"]:.1f}%', bg="#ffffff").grid(row=idx, column=5, sticky="nsew")

        for i in range(len(colunas)):
            frame_refs.grid_columnconfigure(i, weight=1)

        # --- Rodapé com botões ---
        frame_botoes = tk.Frame(self.root, bg="#f5f5f5", pady=15)
        frame_botoes.pack(fill="x")

        btn_prosseguir = ttk.Button(frame_botoes, text="Prosseguir", command=self.prosseguir)
        btn_prosseguir.pack(side="right", padx=(0, 20))

        btn_finalizar = ttk.Button(frame_botoes, text="Finalizar", command=self.finalizar)
        btn_finalizar.pack(side="left", padx=20)

        btn_pular = ttk.Button(frame_botoes, text="Pular", command=self.pular)
        btn_pular.pack(side="right", padx=10)

    def pular(self):
        print("Pular.")
        self.root.destroy()
 
    def prosseguir(self):
        selecao = self.var_selecao.get()
        if selecao:
            print(f"Referência selecionada: {selecao}")
            # Aqui você pode chamar a lógica para avançar para o próximo item
        else:
            print("Nenhuma referência selecionada.")

    def finalizar(self):
        print("Processo finalizado.")
        self.root.destroy()


# --- Exemplo de uso ---
if __name__ == "__main__":
    root = tk.Tk()
    item_orcamento = "Execução de alvenaria com blocos cerâmicos 9x19x39 cm"
    referencias = [
        {"id": "REF001", "descricao": "Alvenaria cerâmica estrutural", "unidade": "m²",
         "mao_obra": 35.50, "material": 48.30, "similaridade": 92.4},
        {"id": "REF002", "descricao": "Alvenaria de vedação com bloco cerâmico", "unidade": "m²",
         "mao_obra": 30.00, "material": 42.75, "similaridade": 87.6},
        {"id": "REF003", "descricao": "Alvenaria de tijolo maciço comum", "unidade": "m²",
         "mao_obra": 33.20, "material": 50.10, "similaridade": 84.3}
    ]
    TelaSelecaoReferencia(root, item_orcamento, referencias)
    root.mainloop()
