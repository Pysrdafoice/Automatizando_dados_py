import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime

class View:
    """
    Gerencia a criação de todos os widgets e janelas da interface gráfica.
    """
    def __init__(self, root, controller):
        self.root = root
        self.controller = controller
        self.root.title("Ferramenta de Análise de Orçamentos")
        self.root.geometry("1000x600")
        
        self.btn_carregar_orcamento = None
        self.btn_carregar_referencia = None
        self.btn_mesclar = None
        self.btn_adicionar_cesta = None
        self.btn_ver_cesta = None
        self.btn_gerar_relatorio = None
        self.response_text = None
        self.tree_grupos = None
        self.tree_valores = None
        self.cesta_tree = None
        self.entry_linha_inicial = None
        self.entry_linha_final = None
        
        # Referências de janelas secundárias para evitar duplicação
        self.grupos_window = None
        self.valores_window = None
        self.cesta_window = None
        self.escolha_window = None
        self.prompt_window = None
        
        self.setup_interface()

    def setup_interface(self):
        """
        Cria a interface principal da aplicação.
        """
        main_frame = ttk.Frame(self.root, padding="15")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        
        load_frame = ttk.LabelFrame(main_frame, text="Carregar Planilhas", padding="10")
        load_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=10)
        
        ttk.Label(load_frame, text="Planilha de Orçamento:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.btn_carregar_orcamento = ttk.Button(load_frame, text="Carregar",
                                                 command=lambda: self.controller.carregar_planilha_gui(1))
        self.btn_carregar_orcamento.grid(row=0, column=1, sticky=tk.W, pady=5, padx=5)
        
        ttk.Label(load_frame, text="Planilha de Referência:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.btn_carregar_referencia = ttk.Button(load_frame, text="Carregar",
                                                  command=lambda: self.controller.carregar_planilha_gui(2))
        self.btn_carregar_referencia.grid(row=1, column=1, sticky=tk.W, pady=5, padx=5)
        
        self.btn_mesclar = ttk.Button(load_frame, text="Iniciar Análise",
                                      command=self.controller.mesclar_planilhas_gui, state=tk.DISABLED)
        self.btn_mesclar.grid(row=2, column=0, columnspan=2, pady=10)
        
        ttk.Label(main_frame, text="Log de Atividades:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.response_text = tk.Text(main_frame, height=5, wrap=tk.WORD, state=tk.DISABLED, bg="#f0f0f0")
        self.response_text.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        response_scrollbar = ttk.Scrollbar(main_frame, orient=tk.VERTICAL, command=self.response_text.yview)
        response_scrollbar.grid(row=2, column=1, sticky=tk.N+tk.S)
        self.response_text.config(yscrollcommand=response_scrollbar.set)
        
        main_frame.rowconfigure(2, weight=1)
        
        bottom_frame = ttk.Frame(main_frame, padding="10")
        bottom_frame.grid(row=3, column=0, sticky=(tk.W, tk.E, tk.S))
        
        self.btn_adicionar_cesta = ttk.Button(bottom_frame, text="Adicionar à Cesta",
                                              command=self.controller.escolher_adicionar_metodo, state=tk.DISABLED)
        self.btn_adicionar_cesta.grid(row=0, column=0, padx=5, pady=5)
        
        self.btn_ver_cesta = ttk.Button(bottom_frame, text="Visualizar Cesta",
                                       command=self.controller.visualizar_cesta_gui, state=tk.DISABLED)
        self.btn_ver_cesta.grid(row=0, column=1, padx=5, pady=5)
        
        self.btn_gerar_relatorio = ttk.Button(bottom_frame, text="Gerar Relatório (Excel)",
                                              command=self.controller.gerar_relatorio_gui, state=tk.DISABLED)
        self.btn_gerar_relatorio.grid(row=0, column=2, padx=5, pady=5)

    def update_log(self, message):
        """Adiciona uma mensagem ao log de atividades."""
        self.response_text.config(state=tk.NORMAL)
        log_message = f"{datetime.now().strftime('%H:%M:%S')} - {message}\n"
        self.response_text.insert(tk.END, log_message)
        self.response_text.see(tk.END)
        self.response_text.config(state=tk.DISABLED)
    
    def exibir_janelas_resultados(self, correspondencias_df):
        """Cria e exibe as janelas de resultados."""
        if self.grupos_window and self.grupos_window.winfo_exists():
            self.grupos_window.lift()
        else:
            grupos_window = tk.Toplevel(self.root)
            grupos_window.title("Grupos Encontrados")
            grupos_window.geometry("600x400")
            self.grupos_window = grupos_window
            
            tree_frame_grupos = ttk.Frame(self.grupos_window)
            tree_frame_grupos.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
            
            self.tree_grupos = ttk.Treeview(
                tree_frame_grupos, 
                columns=('descricao_orc', 'similaridade_pontuacao'), 
                show='headings'
            )
            self.tree_grupos.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            self.tree_grupos.heading('descricao_orc', text='Descrição Orçamento', anchor=tk.W)
            self.tree_grupos.heading('similaridade_pontuacao', text='Similaridade (%)', anchor=tk.W)
            self.tree_grupos.column('descricao_orc', width=300, minwidth=250, stretch=tk.YES)
            self.tree_grupos.column('similaridade_pontuacao', width=120, minwidth=100)
            
            scrollbar_y_grupos = ttk.Scrollbar(tree_frame_grupos, orient=tk.VERTICAL, command=self.tree_grupos.yview)
            scrollbar_y_grupos.pack(side=tk.RIGHT, fill=tk.Y)
            self.tree_grupos.configure(yscrollcommand=scrollbar_y_grupos.set)
            
            for i in self.tree_grupos.get_children():
                self.tree_grupos.delete(i)
                
            for _, row in correspondencias_df.iterrows():
                self.tree_grupos.insert(
                    '', tk.END, values=(row['Descricao_Orcamento'], f"{row['Similaridade_Pontuacao']:.2f}%"))
        
        if self.valores_window and self.valores_window.winfo_exists():
            self.valores_window.lift()
        else:
            valores_window = tk.Toplevel(self.root)
            valores_window.title("Valores da Análise")
            valores_window.geometry("1000x600")
            self.valores_window = valores_window

            tree_frame_valores = ttk.Frame(self.valores_window)
            tree_frame_valores.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

            self.tree_valores = ttk.Treeview(
                tree_frame_valores,
                columns=('num_item', 'descricao_orc', 'unidade_orc', 'quantidade_orc',
                          'materiais_referencia', 'maodeobra_referencia', 'status_correspondencia'),
                show='headings'
            )
            self.tree_valores.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            self.tree_valores.heading('num_item', text='Nº Item', anchor=tk.W)
            self.tree_valores.heading('descricao_orc', text='Descrição Orçamento', anchor=tk.W)
            self.tree_valores.heading('unidade_orc', text='Unidade Orçamento', anchor=tk.W)
            self.tree_valores.heading('quantidade_orc', text='Quantidade', anchor=tk.W)
            self.tree_valores.heading('materiais_referencia', text='Materiais Ref.', anchor=tk.W)
            self.tree_valores.heading('maodeobra_referencia', text='Mão de Obra Ref.', anchor=tk.W)
            self.tree_valores.heading('status_correspondencia', text='Status', anchor=tk.W)
            
            scrollbar_y_valores = ttk.Scrollbar(tree_frame_valores, orient=tk.VERTICAL, command=self.tree_valores.yview)
            scrollbar_y_valores.pack(side=tk.RIGHT, fill=tk.Y)
            self.tree_valores.configure(yscrollcommand=scrollbar_y_valores.set)
            
            scrollbar_x_valores = ttk.Scrollbar(tree_frame_valores, orient=tk.HORIZONTAL, command=self.tree_valores.xview)
            scrollbar_x_valores.pack(side=tk.BOTTOM, fill=tk.X)
            self.tree_valores.configure(xscrollcommand=scrollbar_x_valores.set)
            
            for i in self.tree_valores.get_children():
                self.tree_valores.delete(i)

            for i, row in correspondencias_df.iterrows():
                self.tree_valores.insert('', tk.END, values=(
                    i + 1, row['Descricao_Orcamento'], row['Unidade_Orcamento'],
                    f"{row['Quantidade_Orcamento']:.2f}",
                    f"{row['Materiais_Referencia']:.2f}",
                    f"{row['MaoDeObra_Referencia']:.2f}",
                    row['Status_Correspondencia']
                ))

    def visualizar_cesta(self, cesta_df, excluir_callback):
        """Exibe a janela da cesta de compras."""
        if self.cesta_window and self.cesta_window.winfo_exists():
            self.cesta_window.lift()
            self.limpar_e_carregar_cesta(cesta_df)
        else:
            cesta_window = tk.Toplevel(self.root)
            cesta_window.title("Cesta de Compras")
            cesta_window.geometry("1000x600")
            self.cesta_window = cesta_window
            
            tree_frame = ttk.Frame(cesta_window)
            tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
            
            self.cesta_tree = ttk.Treeview(
                tree_frame, 
                columns=('numero_linha', 'descricao_orc', 'similaridade_pontuacao', 'unidade_orc',
                          'quantidade_orc', 'materiais_referencia', 'maodeobra_referencia', 'status_correspondencia'),
                show='headings'
            )
            self.cesta_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            self.cesta_tree.heading('numero_linha', text='Nº Item', anchor=tk.W)
            self.cesta_tree.heading('descricao_orc', text='Descrição Orçamento', anchor=tk.W)
            self.cesta_tree.heading('similaridade_pontuacao', text='Similaridade (%)', anchor=tk.W)
            self.cesta_tree.heading('unidade_orc', text='Unidade Orçamento', anchor=tk.W)
            self.cesta_tree.heading('quantidade_orc', text='Quantidade', anchor=tk.W)
            self.cesta_tree.heading('materiais_referencia', text='Materiais Ref.', anchor=tk.W)
            self.cesta_tree.heading('maodeobra_referencia', text='Mão de Obra Ref.', anchor=tk.W)
            self.cesta_tree.heading('status_correspondencia', text='Status', anchor=tk.W)

            self.limpar_e_carregar_cesta(cesta_df)
            
            scrollbar_y = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.cesta_tree.yview)
            scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
            self.cesta_tree.configure(yscrollcommand=scrollbar_y.set)
            
            scrollbar_x = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL, command=self.cesta_tree.xview)
            scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)
            self.cesta_tree.configure(xscrollcommand=scrollbar_x.set)
            
            btn_excluir = ttk.Button(cesta_window, text="Excluir Selecionado", command=excluir_callback)
            btn_excluir.pack(pady=5)
            
    def limpar_e_carregar_cesta(self, cesta_df):
        """Limpa e recarrega o Treeview da cesta."""
        if self.cesta_tree:
            for item in self.cesta_tree.get_children():
                self.cesta_tree.delete(item)
                
            for _, row in cesta_df.iterrows():
                valores_formatados = (
                    row['Numero_Linha'], row['Descricao_Orcamento'],
                    f"{row['Similaridade_Pontuacao']:.2f}%", row['Unidade_Orcamento'],
                    f"{row['Quantidade_Orcamento']:.2f}", f"{row['Materiais_Referencia']:.2f}",
                    f"{row['MaoDeObra_Referencia']:.2f}", row['Status_Correspondencia']
                )
                self.cesta_tree.insert('', tk.END, values=valores_formatados)
                
    def prompt_adicionar_por_intervalo(self, adicionar_callback, fechar_callback):
        """Cria a janela de prompt para adicionar por intervalo."""
        if self.prompt_window and self.prompt_window.winfo_exists():
            self.prompt_window.lift()
        else:
            prompt_window = tk.Toplevel(self.root)
            prompt_window.title("Adicionar por Intervalo de Linhas")
            prompt_window.geometry("400x200")
            prompt_window.resizable(False, False)
            self.prompt_window = prompt_window
            
            frame = ttk.Frame(prompt_window, padding="15")
            frame.pack(expand=True)
            
            ttk.Label(frame, text="Digite o intervalo de linhas que deseja adicionar:", 
                      font=("Arial", 10)).grid(row=0, column=0, columnspan=2, sticky=tk.W, pady=5)
            
            ttk.Label(frame, text="Linha Inicial:").grid(row=1, column=0, sticky=tk.W, pady=10)
            self.entry_linha_inicial = ttk.Entry(frame, width=15)
            self.entry_linha_inicial.grid(row=1, column=1, pady=10, padx=5)
            
            ttk.Label(frame, text="Linha Final:").grid(row=2, column=0, sticky=tk.W, pady=10)
            self.entry_linha_final = ttk.Entry(frame, width=15)
            self.entry_linha_final.grid(row=2, column=1, pady=10, padx=5)
            
            btn_frame = ttk.Frame(frame)
            btn_frame.grid(row=3, column=0, columnspan=2, pady=10)
            
            btn_confirmar = ttk.Button(btn_frame, text="Adicionar", 
                                       command=lambda: adicionar_callback(self.entry_linha_inicial.get(), self.entry_linha_final.get(), self.prompt_window))
            btn_confirmar.pack(side=tk.LEFT, padx=5)
            
            btn_cancelar = ttk.Button(btn_frame, text="Cancelar", command=fechar_callback)
            btn_cancelar.pack(side=tk.LEFT, padx=5)
