import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from rapidfuzz import process, fuzz
from datetime import datetime

class AnaliseOrcamentosApp:
    """
    Classe principal para a aplicação de Análise e Auditoria de Orçamentos com GUI.
    """
    def __init__(self, root):
        self.root = root
        self.root.title("Ferramenta de Análise de Orçamentos")
        self.root.geometry("1000x600")
        
        # Variáveis de estado para armazenar os DataFrames e resultados
        self.df_orcamento = None
        self.df_referencia = None
        self.correspondencias_df = pd.DataFrame()
        self.cesta_df = pd.DataFrame()
        
        # --- NOVAS VARIÁVEIS PARA REFERÊNCIAS DE JANELAS SECUNDÁRIAS ---
        # Isso permite verificar se uma janela já está aberta.
        self.grupos_window = None
        self.valores_window = None
        self.cesta_window = None
        self.escolha_window = None
        self.prompt_window = None
        
        # Variáveis para as caixas de entrada do intervalo de linhas
        self.entry_linha_inicial = None
        self.entry_linha_final = None
        
        # Limiar de similaridade para o fuzzy matching
        self.match_threshold = 85
        
        # Configura a interface gráfica
        self.setup_interface()
        
    def setup_interface(self):
        """
        Cria e organiza todos os widgets da interface gráfica.
        """
        # Frame principal
        main_frame = ttk.Frame(self.root, padding="15")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        
        # --- Seção de Carregamento de Arquivos ---
        load_frame = ttk.LabelFrame(main_frame, text="Carregar Planilhas", padding="10")
        load_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=10)
        
        ttk.Label(load_frame, text="Planilha de Orçamento:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.btn_carregar_orcamento = ttk.Button(load_frame, text="Carregar",
                                                 command=lambda: self.carregar_planilha(1))
        self.btn_carregar_orcamento.grid(row=0, column=1, sticky=tk.W, pady=5, padx=5)
        
        ttk.Label(load_frame, text="Planilha de Referência:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.btn_carregar_referencia = ttk.Button(load_frame, text="Carregar",
                                                  command=lambda: self.carregar_planilha(2))
        self.btn_carregar_referencia.grid(row=1, column=1, sticky=tk.W, pady=5, padx=5)
        
        self.btn_mesclar = ttk.Button(load_frame, text="Iniciar Análise",
                                      command=self.mesclar_planilhas, state=tk.DISABLED)
        self.btn_mesclar.grid(row=2, column=0, columnspan=2, pady=10)
        
        # --- Seção de Respostas/Log ---
        ttk.Label(main_frame, text="Log de Atividades:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.response_text = tk.Text(main_frame, height=5, wrap=tk.WORD, state=tk.DISABLED, bg="#f0f0f0")
        self.response_text.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        response_scrollbar = ttk.Scrollbar(main_frame, orient=tk.VERTICAL, command=self.response_text.yview)
        response_scrollbar.grid(row=2, column=1, sticky=tk.N+tk.S)
        self.response_text.config(yscrollcommand=response_scrollbar.set)
        
        main_frame.rowconfigure(2, weight=1) # Permite que o Textbox se expanda
        
        # --- Seção da Cesta de Compras e Relatório ---
        bottom_frame = ttk.Frame(main_frame, padding="10")
        bottom_frame.grid(row=3, column=0, sticky=(tk.W, tk.E, tk.S))
        
        self.btn_adicionar_cesta = ttk.Button(bottom_frame, text="Adicionar à Cesta",
                                              command=self.escolher_adicionar_metodo, state=tk.DISABLED)
        self.btn_adicionar_cesta.grid(row=0, column=0, padx=5, pady=5)
        
        self.btn_ver_cesta = ttk.Button(bottom_frame, text="Visualizar Cesta",
                                       command=self.visualizar_cesta, state=tk.DISABLED)
        self.btn_ver_cesta.grid(row=0, column=1, padx=5, pady=5)
        
        self.btn_gerar_relatorio = ttk.Button(bottom_frame, text="Gerar Relatório (Excel)",
                                              command=self.gerar_relatorio, state=tk.DISABLED)
        self.btn_gerar_relatorio.grid(row=0, column=2, padx=5, pady=5)
        
    def carregar_planilha(self, tipo):
        """
        Abre uma caixa de diálogo para carregar um arquivo Excel.
        Args:
            tipo (int): 1 para orçamento, 2 para referência.
        """
        file_path = filedialog.askopenfilename(
            title=f"Selecione a {'Planilha de Orçamento' if tipo == 1 else 'Planilha de Referência'}",
            filetypes=[("Arquivos Excel", "*.xlsx")]
        )
        
        if file_path:
            try:
                # O parâmetro header=None garante que a primeira linha não seja tratada como cabeçalho
                df = pd.read_excel(file_path, header=None)
                if tipo == 1:
                    self.df_orcamento = df
                    self.update_log(f"Planilha de Orçamento carregada: '{file_path.split('/')[-1]}'")
                else:
                    self.df_referencia = df
                    self.update_log(f"Planilha de Referência carregada: '{file_path.split('/')[-1]}'")
                
                # Habilita o botão de análise se ambas as planilhas estiverem carregadas
                if self.df_orcamento is not None and self.df_referencia is not None:
                    self.btn_mesclar.config(state=tk.NORMAL)
                    self.update_log("Pronto para iniciar a análise. Clique em 'Iniciar Análise'.")
            except Exception as e:
                messagebox.showerror("Erro de Leitura", f"Erro ao ler o arquivo: {str(e)}")
                self.update_log(f"ERRO: Não foi possível carregar o arquivo. Detalhes: {e}")
                
    def mesclar_planilhas(self):
        """
        Executa a lógica de mesclagem e correspondência dos dados e exibe os resultados em janelas separadas.
        """
        if self.df_orcamento is None or self.df_referencia is None:
            messagebox.showerror("Erro", "Por favor, carregue ambas as planilhas.")
            return

        self.update_log("Iniciando a análise e busca por correspondências...")
        
        # Extrai as descrições para a correspondência, ignorando a primeira linha
        descricoes_orc = self.df_orcamento.iloc[1:, 1].astype(str).dropna().tolist()
        descricoes_ref = self.df_referencia.iloc[1:, 1].astype(str).dropna().tolist()
        
        resultados = []
        
        for idx_ref, descricao_ref in enumerate(descricoes_ref):
            if descricao_ref and descricao_ref.strip() != 'nan':
                melhor_match = process.extractOne(
                    descricao_ref,
                    descricoes_orc,
                    scorer=fuzz.WRatio,
                    score_cutoff=self.match_threshold
                )
                
                if melhor_match:
                    descricao_orc, pontuacao, idx_orc = melhor_match
                    
                    # As linhas originais estão 1 a mais que o índice da lista (devido ao cabeçalho)
                    linha_orc_original = self.df_orcamento.iloc[idx_orc + 1]
                    linha_ref_original = self.df_referencia.iloc[idx_ref + 1]
                    
                    # Garantindo que os valores sejam numéricos antes de formatar
                    materiais = linha_ref_original.iloc[4]
                    maodeobra = linha_ref_original.iloc[5]
                    quantidade = linha_orc_original.iloc[3]
                    
                    # Dicionário com todas as informações necessárias.
                    # A numeração de linha agora é um contador simples para atender à solicitação do usuário.
                    resultados.append({
                        "Numero_Linha": len(resultados) + 1, 
                        "Descricao_Orcamento": str(linha_orc_original.iloc[1]),
                        "Similaridade_Pontuacao": float(pontuacao),
                        "Unidade_Orcamento": str(linha_orc_original.iloc[2]),
                        "Quantidade_Orcamento": float(quantidade) if pd.notna(quantidade) else 0.0,
                        "Materiais_Referencia": float(materiais) if pd.notna(materiais) else 0.0,
                        "MaoDeObra_Referencia": float(maodeobra) if pd.notna(maodeobra) else 0.0,
                        "Status_Correspondencia": "Correspondência Encontrada"
                    })
        
        self.correspondencias_df = pd.DataFrame(resultados)
        self.exibir_janelas_resultados()
        
        self.update_log(f"Análise concluída. {len(self.correspondencias_df)} correspondências encontradas.")
        
        # Habilita os botões para as próximas acciones
        if not self.correspondencias_df.empty:
            self.btn_adicionar_cesta.config(state=tk.NORMAL)
            self.btn_gerar_relatorio.config(state=tk.NORMAL)
            
    def exibir_janelas_resultados(self):
        """
        Cria e exibe as duas janelas de resultados (Grupos e Valores da Análise).
        Adiciona a verificação para evitar duplicação.
        """
        # --- Janela de Grupos Encontrados ---
        # Verifica se a janela já existe e está aberta
        if self.grupos_window and self.grupos_window.winfo_exists():
            self.grupos_window.lift() # Traz a janela para a frente
        else:
            grupos_window = tk.Toplevel(self.root)
            grupos_window.title("Grupos Encontrados")
            grupos_window.geometry("600x400")
            self.grupos_window = grupos_window # Armazena a referência
        
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
            
            # Limpa o Treeview antes de inserir novos dados
            for i in self.tree_grupos.get_children():
                self.tree_grupos.delete(i)
                
            for _, row in self.correspondencias_df.iterrows():
                self.tree_grupos.insert(
                    '', 
                    tk.END, 
                    values=(
                        row['Descricao_Orcamento'],
                        f"{row['Similaridade_Pontuacao']:.2f}%"
                    )
                )

        # --- Janela de Valores da Análise ---
        # Verifica se a janela já existe e está aberta
        if self.valores_window and self.valores_window.winfo_exists():
            self.valores_window.lift()
        else:
            valores_window = tk.Toplevel(self.root)
            valores_window.title("Valores da Análise")
            valores_window.geometry("1000x600")
            self.valores_window = valores_window # Armazena a referência

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

            self.tree_valores.column('num_item', width=60)
            self.tree_valores.column('descricao_orc', width=250, minwidth=200, stretch=tk.YES)
            self.tree_valores.column('unidade_orc', width=100)
            self.tree_valores.column('quantidade_orc', width=100)
            self.tree_valores.column('materiais_referencia', width=120)
            self.tree_valores.column('maodeobra_referencia', width=120)
            self.tree_valores.column('status_correspondencia', width=150)
            
            scrollbar_y_valores = ttk.Scrollbar(tree_frame_valores, orient=tk.VERTICAL, command=self.tree_valores.yview)
            scrollbar_y_valores.pack(side=tk.RIGHT, fill=tk.Y)
            self.tree_valores.configure(yscrollcommand=scrollbar_y_valores.set)
            
            scrollbar_x_valores = ttk.Scrollbar(tree_frame_valores, orient=tk.HORIZONTAL, command=self.tree_valores.xview)
            scrollbar_x_valores.pack(side=tk.BOTTOM, fill=tk.X)
            self.tree_valores.configure(xscrollcommand=scrollbar_x_valores.set)
            
            # Limpa o Treeview antes de inserir novos dados
            for i in self.tree_valores.get_children():
                self.tree_valores.delete(i)

            for i, row in self.correspondencias_df.iterrows():
                self.tree_valores.insert(
                    '', 
                    tk.END, 
                    values=(
                        i + 1,
                        row['Descricao_Orcamento'],
                        row['Unidade_Orcamento'],
                        f"{row['Quantidade_Orcamento']:.2f}",
                        f"{row['Materiais_Referencia']:.2f}",
                        f"{row['MaoDeObra_Referencia']:.2f}",
                        row['Status_Correspondencia']
                    )
                )
                
    def escolher_adicionar_metodo(self):
        """
        Abre uma janela pop-up para o usuário escolher o método de adição à cesta.
        Adiciona a verificação para evitar duplicação.
        """
        if self.escolha_window and self.escolha_window.winfo_exists():
            self.escolha_window.lift()
        else:
            escolha_window = tk.Toplevel(self.root)
            escolha_window.title("Escolher Método")
            escolha_window.geometry("300x120")
            escolha_window.resizable(False, False)
            self.escolha_window = escolha_window

            frame = ttk.Frame(escolha_window, padding="15")
            frame.pack(expand=True)
            
            ttk.Label(frame, text="Como deseja adicionar os itens?", font=("Arial", 10)).pack(pady=5)
            
            btn_grupos = ttk.Button(frame, text="Adicionar por Grupos (Seleção Atual)", command=self.adicionar_por_grupos)
            btn_grupos.pack(pady=5, fill=tk.X)
            
            btn_intervalo = ttk.Button(frame, text="Adicionar por Intervalo de Linhas", command=self.prompt_adicionar_por_intervalo)
            btn_intervalo.pack(pady=5, fill=tk.X)
            
    def adicionar_por_grupos(self):
        """
        Adiciona os itens selecionados do Treeview de grupos à cesta.
        """
        if self.tree_grupos is None or not self.tree_grupos.winfo_exists():
            messagebox.showerror("Erro", "A janela de grupos não está aberta ou foi fechada.")
            return
            
        selecionados = self.tree_grupos.selection()
        if not selecionados:
            messagebox.showwarning("Aviso", "Por favor, selecione os itens que deseja adicionar à cesta.")
            return

        novos_itens_values = [self.tree_grupos.item(item_id, 'values')[0] for item_id in selecionados]
        
        df_a_adicionar = self.correspondencias_df[self.correspondencias_df['Descricao_Orcamento'].isin(novos_itens_values)]
        self.adicionar_ao_cesta_df(df_a_adicionar)
        self.update_log(f"{len(novos_itens_values)} item(s) do grupo adicionados à cesta.")
        
    def prompt_adicionar_por_intervalo(self):
        """
        Abre a janela de prompt para adicionar itens por intervalo de linhas.
        Adiciona a verificação para evitar duplicação.
        """
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
                                       command=lambda: self.adicionar_por_intervalo(prompt_window))
            btn_confirmar.pack(side=tk.LEFT, padx=5)
            
            btn_cancelar = ttk.Button(btn_frame, text="Cancelar", 
                                       command=prompt_window.destroy)
            btn_cancelar.pack(side=tk.LEFT, padx=5)
            
    def adicionar_por_intervalo(self, window):
        """
        Adiciona os itens com base no intervalo de linhas especificado.
        """
        if self.entry_linha_inicial is None or self.entry_linha_final is None:
            messagebox.showerror("Erro", "Campos de entrada não inicializados. Tente novamente.")
            return

        linha_inicial_str = self.entry_linha_inicial.get().strip()
        linha_final_str = self.entry_linha_final.get().strip()
        
        # Verifica se ambos os campos estão preenchidos
        if not linha_inicial_str or not linha_final_str:
            messagebox.showerror("Erro", "Por favor, preencha ambos os campos (linha inicial e final).")
            return
            
        try:
            linha_inicial = int(linha_inicial_str)
            linha_final = int(linha_final_str)
        except ValueError:
            messagebox.showerror("Erro", "Por favor, digite números válidos para as linhas.")
            return
            
        # Verifica se a linha inicial é menor ou igual à linha final
        if linha_inicial > linha_final:
            messagebox.showerror("Erro", "A linha inicial deve ser menor ou igual à linha final.")
            return
            
        # Verifica se as linhas estão dentro do intervalo válido
        if self.correspondencias_df.empty:
            messagebox.showerror("Erro", "Nenhuma correspondência encontrada para adicionar.")
            return
            
        min_linha = self.correspondencias_df['Numero_Linha'].min()
        max_linha = self.correspondencias_df['Numero_Linha'].max()
        
        if linha_inicial < min_linha or linha_final > max_linha:
            messagebox.showerror("Erro", f"As linhas devem estar entre {min_linha} e {max_linha}.")
            return
            
        # Filtra o DataFrame para encontrar os itens no intervalo especificado
        df_a_adicionar = self.correspondencias_df[
            (self.correspondencias_df['Numero_Linha'] >= linha_inicial) & 
            (self.correspondencias_df['Numero_Linha'] <= linha_final)
        ]
        
        if not df_a_adicionar.empty:
            self.adicionar_ao_cesta_df(df_a_adicionar)
            self.update_log(f"{len(df_a_adicionar)} item(s) do intervalo de linhas {linha_inicial} a {linha_final} adicionados à cesta.")
            window.destroy()
        else:
            messagebox.showerror("Erro", f"Não foi possível encontrar itens no intervalo de linhas {linha_inicial} a {linha_final}.")
            
    def adicionar_ao_cesta_df(self, df_a_adicionar):
        """
        Adiciona uma DataFrame à cesta de compras, tratando duplicatas.
        """
        if df_a_adicionar.empty:
            return

        # Garante que as colunas 'Unidade_Orcamento', 'Quantidade_Orcamento',
        # 'Materiais_Referencia', 'MaoDeObra_Referencia'
        # estejam no formato numérico para evitar erros na concatenação
        df_a_adicionar['Unidade_Orcamento'] = pd.to_numeric(df_a_adicionar['Unidade_Orcamento'], errors='coerce').fillna(0)
        df_a_adicionar['Quantidade_Orcamento'] = pd.to_numeric(df_a_adicionar['Quantidade_Orcamento'], errors='coerce').fillna(0)
        df_a_adicionar['Materiais_Referencia'] = pd.to_numeric(df_a_adicionar['Materiais_Referencia'], errors='coerce').fillna(0)
        df_a_adicionar['MaoDeObra_Referencia'] = pd.to_numeric(df_a_adicionar['MaoDeObra_Referencia'], errors='coerce').fillna(0)

        # Concatenar e remover duplicatas
        self.cesta_df = pd.concat([self.cesta_df, df_a_adicionar], ignore_index=True)
        self.cesta_df.drop_duplicates(subset=['Descricao_Orcamento'], inplace=True)
        
        self.btn_ver_cesta.config(state=tk.NORMAL)

    def visualizar_cesta(self):
        """
        Exibe o conteúdo da cesta de compras em uma nova janela com opção de exclusão.
        Adiciona a verificação para evitar duplicação.
        """
        if self.cesta_df.empty:
            messagebox.showinfo("Cesta Vazia", "A cesta de compras está vazia. Adicione itens para visualizá-los.")
            return
        
        if self.cesta_window and self.cesta_window.winfo_exists():
            self.cesta_window.lift()
            # Limpa e recarrega os dados caso a cesta tenha sido alterada
            self.limpar_e_carregar_cesta()
        else:
            cesta_window = tk.Toplevel(self.root)
            cesta_window.title("Cesta de Compras")
            cesta_window.geometry("1000x600")
            self.cesta_window = cesta_window
            
            tree_frame = ttk.Frame(cesta_window)
            tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
            
            cesta_tree = ttk.Treeview(
                tree_frame, 
                columns=('numero_linha', 'descricao_orc', 'similaridade_pontuacao', 'unidade_orc',
                          'quantidade_orc', 'materiais_referencia', 'maodeobra_referencia', 'status_correspondencia'),
                show='headings'
            )
            cesta_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            self.cesta_tree = cesta_tree # Armazena a referência para o Treeview
            
            cesta_tree.heading('numero_linha', text='Nº Item', anchor=tk.W)
            cesta_tree.heading('descricao_orc', text='Descrição Orçamento', anchor=tk.W)
            cesta_tree.heading('similaridade_pontuacao', text='Similaridade (%)', anchor=tk.W)
            cesta_tree.heading('unidade_orc', text='Unidade Orçamento', anchor=tk.W)
            cesta_tree.heading('quantidade_orc', text='Quantidade', anchor=tk.W)
            cesta_tree.heading('materiais_referencia', text='Materiais Ref.', anchor=tk.W)
            cesta_tree.heading('maodeobra_referencia', text='Mão de Obra Ref.', anchor=tk.W)
            cesta_tree.heading('status_correspondencia', text='Status', anchor=tk.W)

            cesta_tree.column('numero_linha', width=100)
            cesta_tree.column('descricao_orc', width=250, minwidth=200, stretch=tk.YES)
            cesta_tree.column('similaridade_pontuacao', width=100)
            cesta_tree.column('unidade_orc', width=100)
            cesta_tree.column('quantidade_orc', width=100)
            cesta_tree.column('materiais_referencia', width=120)
            cesta_tree.column('maodeobra_referencia', width=120)
            cesta_tree.column('status_correspondencia', width=150)
                
            self.limpar_e_carregar_cesta() # Carrega os dados na primeira vez
                
            scrollbar_y = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=cesta_tree.yview)
            scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
            cesta_tree.configure(yscrollcommand=scrollbar_y.set)
            
            scrollbar_x = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL, command=cesta_tree.xview)
            scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)
            cesta_tree.configure(xscrollcommand=scrollbar_x.set)
            
            def excluir_selecionados():
                selecionados = self.cesta_tree.selection()
                if not selecionados:
                    messagebox.showwarning("Aviso", "Por favor, selecione os itens que deseja excluir.")
                    return

                confirmar = messagebox.askyesno("Confirmar Exclusão", 
                                                f"Tem certeza que deseja excluir {len(selecionados)} item(s) da cesta?")
                if confirmar:
                    descricoes_para_excluir = [self.cesta_tree.item(item)['values'][1] for item in selecionados]
                    
                    self.cesta_df = self.cesta_df[~self.cesta_df['Descricao_Orcamento'].isin(descricoes_para_excluir)]
                    
                    for item in selecionados:
                        self.cesta_tree.delete(item)
                        
                    self.update_log(f"{len(selecionados)} item(s) excluídos da cesta.")
            
            btn_excluir = ttk.Button(cesta_window, text="Excluir Selecionado", command=excluir_selecionados)
            btn_excluir.pack(pady=5)

    def limpar_e_carregar_cesta(self):
        """
        Limpa o Treeview da cesta e recarrega os dados mais recentes.
        """
        if self.cesta_tree:
            for item in self.cesta_tree.get_children():
                self.cesta_tree.delete(item)
                
            for _, row in self.cesta_df.iterrows():
                valores_formatados = (
                    row['Numero_Linha'],
                    row['Descricao_Orcamento'],
                    f"{row['Similaridade_Pontuacao']:.2f}%",
                    row['Unidade_Orcamento'],
                    f"{row['Quantidade_Orcamento']:.2f}",
                    f"{row['Materiais_Referencia']:.2f}",
                    f"{row['MaoDeObra_Referencia']:.2f}",
                    row['Status_Correspondencia']
                )
                self.cesta_tree.insert('', tk.END, values=valores_formatados)
                
    def gerar_relatorio(self):
        """
        Salva o conteúdo da cesta de compras em um arquivo Excel.
        """
        if self.cesta_df.empty:
            messagebox.showwarning("Aviso", "A cesta de compras está vazia. Adicione itens antes de gerar o relatório.")
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Arquivos Excel", "*.xlsx"), ("Todos os Arquivos", "*.*")]
        )
        
        if file_path:
            try:
                self.cesta_df.to_excel(file_path, index=False)
                self.update_log(f"Relatório gerado com sucesso: '{file_path.split('/')[-1]}'")
                messagebox.showinfo("Sucesso", "Relatório gerado com sucesso!")
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao gerar relatório: {str(e)}")
                self.update_log(f"ERRO: Não foi possível gerar o relatório. Detalhes: {e}")
                
    def update_log(self, message):
        """
        Adiciona uma mensagem ao log de atividades.
        """
        self.response_text.config(state=tk.NORMAL)
        log_message = f"{datetime.now().strftime('%H:%M:%S')} - {message}\n"
        self.response_text.insert(tk.END, log_message)
        self.response_text.see(tk.END) # Rola para a última linha
        self.response_text.config(state=tk.DISABLED)

if __name__ == "__main__":
    root = tk.Tk()
    app = AnaliseOrcamentosApp(root)
    root.mainloop()
