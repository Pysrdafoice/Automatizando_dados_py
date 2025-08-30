from data_model import DataModel
from view import View
from tkinter import filedialog, messagebox
import tkinter as tk
from tkinter import ttk

class Controller:
    """
    Controla o fluxo da aplicação e a interação entre a View e o DataModel.
    """
    def __init__(self, root):
        self.model = DataModel()
        self.view = View(root, self)

    def carregar_planilha_gui(self, tipo):
        """
        Inicia o processo de carregamento de planilha via GUI.
        """
        file_path = filedialog.askopenfilename(
            title=f"Selecione a {'Planilha de Orçamento' if tipo == 1 else 'Planilha de Referência'}",
            filetypes=[("Arquivos Excel", "*.xlsx")]
        )
        
        if file_path:
            tipo_str = 'orcamento' if tipo == 1 else 'referencia'
            if self.model.carregar_planilha(file_path, tipo_str):
                self.view.update_log(f"Planilha de {'Orçamento' if tipo == 1 else 'Referência'} carregada: '{file_path.split('/')[-1]}'")
                if self.model.df_orcamento is not None and self.model.df_referencia is not None:
                    self.view.btn_mesclar.config(state=tk.NORMAL)
                    self.view.update_log("Pronto para iniciar a análise. Clique em 'Iniciar Análise'.")
    
    def mesclar_planilhas_gui(self):
        """
        Inicia o processo de mesclagem e atualiza a GUI.
        """
        self.view.update_log("Iniciando a análise e busca por correspondências...")
        if self.model.mesclar_planilhas():
            self.view.exibir_janelas_resultados(self.model.correspondencias_df)
            self.view.update_log(f"Análise concluída. {len(self.model.correspondencias_df)} correspondências encontradas.")
            if not self.model.correspondencias_df.empty:
                self.view.btn_adicionar_cesta.config(state=tk.NORMAL)
                self.view.btn_gerar_relatorio.config(state=tk.NORMAL)
        else:
            messagebox.showerror("Erro", "Por favor, carregue ambas as planilhas.")

    def escolher_adicionar_metodo(self):
        """
        Abre a janela para escolher o método de adição à cesta.
        """
        if self.view.escolha_window and self.view.escolha_window.winfo_exists():
            self.view.escolha_window.lift()
        else:
            escolha_window = tk.Toplevel(self.view.root)
            escolha_window.title("Escolher Método")
            escolha_window.geometry("300x120")
            escolha_window.resizable(False, False)
            self.view.escolha_window = escolha_window

            frame = ttk.Frame(escolha_window, padding="15")
            frame.pack(expand=True)
            
            ttk.Label(frame, text="Como deseja adicionar os itens?", font=("Arial", 10)).pack(pady=5)
            
            btn_grupos = ttk.Button(frame, text="Adicionar por Grupos (Seleção Atual)", command=self.adicionar_por_grupos)
            btn_grupos.pack(pady=5, fill=tk.X)
            
            btn_intervalo = ttk.Button(frame, text="Adicionar por Intervalo de Linhas", command=lambda: self.view.prompt_adicionar_por_intervalo(self.adicionar_por_intervalo, escolha_window.destroy))
            btn_intervalo.pack(pady=5, fill=tk.X)

    def adicionar_por_grupos(self):
        """
        *** PROBLEMA AQUI - PRECISO DO SEU CÓDIGO A SER CORRIGIDO ***
        Adiciona os itens selecionados do Treeview de grupos à cesta.
        O seu código original está com um problema que limita a seleção a 3 itens.
        Por favor, me envie o código que você implementou para esta função.
        """
        # Placeholder para o código a ser corrigido
        messagebox.showinfo("Aviso", "Por favor, me envie o código para a função 'adicionar_por_grupos' para que eu possa corrigi-lo.")

    def adicionar_por_intervalo(self, linha_inicial_str, linha_final_str, window):
        """
        Adiciona os itens com base no intervalo de linhas especificado.
        """
        linha_inicial = 0
        linha_final = 0
        
        try:
            linha_inicial = int(linha_inicial_str)
            linha_final = int(linha_final_str)
        except ValueError:
            messagebox.showerror("Erro", "Por favor, digite números válidos para as linhas.")
            return

        if linha_inicial > linha_final:
            messagebox.showerror("Erro", "A linha inicial deve ser menor ou igual à linha final.")
            return

        if self.model.correspondencias_df.empty:
            messagebox.showerror("Erro", "Nenhuma correspondência encontrada para adicionar.")
            return
            
        min_linha = self.model.correspondencias_df['Numero_Linha'].min()
        max_linha = self.model.correspondencias_df['Numero_Linha'].max()
        
        if linha_inicial < min_linha or linha_final > max_linha:
            messagebox.showerror("Erro", f"As linhas devem estar entre {min_linha} e {max_linha}.")
            return
            
        df_a_adicionar = self.model.correspondencias_df[
            (self.model.correspondencias_df['Numero_Linha'] >= linha_inicial) & 
            (self.model.correspondencias_df['Numero_Linha'] <= linha_final)
        ]
        
        if not df_a_adicionar.empty:
            self.model.adicionar_ao_cesta_df(df_a_adicionar)
            self.view.update_log(f"{len(df_a_adicionar)} item(s) do intervalo de linhas {linha_inicial} a {linha_final} adicionados à cesta.")
            window.destroy()
            self.view.btn_ver_cesta.config(state=tk.NORMAL)
        else:
            messagebox.showerror("Erro", f"Não foi possível encontrar itens no intervalo de linhas {linha_inicial} a {linha_final}.")
            
    def visualizar_cesta_gui(self):
        """
        Controla a exibição da cesta de compras.
        """
        if self.model.cesta_df.empty:
            messagebox.showinfo("Cesta Vazia", "A cesta de compras está vazia. Adicione itens para visualizá-los.")
            return
        
        self.view.visualizar_cesta(self.model.cesta_df, self.excluir_selecionados_cesta)

    def excluir_selecionados_cesta(self):
        """
        Exclui os itens selecionados da cesta.
        """
        selecionados = self.view.cesta_tree.selection()
        if not selecionados:
            messagebox.showwarning("Aviso", "Por favor, selecione os itens que deseja excluir.")
            return

        confirmar = messagebox.askyesno("Confirmar Exclusão", 
                                        f"Tem certeza que deseja excluir {len(selecionados)} item(s) da cesta?")
        if confirmar:
            descricoes_para_excluir = [self.view.cesta_tree.item(item)['values'][1] for item in selecionados]
            
            self.model.cesta_df = self.model.cesta_df[~self.model.cesta_df['Descricao_Orcamento'].isin(descricoes_para_excluir)]
            
            for item in selecionados:
                self.view.cesta_tree.delete(item)
                
            self.view.update_log(f"{len(selecionados)} item(s) excluídos da cesta.")

    def gerar_relatorio_gui(self):
        """
        Salva o conteúdo da cesta de compras em um arquivo Excel.
        """
        if self.model.cesta_df.empty:
            messagebox.showwarning("Aviso", "A cesta de compras está vazia. Adicione itens antes de gerar o relatório.")
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Arquivos Excel", "*.xlsx"), ("Todos os Arquivos", "*.*")]
        )
        
        if file_path:
            try:
                self.model.cesta_df.to_excel(file_path, index=False)
                self.view.update_log(f"Relatório gerado com sucesso: '{file_path.split('/')[-1]}'")
                messagebox.showinfo("Sucesso", "Relatório gerado com sucesso!")
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao gerar relatório: {str(e)}")
                self.view.update_log(f"ERRO: Não foi possível gerar o relatório. Detalhes: {e}")
