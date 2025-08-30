import pandas as pd
from rapidfuzz import process, fuzz
from tkinter import messagebox

class DataModel:
    """
    Gerencia o estado dos dados da aplicação, incluindo os DataFrames
    e a lógica de correspondência.
    """
    def __init__(self):
        self.df_orcamento = None
        self.df_referencia = None
        self.correspondencias_df = pd.DataFrame()
        self.cesta_df = pd.DataFrame()
        self.match_threshold = 85

    def carregar_planilha(self, file_path, tipo):
        """
        Carrega um arquivo Excel para o DataFrame correspondente.
        Args:
            file_path (str): O caminho do arquivo.
            tipo (str): 'orcamento' ou 'referencia'.
        Returns:
            bool: True se o carregamento foi bem-sucedido, False caso contrário.
        """
        try:
            df = pd.read_excel(file_path, header=None)
            if tipo == 'orcamento':
                self.df_orcamento = df
            else:
                self.df_referencia = df
            return True
        except Exception as e:
            messagebox.showerror("Erro de Leitura", f"Erro ao ler o arquivo: {str(e)}")
            return False

    def mesclar_planilhas(self):
        """
        Executa a correspondência fuzzy entre as planilhas de orçamento e referência.
        """
        if self.df_orcamento is None or self.df_referencia is None:
            return False

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
                    
                    linha_orc_original = self.df_orcamento.iloc[idx_orc + 1]
                    linha_ref_original = self.df_referencia.iloc[idx_ref + 1]
                    
                    materiais = linha_ref_original.iloc[4]
                    maodeobra = linha_ref_original.iloc[5]
                    quantidade = linha_orc_original.iloc[3]
                    
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
        return True

    def adicionar_ao_cesta_df(self, df_a_adicionar):
        """
        Adiciona itens à cesta de compras, tratando duplicatas.
        """
        if df_a_adicionar.empty:
            return

        # Garante que as colunas numéricas estão no formato correto
        for col in ['Unidade_Orcamento', 'Quantidade_Orcamento', 'Materiais_Referencia', 'MaoDeObra_Referencia']:
            df_a_adicionar[col] = pd.to_numeric(df_a_adicionar[col], errors='coerce').fillna(0)
            
        self.cesta_df = pd.concat([self.cesta_df, df_a_adicionar], ignore_index=True)
        self.cesta_df.drop_duplicates(subset=['Descricao_Orcamento'], inplace=True)
