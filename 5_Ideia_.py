import pandas as pd
import numpy as np
from rapidfuzz import process, fuzz


file_orcamento = r'C:\Users\55719\OneDrive\Documentos\Automação_planilhas\PLANILHA ORC PADRAO - VAZIO.xlsx'
file_referencia = r'C:\Users\55719\OneDrive\Documentos\Automação_planilhas\PLANILHA ORC PADRAO - REFERÊNCIA DE PREÇOS.xlsx'

try:
    df_orcamento = pd.read_excel(file_orcamento, sheet_name='PLANILHA VENDA', header=None)
    df_referencia = pd.read_excel(file_referencia, sheet_name='PLANILHA VENDA', header=None)
except FileNotFoundError as e:
    print(f"Erro: Arquivo não encontrado. Verifique o caminho especificado. Detalhes: {e}")
    exit()

# --- Extração de descrições para a busca de correspondência ---

descricoes_orcamento = df_orcamento[1].astype(str).tolist()[1:]  # Ignora o cabeçalho
descricoes_referencia = df_referencia[1].astype(str).tolist()[1:] # Ignora o cabeçalho

# --- Configurações de correspondência

MATCH_THRESHOLD = 85  # Limite de similaridade (0-100)

# --- Lista para armazenamento dos objetos correspondentes
objetos_correspondentes = []

# --- Loop principal: itera sobre a planilha de referência e busca na de orçamento ---
print("Iniciando a busca por correspondências...")
for idx_ref, descricao_ref in enumerate(descricoes_referencia):
    # A função process.extractOne busca a melhor correspondência
    melhor_correspondencia = process.extractOne(
        descricao_ref,
        descricoes_orcamento,
        scorer=fuzz.WRatio,  # Usa Weighted Ratio para melhor precisão
        score_cutoff=MATCH_THRESHOLD # Garante que a pontuação é suficiente
    )

    # Se uma correspondência válida for encontrada
    if melhor_correspondencia:
        descricao_orc, pontuacao, idx_orc = melhor_correspondencia
        
        # Cria o objeto (dicionário) com os dados da planilha de REFERÊNCIA
        objeto = {
            "Descricao": df_referencia.iloc[idx_ref + 1, 1],
            "UnidadeDeMedida": df_referencia.iloc[idx_ref + 1, 2],
            "ValorDaMedida": df_referencia.iloc[idx_ref + 1, 3],
            "MaoDeObra": df_referencia.iloc[idx_ref + 1, 4],
            "Materiais": df_referencia.iloc[idx_ref + 1, 5]
        }
        
        objetos_correspondentes.append(objeto)
        
        # Log para depuração
        print(f"Correspondência encontrada: '{descricao_ref}' na referência com '{descricao_orc}' no orçamento (Pontuação: {pontuacao:.2f})")

# --- Exibição do resultado ---
print("\n--- Resultado Final ---")
if objetos_correspondentes:
    # Exibir o resultado como DataFrame
    
    df_resultado = pd.DataFrame(objetos_correspondentes)
    print(df_resultado)
else:
    print("Nenhuma correspondência válida foi encontrada.")

print("\nProcesso concluído.")
