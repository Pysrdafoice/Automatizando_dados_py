import pandas as pd
import numpy as np
from rapidfuzz import process, fuzz
import sys

# --- Configurações de Arquivo ---
# Defina aqui os nomes dos arquivos de entrada e saída.
ARQUIVO_ORCAMENTO = 'PlanilhaOrçamento.xlsx'
ARQUIVO_REFERENCIA = 'PlanilhaReferencia.xlsx'
MATCH_THRESHOLD = 85  # Limiar de similaridade (0-100) para aceitar uma correspondência.

# -----------------------------
# Funções de Leitura e Preparação de Dados
# -----------------------------

def carregar_planilhas(caminho_orcamento, caminho_referencia):
    """
    Carrega as planilhas de orçamento e referência a partir de arquivos Excel.
    
    Args:
        caminho_orcamento (str): Caminho para o arquivo de orçamento.
        caminho_referencia (str): Caminho para o arquivo de referência.
        
    Returns:
        tuple: DataFrames (df_orcamento, df_referencia).
        
    Raises:
        SystemExit: Se algum arquivo não for encontrado.
    """
    try:
        # Puxa a primeira aba de cada arquivo.
        df_orcamento = pd.read_excel(caminho_orcamento)
        df_referencia = pd.read_excel(caminho_referencia)
        return df_orcamento, df_referencia
    except FileNotFoundError as e:
        print(f"Erro: Arquivo não encontrado. Verifique os caminhos especificados. Detalhes: {e}")
        sys.exit()

def extrair_descricoes(df):
    """
    Extrai as descrições da segunda coluna (índice 1) de um DataFrame,
    ignorando a primeira linha (cabeçalho).
    
    Args:
        df (DataFrame): DataFrame de entrada.
        
    Returns:
        list: Lista de descrições como strings.
    """
    # Verifica se o DataFrame não está vazio para evitar erros
    if df.empty or df.shape[1] < 2:
        return []
    # Retorna a coluna de descrição a partir da segunda linha
    return df.iloc[1:, 1].astype(str).tolist()

# -----------------------------
# Funções de Análise e Matching
# -----------------------------

def encontrar_correspondencias(df_orcamento, df_referencia, threshold=MATCH_THRESHOLD):
    """
    Encontra a melhor correspondência para cada item da planilha de referência
    na planilha de orçamento, com base na similaridade de descrição.
    
    Args:
        df_orcamento (DataFrame): DataFrame de orçamento.
        df_referencia (DataFrame): DataFrame de referência.
        threshold (int): Pontuação mínima de similaridade para aceitar um match.
        
    Returns:
        DataFrame: Um DataFrame com todas as correspondências e seus detalhes.
    """
    descricoes_orcamento = extrair_descricoes(df_orcamento)
    descricoes_referencia = extrair_descricoes(df_referencia)

    resultados_correspondencias = []
    
    print("Iniciando a busca por correspondências...")
    
    for idx_ref, descricao_ref in enumerate(descricoes_referencia):
        # A busca de correspondência é feita apenas se a descrição não for vazia ou "nan"
        if descricao_ref.strip().lower() != "nan" and descricao_ref.strip() != "":
            melhor_correspondencia = process.extractOne(
                descricao_ref,
                descricoes_orcamento,
                scorer=fuzz.WRatio,
                score_cutoff=threshold
            )
            
            # Pega a linha original do DataFrame, ajustando para o cabeçalho.
            linha_ref_original = df_referencia.iloc[idx_ref + 1]
            
            if melhor_correspondencia:
                descricao_orc, pontuacao, idx_orc = melhor_correspondencia
                linha_orc_original = df_orcamento.iloc[idx_orc + 1]
                
                resultados_correspondencias.append({
                    "Descricao_Referencia": descricao_ref,
                    "Descricao_Orcamento": descricao_orc,
                    "Similaridade_Pontuacao": f"{pontuacao:.2f}%",
                    "Unidade_Orcamento": linha_orc_original.iloc[2],
                    "Quantidade_Orcamento": linha_orc_original.iloc[3],
                    "Materiais_Referencia": linha_ref_original.iloc[4],
                    "MaoDeObra_Referencia": linha_ref_original.iloc[5],
                    "Status_Correspondencia": "Encontrado"
                })
            
    # Cria o DataFrame final com os resultados.
    return pd.DataFrame(resultados_correspondencias)

def detectar_grupos(df, col_index=1):
    """
    Identifica grupos de itens em um DataFrame baseando-se em linhas vazias
    na coluna de descrição (col_index).
    
    Args:
        df (DataFrame): O DataFrame a ser analisado.
        col_index (int): O índice da coluna de descrição.
        
    Returns:
        list: Uma lista de dicionários, onde cada um representa um grupo
              com seu cabeçalho, início e fim.
    """
    groups = []
    start_idx = None
    header_name = None
    
    # A iteração começa na segunda linha para ignorar o cabeçalho
    for i in range(1, df.shape[0]):
        val = str(df.iloc[i, col_index]).strip()
        is_blank_or_nan = not val or val.lower() == "nan"

        if not is_blank_or_nan:
            if start_idx is None:
                start_idx = i
                header_name = val
        elif is_blank_or_nan and start_idx is not None:
            end_idx = i - 1
            groups.append({"header": header_name, "start_row": start_idx, "end_row": end_idx})
            start_idx = None
            header_name = None
    
    # Adiciona o último grupo se o arquivo não terminar com uma linha em branco
    if start_idx is not None:
        groups.append({"header": header_name, "start_row": start_idx, "end_row": df.shape[0] - 1})
        
    return groups

# -----------------------------
# Função Principal do Fluxo de Trabalho
# -----------------------------

def main():
    """
    Orquestra o processo de mesclagem, análise e interação com o usuário.
    """
    print("--- Iniciando o Processo de Mesclagem e Análise ---")
    
    # 1. Carregar as planilhas
    df_orcamento, df_referencia = carregar_planilhas(ARQUIVO_ORCAMENTO, ARQUIVO_REFERENCIA)
    
    # 2. Encontrar todas as correspondências
    df_resultados_gerais = encontrar_correspondencias(df_orcamento, df_referencia, MATCH_THRESHOLD)
    
    # 3. Detectar os grupos na planilha de referência
    grupos_referencia = detectar_grupos(df_referencia)
    
    # Cria um dicionário para busca rápida de grupos por nome ou número
    grupos_dict = {str(i + 1): grupo for i, grupo in enumerate(grupos_referencia)}
    grupos_dict.update({grupo['header'].lower(): grupo for grupo in grupos_referencia})
    
    cesta_de_compras = []

    # Loop de interação com o usuário
    while True:
        print("\n" + "="*50)
        print("Grupos disponíveis para análise:")
        for i, grupo in enumerate(grupos_referencia):
            print(f"  [{i + 1}] - {grupo['header']}")
        
        escolha = input("\nDigite o NÚMERO, NOME do grupo ou um INTERVALO (ex: '2-5'). Digite 'sair' para encerrar: ").strip().lower()

        if escolha == 'sair':
            print("Programa encerrado.")
            break

        grupos_para_analisar = []

        # Tenta interpretar a entrada do usuário
        if '-' in escolha:
            try:
                start_num, end_num = map(int, escolha.split('-'))
                if start_num > end_num:
                    start_num, end_num = end_num, start_num # Inverte se o usuário digitar ao contrário
                
                # Valida os números do intervalo
                if 1 <= start_num <= len(grupos_referencia) and 1 <= end_num <= len(grupos_referencia):
                    grupos_para_analisar = grupos_referencia[start_num - 1:end_num]
                else:
                    print("Intervalo inválido. Os números devem corresponder aos grupos disponíveis.")
            except ValueError:
                print("Entrada inválida para o intervalo. Use o formato 'número-número' (ex: '2-5').")
        else:
            # Tenta encontrar um único grupo
            grupo_encontrado = grupos_dict.get(escolha)
            if grupo_encontrado:
                grupos_para_analisar.append(grupo_encontrado)
            else:
                print(f"Grupo ou número '{escolha}' não encontrado. Por favor, tente novamente.")
        
        if grupos_para_analisar:
            for grupo_encontrado in grupos_para_analisar:
                print(f"\n--- Sumário Detalhado para o Grupo: {grupo_encontrado['header']} ---")
                
                # Filtra os resultados para este grupo específico
                df_grupo = df_resultados_gerais.iloc[grupo_encontrado["start_row"]:grupo_encontrado["end_row"] + 1].copy()
                
                # Exibe o sumário no console, se houver itens
                if not df_grupo.empty:
                    # Mostra todas as colunas, exceto "Descricao_Referencia"
                    cols_to_display = [
                        "Descricao_Orcamento",
                        "Similaridade_Pontuacao",
                        "Unidade_Orcamento",
                        "Quantidade_Orcamento",
                        "Materiais_Referencia",
                        "MaoDeObra_Referencia",
                        "Status_Correspondencia"
                    ]
                    # Adiciona uma coluna de índice para o usuário selecionar os itens
                    df_grupo_display = df_grupo.reset_index(drop=True)
                    df_grupo_display.index = df_grupo_display.index + 1
                    
                    print(df_grupo_display[cols_to_display].to_string())

                    while True:
                        adicionar_itens = input("\nDeseja analisar algum item separadamente? (sim/nao): ").strip().lower()
                        if adicionar_itens == 'nao':
                            break
                        elif adicionar_itens == 'sim':
                            try:
                                indices_str = input("Digite os NÚMEROS dos itens separados por vírgula (ex: 1, 3, 5): ").strip()
                                indices = [int(i.strip()) for i in indices_str.split(',') if i.strip()]
                                
                                for idx in indices:
                                    if 1 <= idx <= len(df_grupo_display):
                                        item_selecionado = df_grupo_display.iloc[idx - 1].to_dict()
                                        cesta_de_compras.append(item_selecionado)
                                        print(f"Item {idx} adicionado à cesta.")
                                    else:
                                        print(f"Número de item inválido: {idx}. Ignorando.")
                                break
                            except ValueError:
                                print("Entrada inválida. Por favor, digite apenas números separados por vírgula.")
                        else:
                            print("Opção inválida. Por favor, digite 'sim' ou 'nao'.")

                else:
                    print("Nenhum item correspondente encontrado para este grupo.")
            
    
    # Após o loop principal, verifica se há itens na cesta para exibir.
    if cesta_de_compras:
        print("\n--- Analisando Cesta de Compras... ---")
        df_cesta = pd.DataFrame(cesta_de_compras)
        
        # Mostra o DataFrame final no console
        print("\n" + "="*50)
        print("RESUMO DA CESTA DE COMPRAS")
        print("="*50)
        print(df_cesta.to_string(index=False))
        print("\nAnálise da cesta de compras concluída.")

if __name__ == "__main__":
    main()