#!/usr/bin/env python3
"""
ARQUIVO DE TESTE - Atalho para testar a interface de processamento
Integração com processamento.py para testes rápidos da interface

Uso:
    python tela_correlacao_teste.py

Este arquivo executa a interface real de processamento com parâmetros 
de teste pré-configurados para agilizar testes sem passar por todos 
os formulários de entrada.
"""

import tkinter as tk
from pathlib import Path
import logging
import sys

# Configurar logging ANTES de importar o resto
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s | %(levelname)-8s | %(funcName)s():%(lineno)d | %(message)s',
    stream=sys.stdout
)
logger = logging.getLogger(__name__)

from parametrosPlanilha import ParametrosPlanilhas
from OperacaoCorrelacao import OperacaoCorrelacao
from ParametrosProcessamento import ParametrosProcessamento
from processamento import TelaProcessamento


def executar_teste():
    """Executa a tela de processamento com parâmetros de teste"""
    
    # Configurar caminhos de teste
    BASE_DIR = Path(__file__).parent
    CAMINHO_REFERENCIA = str(BASE_DIR / "ArquivosDados" / "PlanilhaReferencia.xlsx")
    CAMINHO_ORCAMENTO = str(BASE_DIR / "ArquivosDados" / "PlanilhaOrçamento.xlsx")
    
    print(f"\n{'='*80}")
    print(" TESTE: tela_correlacao_teste.py".center(80))
    print(f"{'='*80}")
    print(f"\n✓ BASE_DIR: {BASE_DIR}")
    print(f"✓ CAMINHO_REFERENCIA: {CAMINHO_REFERENCIA}")
    print(f"✓ CAMINHO_ORCAMENTO: {CAMINHO_ORCAMENTO}")
    print(f"\n✓ Arquivo ref existe? {Path(CAMINHO_REFERENCIA).exists()}")
    print(f"✓ Arquivo orc existe? {Path(CAMINHO_ORCAMENTO).exists()}")
    print(f"{'='*80}\n")
    
    # ===== CONFIGURAÇÃO DE TESTE =====
    # Ajuste estes parâmetros conforme necessário
    parametros = ParametrosProcessamento(
        referencia=ParametrosPlanilhas(
            caminho_planilha=CAMINHO_REFERENCIA,
            aba="Planilha de Custo",                    # ← Ajustar nome da aba conforme planilha
            coluna_descrição="B",             # ← Coluna com descrição da referência
            coluna_material="E",              # ← Coluna com valor de material
            coluna_mao_de_obra="F",           # ← Coluna com valor de mão de obra
            coluna_unidade_medida="C"         # ← Coluna com unidade de medida
        ),
        orcamento=ParametrosPlanilhas(
            caminho_planilha=CAMINHO_ORCAMENTO,
            aba="Planilha de Custo",                 # ← Ajustar nome da aba conforme planilha
            coluna_descrição="B",             # ← Coluna com descrição do orçamento
            coluna_material="E",              # ← Coluna com valor de material
            coluna_mao_de_obra="F",           # ← Coluna com valor de mão de obra
            coluna_unidade_medida="C"         # ← Coluna com unidade de medida
        ),
        pesquisa=OperacaoCorrelacao(
            ComecoPesquisa=2,                # ← Primeira linha a processar
            TerminoPesquisa=20,              # ← Última linha a processar
            TaxaSimilaridade=0.70            # ← Taxa mínima de similaridade (0.0 a 1.0)
        ),
        aba_pesquisa="Planilha de Custo"  # ← Aba do orçamento a ser processada
    )
    
    print("[TESTE] Criando tk.Tk()...")
    # Criar e executar interface
    root = tk.Tk()
    root.title("Teste: Tela de Processamento")
    root.geometry("1200x700")
    
    print("[TESTE] Criando TelaProcessamento (ÚNICA instância)...")
    try:
        TelaProcessamento(root, parametros)
        print("[TESTE] ✓ TelaProcessamento criada com SUCESSO")
    except Exception as e:
        logger.error(f"❌ ERRO ao criar TelaProcessamento: {e}")
        import traceback
        traceback.print_exc()
        return
    
    print("[TESTE] Iniciando mainloop...")
    print("[TESTE] 👉 Clique em um item do orçamento para selecionar uma referência")
    print("[TESTE] 👉 Clique em 'Prosseguir' para ir para a tela de confirmação")
    print(f"{'='*80}\n")
    
    try:
        root.mainloop()
        print("\n[TESTE] ✓ Janela fechou normalmente")
    except Exception as e:
        logger.error(f"❌ ERRO durante mainloop: {e}")
        import traceback
        traceback.print_exc()
    
    print(f"{'='*80}")
    print("[TESTE] CONCLUÍDO".center(80))
    print(f"{'='*80}\n")


if __name__ == "__main__":
    
    try:
        executar_teste()
    except Exception as e:
        print(f"\n❌ [ERROR] Erro ao executar teste: {e}")
        import traceback
        traceback.print_exc()
    finally:
        print()
        print("[INFO] Verifique logs/automacao_*.log para diagnóstico detalhado (se houver)")
        print("[INFO] Se a tela fecha sozinha, procure a mensagem de erro acima ⬆️")
