import pandas as pd
from difflib import SequenceMatcher
import tkinter as tk
from tkinter import ttk, messagebox
from typing import List, Dict
import logging
import sys
from pathlib import Path
from datetime import datetime
from ParametrosProcessamento import ParametrosProcessamento
from tela_checkin import ItemCheckin

# Configurar logging
Path("logs").mkdir(exist_ok=True)
logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)

console_handler = logging.StreamHandler(sys.stdout)
file_handler = logging.FileHandler(f"logs/automacao_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log")

formatter = logging.Formatter(
    '%(asctime)s | %(levelname)-8s | %(funcName)s():%(lineno)d | %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)

console_handler.setFormatter(formatter)
file_handler.setFormatter(formatter)

if not logger.handlers:
    logger.addHandler(console_handler)
    logger.addHandler(file_handler)

class ProcessamentoBase:
    """Classe base para processamento de dados"""
    def __init__(self, parametros, janela_progresso=None):
        self.parametros = parametros
        self.janela_progresso = janela_progresso

    @staticmethod
    def similaridade(a, b):
        return SequenceMatcher(None, a.lower(), b.lower()).ratio()

    @staticmethod
    def transformacaoIndiceColuna(coluna: str) -> int:
        coluna = coluna.upper()
        indice = 0
        for char in coluna:
            indice = indice * 26 + (ord(char) - ord('A') + 1)
        return indice - 1
    
    @staticmethod
    def validadorDeLinhasReferencia(linha: pd.Series, IndiceColunadescricao: int, IndiceColunamaterial: int, IndiceColuna_mao_de_obra: int) -> bool:
        descricaoValida = not pd.isna(linha.iloc[IndiceColunadescricao]) and str(linha.iloc[IndiceColunadescricao]).strip() != ""
        materialValido = not pd.isna(linha.iloc[IndiceColunamaterial]) and str(linha.iloc[IndiceColunamaterial]).strip() != ""
        maoDeObraValida = not pd.isna(linha.iloc[IndiceColuna_mao_de_obra]) and str(linha.iloc[IndiceColuna_mao_de_obra]).strip() != ""
        return descricaoValida and materialValido and maoDeObraValida
    
    @staticmethod
    def validadorDeLinhasOrcamento(linha: pd.Series, IndiceColuunadescricao: int, IndiceColuna_unidade_medida: int) -> bool:
        descricaoValida = not pd.isna(linha.iloc[IndiceColuunadescricao]) and str(linha.iloc[IndiceColuunadescricao]).strip() != ""
        unidadeMedidaValida = not pd.isna(linha.iloc[IndiceColuna_unidade_medida]) and str(linha.iloc[IndiceColuna_unidade_medida]).strip() != ""
        return descricaoValida and unidadeMedidaValida

    def processar_dados(self) -> Dict[str, List[Dict]]:
        """Processa os dados das planilhas e retorna os resultados"""
        try:
            colunaDescricaoOrcamento = self.parametros.orcamento.coluna_descrição
            indiceDaColunaOrcamento = self.transformacaoIndiceColuna(colunaDescricaoOrcamento)
            
            colunaDescricaoReferencia = self.parametros.referencia.coluna_descrição
            indiceDaColunaReferencia = self.transformacaoIndiceColuna(colunaDescricaoReferencia)

            colunaMaterialReferencia = self.parametros.referencia.coluna_material
            indiceDaColunaMaterialReferencia = self.transformacaoIndiceColuna(colunaMaterialReferencia)
            colunaMaoDeObraReferencia = self.parametros.referencia.coluna_mao_de_obra
            indiceDaColunaMaoDeObraReferencia = self.transformacaoIndiceColuna(colunaMaoDeObraReferencia)
            colunaUnidadeMedidaReferencia = self.parametros.referencia.coluna_unidade_medida
            indiceDaColunaUnidadeMedidaReferencia = self.transformacaoIndiceColuna(colunaUnidadeMedidaReferencia)

            linha_inicio = self.parametros.pesquisa.ComecoPesquisa
            linha_fim = self.parametros.pesquisa.TerminoPesquisa

            referencia = pd.read_excel(self.parametros.referencia.caminho_planilha, sheet_name=self.parametros.referencia.aba)
            orcamento = pd.read_excel(self.parametros.orcamento.caminho_planilha, sheet_name=self.parametros.orcamento.aba)

            # Arredonda os valores numéricos nas colunas de material e mão de obra (2 casas decimais)
            try:
                referencia.iloc[:, indiceDaColunaMaterialReferencia] = referencia.iloc[:, indiceDaColunaMaterialReferencia].round(2)
                referencia.iloc[:, indiceDaColunaMaoDeObraReferencia] = referencia.iloc[:, indiceDaColunaMaoDeObraReferencia].round(2)
            except Exception:
                # Se as colunas não forem numéricas, tentar converter antes de arredondar
                referencia.iloc[:, indiceDaColunaMaterialReferencia] = pd.to_numeric(referencia.iloc[:, indiceDaColunaMaterialReferencia], errors='coerce').round(2)
                referencia.iloc[:, indiceDaColunaMaoDeObraReferencia] = pd.to_numeric(referencia.iloc[:, indiceDaColunaMaoDeObraReferencia], errors='coerce').round(2)

            # Função que filtra as linhas do orçamento usando offset (compatibilidade com cabeçalhos)
            offsetPandas = int(2)
            linhasFiltradas = orcamento[
                (orcamento.index + offsetPandas >= linha_inicio) &
                (orcamento.index + offsetPandas <= linha_fim)
            ]
            
            # IMPORTANTE: Guardar índices originais ANTES de reset_index para calcular linha correta
            indices_originais = linhasFiltradas.index.tolist()
            orcamento = linhasFiltradas.reset_index(drop=True)

            resultados = []
            for idx_novo, row_orc in orcamento.iterrows():
                # Atualizar janela de progresso
                if self.janela_progresso:
                    self.janela_progresso.update()
                
                if not self.validadorDeLinhasOrcamento(row_orc, indiceDaColunaOrcamento, indiceDaColunaUnidadeMedidaReferencia):
                    continue
                descricao_orc = str(row_orc.iloc[indiceDaColunaOrcamento]).strip()  # Normalizar: remover espaços
                unidade_orc = str(row_orc.iloc[indiceDaColunaUnidadeMedidaReferencia]).strip()  # Normalizar também a unidade
                
                # Calcular número de linha na planilha usando índice original
                # indices_originais[idx_novo] é o índice no DataFrame original (antes do reset)
                # +2 porque Pandas pula header (linha 0 Excel = linha 1 lógica, linha 1 Excel = linha 2 lógica)
                numero_linha_planilha = indices_originais[idx_novo] + 2
                
                for idx_ref, row_ref in referencia.iterrows():
                    if self.validadorDeLinhasReferencia(
                        row_ref, 
                        indiceDaColunaReferencia,
                        indiceDaColunaMaterialReferencia,
                        indiceDaColunaMaoDeObraReferencia
                    ):
                        descricao_ref = str(row_ref.iloc[indiceDaColunaReferencia]).strip()  # Normalizar também
                        score = self.similaridade(descricao_orc, descricao_ref)
                        
                        if score > self.parametros.pesquisa.TaxaSimilaridade:
                            # TRATAMENTO DE ERRO: Valores podem estar como strings ou NaN
                            try:
                                valor_material = float(row_ref.iloc[indiceDaColunaMaterialReferencia])
                                valor_mao_de_obra = float(row_ref.iloc[indiceDaColunaMaoDeObraReferencia])
                            except (ValueError, TypeError):
                                # Se não conseguir converter, pula esta correlação
                                print(f"[WARNING] Não conseguiu converter valores numéricos para referência: {descricao_ref}")
                                continue
                            
                            # Validação: valores não podem ser NaN
                            if pd.isna(valor_material) or pd.isna(valor_mao_de_obra):
                                print(f"[WARNING] Valores NaN encontrados para referência: {descricao_ref}")
                                continue
                            
                            valor_total = valor_material + valor_mao_de_obra
                            
                            resultados.append({
                                "item": descricao_orc,
                                "numero_linha": numero_linha_planilha,
                                "unidade": unidade_orc,
                                "referencia": descricao_ref,
                                "similaridade": score * 100,
                                "valor_material": valor_material,
                                "valor_mao_de_obra": valor_mao_de_obra,
                                "valor_total": valor_total
                            })
            
            return resultados
        
        except Exception as e:
            raise e
    
    @staticmethod
    def agrupar_correlacoes_por_item(resultados: List[Dict]) -> Dict[str, List[Dict]]:
        """Agrupa correlações por item do orçamento (removendo duplicatas com normalização)"""
        grupos = {}
        for resultado in resultados:
            # Normalizar chave: remover espaços extras e converter para case padrão
            chave_grupo = resultado["item"].strip()  # Remove espaços no início/fim
            
            if chave_grupo not in grupos:
                grupos[chave_grupo] = []
            grupos[chave_grupo].append(resultado)
        return grupos
    
    @staticmethod
    def consolidar_referencias_por_similaridade(referencias: List[Dict]) -> tuple:
        """
        Consolida referências com mesma similaridade, removendo duplicatas visuais.
        Retorna: (referências_consolidadas, mapa_linhas)
        
        """
        # Dicionário para agrupar por (referencia, similaridade)
        grupos = {}
        mapa_linhas = {}
        
        for ref_data in referencias:
            chave = (ref_data["referencia"], round(ref_data["similaridade"], 1))
            
            if chave not in grupos:
                grupos[chave] = ref_data.copy()
                mapa_linhas[chave] = []
            
            # Adicionar número_linha ao mapa
            mapa_linhas[chave].append(ref_data["numero_linha"])
        
        # Construir lista consolidada com contagem de ocorrências
        consolidadas = []
        for (ref_name, similaridade), ref_data in grupos.items():
            ref_consolidada = ref_data.copy()
            ocorrencias = len(mapa_linhas[(ref_name, similaridade)])
            ref_consolidada["ocorrencias"] = ocorrencias
            consolidadas.append(ref_consolidada)
        
        return consolidadas, mapa_linhas

def criar_interface(root, parametros: ParametrosProcessamento):
    root.title("Processamento - Automação de Orçamento")
    root.geometry("1150x700")
    root.configure(bg="#f5f5f5")

    selecoes = {}
    mapa_consolidado = {}  # Mapeia (item_nome, referencia) → lista de números_linha
    dados_orcamento = []
    dados_agrupados = {}
    processador = ProcessamentoBase(parametros, janela_progresso=None)  # Criar instância para usar métodos estáticos

    try:
        if parametros is None:
            raise ValueError("Parâmetros não foram fornecidos")
        if parametros.referencia is None:
            raise ValueError("Parâmetros de referência não foram fornecidos")
        if parametros.orcamento is None:
            raise ValueError("Parâmetros de orçamento não foram fornecidos")
        if parametros.pesquisa is None:
            raise ValueError("Parâmetros de pesquisa não foram fornecidos")

        logger.debug("1. Iniciando processamento...")
        
        # Criar janela de notificação (NÃO-MODAL)
        janela_prog = tk.Toplevel(root)
        janela_prog.title("Carregando Dados")
        janela_prog.geometry("300x100")
        janela_prog.resizable(False, False)
        
        # Centralizar na tela
        janela_prog.transient(root)
        janela_prog.update_idletasks()
        x = root.winfo_x() + (root.winfo_width() - janela_prog.winfo_width()) // 2
        y = root.winfo_y() + (root.winfo_height() - janela_prog.winfo_height()) // 2
        janela_prog.geometry(f"+{x}+{y}")
        
        # Adicionar label e progressbar
        tk.Label(janela_prog, text="Processando…", font=("Segoe UI", 11, "bold")).pack(pady=10)
        progress = ttk.Progressbar(janela_prog, mode="indeterminate", length=250)
        progress.pack(pady=5, padx=20)
        progress.start()
        
        # Se fechar a janela de progresso, fechar a tela inteira
        def ao_fechar_janela():
            logger.warning("Processamento interrompido pelo usuário")
            root.destroy()
        
        janela_prog.protocol("WM_DELETE_WINDOW", ao_fechar_janela)
        
        # Processar mantendo responsividade visual
        processador = ProcessamentoBase(parametros, janela_progresso=janela_prog)  # Atualizar com janela_prog
        dados = processador.processar_dados()
        
        # Fechar janela de progresso (se ainda estiver aberta)
        try:
            progress.stop()
            janela_prog.destroy()
        except:
            pass  # Janela já foi fechada pelo usuário
        
        logger.debug(f"2. Dados retornados: {len(dados)} itens")
        logger.debug(f"3. Amostra de dados: {dados[:2] if dados else 'VAZIO'}")
        
        dados_agrupados = processador.agrupar_correlacoes_por_item(dados)
        logger.debug(f"4. Dados agrupados: {len(dados_agrupados)} grupos")
        logger.debug(f"5. Chaves dos grupos: {list(dados_agrupados.keys())}")
        
        if not dados_agrupados:
            logger.error("dados_agrupados está vazio!")
            messagebox.showwarning(
                "Nenhum Resultado",
                "O processamento não encontrou correlações.\n\n"
                "Possíveis causas:\n"
                "1. As colunas especificadas não correspondem aos dados\n"
                "2. A taxa de similaridade é muito alta\n"
                "3. Os dados das planilhas estão vazios"
            )
            return

        # Preparar dados de orçamento para o grid superior
        logger.debug("6. Preparando dados de orçamento...")
        itens_inseridos = set()  # Rastreador de itens já inseridos (evitar duplicatas)
        for item, opcoes in dados_agrupados.items():
            logger.debug(f"   - Item: {item}, Opções: {len(opcoes)}")
            if opcoes and item not in itens_inseridos:  # Verificar se já não foi inserido
                dados_orcamento.append({
                    "item": item,
                    "unidade": opcoes[0].get("unidade", ""),
                    "qtd": len(opcoes)
                })
                itens_inseridos.add(item)  # Marcar como inserido
        
        logger.debug(f"7. Dados de orçamento preparados: {len(dados_orcamento)} itens")
        logger.debug(f"8. Amostra: {dados_orcamento[:2] if dados_orcamento else 'VAZIO'}")

    except Exception as e:
        logger.error(f"Erro ao processar: {e}")
        import traceback
        erro_completo = traceback.format_exc()
        logger.error(erro_completo)
        messagebox.showerror(
            "Erro no Processamento",
            f"Erro ao processar os dados:\n\n{str(e)}\n\n"
            f"Detalhes:\n{erro_completo}\n\n"
            "Verifique o console para mais informações."
        )
        return

    # ===== FRAME TOPO (Cabeçalho Azul) =====
    frame_topo = tk.Frame(root, bg="#004c99", height=80)
    frame_topo.grid(row=0, column=0, sticky="nsew")
    frame_topo.grid_columnconfigure(0, weight=1)

    lbl_titulo = tk.Label(
        frame_topo,
        text="🔵 Processamento - Automação de Orçamento",
        font=("Segoe UI", 14, "bold"),
        bg="#004c99",
        fg="white"
    )
    lbl_titulo.pack(anchor="w", padx=15, pady=(10, 5))

    frame_info = tk.Frame(frame_topo, bg="#004c99")
    frame_info.pack(anchor="w", padx=15, pady=(0, 10))

    tk.Label(
        frame_info,
        text=f"Planilha de Referência: {parametros.referencia.caminho_planilha}",
        bg="#004c99",
        fg="white",
        font=("Segoe UI", 9)
    ).pack(anchor="w")
    
    tk.Label(
        frame_info,
        text=f"Planilha de Orçamento: {parametros.orcamento.caminho_planilha}",
        bg="#004c99",
        fg="white",
        font=("Segoe UI", 9)
    ).pack(anchor="w")

    # ===== CONFIGURAR LAYOUT PRINCIPAL COM GRID =====
    root.grid_rowconfigure(0, weight=0)  # Cabeçalho
    root.grid_rowconfigure(1, weight=0)  # Pesquisa
    root.grid_rowconfigure(2, weight=1)  # Conteúdo (grids)
    root.grid_rowconfigure(3, weight=0)  # Botões
    root.grid_columnconfigure(0, weight=1)

    # ===== FRAME PESQUISA =====
    frame_pesquisa = tk.Frame(root, bg="white", padx=15, pady=10)
    frame_pesquisa.grid(row=1, column=0, sticky="nsew")

    tk.Label(
        frame_pesquisa,
        text="Filtrar itens do orçamento:",
        font=("Segoe UI", 9),
        bg="white"
    ).pack(anchor="w")

    entrada_pesquisa = tk.Entry(frame_pesquisa, font=("Segoe UI", 10), width=50)
    entrada_pesquisa.pack(anchor="w", pady=(5, 0))

    # ===== FRAME CORPO (Contém as 2 grids com layout grid) =====
    frame_corpo = tk.Frame(root, bg="#f5f5f5")
    frame_corpo.grid(row=2, column=0, sticky="nsew", padx=10, pady=10)
    
    # Configurar as linhas: 40% grid superior, 60% grid inferior
    frame_corpo.grid_rowconfigure(0, weight=0)  # Label superior
    frame_corpo.grid_rowconfigure(1, weight=40) # Grid superior
    frame_corpo.grid_rowconfigure(2, weight=0)  # Label inferior
    frame_corpo.grid_rowconfigure(3, weight=60) # Grid inferior
    frame_corpo.grid_columnconfigure(0, weight=1)

    # ===== GRID SUPERIOR (Orçamento) =====
    lbl_superior = tk.Label(
        frame_corpo,
        text="Itens do Orçamento",
        font=("Segoe UI", 10, "bold"),
        bg="#f5f5f5"
    )
    lbl_superior.grid(row=0, column=0, sticky="w", padx=10, pady=(10, 5))

    frame_grid_superior = tk.Frame(frame_corpo, bg="white")
    frame_grid_superior.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 15))
    frame_grid_superior.grid_columnconfigure(0, weight=1)
    frame_grid_superior.grid_rowconfigure(0, weight=1)

    tree_orc = ttk.Treeview(
        frame_grid_superior,
        columns=("item", "unidade", "qty"),
        show="headings",
        height=6
    )
    tree_orc.heading("#0", text="")
    tree_orc.heading("item", text="Item do Orçamento")
    tree_orc.heading("unidade", text="Unidade")
    tree_orc.heading("qty", text="Ocorrências")

    tree_orc.column("#0", width=0, stretch=False)
    tree_orc.column("item", width=500, anchor="w")  # Reduzido de 700 para deixar espaço
    tree_orc.column("unidade", width=150, anchor="center")  # Reduzido de 200
    tree_orc.column("qty", width=100, anchor="center")

    scrolly_sup = ttk.Scrollbar(frame_grid_superior, orient="vertical", command=tree_orc.yview)
    scrollx_sup = ttk.Scrollbar(frame_grid_superior, orient="horizontal", command=tree_orc.xview)
    tree_orc.configure(yscrollcommand=scrolly_sup.set, xscrollcommand=scrollx_sup.set)

    tree_orc.grid(row=0, column=0, sticky="nsew")
    scrolly_sup.grid(row=0, column=1, sticky="ns")
    scrollx_sup.grid(row=1, column=0, sticky="ew")

    # ===== SISTEMA DE ORDENAÇÃO PARA GRID SUPERIOR =====
    estado_ordenacao = {
        "coluna_ativa": None,  # Qual coluna está sendo ordenada
        "direcao": "asc"       # "asc" para A→Z/crescente, "desc" para Z→A/decrescente
    }
    
    def ordenar_grid_superior(coluna: str):
        """Ordena os dados em dados_orcamento e reinsere na grid superior"""
        nonlocal dados_orcamento, estado_ordenacao
        
        logger.debug(f"Ordenação iniciada | coluna={coluna}, estado_anterior={estado_ordenacao}")
        
        try:
            # Se clicou na mesma coluna, inverter direção
            if estado_ordenacao["coluna_ativa"] == coluna:
                estado_ordenacao["direcao"] = "desc" if estado_ordenacao["direcao"] == "asc" else "asc"
            else:
                # Se clicou em coluna diferente, começar com ascendente
                estado_ordenacao["coluna_ativa"] = coluna
                estado_ordenacao["direcao"] = "asc"
            
            logger.debug(f"Novo estado_ordenacao: {estado_ordenacao}")
            
            # Mapear coluna para índice e função de sort
            if coluna == "item":
                # Ordenação alfabética
                reverse = (estado_ordenacao["direcao"] == "desc")
                dados_orcamento = sorted(dados_orcamento, key=lambda x: x["item"].lower(), reverse=reverse)
                logger.debug(f"Ordenado por ITEM | direção={estado_ordenacao['direcao']}")
                
            elif coluna == "unidade":
                # Ordenação por unidade (alfabética)
                reverse = (estado_ordenacao["direcao"] == "desc")
                dados_orcamento = sorted(dados_orcamento, key=lambda x: x["unidade"].lower(), reverse=reverse)
                logger.debug(f"Ordenado por UNIDADE | direção={estado_ordenacao['direcao']}")
                
            elif coluna == "qty":
                # Ordenação numérica
                reverse = (estado_ordenacao["direcao"] == "desc")
                dados_orcamento = sorted(dados_orcamento, key=lambda x: x["qtd"], reverse=reverse)
                logger.debug(f"Ordenado por OCORRÊNCIAS | direção={estado_ordenacao['direcao']}")
            
            # Limpar grid superior
            for item in tree_orc.get_children():
                tree_orc.delete(item)
            
            # Reinserir dados ordenados
            for item_data in dados_orcamento:
                tree_orc.insert("", "end", values=(
                    item_data["item"],
                    item_data["unidade"],
                    item_data["qtd"]
                ))
            
            logger.debug(f"Grid reinserida com {len(dados_orcamento)} itens")
            
        except Exception as e:
            logger.error(f"Erro ao ordenar grid: {e}")
            import traceback
            traceback.print_exc()
    
    # Adicionar comando de ordenação aos headings
    tree_orc.heading("item", command=lambda: ordenar_grid_superior("item"))
    tree_orc.heading("unidade", command=lambda: ordenar_grid_superior("unidade"))
    tree_orc.heading("qty", command=lambda: ordenar_grid_superior("qty"))

    # Inserir dados no grid superior
    logger.debug(f"9. Inserindo {len(dados_orcamento)} itens no grid superior...")
    for i, item_data in enumerate(dados_orcamento):
        logger.debug(f"   {i+1}. Inserindo: {item_data}")
        tree_orc.insert("", "end", values=(
            item_data["item"],
            item_data["unidade"],
            item_data["qtd"]
        ))
    logger.debug("10. Inserção concluída!")

    item_selecionado = {"id": None, "nome": None}
    itens_respondidos = {}  # Rastreador de respostas do usuário: {item_nome: True/False}
    debounce_timer_item = None  # Timer para evitar múltiplos disparos de seleção de item

    def on_item_selecionado(event=None):
        """Ao selecionar item na grid superior, atualiza grid inferior"""
        nonlocal debounce_timer_item, selecoes
        
        # Se há um timer pendente, cancela (novo clique chegou)
        if debounce_timer_item is not None:
            root.after_cancel(debounce_timer_item)
        
        # Agenda o processamento com delay de 100ms
        def processar_item():
            nonlocal debounce_timer_item, selecoes
            debounce_timer_item = None
            
            logger.debug("on_item_selecionado disparado!")
            try:
                selection = tree_orc.selection()
                logger.debug(f"Selection: {selection}")
                if selection:
                    item_id = selection[0]
                    valores = tree_orc.item(item_id)["values"]
                    logger.debug(f"Valores extraídos: {valores}")
                    item_selecionado["id"] = item_id
                    item_selecionado["nome"] = valores[0]
                    logger.debug(f"Item selecionado atualizado: {item_selecionado['nome']}")
                    
                    # Chamar processador com estado explícito e receber novo estado
                    selecoes = atualizar_grid_inferior(
                        valores[0],
                        selecoes,
                        dados_agrupados,
                        tree_ref,
                        itens_respondidos,
                        mapa_consolidado
                    )
                    logger.debug(f"Grid inferior atualizado | novo_estado_selecoes={selecoes}")
            except Exception as e:
                logger.error(f"Erro em on_item_selecionado: {e}")
                import traceback
                traceback.print_exc()
        
        debounce_timer_item = root.after(100, processar_item)  # Aguarda 100ms antes de processar

    tree_orc.bind("<<TreeviewSelect>>", on_item_selecionado)

    # ===== GRID INFERIOR (Opções de Correlação) =====
    lbl_inferior = tk.Label(
        frame_corpo,
        text="Referências Correlacionadas",
        font=("Segoe UI", 10, "bold"),
        bg="#f5f5f5"
    )
    lbl_inferior.grid(row=2, column=0, sticky="w", padx=10, pady=(10, 5))

    frame_grid_inferior = tk.Frame(frame_corpo, bg="white")
    frame_grid_inferior.grid(row=3, column=0, sticky="nsew", padx=10, pady=(0, 10))
    frame_grid_inferior.grid_columnconfigure(0, weight=1)
    frame_grid_inferior.grid_rowconfigure(0, weight=1)

    tree_ref = ttk.Treeview(
        frame_grid_inferior,
        columns=("selecao", "referencia", "unidade", "similaridade", "valor", "ocorrencias"),
        show="headings",
        height=8
    )
    tree_ref.heading("selecao", text="✓")
    tree_ref.heading("referencia", text="Referência Sugerida")
    tree_ref.heading("unidade", text="Unidade")
    tree_ref.heading("similaridade", text="Similaridade")
    tree_ref.heading("valor", text="Valor")
    tree_ref.heading("ocorrencias", text="Ocorrências")

    tree_ref.column("selecao", width=40, anchor="center")
    tree_ref.column("referencia", width=350, anchor="w")  # Reduzido para acomodar nova coluna
    tree_ref.column("unidade", width=80, anchor="center")
    tree_ref.column("similaridade", width=100, anchor="center")
    tree_ref.column("valor", width=120, anchor="e")
    tree_ref.column("ocorrencias", width=80, anchor="center")  # Nova coluna

    scrolly_inf = ttk.Scrollbar(frame_grid_inferior, orient="vertical", command=tree_ref.yview)
    scrollx_inf = ttk.Scrollbar(frame_grid_inferior, orient="horizontal", command=tree_ref.xview)
    tree_ref.configure(yscrollcommand=scrolly_inf.set, xscrollcommand=scrollx_inf.set)

    tree_ref.grid(row=0, column=0, sticky="nsew")
    scrolly_inf.grid(row=0, column=1, sticky="ns")
    scrollx_inf.grid(row=1, column=0, sticky="ew")

    # Rastreador de seleção: {item_orcamento: id_tree}
    debounce_timer = None  # Timer para evitar múltiplos disparos rápidos

    def processar_selecao_referencia(
        item_orcamento: str,
        selecoes: dict,
        tree_ref: ttk.Treeview
    ) -> dict:
        """
        Marca uma referência como selecionada na grid e atualiza estado.
        
        Args:
            item_orcamento: Nome do item de orçamento selecionado
            selecoes: Estado atual de seleções
            tree_ref: Widget Treeview para atualizar UI
            
        Returns:
            dict: Estado atualizado de seleções
        """
        selecoes_atualizadas = selecoes.copy()
        
        logger.debug(f"processar_selecao_referencia iniciado | item_orcamento={item_orcamento}")
        
        try:
            # Limpar grid inferior
            logger.debug("Limpando grid inferior...")
            for item in tree_ref.get_children():
                tree_ref.delete(item)

            # Verificar se já existe seleção anterior para este item
            referencia_anterior = selecoes_atualizadas.get(item_orcamento, None)
            logger.debug(f"Seleção anterior para '{item_orcamento}': {referencia_anterior}")
            
            # Adicionar opção "Nenhum" como primeiro item
            logger.debug("Inserindo opção 'NENHUM'...")
            referencia_anterior = selecoes_atualizadas.get(item_orcamento, "NENHUM")
            marcado = "◉" if referencia_anterior == "NENHUM" else "○"
            tree_ref.insert("", "end", values=(
                marcado,
                "NENHUM",
                "-",
                "-",
                "-",
                "-"
            ))
                        
            logger.debug("Verificando se item_nome está em dados_agrupados...")
            logger.debug(f"Chaves disponíveis: {list(dados_agrupados.keys())[:5]}...")
            logger.debug(f"Item procurado: '{item_orcamento}'")
            
            # Preencher com referências do item selecionado
            if item_orcamento in dados_agrupados:
                refs_originais = dados_agrupados[item_orcamento]
                logger.debug(f"Item ENCONTRADO! Referências originais: {len(refs_originais)}")
                
                # Consolidar referências por (nome + similaridade)
                refs_consolidadas, mapa_linhas = processador.consolidar_referencias_por_similaridade(refs_originais)
                logger.debug(f"Referências consolidadas: {len(refs_consolidadas)}")
                
                # Atualizar mapa global de consolidação
                mapa_consolidado[item_orcamento] = mapa_linhas
                
                # Verificar se tem 100% de similaridade
                ref_100_porcento = None
                for ref_data in refs_consolidadas:
                    if ref_data["similaridade"] >= 99.95:
                        ref_100_porcento = ref_data
                        break
                
                # Se encontrou 100%, perguntar ao usuário (se ainda não respondeu)
                if ref_100_porcento and item_orcamento not in itens_respondidos:
                    resposta = messagebox.askyesno(
                        "Correspondência Perfeita",
                        f"Foi encontrada uma correspondência com 100% de similaridade:\n\n"
                        f"Referência: {ref_100_porcento['referencia']}\n"
                        f"Valor: R$ {ref_100_porcento['valor_total']:.2f}\n"
                        f"Ocorrências: {ref_100_porcento['ocorrencias']}\n\n"
                        f"Deseja usar automaticamente?"
                    )
                    
                    # Registrar a resposta
                    itens_respondidos[item_orcamento] = resposta
                    logger.debug(f"Usuário respondeu: {resposta} para item '{item_orcamento}'")
                
                # Se já respondeu SIM antes, aplicar automaticamente
                if ref_100_porcento and itens_respondidos.get(item_orcamento) == True:
                    logger.debug(f"Aplicando automaticamente 100% para '{item_orcamento}'")
                    selecoes_atualizadas[item_orcamento] = ref_100_porcento["referencia"]
                    tree_ref.insert("", "end", values=(
                        "◉",
                        ref_100_porcento["referencia"],
                        ref_100_porcento["unidade"],
                        f"{ref_100_porcento['similaridade']:.1f}%",
                        f"R$ {ref_100_porcento['valor_total']:.2f}",
                        str(ref_100_porcento['ocorrencias'])
                    ))
                    logger.debug(f"100% auto-selecionado e estado atualizado | retornando novo_estado={selecoes_atualizadas}")
                    return selecoes_atualizadas
                
                # Inserir todas as referências consolidadas
                for ref_data in refs_consolidadas:
                    # Verificar se esta referência é a que estava selecionada
                    marcado = "◉" if ref_data["referencia"] == referencia_anterior else "○"
                    tree_ref.insert("", "end", values=(
                        marcado,
                        ref_data["referencia"],
                        ref_data["unidade"],
                        f"{ref_data['similaridade']:.1f}%",
                        f"R$ {ref_data['valor_total']:.2f}",
                        str(ref_data['ocorrencias'])
                    ))
                    logger.debug(f"Inserindo consolidada: {ref_data['referencia']} ({ref_data['ocorrencias']} ocorrências)")
                logger.debug("Inserção de referências consolidadas concluída!")
            else:
                logger.warning(f"Item NÃO ENCONTRADO em dados_agrupados!")
            
            logger.debug(f"Retornando novo_estado | item={item_orcamento} | selecoes_atualizadas={selecoes_atualizadas}")
            return selecoes_atualizadas
                
        except Exception as e:
            logger.error(f"Erro em atualizar_grid_inferior: {e}")
            import traceback
            traceback.print_exc()
            return selecoes_atualizadas  # Retorna estado mesmo em erro

    def on_referencia_selecionada(event=None):
        """Callback ao clicar em uma referência - thin wrapper"""
        nonlocal debounce_timer, selecoes
        
        # Se há um timer pendente, cancela (novo clique chegou)
        if debounce_timer is not None:
            root.after_cancel(debounce_timer)
        
        # Agenda o processamento com delay de 100ms
        # Isso garante que só o ÚLTIMO clique será processado
        def executar():
            nonlocal debounce_timer, selecoes
            debounce_timer = None
            
            logger.debug("on_referencia_selecionada disparado!")
            
            if item_selecionado["nome"]:
                # Chamar processador com estado explícito e receber novo estado
                selecoes = processar_selecao_referencia(
                    item_selecionado["nome"],
                    selecoes,
                    tree_ref
                )
        
        debounce_timer = root.after(100, executar)  # Aguarda 100ms antes de processar

    tree_ref.bind("<<TreeviewSelect>>", on_referencia_selecionada)

    def atualizar_grid_inferior(
        item_nome: str,
        selecoes: dict,
        dados_agrupados: dict,
        tree_ref: ttk.Treeview,
        itens_respondidos: dict,
        mapa_consolidado: dict
    ) -> dict:
        """
        Atualiza grid inferior com referências do item selecionado (consolidadas).
        
        Args:
            item_nome: Nome do item selecionado
            selecoes: Estado atual de seleções
            dados_agrupados: Dados agrupados por item
            tree_ref: Widget Treeview para atualizar
            itens_respondidos: Respostas do usuário a 100% similaridade
            mapa_consolidado: Mapa de (item_nome, referencia) → lista de números_linha
            
        Returns:
            dict: Estado atualizado de seleções
        """
        logger.debug(f"atualizar_grid_inferior chamado com: {item_nome}")
        
        # Copiar estado (imutabilidade)
        selecoes_atualizadas = selecoes.copy()
        
        try:
            # Limpar grid inferior
            logger.debug("Limpando grid inferior...")
            for item in tree_ref.get_children():
                tree_ref.delete(item)

            # Verificar se já existe seleção anterior para este item
            referencia_anterior = selecoes_atualizadas.get(item_nome, None)
            logger.debug(f"Seleção anterior para '{item_nome}': {referencia_anterior}")
            
            # Adicionar opção "Nenhum" como primeiro item
            logger.debug("Inserindo opção 'NENHUM'...")
            referencia_anterior = selecoes_atualizadas.get(item_nome, "NENHUM")
            marcado = "◉" if referencia_anterior == "NENHUM" else "○"
            tree_ref.insert("", "end", values=(
                marcado,
                "NENHUM",
                "-",
                "-",
                "-",
                "-"
            ))
                        
            logger.debug("Verificando se item_nome está em dados_agrupados...")
            logger.debug(f"Chaves disponíveis: {list(dados_agrupados.keys())[:5]}...")
            logger.debug(f"Item procurado: '{item_nome}'")
            
            # Preencher com referências do item selecionado
            if item_nome in dados_agrupados:
                refs_originais = dados_agrupados[item_nome]
                logger.debug(f"Item ENCONTRADO! Referências originais: {len(refs_originais)}")
                
                # Consolidar referências por (nome + similaridade)
                refs_consolidadas, mapa_linhas = processador.consolidar_referencias_por_similaridade(refs_originais)
                logger.debug(f"Referências consolidadas: {len(refs_consolidadas)}")
                
                # Atualizar mapa global de consolidação
                mapa_consolidado[item_nome] = mapa_linhas
                
                # Verificar se tem 100% de similaridade
                ref_100_porcento = None
                for ref_data in refs_consolidadas:
                    if ref_data["similaridade"] >= 99.95:
                        ref_100_porcento = ref_data
                        break
                
                # Se encontrou 100%, perguntar ao usuário (se ainda não respondeu)
                if ref_100_porcento and item_nome not in itens_respondidos:
                    resposta = messagebox.askyesno(
                        "Correspondência Perfeita",
                        f"Foi encontrada uma correspondência com 100% de similaridade:\n\n"
                        f"Referência: {ref_100_porcento['referencia']}\n"
                        f"Valor: R$ {ref_100_porcento['valor_total']:.2f}\n"
                        f"Ocorrências: {ref_100_porcento['ocorrencias']}\n\n"
                        f"Deseja usar automaticamente?"
                    )
                    
                    # Registrar a resposta
                    itens_respondidos[item_nome] = resposta
                    logger.debug(f"Usuário respondeu: {resposta} para item '{item_nome}'")
                
                # Se já respondeu SIM antes, aplicar automaticamente
                if ref_100_porcento and itens_respondidos.get(item_nome) == True:
                    logger.debug(f"Aplicando automaticamente 100% para '{item_orcamento}'")
                    selecoes_atualizadas[item_nome] = ref_100_porcento["referencia"]
                    tree_ref.insert("", "end", values=(
                        "◉",
                        ref_100_porcento["referencia"],
                        ref_100_porcento["unidade"],
                        f"{ref_100_porcento['similaridade']:.1f}%",
                        f"R$ {ref_100_porcento['valor_total']:.2f}",
                        str(ref_100_porcento['ocorrencias'])
                    ))
                    logger.debug(f"100% auto-selecionado e estado atualizado | retornando novo_estado={selecoes_atualizadas}")
                    return selecoes_atualizadas
                
                # Inserir todas as referências consolidadas
                for ref_data in refs_consolidadas:
                    # Verificar se esta referência é a que estava selecionada
                    marcado = "◉" if ref_data["referencia"] == referencia_anterior else "○"
                    tree_ref.insert("", "end", values=(
                        marcado,
                        ref_data["referencia"],
                        ref_data["unidade"],
                        f"{ref_data['similaridade']:.1f}%",
                        f"R$ {ref_data['valor_total']:.2f}",
                        str(ref_data['ocorrencias'])
                    ))
                    logger.debug(f"Inserindo consolidada: {ref_data['referencia']} ({ref_data['ocorrencias']} ocorrências)")
                logger.debug("Inserção de referências consolidadas concluída!")
            else:
                logger.warning(f"Item NÃO ENCONTRADO em dados_agrupados!")
            
            logger.debug(f"Retornando novo_estado | item={item_nome} | selecoes_atualizadas={selecoes_atualizadas}")
            return selecoes_atualizadas
                
        except Exception as e:
            logger.error(f"Erro em atualizar_grid_inferior: {e}")
            import traceback
            traceback.print_exc()
            return selecoes_atualizadas  # Retorna estado mesmo em erro

    def filtrar_itens(event=None):
        texto = entrada_pesquisa.get().lower()
        
        # Limpar grid superior
        for item in tree_orc.get_children():
            tree_orc.delete(item)

        # Reinsert apenas itens que correspondem ao filtro
        for item_data in dados_orcamento:
            if texto in item_data["item"].lower():
                tree_orc.insert("", "end", values=(
                    item_data["item"],
                    item_data["unidade"],
                    item_data["qtd"]
                ))

    entrada_pesquisa.bind('<KeyRelease>', filtrar_itens)

    # ===== SELECIONAR PRIMEIRO ITEM AUTOMATICAMENTE =====
    if dados_orcamento:
        tree_orc.selection_set(tree_orc.get_children()[0])
        # ✅ CORREÇÃO: Passar todos os 6 parâmetros e capturar retorno
        selecoes = atualizar_grid_inferior(
            dados_orcamento[0]["item"],
            selecoes,
            dados_agrupados,
            tree_ref,
            itens_respondidos,
            mapa_consolidado
        )
        item_selecionado["nome"] = dados_orcamento[0]["item"]

    # ===== BOTÕES INFERIORES =====
    frame_btns = tk.Frame(root, bg="white", padx=15, pady=10)
    frame_btns.grid(row=3, column=0, sticky="nsew")
    frame_btns.grid_columnconfigure(0, weight=1)

    def finalizar():
        if messagebox.askyesno("Finalizar", "Deseja finalizar o processamento?"):
            root.destroy()

    def prosseguir():
        # 1. Validar se existe pelo menos uma seleção válida
        logger.debug(f"Estado de selecoes em prosseguir(): {selecoes}")
        if not selecoes:
            messagebox.showwarning(
                "Aviso",
                "Nenhum item foi processado."
            )
            return

        # 2. Montar lista final para a tela de check-in - ENRIQUECER COM DADOS (CONSOLIDADO)
        itens_confirmados = []

        for item_orcamento, referencia_selecionada in selecoes.items():
            logger.debug(f"Processando: item_orcamento={item_orcamento}, referencia={referencia_selecionada}")
            
            # Se a referência for "NENHUM", pular para a próxima
            if referencia_selecionada == "NENHUM":
                logger.debug(f"Item com 'NENHUM' detectado - pulando")
                continue
            
            # Buscar os dados completos da referência em dados_agrupados
            if item_orcamento in dados_agrupados:
                # Encontrar todas as ocorrências com essa referência
                numeros_linha_consolidados = []
                primeiro_item = None
                
                for opcao in dados_agrupados[item_orcamento]:
                    if opcao["referencia"] == referencia_selecionada:
                        numeros_linha_consolidados.append(opcao["numero_linha"])
                        # Guardar primeiro item para usar seus dados
                        if primeiro_item is None:
                            primeiro_item = opcao
                
                if primeiro_item is not None:
                    logger.debug(f"Referência ENCONTRADA! {len(numeros_linha_consolidados)} ocorrência(s)")
                    # Criar um único item para o check-in (consolidado)
                    item_checkin = ItemCheckin(
                        item=item_orcamento,
                        unidade=primeiro_item["unidade"],
                        referencia=referencia_selecionada,
                        similaridade=primeiro_item["similaridade"],
                        valor_total=primeiro_item["valor_total"],
                        numero_linha=primeiro_item["numero_linha"],  # Primeiro número para compatibilidade
                        valor_material=primeiro_item["valor_material"],
                        valor_mao_de_obra=primeiro_item["valor_mao_de_obra"]
                    )
                    # Adicionar campo customizado para rastreamento de todas as ocorrências
                    item_checkin.numeros_linha_todas_ocorrencias = numeros_linha_consolidados
                    item_checkin.ocorrencias = len(numeros_linha_consolidados)
                    
                    itens_confirmados.append(item_checkin)
                    logger.debug(f"Item consolidado adicionado | {len(numeros_linha_consolidados)} ocorrências")
                else:
                    logger.warning(f"Referência '{referencia_selecionada}' não encontrada em dados_agrupados")
            else:
                logger.warning(f"Item '{item_orcamento}' não encontrado em dados_agrupados")

        logger.debug(f"itens_confirmados para checkin: {itens_confirmados}")
        
        if not itens_confirmados:
            messagebox.showwarning(
                "Aviso",
                "Nenhuma referência válida foi selecionada.\n(Todas estão marcadas como 'NENHUM')"
            )
            return
        
        # 3. Função callback para quando item for excluído no checkin
        def processar_exclusao_item_checkin(
            numero_linha: int,
            selecoes: dict,
            dados_agrupados: dict,
            item_selecionado: dict,
            tree_ref: ttk.Treeview
        ) -> dict:
            """
            Processa exclusão de item no checkin.
            
            Args:
                numero_linha: ID único da linha excluída
                selecoes: Estado atual de seleções
                dados_agrupados: Dados agrupados por item
                item_selecionado: Item atualmente selecionado na UI
                tree_ref: Widget Treeview da grid inferior
                
            Returns:
                dict: Estado atualizado de seleções
            """
            selecoes_atualizadas = selecoes.copy()
            
            logger.debug(f"processar_exclusao_item_checkin | numero_linha={numero_linha}, estado_entrada={selecoes_atualizadas}")
            
            try:
                # Encontrar o item correspondente pelo numero_linha (ID único)
                item_para_remover = None
                for item_orc, refs in dados_agrupados.items():
                    for ref in refs:
                        if ref["numero_linha"] == numero_linha:
                            item_para_remover = item_orc
                            break
                    if item_para_remover:
                        break
                
                if item_para_remover:
                    logger.debug(f"Item encontrado para remoção: {item_para_remover}")
                    
                    # Resetar seleção do item para "NENHUM"
                    if item_para_remover in selecoes_atualizadas:
                        selecoes_atualizadas[item_para_remover] = "NENHUM"
                        logger.debug(f"Seleção resetada para NENHUM: {item_para_remover}")
                    
                    # Resetar visual na grid inferior (se o item está selecionado visualmente)
                    if item_selecionado["nome"] == item_para_remover:
                        logger.debug(f"Item excluído estava selecionado, atualizando grid inferior")
                        selecoes_atualizadas = atualizar_grid_inferior(
                            item_para_remover,
                            selecoes_atualizadas,
                            dados_agrupados,
                            tree_ref,
                            itens_respondidos,
                            mapa_consolidado
                        )
                else:
                    logger.warning(f"Item com linha {numero_linha} não encontrado em dados_agrupados")
            
            except Exception as e:
                logger.error(f"Erro em processar_exclusao_item_checkin: {e}")
                import traceback
                traceback.print_exc()
            
            logger.debug(f"Retornando novo_estado_selecoes | {selecoes_atualizadas}")
            return selecoes_atualizadas

        def ao_excluir_item_checkin(numero_linha):
            """Callback executada quando item é excluído no checkin - thin wrapper"""
            nonlocal selecoes
            logger.debug(f"Callback ao_excluir_item_checkin | linha {numero_linha}")
            
            # Chamar processador com estado explícito e receber novo estado
            selecoes = processar_exclusao_item_checkin(
                numero_linha,
                selecoes,
                dados_agrupados,
                item_selecionado,
                tree_ref
            )
        
        # 3. Preparar parâmetros da planilha para atualização
        parametros_planilha = {
            "caminho_orcamento": parametros.orcamento.caminho_planilha,
            "aba_orcamento": parametros.orcamento.aba,
            "coluna_material": parametros.orcamento.coluna_material,
            "coluna_mao_de_obra": parametros.orcamento.coluna_mao_de_obra,
            "coluna_unidade_medida": parametros.orcamento.coluna_unidade_medida
        }
        
        # 4. Abrir tela de check-in COM callback de exclusão E parâmetros de planilha
        from tela_checkin import TelaCheckin
        logger.debug(f"Abrindo TelaCheckin | {len(itens_confirmados)} itens para confirmar")
        
        # IMPORTANTE: Criar instância E chamar .run() para bloquear até usuário responder
        tela_checkin = TelaCheckin(
            root, 
            itens_confirmados, 
            callback_excluir=ao_excluir_item_checkin, 
            parametros_planilha=parametros_planilha
        )
        logger.debug(f"TelaCheckin criada | aguardando resposta do usuário...")
        confirmado = tela_checkin.run()  # ← CRUCIAL: Bloqueia até usuário confirmar/cancelar
        logger.debug(f"TelaCheckin respondida | confirmado={confirmado}")


    btn_finalizar = ttk.Button(frame_btns, text="Finalizar", command=finalizar)
    btn_prosseguir = ttk.Button(frame_btns, text="Prosseguir", command=prosseguir)

    btn_finalizar.pack(side="left", padx=5)
    btn_prosseguir.pack(side="right", padx=5)

class TelaProcessamento:
    def __init__(self, root, parametros: ParametrosProcessamento):
        criar_interface(root, parametros)

