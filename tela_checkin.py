import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from typing import List, Callable, Optional
from dataclasses import dataclass
from pathlib import Path
import logging
import sys
from datetime import datetime

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


@dataclass
class ItemCheckin:
    """Representa um item selecionado no checkin"""
    item: str
    unidade: str
    referencia: str
    similaridade: float
    valor_total: float
    numero_linha: int
    valor_material: float = 0.0
    valor_mao_de_obra: float = 0.0


# ===== FUNÇÕES PROCESSADORAS (Lógica Pura) =====

def gerar_texto_resumo(itens_selecionados: List[ItemCheckin]) -> str:
    """
    Gera texto do resumo com totais.
    
    Args:
        itens_selecionados: Lista de itens a contabilizar
        
    Returns:
        str: Texto formatado do resumo
    """
    logger.debug(f">>> gerar_texto_resumo INICIADO | {len(itens_selecionados)} itens")
    
    try:
        total_itens = len(itens_selecionados)
        valor_total = sum(item_checkin.valor_total for item_checkin in itens_selecionados)
        
        texto = f"Total de itens: {total_itens}  |  Valor total: R$ {valor_total:.2f}"
        logger.debug(f"<<< gerar_texto_resumo CONCLUÍDO | {texto}")
        
        return texto
    
    except Exception as e:
        logger.error(f"ERRO em gerar_texto_resumo: {e}")
        import traceback
        traceback.print_exc()
        return "Erro ao gerar resumo"


def preparar_itens_para_display(itens_selecionados: List[ItemCheckin]) -> List[tuple]:
    """
    Prepara itens para exibição na Treeview com badges de ocorrências.
    
    Args:
        itens_selecionados: Lista de itens a processar
        
    Returns:
        List[tuple]: Lista de tuplas formatadas para exibição (item_nome, unidade, ref_display, similaridade, valor, acao)
    """
    logger.debug(f">>> preparar_itens_para_display INICIADO | {len(itens_selecionados)} itens")
    
    try:
        # Contar quantas vezes cada referência aparece
        referencias_contadas = {}
        for item_checkin in itens_selecionados:
            chave = item_checkin.referencia
            if chave not in referencias_contadas:
                referencias_contadas[chave] = []
            referencias_contadas[chave].append(item_checkin.numero_linha)
        
        logger.debug(f"Referências contadas: {len(referencias_contadas)} tipos")
        
        # Rastrear referências já exibidas (não duplicar na UI)
        referencias_exibidas = set()
        itens_para_exibir = []
        
        # Processar cada item
        for item_checkin in itens_selecionados:
            # Exibir cada referência apenas UMA VEZ na UI
            if item_checkin.referencia not in referencias_exibidas:
                # Contar ocorrências
                qtd_ocorrencias = len(referencias_contadas[item_checkin.referencia])
                
                # Montar display com badge
                if qtd_ocorrencias > 1:
                    display_ref = f"{item_checkin.referencia} ({qtd_ocorrencias} ocorrências)"
                    logger.debug(f"  Badge: {item_checkin.referencia} - {qtd_ocorrencias} ocorrências")
                else:
                    display_ref = item_checkin.referencia
                
                # Adicionar linha à lista de exibição
                linha = (
                    item_checkin.item,
                    item_checkin.unidade,
                    display_ref,
                    f"{item_checkin.similaridade:.1f}%",
                    f"R$ {item_checkin.valor_total:.2f}",
                    "Excluir"
                )
                itens_para_exibir.append(linha)
                referencias_exibidas.add(item_checkin.referencia)
        
        logger.debug(f"<<< preparar_itens_para_display CONCLUÍDO | {len(itens_para_exibir)} linhas para exibir")
        return itens_para_exibir
    
    except Exception as e:
        logger.error(f"ERRO em preparar_itens_para_display: {e}")
        import traceback
        traceback.print_exc()
        return []


def processar_exclusao_item_checkin(
    item_nome: str,
    numero_linha_target: int,
    itens_selecionados: List[ItemCheckin],
    callback_excluir: Optional[Callable]
) -> tuple:
    """
    Processa exclusão de um item do checkin.
    
    Args:
        item_nome: Nome do item a excluir
        numero_linha_target: ID único do item (numero_linha)
        itens_selecionados: Lista atual de itens
        callback_excluir: Callback a chamar antes da exclusão
        
    Returns:
        tuple: (itens_atualizados, sucesso: bool, mensagem: str)
    """
    itens_atualizados = list(itens_selecionados)  # Cópia para imutabilidade
    
    logger.debug(f"processar_exclusao_item_checkin | item_nome={item_nome}, numero_linha={numero_linha_target}")
    
    try:
        # Encontrar item usando numero_linha como ID único
        item_para_remover = None
        idx_para_remover = None
        
        for i, item_checkin in enumerate(itens_atualizados):
            if item_checkin.item == item_nome and item_checkin.numero_linha == numero_linha_target:
                item_para_remover = item_checkin
                idx_para_remover = i
                break
        
        if idx_para_remover is None:
            msg = f"Item não encontrado: {item_nome} (linha {numero_linha_target})"
            logger.error(msg)
            return (itens_atualizados, False, msg)
        
        logger.debug(f"Item encontrado no índice {idx_para_remover}")
        
        # Chamar callback ANTES de remover (sincronização)
        if callback_excluir:
            try:
                callback_excluir(numero_linha_target)
                logger.debug(f"Callback de exclusão executado para: {item_nome}")
            except Exception as e:
                msg = f"Erro no callback de exclusão: {str(e)}"
                logger.error(msg)
                return (itens_atualizados, False, msg)
        
        # Remover após sucesso do callback
        itens_atualizados.pop(idx_para_remover)
        logger.debug(f"Item removido: {item_nome} | Total restante: {len(itens_atualizados)}")
        
        return (itens_atualizados, True, f"Item '{item_nome}' removido com sucesso")
    
    except Exception as e:
        msg = f"Erro em processar_exclusao_item_checkin: {str(e)}"
        logger.error(msg)
        import traceback
        traceback.print_exc()
        return (itens_selecionados, False, msg)


class TelaCheckin:
    """Tela de confirmação (Checkin) para exibir itens selecionados antes de finalizar"""
    
    # ===== CONTROLE DE DUPLICAÇÃO =====
    # Sinalizador de classe para evitar múltiplas instâncias abertas simultaneamente
    _instance_aberta = None  # marker para evitar tela duplicada
    
    def __init__(self, parent, itens_selecionados: List[ItemCheckin], callback_confirmar: Optional[Callable] = None, callback_excluir: Optional[Callable] = None, parametros_planilha: Optional[dict] = None):
        """
        Inicializa a tela de Checkin.
        
        Args:
            parent: janela pai (para modal)
            itens_selecionados: lista de ItemCheckin com dados dos itens selecionados
            callback_confirmar: função opcional para executar ao confirmar (recebe itens como argumento)
            callback_excluir: função opcional chamada quando item é excluído (recebe numero_linha do item)
            parametros_planilha: dict com caminhos e abas da planilha para atualização
        """
        logger.debug("=== INICIANDO TelaCheckin.__init__ ===")
        logger.debug(f"itens_selecionados: {len(itens_selecionados)} itens")
        logger.debug(f"_instance_aberta ANTES: {TelaCheckin._instance_aberta}")
        
        # Blindagem: Se já há tela aberta, trazer para frente e retornar
        if TelaCheckin._instance_aberta is not None and TelaCheckin._instance_aberta.janela.winfo_exists():
            logger.warning("⚠️ TelaCheckin já está aberta - trazendo para frente (RETORNANDO)")
            TelaCheckin._instance_aberta.janela.lift()
            TelaCheckin._instance_aberta.janela.focus()
            return  # ← RETORNO PREMATURO
        
        # Registrar esta instância como a aberta
        TelaCheckin._instance_aberta = self
        logger.debug(f"✓ _instance_aberta registrada: {self}")
        
        self.parent = parent
        self.itens_selecionados = list(itens_selecionados)
        self.callback_confirmar = callback_confirmar
        self.callback_excluir = callback_excluir
        self.parametros_planilha = parametros_planilha
        self.confirmado = False
        self.tree = None
        self.frame_tabela = None
        self.lbl_resumo_texto = None
        logger.debug("✓ Atributos inicializados")
        
        # Criar janela modal
        try:
            logger.debug("Criando tk.Toplevel...")
            self.janela = tk.Toplevel(parent)
            self.janela.title("Confirmação de Atualização - Checkin")
            self.janela.geometry("1100x650")
            self.janela.configure(bg="#f5f5f5")
            logger.debug("✓ Janela Toplevel criada")
            
            # Tornar modal
            logger.debug("Configurando modal...")
            self.janela.grab_set()
            self.janela.transient(parent)
            logger.debug("✓ Modal configurado")
            
            # Criar layout
            logger.debug("Chamando _criar_layout...")
            self._criar_layout()
            logger.debug("✓ _criar_layout concluído")
            
            # Centralizar janela
            logger.debug("Centralizando janela...")
            self.janela.update_idletasks()
            x = parent.winfo_x() + (parent.winfo_width() - self.janela.winfo_width()) // 2
            y = parent.winfo_y() + (parent.winfo_height() - self.janela.winfo_height()) // 2
            self.janela.geometry(f"+{x}+{y}")
            logger.debug("✓ Janela centralizada")
            
            # Protocol para fechar
            logger.debug("Configurando WM_DELETE_WINDOW...")
            self.janela.protocol("WM_DELETE_WINDOW", self._ao_fechar)
            logger.debug("✓ WM_DELETE_WINDOW configurado")
            
            logger.debug("=== TelaCheckin.__init__ CONCLUÍDO COM SUCESSO ===\n")
        
        except Exception as e:
            logger.error(f"ERRO CRÍTICO em TelaCheckin.__init__: {e}")
            import traceback
            traceback.print_exc()
            # Limpar marker se houve erro
            TelaCheckin._instance_aberta = None
            raise
    
    def _criar_layout(self):
        """Cria o layout da tela de Checkin"""
        logger.debug("=== INICIANDO _criar_layout ===")
        
        try:
            # --- Cabeçalho ---
            logger.debug("Criando cabeçalho...")
            frame_cabecalho = tk.Frame(self.janela, bg="#004c99", pady=15)
            frame_cabecalho.pack(fill="x")
            
            lbl_titulo = tk.Label(
                frame_cabecalho,
                text="Confirmação de Atualização",
                fg="white",
                bg="#004c99",
                font=("Segoe UI", 14, "bold")
            )
            lbl_titulo.pack()
            
            lbl_subtitulo = tk.Label(
                frame_cabecalho,
                text="Revise os itens selecionados antes de confirmar (clique em uma linha para excluir)",
                fg="#e0e0e0",
                bg="#004c99",
                font=("Segoe UI", 10)
            )
            lbl_subtitulo.pack()
            logger.debug("✓ Cabeçalho criado")
            
            # --- Área de tabela com itens ---
            logger.debug("Criando frame_tabela...")
            self.frame_tabela = tk.Frame(self.janela, bg="#ffffff", bd=1, relief="solid")
            self.frame_tabela.pack(padx=20, pady=10, fill="both", expand=True)
            
            # Criar Treeview para exibir itens
            logger.debug("Criando Treeview...")
            colunas = ("item", "unidade", "referencia", "similaridade", "valor_total", "acao")
            
            self.tree = ttk.Treeview(
                self.frame_tabela,
                columns=colunas,
                show="headings",
                height=12
            )
            logger.debug("✓ Treeview criado")
            
            headers = {
                "item": "Item do Orçamento",
                "unidade": "Unidade",
                "referencia": "Referência Sugerida",
                "similaridade": "Similaridade",
                "valor_total": "Valor Total",
                "acao": "Ação"
            }
            
            widths = {
                "item": 280,
                "unidade": 80,
                "referencia": 280,
                "similaridade": 100,
                "valor_total": 120,
                "acao": 80
            }
            
            logger.debug("Configurando colunas...")
            for col in colunas:
                self.tree.heading(col, text=headers[col], anchor="w")
                self.tree.column(col, width=widths[col], minwidth=50)
            logger.debug("✓ Colunas configuradas")
            
            # Scrollbars
            logger.debug("Criando scrollbars...")
            scrolly = ttk.Scrollbar(self.frame_tabela, orient="vertical", command=self.tree.yview)
            scrollx = ttk.Scrollbar(self.frame_tabela, orient="horizontal", command=self.tree.xview)
            self.tree.configure(yscrollcommand=scrolly.set, xscrollcommand=scrollx.set)
            
            # Grid layout
            self.tree.grid(row=0, column=0, sticky="nsew")
            scrolly.grid(row=0, column=1, sticky="ns")
            scrollx.grid(row=1, column=0, sticky="ew")
            
            self.frame_tabela.grid_rowconfigure(0, weight=1)
            self.frame_tabela.grid_columnconfigure(0, weight=1)
            logger.debug("✓ Scrollbars criados e grid configurado")
            
            # Preencher tabela com itens
            logger.debug(f"Atualizando tabela com {len(self.itens_selecionados)} itens...")
            self._atualizar_tabela()
            logger.debug("✓ Tabela atualizada")
            
            # Bind de clique duplo para excluir
            self.tree.bind("<Double-1>", self._ao_excluir_item)
            logger.debug("✓ Bind de duplo clique configurado")
            
            # --- Resumo ---
            logger.debug("Criando frame_resumo...")
            frame_resumo = tk.Frame(self.janela, bg="#f5f5f5", pady=10)
            frame_resumo.pack(fill="x", padx=20)
            
            self.lbl_resumo_texto = tk.Label(
                frame_resumo,
                text="",
                bg="#f5f5f5",
                font=("Segoe UI", 10, "bold"),
                fg="#004c99"
            )
            self.lbl_resumo_texto.pack()
            self._atualizar_resumo()
            logger.debug("✓ Resumo criado")
            
            # --- Botões (COM expand=True para ocupar espaço correto) ---
            logger.debug("Criando frame_botoes com 3 botões...")
            frame_botoes = tk.Frame(self.janela, bg="#f5f5f5", pady=15)
            frame_botoes.pack(fill="both", expand=False, padx=20)  # expand=False, não compete com tabela
            
            btn_cancelar = tk.Button(
                frame_botoes,
                text="Cancelar",
                bg="#d32f2f",
                fg="white",
                font=("Segoe UI", 10),
                width=15,
                command=self._ao_fechar
            )
            btn_cancelar.pack(side="left", padx=5, pady=5)
            logger.debug("✓ Botão Cancelar criado")
            
            btn_finalizar_preenchimento = tk.Button(
                frame_botoes,
                text="Finalizar Preenchimento",
                bg="#ff9800",
                fg="white",
                font=("Segoe UI", 10),
                width=20,
                command=self._finalizar_preenchimento
            )
            btn_finalizar_preenchimento.pack(side="left", padx=5, pady=5)
            logger.debug("✓ Botão Finalizar Preenchimento criado")
            
            btn_confirmar = tk.Button(
                frame_botoes,
                text="Confirmar",
                bg="#388e3c",
                fg="white",
                font=("Segoe UI", 10),
                width=15,
                command=self._confirmar
            )
            btn_confirmar.pack(side="right", padx=5, pady=5)
            logger.debug("✓ Botão Confirmar criado")
            
            logger.debug("=== _criar_layout CONCLUÍDO COM SUCESSO ===\n")
        
        except Exception as e:
            logger.error(f"ERRO CRÍTICO em _criar_layout: {e}")
            import traceback
            traceback.print_exc()
            raise  # Re-lançar para não esconder o erro
    
    def _atualizar_tabela(self):
        """Limpa e repopula a tabela com os itens atuais usando processador"""
        logger.debug(f">>> _atualizar_tabela INICIADO | {len(self.itens_selecionados)} itens")
        
        try:
            # Limpar tabela
            logger.debug("  Limpando tabela existente...")
            for item in self.tree.get_children():
                self.tree.delete(item)
            logger.debug("  ✓ Tabela limpa")
            
            # Processar itens para exibição
            logger.debug("  Chamando preparar_itens_para_display...")
            itens_para_exibir = preparar_itens_para_display(self.itens_selecionados)
            logger.debug(f"  ✓ Retorno: {len(itens_para_exibir)} itens para exibir")
            
            # Preencher tabela
            logger.debug("  Inserindo itens na tree...")
            for linha in itens_para_exibir:
                self.tree.insert("", "end", values=linha)
            logger.debug(f"  ✓ {len(itens_para_exibir)} linhas inseridas")
            
            logger.debug(f"<<< _atualizar_tabela CONCLUÍDO")
        
        except Exception as e:
            logger.error(f"ERRO em _atualizar_tabela: {e}")
            import traceback
            traceback.print_exc()
    
    def _atualizar_resumo(self):
        """Atualiza o texto do resumo usando processador"""
        logger.debug(f">>> _atualizar_resumo INICIADO")
        
        try:
            # Processar texto do resumo
            logger.debug("  Chamando gerar_texto_resumo...")
            texto_resumo = gerar_texto_resumo(self.itens_selecionados)
            logger.debug(f"  ✓ Retorno: '{texto_resumo}'")
            
            # Atualizar label
            logger.debug("  Atualizando lbl_resumo_texto...")
            self.lbl_resumo_texto.config(text=texto_resumo)
            logger.debug(f"  ✓ Label atualizado")
            
            logger.debug(f"<<< _atualizar_resumo CONCLUÍDO")
        
        except Exception as e:
            logger.error(f"ERRO em _atualizar_resumo: {e}")
            import traceback
            traceback.print_exc()
    
    def _ao_excluir_item(self, event):
        """Exclui item quando o usuário clica duplo na linha"""
        logger.debug("_ao_excluir_item disparado")
        
        try:
            # Blindagem: Ignorar clique se nenhum item selecionado
            if not self.tree.selection():
                logger.debug("Nenhum item selecionado")
                return
            
            item_id = self.tree.selection()[0]
            valores = self.tree.item(item_id)["values"]
            
            item_nome = valores[0]
            
            # Confirmar exclusão
            confirmado = messagebox.askyesno(
                "Excluir Item",
                f"Deseja excluir o item:\n\n'{item_nome}'?\n\nEsta ação não pode ser desfeita."
            )
            
            if confirmado:
                logger.debug(f"Exclusão confirmada para: {item_nome}")
                
                # Encontrar numero_linha do item
                numero_linha_alvo = None
                for item_checkin in self.itens_selecionados:
                    if item_checkin.item == item_nome:
                        numero_linha_alvo = item_checkin.numero_linha
                        break
                
                if numero_linha_alvo is None:
                    logger.error(f"numero_linha não encontrado para: {item_nome}")
                    messagebox.showerror("Erro", "Item não encontrado na lista.")
                    return
                
                # Chamar processador com estado explícito
                itens_atualizados, sucesso, mensagem = processar_exclusao_item_checkin(
                    item_nome,
                    numero_linha_alvo,
                    self.itens_selecionados,
                    self.callback_excluir
                )
                
                if sucesso:
                    # Atualizar estado
                    self.itens_selecionados = itens_atualizados
                    
                    # Atualizar visualização
                    self._atualizar_tabela()
                    self._atualizar_resumo()
                    
                    logger.debug(f"Item removido com sucesso: {item_nome}")
                    
                    # Se não houver mais itens, avisar
                    if not self.itens_selecionados:
                        messagebox.showinfo(
                            "Aviso",
                            "Todos os itens foram removidos.\n\nPor favor, retorne para fazer novas seleções."
                        )
                else:
                    # Falha na remoção
                    logger.error(f"Falha ao remover item: {mensagem}")
                    messagebox.showerror(
                        "Erro na Exclusão",
                        f"Não foi possível excluir o item:\n\n{mensagem}"
                    )
        
        except Exception as e:
            logger.error(f"Erro em _ao_excluir_item: {e}")
            import traceback
            traceback.print_exc()
    
    def _finalizar_preenchimento(self):
        """Atualiza a planilha com os itens selecionados"""
        if not self.itens_selecionados:
            messagebox.showwarning(
                "Aviso",
                "Não há itens para preencher."
            )
            return
        
        if not self.parametros_planilha:
            messagebox.showerror(
                "Erro",
                "Parâmetros da planilha não foram fornecidos."
            )
            return
        
        # Pedir ao usuário onde salvar
        caminho_destino = filedialog.asksaveasfilename(
            title="Salvar Planilha Preenchida",
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")],
            initialdir=str(Path(self.parametros_planilha["caminho_orcamento"]).parent)
        )
        
        if not caminho_destino:
            logger.debug("Salvamento cancelado pelo usuário")
            return
        
        try:
            from atualizador_planilha import AtualizadorPlanilha
            
            logger.debug("Iniciando preenchimento da planilha...")
            
            # Preparar dados para atualização
            selecoes_para_atualizar = []
            for item_checkin in self.itens_selecionados:
                selecoes_para_atualizar.append({
                    "item": item_checkin.item,
                    "numero_linha": item_checkin.numero_linha,
                    "referencia": item_checkin.referencia,
                    "valor_material": item_checkin.valor_material,
                    "valor_mao_de_obra": item_checkin.valor_mao_de_obra,
                    "unidade": item_checkin.unidade
                })
            
            # Atualizar planilha
            atualizador = AtualizadorPlanilha(
                caminho_planilha=self.parametros_planilha["caminho_orcamento"],
                aba=self.parametros_planilha["aba_orcamento"],
                parametros={
                    "coluna_material": self.parametros_planilha["coluna_material"],
                    "coluna_mao_de_obra": self.parametros_planilha["coluna_mao_de_obra"],
                    "coluna_unidade_medida": self.parametros_planilha["coluna_unidade_medida"]
                }
            )
            
            caminho_salvo = atualizador.atualizar_com_selecoes(
                selecoes_para_atualizar,
                caminho_destino
            )
            
            messagebox.showinfo(
                "Sucesso",
                f"Planilha preenchida com sucesso!\n\n"
                f"Arquivo: {Path(caminho_salvo).name}\n"
                f"Localização: {Path(caminho_salvo).parent}"
            )
            
            logger.debug(f"Planilha salva em: {caminho_salvo}")
            
        except Exception as e:
            logger.error(f"Erro ao preencher planilha: {e}")
            import traceback
            traceback.print_exc()
            messagebox.showerror(
                "Erro",
                f"Erro ao preencher planilha:\n\n{str(e)}"
            )
    
    def _confirmar(self):
        """Executa callback de confirmação e fecha a janela"""
        logger.debug(">>> _confirmar INICIADO")
        self.confirmado = True
        if self.callback_confirmar:
            try:
                logger.debug("  Executando callback_confirmar...")
                self.callback_confirmar(self.itens_selecionados)
                logger.debug("  ✓ callback_confirmar executado com sucesso")
            except Exception as e:
                logger.error(f"ERRO ao executar callback_confirmar: {e}")
        logger.debug("  Chamando _ao_fechar...")
        self._ao_fechar()  # Fechar após confirmar
        logger.debug("<<< _confirmar CONCLUÍDO")
    
    def _ao_fechar(self):
        """Fecha a janela sem confirmar e limpa o marker de instância"""
        logger.debug(">>> _ao_fechar INICIADO")
        logger.debug(f"  _instance_aberta ANTES: {TelaCheckin._instance_aberta is self}")
        
        # Limpar o marker para permitir nova instância
        if TelaCheckin._instance_aberta is self:
            logger.debug("  ✓ Limpando _instance_aberta")
            TelaCheckin._instance_aberta = None
        
        self.confirmado = False
        logger.debug("  Destruindo janela...")
        self.janela.destroy()
        logger.debug("✓ _ao_fechar CONCLUÍDO - janela destruída")
    
    def run(self) -> bool:
        """
        Exibe a tela modal e retorna True se confirmou, False caso contrário.
        
        Returns:
            bool: True se usuário confirmou, False caso contrário
        """
        self.janela.wait_window()
        return self.confirmado
