import tkinter as tk
from tkinter import ttk, filedialog
from dataclasses import dataclass

from parametrosPlanilha import ParametrosPlanilhas
from OperacaoCorrelacao import OperacaoCorrelacao
from parametros_planilhas_pesquisa import ParametrosPlanilhasPesquisa
from ParametrosProcessamento import ParametrosProcessamento
from processamento import TelaProcessamento, ProcessamentoBase

class FormParametrosPesquisa:
    def __init__(self, parametros_referencia_teste: ParametrosPlanilhas, parametros_orcamento_teste: ParametrosPlanilhas, parent=None):
        """
        Se parent for fornecido, cria uma Toplevel modal ligada ao parent.
        Caso contrário, cria uma janela Tk principal.
        """
        self._is_root = parent is None
        if self._is_root:
            self.janela = tk.Tk()
        else:
            self.janela = tk.Toplevel(parent)
            # relaciona com a janela pai e torna modal
            self.janela.transient(parent)
            try:
                self.janela.grab_set()
            except Exception:
                pass

        self.janela.title("Busca Automática de Composições")
        self.janela.geometry("700x450")

        for i in range(6):
            self.janela.grid_columnconfigure(i, weight=1)

        self.parametros_referencia_teste = parametros_referencia_teste
        self.parametros_orcamento_teste = parametros_orcamento_teste

        parametros = tk.Label(self.janela, text="Parametros Escolhidos")
        parametros.grid(row=0, column=0, columnspan=6, padx=10, pady=5, sticky="n")

        referencia = tk.Label(self.janela, text="Planilha Referência")
        referencia.grid(row=1, column=2, columnspan=2, padx=10, pady=5, sticky="n")

        orcamento = tk.Label(self.janela, text="Planilha Orçamento")
        orcamento.grid(row=1, column=4, columnspan=2, padx=10, pady=5, sticky="n")

        aba_usada = tk.Label(self.janela, text="Aba")
        aba_usada.grid(row=2, column=0, columnspan=2, padx=10, pady=5, sticky="n")
        aba_planilha_referencia = tk.Label(self.janela, text=self.parametros_referencia_teste.aba)
        aba_planilha_referencia.grid(row=2, column=2, columnspan=2, sticky="we", padx=10, pady=5)
        aba_planilha_preencher = tk.Label(self.janela, text=self.parametros_orcamento_teste.aba)
        aba_planilha_preencher.grid(row=2, column=4, columnspan=2, sticky="we", padx=10, pady=5)

        descricao = tk.Label(self.janela, text="Descrição")
        descricao.grid(row=3, column=0, columnspan=2, padx=10, pady=5, sticky="n")
        descricao_planilha_referencia = tk.Label(self.janela, text=self.parametros_referencia_teste.coluna_descrição)
        descricao_planilha_referencia.grid(row=3, column=2, columnspan=2, sticky="we", padx=10, pady=5)
        descricao_planilha_preencher = tk.Label(self.janela, text=self.parametros_orcamento_teste.coluna_descrição)
        descricao_planilha_preencher.grid(row=3, column=4, columnspan=2, sticky="we", padx=10, pady=5)

        material = tk.Label(self.janela, text="Material")
        material.grid(row=4, column=0, columnspan=2, padx=10, pady=5, sticky="n")
        material_planilha_referencia = tk.Label(self.janela, text=self.parametros_referencia_teste.coluna_material)
        material_planilha_referencia.grid(row=4, column=2, columnspan=2, sticky="we", padx=10, pady=5)
        material_planilha_preencher = tk.Label(self.janela, text=self.parametros_orcamento_teste.coluna_material)
        material_planilha_preencher.grid(row=4, column=4, columnspan=2, sticky="we", padx=10, pady=5)

        mao_de_obra = tk.Label(self.janela, text="Mão de Obra")
        mao_de_obra.grid(row=5, column=0, columnspan=2, padx=10, pady=5, sticky="n")
        mao_de_obra_planilha_referencia = tk.Label(self.janela, text=self.parametros_referencia_teste.coluna_mao_de_obra)
        mao_de_obra_planilha_referencia.grid(row=5, column=2, columnspan=2, sticky="we", padx=10, pady=5)
        mao_de_obra_planilha_preencher = tk.Label(self.janela, text=self.parametros_orcamento_teste.coluna_mao_de_obra)
        mao_de_obra_planilha_preencher.grid(row=5, column=4, columnspan=2, sticky="we", padx=10, pady=5)

        unidade_medida = tk.Label(self.janela, text="Unidade de Medida")
        unidade_medida.grid(row=6, column=0, columnspan=2, padx=10, pady=5, sticky="n")
        unidade_medida_planilha_referencia = tk.Label(self.janela, text=self.parametros_referencia_teste.coluna_unidade_medida)
        unidade_medida_planilha_referencia.grid(row=6, column=2, columnspan=2, sticky="we", padx=10, pady=5)
        unidade_medida_planilha_preencher = tk.Label(self.janela, text=self.parametros_orcamento_teste.coluna_unidade_medida)
        unidade_medida_planilha_preencher.grid(row=6, column=4, columnspan=2, sticky="we", padx=10, pady=5)

        vazio = tk.Label(self.janela, text="")
        vazio.grid(row=7)

        #Validador para limitar numeros inteiros  e float
        validadorInteiro = (self.janela.register(self.somenteNumerosInteiros), '%P')
        validadorFloat = (self.janela.register(self.somenteNumerosFloat), '%P')

        parametros_pesquisa = tk.Label(self.janela, text="Parametros para Pesquisa na Planilha de Orçamento")
        parametros_pesquisa.grid(row=8, column=0, columnspan=6, padx=10, pady=5, sticky="n")

        comeco_pesquisa = tk.Label(self.janela, text="Começa em:")
        comeco_pesquisa.grid(row=9, column=0, columnspan=2, padx=10, pady=5, sticky="n")
        self.comeco_pesquisa_preencher = tk.Entry(
            self.janela,
            validate="key",
            validatecommand=validadorInteiro
        )
        self.comeco_pesquisa_preencher.grid(row=9, column=2, columnspan=4, sticky="we", padx=10, pady=5)

        termino_pesquisa = tk.Label(self.janela, text="Termina em:")
        termino_pesquisa.grid(row=10, column=0, columnspan=2, padx=10, pady=5, sticky="n")
        self.termino_pesquisa_preencher = tk.Entry(
            self.janela,
            validate="key",
            validatecommand=validadorInteiro
        )
        self.termino_pesquisa_preencher.grid(row=10, column=2, columnspan=4, sticky="we", padx=10, pady=5)

        taxa_similaridade = tk.Label(self.janela, text="Taxa de Similaridade (0 a 1):")
        taxa_similaridade.grid(row=11, column=0, columnspan=2, padx=10, pady=5, sticky="n")
        self.entrada_taxa_similaridade = tk.Entry(
            self.janela,
            validate="key",
            validatecommand=validadorFloat
        )
        self.entrada_taxa_similaridade.grid(row=11, column=2, columnspan=4, sticky="we", padx=10, pady=5)

        frame_botoes = tk.Frame(self.janela)
        frame_botoes.grid(row=12, column=4, columnspan=2, padx=10, pady=5, sticky="nsew")
        frame_botoes.grid_columnconfigure(0, weight=1)
        frame_botoes.grid_columnconfigure(1, weight=1)

        botao_limpar = tk.Button(frame_botoes, text="Limpar", bg="lightblue", command=self.limpar)
        botao_limpar.grid(row=0, column=0, sticky="nsew", padx=(0, 5))

        botao_iniciar = tk.Button(frame_botoes, text="Avançar", command=self.avancar, bg="lightblue")
        botao_iniciar.grid(row=0, column=1, sticky="nsew", padx=(5, 0))

    def limpar(self):
        self.comeco_pesquisa_preencher.delete(0, tk.END)
        self.termino_pesquisa_preencher.delete(0, tk.END)
        self.entrada_taxa_similaridade.delete(0, tk.END)

    def avancar(self):
        # Dados da tela atual
        operacaoAtual = OperacaoCorrelacao(
            ComecoPesquisa=int(self.comeco_pesquisa_preencher.get()),
            TerminoPesquisa=int(self.termino_pesquisa_preencher.get()),
            TaxaSimilaridade=float(self.entrada_taxa_similaridade.get())
        )

        # Cria um objeto ParametrosProcessamento tipado
        parametros = ParametrosProcessamento(
            referencia=self.parametros_referencia_teste,
            orcamento=self.parametros_orcamento_teste,
            pesquisa=operacaoAtual
        )

        # Criar janela de processamento (sempre Toplevel ligada à janela atual)
        janela_processamento = tk.Toplevel(self.janela)
        processamento = TelaProcessamento(janela_processamento, parametros)
        # Não chamamos mainloop() aqui porque já estamos dentro de um loop Tkinter

    @staticmethod
    def somenteNumerosInteiros(texto):
        if texto == "":
            return True
        return texto.isdigit()

    @staticmethod
    def somenteNumerosFloat(texto):
        if texto == "":
            return True
        try:
            valor = float(texto)
            return 0 <= valor <= 1
        except ValueError:
            return False

    def run(self):
        # Se esta classe criou um Tk, executa mainloop normalmente.
        # Caso contrário (Toplevel), espera a janela ser fechada (modal behavior).
        if self._is_root:
            self.janela.mainloop()
        else:
            try:
                self.janela.wait_window()
            except Exception:
                pass