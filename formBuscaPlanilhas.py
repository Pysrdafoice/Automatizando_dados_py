
import tkinter as tk
from tkinter import ttk, filedialog
import pandas as pd
from dataclasses import dataclass
from parametrosPlanilha import ParametrosPlanilhas
from formParametrosPesquisa import FormParametrosPesquisa
from formSelecaoAba import FormSelecaoAba
from ParametrosProcessamento import ParametrosProcessamento

class FormBuscaPlanilhas:
    def __init__(self):
        self.janela = tk.Tk()
        self.janela.title("Busca Automática de Composições")
        self.janela.geometry("700x450")

        for i in range(6):
            self.janela.grid_columnconfigure(i, weight=1)
        # referência para evitar duplicidade de FormParametrosPesquisa
        self._form_parametros = None
        # lista de campos obrigatórios (vai ser preenchida após criação dos widgets)
        self._required_entries = []

        # CHAMADA PLANILHA REFERÊNCIA
        planilha_referencia = tk.Label(self.janela, text="Planilha de Preços Referência")
        planilha_referencia.grid(row=0, column=0, columnspan=6, padx=10, pady=5, sticky="n")

        self.entrada_planilha_referencia = tk.Entry(self.janela)
        self.entrada_planilha_referencia.grid(row=1, column=1, columnspan=5, sticky="we", padx=10, pady=5)

        buscar_planilha_referencia = tk.Button(
            self.janela,
            text="Buscar",
            bg="lightblue",
            command=lambda: self.buscar_planilhas(self.entrada_planilha_referencia, "Selecione a Planilha Referência"))
        buscar_planilha_referencia.grid(row=1, column=0, sticky="ew", padx=10, pady=5)

        # CHAMADA PLANILHA ORÇAMENTO
        planilha_preencher = tk.Label(self.janela, text="Planilha a Ser Preenchida")
        planilha_preencher.grid(row=2, column=0, columnspan=6, padx=10, pady=5, sticky="n")

        self.entrada_planilha_preencher = tk.Entry(self.janela)
        self.entrada_planilha_preencher.grid(row=3, column=1, columnspan=5, sticky="we", padx=10, pady=5)

        buscar_planilha_preencher = tk.Button(
            self.janela,
            text="Buscar",
            bg="lightblue",
            command=lambda: self.buscar_planilhas(self.entrada_planilha_preencher, "Selecione a Planilha a Ser Preenchida"))
        buscar_planilha_preencher.grid(row=3, column=0, sticky="ew", padx=10, pady=5)

        vazio = tk.Label(self.janela, text="")
        vazio.grid(row=4)

        # PARÂMETROS
        parametros = tk.Label(self.janela, text="Parâmetros")
        parametros.grid(row=5, column=0, columnspan=6, padx=10, pady=5, sticky="n")

        referencia = tk.Label(self.janela, text="Planilha Referência")
        referencia.grid(row=6, column=2, columnspan=2, padx=10, pady=5, sticky="n")

        orcamento = tk.Label(self.janela, text="Planilha Orçamento")
        orcamento.grid(row=6, column=4, columnspan=2, padx=10, pady=5, sticky="n")

        aba_usada = tk.Label(self.janela, text="Aba")
        aba_usada.grid(row=7, column=0, columnspan=2, padx=10, pady=5, sticky="n")
        
        # Frame para entrada e botão da aba de referência
        frame_aba_ref = tk.Frame(self.janela)
        frame_aba_ref.grid(row=7, column=2, columnspan=2, sticky="we", padx=10, pady=5)
        frame_aba_ref.grid_columnconfigure(0, weight=1)
        
        self.aba_planilha_referencia = tk.Entry(frame_aba_ref)
        self.aba_planilha_referencia.grid(row=0, column=0, sticky="we", padx=(0, 5))
        
        btn_selecionar_aba_ref = ttk.Button(
            frame_aba_ref,
            text="...",
            width=3,
            command=lambda: self.selecionar_aba(self.entrada_planilha_referencia, self.aba_planilha_referencia)
        )
        btn_selecionar_aba_ref.grid(row=0, column=1)
        
        # Frame para entrada e botão da aba de preenchimento
        frame_aba_preen = tk.Frame(self.janela)
        frame_aba_preen.grid(row=7, column=4, columnspan=2, sticky="we", padx=10, pady=5)
        frame_aba_preen.grid_columnconfigure(0, weight=1)
        
        self.aba_planilha_preencher = tk.Entry(frame_aba_preen)
        self.aba_planilha_preencher.grid(row=0, column=0, sticky="we", padx=(0, 5))
        
        btn_selecionar_aba_preen = ttk.Button(
            frame_aba_preen,
            text="...",
            width=3,
            command=lambda: self.selecionar_aba(self.entrada_planilha_preencher, self.aba_planilha_preencher)
        )
        btn_selecionar_aba_preen.grid(row=0, column=1)

        descricao = tk.Label(self.janela, text="Descrição")
        descricao.grid(row=8, column=0, columnspan=2, padx=10, pady=5, sticky="n")
        self.descricao_planilha_referencia = tk.Entry(self.janela)
        self.descricao_planilha_referencia.grid(row=8, column=2, columnspan=2, sticky="we", padx=10, pady=5)
        self.descricao_planilha_preencher = tk.Entry(self.janela)
        self.descricao_planilha_preencher.grid(row=8, column=4, columnspan=2, sticky="we", padx=10, pady=5)

        material = tk.Label(self.janela, text="Material")
        material.grid(row=9, column=0, columnspan=2, padx=10, pady=5, sticky="n")
        self.material_planilha_referencia = tk.Entry(self.janela)
        self.material_planilha_referencia.grid(row=9, column=2, columnspan=2, sticky="we", padx=10, pady=5)
        self.material_planilha_preencher = tk.Entry(self.janela)
        self.material_planilha_preencher.grid(row=9, column=4, columnspan=2, sticky="we", padx=10, pady=5)

        mao_de_obra = tk.Label(self.janela, text="Mão de Obra")
        mao_de_obra.grid(row=10, column=0, columnspan=2, padx=10, pady=5, sticky="n")
        self.mao_de_obra_planilha_referencia = tk.Entry(self.janela)
        self.mao_de_obra_planilha_referencia.grid(row=10, column=2, columnspan=2, sticky="we", padx=10, pady=5)
        self.mao_de_obra_planilha_preencher = tk.Entry(self.janela)
        self.mao_de_obra_planilha_preencher.grid(row=10, column=4, columnspan=2, sticky="we", padx=10, pady=5)

        unidade_medida = tk.Label(self.janela, text="Unidade de Medida")
        unidade_medida.grid(row=11, column=0, columnspan=2, padx=10, pady=5, sticky="n")
        self.unidade_medida_planilha_referencia = tk.Entry(self.janela)
        self.unidade_medida_planilha_referencia.grid(row=11, column=2, columnspan=2, sticky="we", padx=10, pady=5)
        self.unidade_medida_planilha_preencher = tk.Entry(self.janela)
        self.unidade_medida_planilha_preencher.grid(row=11, column=4, columnspan=2, sticky="we", padx=10, pady=5)

        # registrar campos obrigatórios para validação e highlight
        self._required_entries = [
            self.entrada_planilha_referencia,
            self.entrada_planilha_preencher,
            self.aba_planilha_referencia,
            self.aba_planilha_preencher,
            self.descricao_planilha_referencia,
            self.descricao_planilha_preencher,
            self.material_planilha_referencia,
            self.material_planilha_preencher,
            self.mao_de_obra_planilha_referencia,
            self.mao_de_obra_planilha_preencher,
            self.unidade_medida_planilha_referencia,
            self.unidade_medida_planilha_preencher,
        ]

        # ao digitar em um campo obrigatório, remover highlight
        for e in self._required_entries:
            try:
                e.bind('<KeyRelease>', lambda ev, widget=e: widget.config(bg='white'))
            except Exception:
                pass

        vazio = tk.Label(self.janela, text="")
        vazio.grid(row=12)

        # Botão para iniciar pesquisa
        frame_botoes = tk.Frame(self.janela)
        frame_botoes.grid(row=13, column=4, columnspan=2, padx=10, pady=5, sticky="nsew")
        frame_botoes.grid_columnconfigure(0, weight=1)
        frame_botoes.grid_columnconfigure(1, weight=1)

        botao_limpar = tk.Button(frame_botoes, text="Limpar", bg="lightblue", command=self.limpar)
        botao_limpar.grid(row=0, column=0, sticky="nsew", padx=(0, 5))

        botao_iniciar = tk.Button(frame_botoes, text="Avançar", command=self.avancar, bg="lightblue")
        botao_iniciar.grid(row=0, column=1, sticky="nsew", padx=(5, 0))

    def buscar_planilhas(self, entrada_destino, titulo_janela):
        caminho_arquivo = filedialog.askopenfilename(
            title=titulo_janela,
            filetypes=[("Planilhas Excel", "*.xlsx *.xls")]
        )
        if caminho_arquivo:
            entrada_destino.delete(0, tk.END)
            entrada_destino.insert(0, caminho_arquivo)

            # remove highlight se havia marcado antes
            try:
                entrada_destino.config(bg='white')
            except Exception:
                pass

    def limpar(self):
        self.entrada_planilha_referencia.delete(0, tk.END)
        self.entrada_planilha_preencher.delete(0, tk.END)
        
    def selecionar_aba(self, entrada_planilha, entrada_aba):
        """
        Abre o diálogo de seleção de aba para uma planilha específica
        
        Args:
            entrada_planilha: Entry contendo o caminho da planilha
            entrada_aba: Entry onde será colocada a aba selecionada
        """
        caminho_planilha = entrada_planilha.get()
        if not caminho_planilha:
            tk.messagebox.showwarning(
                "Aviso",
                "Por favor, selecione uma planilha primeiro."
            )
            return
            
        try:
            # Lê as abas da planilha
            excel = pd.ExcelFile(caminho_planilha)
        except FileNotFoundError:
            tk.messagebox.showerror("Erro", "Arquivo não encontrado. Selecione um arquivo válido.")
            return
        except (ValueError, OSError) as e:
            tk.messagebox.showerror("Erro", f"Erro ao abrir a planilha: {e}")
            return
        except Exception as e:
            tk.messagebox.showerror("Erro", f"Erro inesperado ao ler a planilha: {e}")
            return

        abas = excel.sheet_names
        if not abas:
            tk.messagebox.showwarning("Aviso", "A planilha não contém nenhuma aba.")
            return

        # Cria parâmetros temporários para o diálogo
        params = ParametrosProcessamento(None, None, None)

        # Mostra o diálogo de seleção
        aba_selecionada = FormSelecaoAba.mostrar_dialogo(self.janela, params, abas)

        if aba_selecionada:
            entrada_aba.delete(0, tk.END)
            entrada_aba.insert(0, aba_selecionada)
            try:
                entrada_aba.config(bg='white')
            except Exception:
                pass

    def avancar(self):
        parametros_referencia = ParametrosPlanilhas(
            caminho_planilha=self.entrada_planilha_referencia.get(),
            aba=self.aba_planilha_referencia.get(),
            coluna_descrição=self.descricao_planilha_referencia.get(),
            coluna_material=self.material_planilha_referencia.get(),
            coluna_mao_de_obra=self.mao_de_obra_planilha_referencia.get(),
            coluna_unidade_medida=self.unidade_medida_planilha_referencia.get()
        )

        parametros_orcamento = ParametrosPlanilhas(
            caminho_planilha=self.entrada_planilha_preencher.get(),
            aba=self.aba_planilha_preencher.get(),
            coluna_descrição=self.descricao_planilha_preencher.get(),
            coluna_material=self.material_planilha_preencher.get(),
            coluna_mao_de_obra=self.mao_de_obra_planilha_preencher.get(),
            coluna_unidade_medida=self.unidade_medida_planilha_preencher.get()
        )

        # Validação simples: todos os campos obrigatórios não devem estar vazios
        campos = [
            self.entrada_planilha_referencia.get().strip(),
            self.entrada_planilha_preencher.get().strip(),
            self.aba_planilha_referencia.get().strip(),
            self.aba_planilha_preencher.get().strip(),
            self.descricao_planilha_referencia.get().strip(),
            self.descricao_planilha_preencher.get().strip(),
            self.material_planilha_referencia.get().strip(),
            self.material_planilha_preencher.get().strip(),
            self.mao_de_obra_planilha_referencia.get().strip(),
            self.mao_de_obra_planilha_preencher.get().strip(),
            self.unidade_medida_planilha_referencia.get().strip(),
            self.unidade_medida_planilha_preencher.get().strip(),
        ]

        if any(not v for v in campos):
            # highlight missing fields
            for entry in self._required_entries:
                try:
                    if not entry.get().strip():
                        entry.config(bg='#ffcccc')
                    else:
                        entry.config(bg='white')
                except Exception:
                    pass

            tk.messagebox.showerror("Erro", "Preencha as informações necessárias para continuar")
            return

        # Evita abrir duplicidade: se já existe traz para frente
        if self._form_parametros is not None and getattr(self._form_parametros, 'janela', None) is not None:
            try:
                if self._form_parametros.janela.winfo_exists():
                    self._form_parametros.janela.deiconify()
                    self._form_parametros.janela.lift()
                    self._form_parametros.janela.focus_force()
                    return
            except Exception:
                # se algo deu errado com a referência, vamos criar uma nova
                self._form_parametros = None

        # Cria FormParametrosPesquisa como Toplevel modal ligado à janela principal
        self._form_parametros = FormParametrosPesquisa(parametros_referencia, parametros_orcamento, parent=self.janela)
        # garantir que, ao fechar a janela de parâmetros, removemos a referência para evitar duplicidade
        def _on_param_close():
            try:
                if self._form_parametros and getattr(self._form_parametros, 'janela', None):
                    try:
                        self._form_parametros.janela.destroy()
                    except Exception:
                        pass
            finally:
                self._form_parametros = None

        try:
            self._form_parametros.janela.protocol("WM_DELETE_WINDOW", _on_param_close)
        except Exception:
            pass

        # run modal
        self._form_parametros.run()

    def run(self):
        self.janela.mainloop()