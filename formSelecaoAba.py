import tkinter as tk
from tkinter import ttk
from ParametrosProcessamento import ParametrosProcessamento

class FormSelecaoAba:
    def __init__(self, parent, parametros: ParametrosProcessamento):
        self.parent = parent
        self.parametros = parametros
        
        # Criar janela
        self.window = tk.Toplevel(parent)
        self.window.title("Seleção de Aba para Pesquisa")
        self.window.geometry("400x200")
        self.window.configure(bg="#f5f5f5")
        
        # Frame principal
        main_frame = tk.Frame(self.window, bg="#f5f5f5", padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Label para o combobox
        lbl_aba = tk.Label(
            main_frame,
            text="Selecione a Aba para Pesquisa:",
            font=("Segoe UI", 10),
            bg="#f5f5f5"
        )
        lbl_aba.pack(pady=(0, 5), anchor="w")
        
        # Combobox para seleção da aba
        self.combo_aba = ttk.Combobox(
            main_frame,
            state="readonly",
            font=("Segoe UI", 10),
            width=30
        )
        self.combo_aba.pack(pady=(0, 20), fill="x")
        
        # Frame para os botões
        btn_frame = tk.Frame(main_frame, bg="#f5f5f5")
        btn_frame.pack(fill="x", pady=(10, 0))
        
        # Botão Confirmar
        self.btn_confirmar = ttk.Button(
            btn_frame,
            text="Confirmar",
            command=self.confirmar
        )
        self.btn_confirmar.pack(side="right", padx=5)
        
        # Botão Cancelar
        self.btn_cancelar = ttk.Button(
            btn_frame,
            text="Cancelar",
            command=self.cancelar
        )
        self.btn_cancelar.pack(side="right", padx=5)
        
        # Centralizar a janela
        self.window.transient(parent)
        self.window.grab_set()
        self.center_window()
        
    def center_window(self):
        """Centraliza a janela na tela"""
        self.window.update_idletasks()
        width = self.window.winfo_width()
        height = self.window.winfo_height()
        x = (self.window.winfo_screenwidth() // 2) - (width // 2)
        y = (self.window.winfo_screenheight() // 2) - (height // 2)
        self.window.geometry(f'{width}x{height}+{x}+{y}')
        
    def set_abas(self, abas: list):
        """Define as abas disponíveis no combobox"""
        self.combo_aba['values'] = abas
        if abas:
            self.combo_aba.set(abas[0])
            
    def confirmar(self):
        """Callback do botão confirmar"""
        aba_selecionada = self.combo_aba.get()
        if aba_selecionada:
            self.parametros.aba_pesquisa = aba_selecionada
            self.window.destroy()
            
    def cancelar(self):
        """Callback do botão cancelar"""
        self.window.destroy()
        
    @staticmethod
    def mostrar_dialogo(parent, parametros: ParametrosProcessamento, abas: list) -> str:
        """
        Método estático para criar e mostrar o diálogo de seleção de aba.
        
        Args:
            parent: Widget pai
            parametros: Instância de ParametrosProcessamento
            abas: Lista de abas disponíveis
            
        Returns:
            Nome da aba selecionada ou None se cancelado
        """
        dialog = FormSelecaoAba(parent, parametros)
        dialog.set_abas(abas)
        dialog.window.wait_window()
        return parametros.aba_pesquisa