# Automatizando_dados_py

# Sistema de Cálculo de Materiais

Uma aplicação em **Python** com interface gráfica que permite carregar planilhas Excel (descrições e medidas de materiais), buscar materiais por aproximação e calcular quantidades necessárias com base em medi‑das.

---

## 🧭 Visão Geral

Este projeto usa os seguintes recursos:

* **Objetos e Classes (OOP)** para organizar o programa de forma clara e modular.
* **Pandas** para ler e manipular dados de planilhas Excel.
* **Tkinter + ttk** para criar a interface gráfica interativa.
* **difflib.get\_close\_matches** para busca por similaridade de texto.

---

## ✅ Funcionalidades

1. Carrega duas planilhas:

   * **Descrições**: contém coluna `Descrição`.
   * **Medidas**: contém colunas `Descrição`, `Medida`, `Quantidade`.
2. Botão **“Carregar Planilhas”**: permite ao usuário escolher ambos os arquivos Excel.
3. Campo de busca: digite uma descrição (ex.: "cimento") e clique em **Buscar**.
4. **Listbox** exibe até 10 materiais similares a partir da coluna `Descrição`.
5. Clique sobre um item na lista para abrir calculadora.
6. Nova janela com:

   * Seleção de **medida**.
   * Entrada de **quantidade necessária**.
   * Botão **Calcular** gera o número de unidades necessárias.
7. Adiciona ao DataFrame de descrições:

   * `Medida Utilizada`
   * `Quantidade Calculada`

---

## 📁 Estrutura do Código

```text
- import pandas as pd
- import tkinter as tk
- from tkinter import filedialog, messagebox, ttk
- from difflib import get_close_matches

class SistemaMateriais:
    __init__              # inicializa janela principal
    criar_interface       # constrói a interface com botões e campos
    carregar_planilhas     # carrega arquivos Excel via diálogo
    buscar_material        # busca por similaridade na coluna 'Descrição'
    selecionar_material    # abre janela de cálculo ao selecionar item na lista
    calcular_quantidade    # calcula e atualiza os DataFrames
```

### Explicação OOP (Programação Orientada a Objetos)

* A classe `SistemaMateriais` encapsula todo o comportamento da aplicação.
* O método `__init__` configura tudo ao iniciar (janela, variáveis de dados e interface).
* Cada funcionalidade (interface, carregamento, busca, cálculo) é implementada como método independente — o que facilita manutenção e entendimento.
* `self` refere-se ao objeto que mantém o estado da aplicação (como `self.df_descricoes`).

---

## 🚀 Como Cada Parte Funciona

### 1. Interface Gráfica (`criar_interface`)

* Usa **Frame** para organizar os widgets.
* **Button** com comando `self.carregar_planilhas`.
* **Entry** para busca, inicialmente desativada (`state='disabled'`).
* **Listbox** com scrollbar para mostrar resultados.
* Ao selecionar um item, chama `selecionar_material` via evento `<<ListboxSelect>>`.

#### Por que usar `ttk` e `messagebox`?

* `ttk.Combobox` exige `from tkinter import ttk`.
* Caixas de mensagem (`showinfo`, `showerror`) exigem `messagebox`.
  Esses widgets fazem parte dos submódulos de Tkinter e precisam ser importados corretamente ([linuxconfig.org][1], [geeksforgeeks.org][2], [geeksforgeeks.org][3]).

---

### 2. Carregamento das Planilhas (`carregar_planilhas`)

* Usa `filedialog.askopenfilename` para abrir diálogo de seleção de arquivos.
* Utiliza **pandas** (`pd.read_excel`) para ler os dados.
* Após carregar com sucesso, ativa os widgets de busca.

O uso de pandas permite acessar colunas como `df_descricoes['Descrição']`, transformando os textos para `lower()` e comparações.

---

### 3. Busca Aproxima da Descrição (`buscar_material`)

* Lê o texto digitado (`entry_busca`), converte para minúsculo e busca por similaridade com `get_close_matches`.
* O parâmetro `cutoff=0.6` indica tolerância mínima de 60% de semelhança.
* Retorna até 10 sugestões.
* Lista de resultados preenchida na `Listbox`.

---

### 4. Seleção e Cálculo (`selecionar_material` e `calcular_quantidade`)

#### Seleção:

* Captura o item clicado na lista.
* Abre janela `Toplevel`, filtra no DataFrame `df_medidas` pela descrição selecionada.

#### Cálculo:

* Recupera medida escolhida (via `ttk.Combobox`) e quantidade desejada.
* Calcula unidades necessárias como `quantidade_necessaria / quantidade_por_unidade` e arredonda para cima.
* Atualiza `df_descricoes` adicionando colunas com `at[idx]`.
* Exibe resultado em `messagebox.showinfo`.

---

## 👩‍🏫 Explicação para Iniciantes

* **Classe** é como um modelo de criação de programas: ela armazena dados (DataFrames) e ações (métodos).
* **Métodos** são funções dentro da classe: cada tarefa tem seu próprio método.
* **Self** é o objeto vivo da classe que guarda tudo: janela, dados, funções.
* **Pandas** é como uma planilha dentro do programa: permite ler e processar dados facilmente.
* **Tkinter** cria elementos visuais como botões, caixas de texto e janelas. É comum usar `grid()` ou `pack()` para organizar.
* **Combobox** (via `ttk`) é uma caixa de seleção moderna.
* **Get\_close\_matches** ajuda a encontrar descrições semelhantes mesmo se o texto não for exato.

---

## 🧪 Como Usar

1. Instale dependências:

   ```bash
   pip install pandas openpyxl
   ```
2. Execute o programa:

   ```bash
   python sistema_materiais.py
   ```
3. Na janela que aparece:

   * Clique em **Carregar Planilhas** e selecione os arquivos Excel.
   * Digite a descrição e clique em **Buscar**.
   * Selecione o material na lista.
   * Escolha a medida, digite a quantidade e clique em **Calcular**.

---

## 🎓 Conceitos Aprendidos

| Conceito          | Descrição simples                                           |
| ----------------- | ----------------------------------------------------------- |
| **Classe/Objeto** | Estrutura que junta dados e ações                           |
| **Método**        | Função dentro de classe executa tarefa                      |
| **Pandas**        | Biblioteca que lê Excel e manipula dados                    |
| **Tkinter/ttk**   | Ferramenta para criar interface gráfica                     |
| **Evento GUI**    | Ações como clicar, digitar ou selecionar disparando funções |

---

## ✍️ Próximos Passos

* Adicionar validação de colunas obrigatórias nas planilhas.
* Salvar alterações (DataFrame) de volta para Excel.
* Melhorar layout com `grid()` responsivo.
* Adicionar temas modernos com `ttk.Style` ou bibliotecas como Sun‑Valley ([geeksforgeeks.org][4], [reddit.com][5]).


[1]: https://linuxconfig.org/how-to-build-a-tkinter-application-using-an-object-oriented-approach?utm_source=chatgpt.com "How to build a Tkinter application using an object oriented approach"
[2]: https://www.geeksforgeeks.org/git/what-is-readme-md-file/?utm_source=chatgpt.com "What is README.md File? - GeeksforGeeks"
[3]: https://www.geeksforgeeks.org/python/python-gui-tkinter/?utm_source=chatgpt.com "Python Tkinter - GeeksforGeeks"
[4]: https://www.geeksforgeeks.org/python/python-tkinter-tutorial/?utm_source=chatgpt.com "Python Tkinter Tutorial - GeeksforGeeks"
[5]: https://www.reddit.com/r/Python/comments/xbgyov/make_your_tkinter_app_look_truly_modern_with_a/?utm_source=chatgpt.com "Make your Tkinter app look truly modern with a single line of code!"
