# Automatizando_dados_py(AUTALIZAÇÕES...)

# Sistema de Cálculo de Materiais

Uma aplicação em **Python** com interface gráfica que permite carregar planilhas Excel (descrições e medidas de materiais), buscar materiais por aproximação e calcular quantidades necessárias com base em medi‑das.

---

## 🧭 Visão Geral

Este projeto usa os seguintes recursos:

---

### Automação de Correspondência de Planilhas com RapidFuzz

Este script Python foi desenvolvido para automatizar a busca e extração de dados entre duas planilhas do Excel. Diferente de um preenchimento direto, ele age como um motor de busca, que percorre uma planilha de referência (`file_referencia`), encontra a melhor correspondência em uma planilha de orçamento (`file_orcamento`) e, para cada item encontrado, cria um objeto estruturado com as informações correspondentes.

#### Como o Script Funciona: O Passo a Passo

1.  **Configuração Inicial**: O script começa importando as bibliotecas `pandas` (para leitura e manipulação das planilhas) e `rapidfuzz` (para o algoritmo de correspondência de texto). Ele também define os caminhos dos arquivos e trata possíveis erros caso as planilhas não sejam encontradas.
2.  **Preparação dos Dados**: As planilhas são lidas e carregadas em DataFrames do pandas. Em seguida, as descrições da **coluna B** de ambas as planilhas são extraídas em listas separadas. Essa etapa é crucial, pois essas listas são a base para a comparação de texto.
3.  **Processo de Correspondência (O "Match")**: Este é o coração do script. O código percorre cada descrição da sua planilha de referência e busca a correspondência mais próxima na lista de descrições da planilha de orçamento.
4.  **Criação dos Objetos**: Para cada correspondência que atende a um critério de similaridade, o script coleta os dados específicos (Descrição, Unidade de Medida, Valores, etc.) da linha correspondente na planilha de **referência** e armazena-os em um dicionário. Cada um desses dicionários é um "objeto" que contém as informações que você solicitou.
5.  **Resultado Final**: Ao final do processo, todos os objetos criados são exibidos de forma organizada na tela como um DataFrame do pandas, fornecendo uma visualização clara dos resultados da correspondência.

---

### Entendendo a Biblioteca RapidFuzz e o "Match"

A correspondência de texto, ou "match", é um desafio quando os textos não são idênticos. O RapidFuzz resolve isso com algoritmos inteligentes.

#### O Papel de `rapidfuzz.process.extractOne`

-   A função `process.extractOne` é a responsável por buscar a **melhor correspondência** (`one`) para uma string.
-   Ela precisa de três informações principais:
    1.  `descricao_ref`: A string que você quer encontrar (o item da planilha de referência).
    2.  `descricoes_orcamento`: A lista de strings onde você vai procurar (todas as descrições da planilha de orçamento).
    3.  `scorer`: O algoritmo de similaridade que será usado para a comparação.

#### Como Funciona o `fuzz.WRatio` (O Weighted Ratio)

No nosso script, usamos o **`fuzz.WRatio`**. Este não é um algoritmo simples; é uma pontuação heurística avançada que combina várias métricas de similaridade. Ele é projetado para lidar com strings de comprimentos e formatos diferentes de forma mais eficaz do que métricas mais simples.

O `WRatio` leva em consideração:

-   **Similaridade de sub-strings**: Ele pontua alto se uma string é uma sub-string de outra (ex: "Instalação de Válvula" e "Válvula").
-   **Similaridade de ordenação**: Ele penaliza menos por palavras fora de ordem.
-   **Tamanho da string**: Ele ajusta a pontuação para evitar que strings muito curtas e idênticas recebam pontuações artificialmente altas.

O resultado do `WRatio` é uma pontuação entre `0` e `100`, onde `100` significa uma correspondência perfeita.

#### O Papel do `MATCH_THRESHOLD`

O **`MATCH_THRESHOLD = 85`** é o nosso filtro de qualidade. Ele garante que o script só considere um "match" como válido se a pontuação de similaridade do `WRatio` for **igual ou superior a 85**. Isso evita que correspondências fracas ou incorretas sejam utilizadas, garantindo que apenas os resultados mais confiáveis sejam processados.

A combinação de `extractOne`, `fuzz.WRatio` e o `MATCH_THRESHOLD` é o que permite ao seu script encontrar e processar as correspondências de forma robusta e precisa, mesmo em cenários com erros de digitação, pequenas variações ou diferenças na formatação do texto.
[1]: https://linuxconfig.org/how-to-build-a-tkinter-application-using-an-object-oriented-approach?utm_source=chatgpt.com "How to build a Tkinter application using an object oriented approach"
[2]: https://www.geeksforgeeks.org/git/what-is-readme-md-file/?utm_source=chatgpt.com "What is README.md File? - GeeksforGeeks"
[3]: https://www.geeksforgeeks.org/python/python-gui-tkinter/?utm_source=chatgpt.com "Python Tkinter - GeeksforGeeks"
[4]: https://www.geeksforgeeks.org/python/python-tkinter-tutorial/?utm_source=chatgpt.com "Python Tkinter Tutorial - GeeksforGeeks"
[5]: https://www.reddit.com/r/Python/comments/xbgyov/make_your_tkinter_app_look_truly_modern_with_a/?utm_source=chatgpt.com "Make your Tkinter app look truly modern with a single line of code!"
