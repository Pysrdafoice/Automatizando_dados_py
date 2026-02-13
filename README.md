# 💰 Automação de Orçamento

**Data:** 12 de Fevereiro de 2026
**Versão:** 2.1
**Status:** Funcional com ordenação visual dinâmica

---

# 📑 Índice

* [Visão Geral](#-visão-geral)
* [Arquitetura do Sistema](#-arquitetura-do-sistema)
* [Fluxo Principal](#-fluxo-principal)
* [Detalhamento de Componentes](#-detalhamento-de-componentes)
* [Sistema de Ordenação (v2.1)](#-sistema-de-ordenação-v21)
* [Estrutura de Dados](#-estrutura-de-dados)
* [Fluxo de Integração entre Arquivos](#-fluxo-de-integração-entre-arquivos)
* [Processos Críticos](#-processos-críticos)
* [Estados e Transições](#-estados-e-transições)
* [Exemplo Completo de Execução](#-exemplo-completo-de-execução)
* [Tratamento de Erros](#-tratamento-de-erros)
* [Resumo de Funcionalidades](#-resumo-de-funcionalidades)
* [Conclusão](#-conclusão)
* [Histórico de Versões](#-histórico-de-versões)

---

# 🎯 Visão Geral

O sistema **Automação de Orçamento** é uma aplicação desktop desenvolvida com **Python + Tkinter** que automatiza o processo de correlação entre itens de orçamento e referências de preços.

Permite preenchimento automatizado com validação manual e atualização inteligente de planilhas Excel, preservando fórmulas existentes.

## Objetivos Principais

* ✅ Correlacionar itens de orçamento com referências de preço
* ✅ Exibir similaridade entre descrições
* ✅ Permitir seleção manual de referências
* ✅ Ordenar itens na grid superior (A→Z, por unidade, crescente/decrescente)
* ✅ Confirmar seleções com interface de check-in
* ✅ Atualizar planilha Excel com dados correlacionados
* ✅ Preservar fórmulas Excel durante atualização

---

# 🏗️ Arquitetura do Sistema

## Camadas da Aplicação

```
Interface do Usuário
(FormBuscaPlanilhas → TelaProcessamento → TelaCheckin)
        ↓
Processamento de Dados
(ProcessamentoBase → Correlação + Agrupamento)
        ↓
Persistência (I/O)
(Leitura Excel → Processamento → AtualizadorPlanilha)
        ↓
Configuração e Parâmetros
(ParametrosProcessamento → ParametrosPlanilhas)
```

## Stack Tecnológico

| Componente    | Tecnologia              | Uso                          |
| ------------- | ----------------------- | ---------------------------- |
| GUI           | Tkinter (ttk)           | Interface gráfica            |
| Processamento | Pandas                  | Leitura e manipulação Excel  |
| Similaridade  | difflib.SequenceMatcher | Correlação textual           |
| Excel Output  | openpyxl                | Escrita preservando fórmulas |
| Configuração  | dataclasses             | Estruturação tipada          |
| Versionamento | datetime                | Timestamp automático         |

---

# 🔄 Fluxo Principal

## 1️⃣ Inicialização

* `main.py` executa
* Abre `FormBuscaPlanilhas`

## 2️⃣ Coleta de Parâmetros

Usuário define:

* Planilha de referência
* Planilha de orçamento
* Colunas relevantes
* Intervalo de linhas
* Taxa mínima de similaridade

→ Gera `ParametrosProcessamento`

---

## 3️⃣ Processamento

Classe: `ProcessamentoBase`

### `processar_dados()`

1. Converte índices de coluna (A→0, B→1...)
2. Lê planilhas com Pandas
3. Filtra intervalo
4. Calcula similaridade entre descrições
5. Retorna lista de correlações

### Estrutura retornada:

```python
{
    "item": "Parafuso M10",
    "numero_linha": 5,
    "unidade": "UN",
    "referencia": "Parafuso Inox M10",
    "similaridade": 0.92,
    "valor_material": 12.50,
    "valor_mao_de_obra": 2.30,
    "valor_total": 14.80
}
```

---

## 4️⃣ Interface Principal — TelaProcessamento

### Layout

```
Cabeçalho Azul
Pesquisa
Grid Superior (Itens do orçamento)
Grid Inferior (Referências correlacionadas)
Botões: [Finalizar] [Prosseguir]
```

### Eventos principais

* `on_item_selecionado()`
* `on_referencia_selecionada()`
* `filtrar_itens()`
* `ordenar_grid_superior()` (v2.1)
* `prosseguir()`

---

# 🆕 Sistema de Ordenação (v2.1)

## Estado de Ordenação

```python
estado_ordenacao = {
    "coluna_ativa": None,
    "direcao": "asc"
}
```

## Função principal

```python
def ordenar_grid_superior(coluna: str):
```

### Regras

* Clique na mesma coluna → inverte direção
* Clique em coluna diferente → inicia ascendente
* Ordenação case-insensitive
* Apenas visual (não reprocessa dados)

### Tipos

* `"item"` → alfabético
* `"unidade"` → alfabético
* `"qty"` → numérico

---

# 📊 Estrutura de Dados

## Entrada

```python
ParametrosProcessamento
```

## Intermediário

Lista de dicts com correlações.

## Agrupado

```python
{
  "Item A": [{...}, {...}],
  "Item B": [{...}]
}
```

## Seleção do Usuário

```python
{
  "Item A": "Referência X"
}
```

## Confirmado

```python
List[ItemCheckin]
```

---

# 🔧 Detalhamento de Componentes

## FormBuscaPlanilhas

Responsável por coletar:

* Caminhos das planilhas
* Abas
* Colunas
* Intervalo
* Taxa de similaridade

---

## ProcessamentoBase

Responsável por:

* Ler planilhas
* Calcular similaridade
* Agrupar resultados

---

## ItemCheckin

```python
@dataclass
class ItemCheckin:
    item: str
    unidade: str
    referencia: str
    similaridade: float
    valor_total: float
    numero_linha: int
    valor_material: float
    valor_mao_de_obra: float
```

---

## TelaCheckin

Permite:

* Revisão dos itens
* Exclusão via double-click
* Resumo total
* Finalização do preenchimento

---

## AtualizadorPlanilha

Responsável por:

* Carregar Excel com openpyxl
* Atualizar células específicas
* Preservar fórmulas
* Gerar arquivo com timestamp

Exemplo de saída:

```
orc_PREENCHIDA_20260110_143025.xlsx
```

---

# ⚙️ Processos Críticos

## 1️⃣ Similaridade

```python
SequenceMatcher(None, a.lower(), b.lower()).ratio()
```

---

## 2️⃣ Conversão de Coluna Excel

```python
"A" → 0
"B" → 1
"AA" → 26
```

---

## 3️⃣ Mapeamento Pandas → Excel

```python
numero_linha = idx_orc + 3
```

---

## 4️⃣ Preservação de Fórmulas

Uso correto:

```python
wb = load_workbook(caminho)
ws.cell(row=5, column=3).value = 12.50
```

Fórmulas são mantidas intactas pelo openpyxl.

---

# 🔁 Estados e Transições

## TelaProcessamento

```
Inicial
↓
Grid Ordenada
↓
Direção Invertida
↓
Item Selecionado
↓
Referência Selecionada
↓
Abrir Checkin
```

## TelaCheckin

```
Exibição
↓
Exclusão (double-click)
↓
Finalizar
↓
Atualização Excel
↓
Finalizado
```

---

# 📈 Exemplo Completo de Execução

## Entrada

Planilha Referência:

| Descrição    | Material | MO   |
| ------------ | -------- | ---- |
| Parafuso M10 | 10.00    | 2.00 |

Planilha Orçamento:

| Item         | Unidade |
| ------------ | ------- |
| Parafuso M10 | UN      |

## Processamento

Similaridade ≥ 80%

## Seleção

Usuário escolhe referência.

## Resultado

Planilha preenchida automaticamente com:

| Item         | Un | Material | MO   | Total |
| ------------ | -- | -------- | ---- | ----- |
| Parafuso M10 | UN | 10.00    | 2.00 | 12.00 |

---

# 🛡 Tratamento de Erros

| Cenário                 | Ação          |
| ----------------------- | ------------- |
| Arquivo não encontrado  | Dialog erro   |
| Coluna inválida         | Debug + aviso |
| Planilha vazia          | Aviso         |
| Nenhuma correlação      | Aviso         |
| Sem seleções            | Bloqueio      |
| Cancelamento salvamento | Tratado       |

---

# 📊 Resumo de Funcionalidades

| Funcionalidade        | Status   |
| --------------------- | -------- |
| Seleção de arquivos   | ✅        |
| Processamento         | ✅        |
| Exibição em grids     | ✅        |
| Filtragem             | ✅        |
| Ordenação dinâmica    | ✅ (v2.1) |
| Exclusão com callback | ✅        |
| Atualização Excel     | ✅        |
| Preservação fórmulas  | ✅        |
| Versionamento arquivo | ✅        |

---

# 🎯 Conclusão

O sistema funciona como um pipeline estruturado:

```
Coleta → Processamento → Validação → Atualização Excel
```

A versão 2.1 introduz ordenação visual dinâmica, melhorando significativamente a experiência do usuário sem impactar o desempenho do processamento.

O uso de `@dataclass` torna o código mais seguro, legível e sustentável.

---

# 📝 Histórico de Versões

| Versão | Data       | Mudanças                             |
| ------ | ---------- | ------------------------------------ |
| 1.0    | 01/01/2026 | Processamento base                   |
| 2.0    | 10/01/2026 | Refatoração com ItemCheckin          |
| 2.1    | 12/02/2026 | Sistema de ordenação visual dinâmica |

---

**Autor:** Anderson
**Versão do Documento:** 2.1
**Última Atualização:** 12 de Fevereiro de 2026.
