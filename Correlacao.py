from dataclasses import dataclass

@dataclass
class ResultadoCorrelacao:
    NumeroLinhaReferencia: int
    Descricao: str
    UnidadeMedida: str
    ValorMaterial: float
    ValorMaoDeObra: float
    Score: float

    def __init__(self, numeroLinha: int, descricao: str, unidadeMedida: str, valorMaterial: float, valorMaoDeObra: float, score: float):
        if numeroLinha <= 0:
            raise ValueError("O número da linha de referência deve ser maior que zero")
        if not descricao or descricao.strip() == "":
            raise ValueError("A descrição não pode ser vazia")
        if not unidadeMedida or unidadeMedida.strip() == "":
            raise ValueError("A unidade de medida não pode ser vazia")
        if score < 0 or score > 1:
            raise ValueError("O score deve estar entre 0 e 1")
        if valorMaterial <= 0 and valorMaoDeObra <= 0:
            raise ValueError("Pelo menos um dos valores (material ou mão de obra) deve ser maior que zero")
        
        self.NumeroLinhaReferencia = numeroLinha
        self.Descricao = descricao
        self.UnidadeMedida = unidadeMedida
        self.ValorMaterial = valorMaterial
        self.ValorMaoDeObra = valorMaoDeObra
        self.Score = score

@dataclass
class Correlacao:
    NumeroLinhaOrcamento: int
    DescricaoItemOrcamento: str
    ResultadosEncontrados: list[ResultadoCorrelacao]

    def __init__(self, numeroLinha: int, descricao: str, resultados: list[ResultadoCorrelacao] | None = None):
        if numeroLinha <= 0:
            raise ValueError("O número da linha do orçamento deve ser maior que zero")
        if not descricao or descricao.strip() == "":
            raise ValueError("A descrição do item do orçamento não pode ser vazia")
        
        self.NumeroLinhaOrcamento = numeroLinha
        self.DescricaoItemOrcamento = descricao
        self.ResultadosEncontrados = resultados if resultados is not None else []
    