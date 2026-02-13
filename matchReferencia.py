from dataclasses import dataclass

@dataclass
class MatchReferencia:
    descricao: str
    unidadeMedida: str
    valorMedida : float
    valorMaterial: float
    ValorMaoDeObra: float
    linhaPlanilhaRefencia: int
    score: float