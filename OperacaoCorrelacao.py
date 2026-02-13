from dataclasses import dataclass

@dataclass
class OperacaoCorrelacao:
    ComecoPesquisa: int
    TerminoPesquisa: int
    TaxaSimilaridade: float
