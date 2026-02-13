from dataclasses import dataclass

@dataclass
class ParametrosPlanilhasPesquisa:
    comeco_pesquisa: int
    termino_pesquisa: int
    taxaSimilaridade: float = 0.8
