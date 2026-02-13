from dataclasses import dataclass
from OperacaoCorrelacao import OperacaoCorrelacao
from parametrosPlanilha import ParametrosPlanilhas

@dataclass
class ParametrosProcessamento:
    referencia: ParametrosPlanilhas
    orcamento: ParametrosPlanilhas
    pesquisa: OperacaoCorrelacao
    aba_pesquisa: str = None  # Armazena a aba selecionada para pesquisa