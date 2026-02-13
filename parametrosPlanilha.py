from dataclasses import dataclass

@dataclass
class ParametrosPlanilhas:
    caminho_planilha: str
    aba: str
    coluna_descrição: str
    coluna_material: str
    coluna_mao_de_obra: str
    coluna_unidade_medida: str
