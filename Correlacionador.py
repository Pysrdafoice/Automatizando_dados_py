import pandas as pd
from difflib import SequenceMatcher
from ParametrosProcessamento import ParametrosProcessamento
from Correlacao import Correlacao, ResultadoCorrelacao
from ConversorMedidas import ConversorMedidas

class Correlacionador:
    def __init__(self, parametros: ParametrosProcessamento):
        self.parametros = parametros
        self.planilha_orcamento = pd.read_excel(parametros.orcamento.caminho_planilha)
        self.planilha_referencia = pd.read_excel(parametros.referencia.caminho_planilha)
        self.conversor_medidas = ConversorMedidas()
        self.indices = {
            'orcamento': {
                'descricao': self.transformar_indice_coluna(parametros.orcamento.coluna_descrição),
                'unidade': self.transformar_indice_coluna(parametros.orcamento.coluna_unidade_medida)
            },
            'referencia': {
                'descricao': self.transformar_indice_coluna(parametros.referencia.coluna_descrição),
                'unidade': self.transformar_indice_coluna(parametros.referencia.coluna_unidade_medida),
                'material': self.transformar_indice_coluna(parametros.referencia.coluna_material),
                'mao_obra': self.transformar_indice_coluna(parametros.referencia.coluna_mao_de_obra)
            }
        }
    
    @staticmethod
    def similaridade(a: str, b: str) -> float:
        return SequenceMatcher(None, a.lower(), b.lower()).ratio()

    @staticmethod
    def transformar_indice_coluna(coluna: str) -> int:
        coluna = coluna.upper()
        indice = 0
        for char in coluna:
            indice = indice * 26 + (ord(char) - ord('A') + 1)
        return indice - 1

    def validar_linha_orcamento(self, linha_orcamento) -> bool:
        indice_descricao = self.transformar_indice_coluna(self.parametros.orcamento.coluna_descrição)
        indice_unidade = self.transformar_indice_coluna(self.parametros.orcamento.coluna_unidade_medida)

        descricao_valida = not pd.isna(linha_orcamento.iloc[indice_descricao]) and str(linha_orcamento.iloc[indice_descricao]).strip() != ""
        unidade_valida = not pd.isna(linha_orcamento.iloc[indice_unidade]) and str(linha_orcamento.iloc[indice_unidade]).strip() != ""
        
        return descricao_valida and unidade_valida

    def validar_linha_referencia(self, linha_orcamento, linha_referencia) -> bool:
        # Validar se há valores válidos (material ou mão de obra)
        valor_material = linha_referencia.iloc[self.indices['referencia']['material']]
        valor_mao_obra = linha_referencia.iloc[self.indices['referencia']['mao_obra']]
        
        tem_valor_valido = (
            (not pd.isna(valor_material) and float(valor_material) > 0) or
            (not pd.isna(valor_mao_obra) and float(valor_mao_obra) > 0)
        )
        
        if not tem_valor_valido:
            return False

        # Validar compatibilidade das unidades de medida
        try:
            unidade_orcamento = str(linha_orcamento.iloc[self.indices['orcamento']['unidade']]).strip()
            unidade_referencia = str(linha_referencia.iloc[self.indices['referencia']['unidade']]).strip()
            
            grandeza_orcamento = self.conversor_medidas.identificar_grandeza(unidade_orcamento)
            grandeza_referencia = self.conversor_medidas.identificar_grandeza(unidade_referencia)
            
            return grandeza_orcamento == grandeza_referencia and grandeza_orcamento != self.conversor_medidas.GrandezaFisica.NAO_IDENTIFICADA
            
        except Exception:
            return False

    def buscar_correlacoes(self) -> list[Correlacao]:
        correlacoes = []
        indices_colunas = {
            'orcamento': {
                'descricao': self.transformar_indice_coluna(self.parametros.orcamento.coluna_descrição),
                'unidade': self.transformar_indice_coluna(self.parametros.orcamento.coluna_unidade_medida)
            },
            'referencia': {
                'descricao': self.transformar_indice_coluna(self.parametros.referencia.coluna_descrição),
                'unidade': self.transformar_indice_coluna(self.parametros.referencia.coluna_unidade_medida),
                'material': self.transformar_indice_coluna(self.parametros.referencia.coluna_material),
                'mao_obra': self.transformar_indice_coluna(self.parametros.referencia.coluna_mao_de_obra)
            }
        }

        linhas_filtradas = self.planilha_orcamento[
            (self.planilha_orcamento.index + 2 >= self.parametros.pesquisa.ComecoPesquisa) &
            (self.planilha_orcamento.index + 2 <= self.parametros.pesquisa.TerminoPesquisa)
        ]

        for idx_orc, linha_orc in linhas_filtradas.iterrows():
            if not self.validar_linha_orcamento(linha_orc):
                continue

            descricao_orcamento = str(linha_orc.iloc[indices_colunas['orcamento']['descricao']])
            resultados_encontrados = []

            for idx_ref, linha_ref in self.planilha_referencia.iterrows():
                if not self.validar_linha_referencia(linha_orc, linha_ref):
                    continue

                descricao_referencia = str(linha_ref.iloc[indices_colunas['referencia']['descricao']])
                score = self.similaridade(descricao_orcamento, descricao_referencia)

                if score > self.parametros.pesquisa.TaxaSimilaridade:
                    resultado = ResultadoCorrelacao(
                        numeroLinha=idx_ref + 2,
                        descricao=descricao_referencia,
                        unidadeMedida=str(linha_ref.iloc[indices_colunas['referencia']['unidade']]),
                        valorMaterial=float(linha_ref.iloc[indices_colunas['referencia']['material']]),
                        valorMaoDeObra=float(linha_ref.iloc[indices_colunas['referencia']['mao_obra']]),
                        score=score
                    )
                    resultados_encontrados.append(resultado)

            if resultados_encontrados:
                correlacao = Correlacao(
                    numeroLinha=idx_orc + 2,
                    descricao=descricao_orcamento,
                    resultados=resultados_encontrados
                )
                correlacoes.append(correlacao)

        return correlacoes
