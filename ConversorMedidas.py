from enum import Enum
from typing import Dict, Tuple

class GrandezaFisica(Enum):
    MASSA = "MASSA"
    VOLUME = "VOLUME"
    COMPRIMENTO = "COMPRIMENTO"
    AREA = "AREA"
    UNIDADE = "UNIDADE"
    NAO_IDENTIFICADA = "NAO_IDENTIFICADA"

class ConversorMedidas:
    def __init__(self):
        # Mapeamento de variações de escrita para unidades padrão
        self._variacoes_unidades: Dict[str, str] = {
            # Variações de MASSA
            "KG": "KG",
            "KILO": "KG",
            "KILOS": "KG",
            "QUILOS": "KG",
            "QUILO": "KG",
            "KILOGRAMAS": "KG",
            "QUILOGRAMAS": "KG",
            "KILOGRAM": "KG",
            "KILOGRAMA": "KG",
            "QUILOGRAMA": "KG",
            
            "G": "G",
            "GRAMA": "G",
            "GRAMAS": "G",
            "GRAM": "G",
            "GRAMS": "G",
            
            "MG": "MG",
            "MILIGRAMA": "MG",
            "MILIGRAMAS": "MG",
            "MILLIGRAM": "MG",
            "MILLIGRAMS": "MG",
            
            "T": "T",
            "TON": "T",
            "TONELADA": "T",
            "TONELADAS": "T",
            "TONS": "T",
            
            # Variações de VOLUME
            "M3": "M3",
            "M³": "M3",
            "METRO CUBICO": "M3",
            "METROS CUBICOS": "M3",
            "METRO CÚBICO": "M3",
            "METROS CÚBICOS": "M3",
            "CUBIC METER": "M3",
            "CUBIC METERS": "M3",
            
            "CM3": "CM3",
            "CM³": "CM3",
            "CENTIMETRO CUBICO": "CM3",
            "CENTIMETROS CUBICOS": "CM3",
            "CENTÍMETRO CÚBICO": "CM3",
            "CENTÍMETROS CÚBICOS": "CM3",
            
            "L": "L",
            "LITRO": "L",
            "LITROS": "L",
            "LITER": "L",
            "LITERS": "L",
            
            "ML": "ML",
            "MILILITRO": "ML",
            "MILILITROS": "ML",
            "MILLILITER": "ML",
            "MILLILITERS": "ML",
            
            # Variações de COMPRIMENTO
            "M": "M",
            "METRO": "M",
            "METROS": "M",
            "METER": "M",
            "METERS": "M",
            
            "CM": "CM",
            "CENTIMETRO": "CM",
            "CENTIMETROS": "CM",
            "CENTÍMETRO": "CM",
            "CENTÍMETROS": "CM",
            "CENTIMETER": "CM",
            "CENTIMETERS": "CM",
            
            "MM": "MM",
            "MILIMETRO": "MM",
            "MILIMETROS": "MM",
            "MILÍMETRO": "MM",
            "MILÍMETROS": "MM",
            "MILLIMETER": "MM",
            "MILLIMETERS": "MM",
            
            "KM": "KM",
            "QUILOMETRO": "KM",
            "QUILOMETROS": "KM",
            "QUILÔMETRO": "KM",
            "QUILÔMETROS": "KM",
            "KILOMETER": "KM",
            "KILOMETERS": "KM",
            
            # Variações de UNIDADE
            "UN": "UN",
            "UND": "UN",
            "UNID": "UN",
            "UNIDADE": "UN",
            "UNIDADES": "UN",
            "UNIT": "UN",
            "UNITS": "UN",

            # Variações de AREA
            "M2": "M2",
            "M²": "M2",
            "METRO QUADRADO": "M2",
            "METROS QUADRADOS": "M2",
            "METRO²": "M2",
            "METROS²": "M2",
            "SQUARE METER": "M2",
            "SQUARE METERS": "M2",

            "CM2": "CM2",
            "CM²": "CM2",
            "CENTIMETRO QUADRADO": "CM2",
            "CENTIMETROS QUADRADOS": "CM2",
            "CENTÍMETRO QUADRADO": "CM2",
            "CENTÍMETROS QUADRADOS": "CM2",
            "CENTIMETRO²": "CM2",
            "CENTIMETROS²": "CM2",
            "SQUARE CENTIMETER": "CM2",
            "SQUARE CENTIMETERS": "CM2",

            "MM2": "MM2",
            "MM²": "MM2",
            "MILIMETRO QUADRADO": "MM2",
            "MILIMETROS QUADRADOS": "MM2",
            "MILÍMETRO QUADRADO": "MM2",
            "MILÍMETROS QUADRADOS": "MM2",
            "MILIMETRO²": "MM2",
            "MILIMETROS²": "MM2",
            "SQUARE MILLIMETER": "MM2",
            "SQUARE MILLIMETERS": "MM2",

            "KM2": "KM2",
            "KM²": "KM2",
            "QUILOMETRO QUADRADO": "KM2",
            "QUILOMETROS QUADRADOS": "KM2",
            "QUILÔMETRO QUADRADO": "KM2",
            "QUILÔMETROS QUADRADOS": "KM2",
            "QUILOMETRO²": "KM2",
            "QUILOMETROS²": "KM2",
            "SQUARE KILOMETER": "KM2",
            "SQUARE KILOMETERS": "KM2",

            "HA": "HA",
            "HECTARE": "HA",
            "HECTARES": "HA",
            "HECTAR": "HA",
            "HECTARES": "HA"
        }

        # Fatores de conversão para MASSA (em relação ao kg)
        self._fatores_massa: Dict[str, float] = {
            "KG": 1.0,
            "G": 0.001,
            "MG": 0.000001,
            "T": 1000.0
        }

        # Fatores de conversão para VOLUME (em relação ao m³)
        self._fatores_volume: Dict[str, float] = {
            "M3": 1.0,
            "CM3": 0.000001,
            "L": 0.001,
            "ML": 0.000001
        }

        # Fatores de conversão para COMPRIMENTO (em relação ao metro)
        self._fatores_comprimento: Dict[str, float] = {
            "M": 1.0,
            "CM": 0.01,
            "MM": 0.001,
            "KM": 1000.0
        }

        # Fatores de conversão para AREA (em relação ao m²)
        self._fatores_area: Dict[str, float] = {
            "M2": 1.0,           # metro quadrado
            "CM2": 0.0001,       # centímetro quadrado (1m² = 10000cm²)
            "MM2": 0.000001,     # milímetro quadrado (1m² = 1000000mm²)
            "KM2": 1000000.0,    # quilômetro quadrado (1km² = 1000000m²)
            "HA": 10000.0        # hectare (1ha = 10000m²)
        }

        # Mapeamento de unidades para suas grandezas
        self._mapeamento_unidades: Dict[str, GrandezaFisica] = {
            # Unidades de massa
            "KG": GrandezaFisica.MASSA,
            "G": GrandezaFisica.MASSA,
            "MG": GrandezaFisica.MASSA,
            "T": GrandezaFisica.MASSA,
            # Unidades de volume
            "M3": GrandezaFisica.VOLUME,
            "CM3": GrandezaFisica.VOLUME,
            "L": GrandezaFisica.VOLUME,
            "ML": GrandezaFisica.VOLUME,
            # Unidades de comprimento
            "M": GrandezaFisica.COMPRIMENTO,
            "CM": GrandezaFisica.COMPRIMENTO,
            "MM": GrandezaFisica.COMPRIMENTO,
            "KM": GrandezaFisica.COMPRIMENTO,
            # Unidades de área
            "M2": GrandezaFisica.AREA,
            "CM2": GrandezaFisica.AREA,
            "MM2": GrandezaFisica.AREA,
            "KM2": GrandezaFisica.AREA,
            "HA": GrandezaFisica.AREA,
            # Unidades simples
            "UN": GrandezaFisica.UNIDADE,
            "UNID": GrandezaFisica.UNIDADE,
            "UND": GrandezaFisica.UNIDADE
        }

    def identificar_grandeza(self, unidade: str) -> GrandezaFisica:
        """
        Identifica a grandeza física de uma unidade de medida.
        
        Args:
            unidade: String contendo a unidade de medida
            
        Returns:
            GrandezaFisica correspondente à unidade ou NAO_IDENTIFICADA se não for reconhecida
        """
        unidade_formatada = self._formatar_unidade(unidade)
        return self._mapeamento_unidades.get(unidade_formatada, GrandezaFisica.NAO_IDENTIFICADA)

    def converter_medida(self, valor: float, unidade_origem: str, unidade_destino: str) -> Tuple[bool, float, str]:
        """
        Converte uma medida de uma unidade para outra.
        
        Args:
            valor: Valor numérico a ser convertido
            unidade_origem: Unidade de medida original
            unidade_destino: Unidade de medida desejada
            
        Returns:
            Tupla contendo:
            - bool: Sucesso da conversão
            - float: Valor convertido (ou valor original se falhou)
            - str: Mensagem de erro (vazia se sucesso)
        """
        unidade_origem = self._formatar_unidade(unidade_origem)
        unidade_destino = self._formatar_unidade(unidade_destino)

        grandeza_origem = self.identificar_grandeza(unidade_origem)
        grandeza_destino = self.identificar_grandeza(unidade_destino)

        # Verifica se as unidades são da mesma grandeza
        if grandeza_origem != grandeza_destino:
            return False, valor, "Unidades de grandezas diferentes não podem ser convertidas"

        # Verifica se a grandeza é UNIDADE (não precisa converter)
        if grandeza_origem == GrandezaFisica.UNIDADE:
            return True, valor, ""

        # Verifica se a grandeza é não identificada
        if grandeza_origem == GrandezaFisica.NAO_IDENTIFICADA:
            return False, valor, "Unidade não reconhecida"

        try:
            if grandeza_origem == GrandezaFisica.MASSA:
                fatores = self._fatores_massa
            elif grandeza_origem == GrandezaFisica.VOLUME:
                fatores = self._fatores_volume
            elif grandeza_origem == GrandezaFisica.COMPRIMENTO:
                fatores = self._fatores_comprimento
            else:  # AREA
                fatores = self._fatores_area

            # Converte para a unidade base e depois para a unidade de destino
            valor_base = valor * fatores[unidade_origem]
            valor_convertido = valor_base / fatores[unidade_destino]

            return True, valor_convertido, ""

        except KeyError:
            return False, valor, "Unidade não suportada para conversão"
        except Exception as e:
            return False, valor, f"Erro na conversão: {str(e)}"

    def _formatar_unidade(self, unidade: str) -> str:
        """
        Formata a unidade de medida para um formato padrão, considerando as variações.
        
        Args:
            unidade: String contendo a unidade de medida
            
        Returns:
            String formatada em maiúsculas e sem espaços, convertida para a unidade padrão
        """
        # Primeiro, formata a string removendo espaços extras e convertendo para maiúsculas
        unidade_formatada = unidade.upper().strip()
        
        # Procura no dicionário de variações
        return self._variacoes_unidades.get(unidade_formatada, unidade_formatada)

