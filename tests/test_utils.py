import pytest
from datetime import date, datetime
from utils import (
    normalizar_texto, formatar_cpf, formatar_pis, 
    formatar_data, formatar_telefone_numerico
)

# TESTES DE NORMALIZAÇÃO DE TEXTO
def test_normalizar_texto_basico():
    assert normalizar_texto("João da Silva") == "JOAO DA SILVA"

def test_normalizar_texto_acentos_complexos():
    assert normalizar_texto("Áéíóú Çãõü") == "AEIOU CAOU"

def test_normalizar_texto_espacos():
    assert normalizar_texto("  texto  com  espaço  ") == "TEXTO  COM  ESPACO"

def test_normalizar_texto_none():
    assert normalizar_texto(None) == ""

# TESTES DE FORMATAÇÃO DE CPF
def test_formatar_cpf_limpo():
    assert formatar_cpf("12345678901") == "123.456.789-01"

def test_formatar_cpf_com_pontuacao():
    assert formatar_cpf("123.456.789-01") == "123.456.789-01"

def test_formatar_cpf_invalido_tamanho():
    with pytest.raises(ValueError):
        formatar_cpf("123")

def test_formatar_cpf_none():
    assert formatar_cpf(None) == ""

# TESTES DE FORMATAÇÃO DE PIS
def test_formatar_pis_limpo():
    assert formatar_pis("12345678901") == "12345678901"

def test_formatar_pis_com_pontuacao():
    assert formatar_pis("123.45678.90-1") == "12345678901"

def test_formatar_pis_menor_com_zero_esquerda():
    # Se vier com 10 digitos, deve adicionar 0
    assert formatar_pis("1234567890") == "01234567890"

def test_formatar_pis_invalido():
    with pytest.raises(ValueError):
        formatar_pis("123")

# TESTES DE FORMATAÇÃO DE DATA
def test_formatar_data_string_iso():
    assert formatar_data("2000-12-31") == "31/12/2000"

def test_formatar_data_objeto_date():
    assert formatar_data(date(2023, 1, 1)) == "01/01/2023"

def test_formatar_data_objeto_datetime():
    assert formatar_data(datetime(2023, 1, 1, 14, 30)) == "01/01/2023"

def test_formatar_data_invalida():
    with pytest.raises(ValueError):
        formatar_data("31/12/2000") # Espera formato ISO YYYY-MM-DD se for string

def test_formatar_data_none():
    assert formatar_data(None) == ""

# TESTES DE TELEFONE
def test_formatar_telefone_limpo():
    assert formatar_telefone_numerico("(11) 98765-4321") == "11987654321"

def test_formatar_telefone_curto():
    assert formatar_telefone_numerico("123") == ""

def test_formatar_telefone_none():
    assert formatar_telefone_numerico(None) == ""
