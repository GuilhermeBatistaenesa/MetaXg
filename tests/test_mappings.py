from mappings import (
    MAPA_CARGOS_METAX, MAPA_ESCOLARIDADE, 
    MAPA_ESTADO_CIVIL, MAPA_ESTADO_NATAL, MAPA_SEXO
)

def test_mapa_cargos_nao_vazio():
    assert len(MAPA_CARGOS_METAX) > 0
    assert isinstance(MAPA_CARGOS_METAX, dict)

def test_mapa_escolaridade_conteudo():
    # Testar algumas chaves conhecidas (exemplo genérico, ajuste conforme seu mappings real)
    # Se mappings.py tiver '1': 'ANALFABETO'
    # assert '1' in MAPA_ESCOLARIDADE
    assert isinstance(MAPA_ESCOLARIDADE, dict)

def test_mapa_chaves_sao_strings():
    for k in MAPA_CARGOS_METAX.keys():
        assert isinstance(k, str)

def test_mapa_valores_sao_strings():
    for v in MAPA_CARGOS_METAX.values():
        assert isinstance(v, str)

def test_mapa_estados_brasil():
    assert "SP" in MAPA_ESTADO_NATAL or "SÃO PAULO" in MAPA_ESTADO_NATAL.values()
