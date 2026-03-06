from mappings import (
    MAPA_CARGOS_METAX, MAPA_CARGOS_CODFUNCAO_METAX, MAPA_ESCOLARIDADE,
    MAPA_ESTADO_CIVIL, MAPA_ESTADO_NATAL, MAPA_SEXO, MAPA_CARGOS_OVERRIDE_POR_CONTRATO
)

def test_mapa_cargos_nao_vazio():
    assert len(MAPA_CARGOS_METAX) > 0
    assert isinstance(MAPA_CARGOS_METAX, dict)
    assert len(MAPA_CARGOS_CODFUNCAO_METAX) > 0
    assert isinstance(MAPA_CARGOS_CODFUNCAO_METAX, dict)

def test_mapa_escolaridade_conteudo():
    # Testar algumas chaves conhecidas (exemplo genérico, ajuste conforme seu mappings real)
    # Se mappings.py tiver '1': 'ANALFABETO'
    # assert '1' in MAPA_ESCOLARIDADE
    assert isinstance(MAPA_ESCOLARIDADE, dict)

def test_mapa_chaves_sao_strings():
    for k in MAPA_CARGOS_METAX.keys():
        assert isinstance(k, str)
    for k in MAPA_CARGOS_CODFUNCAO_METAX.keys():
        assert isinstance(k, str)

def test_mapa_valores_sao_strings():
    for v in MAPA_CARGOS_METAX.values():
        assert isinstance(v, str)
    for v in MAPA_CARGOS_CODFUNCAO_METAX.values():
        assert isinstance(v, str)

def test_overrides_por_contrato_sao_dicts_de_strings():
    assert isinstance(MAPA_CARGOS_OVERRIDE_POR_CONTRATO, dict)
    for contrato, overrides in MAPA_CARGOS_OVERRIDE_POR_CONTRATO.items():
        assert isinstance(contrato, str)
        assert isinstance(overrides, dict)
        for origem, destino in overrides.items():
            assert isinstance(origem, str)
            assert isinstance(destino, str)

def test_mapa_estados_brasil():
    assert "SP" in MAPA_ESTADO_NATAL or "SÃO PAULO" in MAPA_ESTADO_NATAL.values()
