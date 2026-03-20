# Contrato de Artefatos - ProcessoMetaX

## Convencoes de nomes
- Log: `execution_<YYYY-MM-DD_HH-mm-ss>__<execution_id>.jsonl`
- Manifest final: `YYYY-MM-DD__HHhMM__<status>__manifest.json`
- Manifest parcial: `YYYY-MM-DD__HHhMM__<status>__manifest_parcial.json`
- Relatorio TXT: `YYYY-MM-DD__HHhMM__<status>__relatorio_execucao.txt`
- Resumo MD: `YYYY-MM-DD__HHhMM__<status>__resumo_execucao.md`
- Relatorio JSON: `YYYY-MM-DD__HHhMM__<status>__relatorio_resumo.json`
- PDF de pendencia: `YYYY-MM-DD__HHhMM__pendencia_<contrato>_<indice>_<nome>_<chapa>.pdf`

## Pastas publicas
- `P:\ProcessoMetaX\logs`
- `P:\ProcessoMetaX\relatorios`
- `P:\ProcessoMetaX\json`
- `P:\ProcessoMetaX\screenshots`
- `P:\ProcessoMetaX\fotos\processados`
- `P:\ProcessoMetaX\fotos\erros`

## Campos minimos de log
- `timestamp`
- `execution_id`
- `run_status`
- `level`
- `event`
- `message`
- `details`

## Classes de erro
- `BUSINESS_VALIDATION`
- `EXTERNAL_DEPENDENCY`
- `INTEGRATION`
- `DATA_QUALITY`
- `INFRA_IO`
- `UNEXPECTED`
