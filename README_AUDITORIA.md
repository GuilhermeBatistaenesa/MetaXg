# Auditoria Excel - MetaX

Este projeto grava automaticamente as execucoes do MetaX em um Excel central de auditoria.

## Arquivos
- `auditoria_excel.py`: modulo reutilizavel com `log_run(...)`.
- `P:\AuditoriaRobos\Auditoria_Robos.xlsx`: Excel central (criado automaticamente se nao existir).
- `P:\AuditoriaRobos\pending`: fallback quando o Excel estiver aberto.

## Integracao no MetaX
A chamada ja esta no final do fluxo em `main.py` e roda dentro do `finally`.

Campos preenchidos automaticamente:
- `run_id`, `robo_nome`, `origem_codigo`, `usuario_execucao`, `host_maquina`, datas/horas e duracao.
- `versao_robo` usa `version.txt` quando disponivel (fallback `1.0.0`).

Campos enviados pelo MetaX:
- `total_processado`, `total_sucesso`, `total_erro`
- `erros_auto_mitigados`, `erros_manuais`
- `ambiente` via `METAX_ENV` (default `PROD`)
- `observacoes` (ex.: skipped)

## Fallback (arquivo aberto)
Se o Excel central estiver aberto e nao for possivel salvar:
- Cria `P:\AuditoriaRobos\pending\Auditoria_Robos__PENDENTE__<timestamp>.xlsx`
- Salva JSON com detalhes: `P:\AuditoriaRobos\pending\run__<run_id>.json`

## Dependencia
- `openpyxl` foi adicionado em `requirements.txt`.

## Uso direto (opcional)
```python
from auditoria_excel import log_run

log_run(
    {
        "run_id": "...",
        "total_processado": 10,
        "total_sucesso": 9,
        "total_erro": 1,
        "erros_auto_mitigados": 0,
        "erros_manuais": 0,
        "ambiente": "PROD",
        "observacoes": ""
    },
    errors=[
        {"etapa": "Cadastro", "tipo_erro": "Tecnico", "mensagem_resumida": "Falha ..."}
    ]
)
```
