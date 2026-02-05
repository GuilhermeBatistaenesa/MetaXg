# Documentacao Oficial - MetaXg

## Visao geral
MetaXg e um robo RPA para integracao entre RM Labore e portal MetaX, com padrao corporativo de operacao, releases e auditoria.

## Pre-requisitos (Windows alvo)
- ODBC Driver 17/18 precisa estar instalado.
- (Opcional) Outlook Desktop instalado para envio de e-mail (usa `pywin32`).

## Modos de execucao
1. Execucao direta do app (`main.py`) em ambiente Python.
2. Execucao via Runner (`runner.py`), com update automatico e fallback GitHub.

## Estrutura de pastas na rede (hub)
Base: `P:\ProcessoMetaX`

Pastas operacionais:
- `Codigo\`
- `em processamento\`
- `processados\`
- `erros\`
- `logs\` (alias)
- `relatorios\` (alias)
- `json\` (alias)
- `07_LOGS\` (padrão legado por data)
- `08_RELATORIOS\` (padrão legado por data)
- `09_JSON\` (padrão legado por data)
- `10_SCREENSHOTS\` (padrão legado por data)
- `releases\`

## Logs e evidencias
Local:
- `logs\` (logs gerais)
- `logs\screenshots\` (evidencias)
- `relatorios\` (relatorios TXT)
- `json\` (manifest e JSON tecnico)

Padrao adicional (em `logs\`):
- `diagnostico_ultima_execucao.txt`
- `relatorio_<data>__<execid>.json`
- `resumo_execucao_<data>__<execid>.md`
- `metax_last_run.log` (runner)

Rede:
Os mesmos artefatos sao gravados em `P:\ProcessoMetaX\07_LOGS`, `08_RELATORIOS`, `09_JSON`, `10_SCREENSHOTS` (padrao legado por data) e em alias `P:\ProcessoMetaX\logs`, `relatorios`, `json`, `logs\screenshots`.

Fotos (pipeline):
- Entrada e processamento: `P:\ProcessoMetaX\em processamento`
- Concluidas/ok: `P:\ProcessoMetaX\processados`
- Com erro: `P:\ProcessoMetaX\erros`
- Observacao: fotos em `processados` e `erros` ficam em subpastas por data (YYYY-MM-DD).

## Runner: atualizacao segura
Fluxo:
1. Le `version.txt` em `C:\MetaXg`.
2. Busca `latest.json` em `P:\ProcessoMetaX\releases`.
3. Se falhar, usa GitHub Releases.
4. Valida SHA256.
5. Extrai em staging e faz swap com backup.
6. Executa o exe dentro de `C:\MetaXg\app\current`.

Rollback:
- Automatico se falhar apos mover `current`.
- Backup mantido em `C:\MetaXg\app\backup`.

Schema do latest.json (ASO):
```json
{
  "version": "1.0.0",
  "package_filename": "MetaXg_1.0.0.zip",
  "sha256_filename": "MetaXg_1.0.0.sha256"
}
```

GitHub fallback (opcional):
- Se `github_repo` estiver vazio/ausente, o runner nao tenta GitHub.

Config (runner):
- `install_dir`
- `prefer_network`
- `allow_prerelease`
- `run_args`
- `log_level`

## Processo de release
1. Build onedir do executavel:
   - `build_windows.bat`
2. Gerar zip + sha256 + latest.json:
   - `build_zip.bat -Bump patch`
3. Publicar em `P:\ProcessoMetaX\releases`:
- `MetaXg_<versao>.zip`
- `MetaXg_<versao>.sha256`
- `latest.json`

## Troubleshooting
- Rede indisponivel: runner executa a versao atual.
- SHA256 divergente: update abortado.
- Exe ausente em current: runner registra erro e retorna codigo 1.
- Falha ao escrever logs na rede: execucao segue, com alerta no diagnostico.

## Checklist de validacao
- Rede disponivel: update pela rede.
- Rede indisponivel: fallback GitHub.
- SHA invalido: update abortado e executa versao atual.
- Rollback acionado: backup restaurado.
- Exe em `current` executado.
