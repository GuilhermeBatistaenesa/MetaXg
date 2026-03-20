# Runner MetaXg

## Objetivo
Executar o aplicativo com atualizacao automatica e segura a partir de `P:\ProcessoMetaX\releases`, com fallback opcional via GitHub.

## Arquivos
- `runner.py`
- `config.json`

## Fluxo
1. Ler `version.txt` em `install_dir`
2. Buscar `latest.json` na rede
3. Se configurado, tentar GitHub Releases como fallback
4. Comparar versao
5. Baixar ZIP e SHA256
6. Validar integridade
7. Extrair em staging
8. Trocar `current` por `backup`
9. Executar `MetaXg.exe`

## Release padrao
Artefatos esperados:
- `MetaXg_<versao>.zip`
- `MetaXg_<versao>.sha256`
- `latest.json`

## latest.json
```json
{
  "version": "1.2.3",
  "package_filename": "MetaXg_1.2.3.zip",
  "sha256_filename": "MetaXg_1.2.3.sha256"
}
```

## Configuracoes principais
- `install_dir`
- `network_release_dir`
- `network_latest_json`
- `github_repo`
- `exe_name`
- `prefer_network`
- `allow_prerelease`
- `run_args`
- `log_level`

## Comportamento importante
- Se a rede falhar, o runner tenta continuar com a versao atual
- Se o SHA256 divergir, o update aborta
- Se o swap falhar, o rollback e automatico

## Ponto de operacao
O runner atualiza o aplicativo. Os artefatos operacionais do robo continuam indo para `P:\ProcessoMetaX`.
