# Runner MetaXg

## Objetivo
Executar o aplicativo com atualizacao automatica, seguindo o padrao de releases na rede e fallback via GitHub.

## Arquivos
- `runner.py`
- `config.json`

## Fluxo
1. Le `version.txt` (default 0.0.0) em `install_dir` (default local `.\_install`).
2. Busca `latest.json` na rede.
3. Se falhar, consulta GitHub Releases (se configurado).
4. Compara versoes (semver).
5. Se houver update:
   - baixa ZIP e SHA256
   - valida SHA256 (obrigatorio)
   - extrai em staging
   - move current -> backup
   - move staging -> current
   - atualiza `version.txt`
6. Executa `exe_name` dentro de `app\current`.

## Schema do latest.json (ASO)
```json
{
  "version": "1.2.3",
  "package_filename": "MetaXg_1.2.3.zip",
  "sha256_filename": "MetaXg_1.2.3.sha256"
}
```

## Config.json (runner)
```json
{
  "app_name": "MetaXg",
  "install_dir": ".\\_install",
  "network_release_dir": ".\\releases",
  "network_latest_json": ".\\releases\\latest.json",
  "github_repo": "GuilhermeBatistaenesa/MetaXg",
  "exe_name": "MetaXg.exe",
  "log_file": ".\\_install\\logs\\metax_last_run.log",
  "prefer_network": false,
  "allow_prerelease": false,
  "run_args": [],
  "log_level": "INFO"
}
```

## Staging / Backup / Rollback
- `app\staging`: recebe o zip extraido.
- `app\backup`: recebe o `current` anterior antes do swap.
- Rollback automatico se falhar apos mexer em `current`.

## Logs
- `metax_last_run.log` em `install_dir\logs`.

## Variaveis de ambiente (opcional)
- `METAX_INSTALL_DIR`
- `METAX_NETWORK_RELEASE_DIR`
- `METAX_NETWORK_LATEST_JSON`
- `METAX_LOG_FILE`

## GitHub fallback (opcional)
- Se `github_repo` estiver vazio/ausente, o runner nao tenta GitHub.
