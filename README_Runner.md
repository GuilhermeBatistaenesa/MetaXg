# Runner MetaXg

## Objetivo
Executar o aplicativo com atualizacao automatica, seguindo o padrao de releases na rede e fallback via GitHub.

## Arquivos
- `runner.py`
- `config.json`

## Fluxo
1. Le `version.txt` (default 0.0.0) em `C:\MetaXg`.
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
  "install_dir": "C:\\MetaXg",
  "network_release_dir": "P:\\ProcessoMetaX\\releases",
  "network_latest_json": "P:\\ProcessoMetaX\\releases\\latest.json",
  "github_repo": "GuilhermeBatistaenesa/MetaXg",
  "exe_name": "MetaXg.exe",
  "log_file": "C:\\MetaXg\\logs\\metax_last_run.log",
  "prefer_network": true,
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
- `metax_last_run.log` em `C:\MetaXg\logs`.

## GitHub fallback (opcional)
- Se `github_repo` estiver vazio/ausente, o runner nao tenta GitHub.
