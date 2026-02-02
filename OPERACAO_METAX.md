# Operação diária - MetaXg

## 1) Atualizar repo
```powershell
cd C:\Users\guilherme.batista\OneDrive\MetaXg
git pull
```

## 2) Rodar
### Dev (Python)
```powershell
. .\.venv\Scripts\Activate.ps1
python main.py
```

### Executável
```powershell
.\dist\MetaXg\MetaXg.exe
```

## 3) Resolver CAPTCHA
- O sistema vai abrir o navegador.
- Resolva o CAPTCHA e finalize o login quando solicitado.

## 4) Conferir resultados
- E-mail automático (se habilitado)
- Relatório em `relatorios/`
- Manifest em `json/`

## 5) Evidências quando falha verificação
- Screenshot: `logs/screenshots/verify_fail_<cpf>_<timestamp>__<execid>.png`
- Debug JSON: `json/verify_debug_<cpf>_<timestamp>__<execid>.json`

## 6) Modo TXT (lista manual)
- Arquivo: `P:\ProcessoMetaX\em processamento\cadastrar_metax.txt`
- Um nome por linha; comentários com `#`
- Se vazio, roda SQL padrão.
