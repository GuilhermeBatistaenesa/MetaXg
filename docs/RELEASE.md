# Release MetaXg

## Modo dev
```powershell
python -m venv .venv
. .\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python -m playwright install chromium
python main.py
```

## Modo executavel
```powershell
build_windows.bat
```

Executavel gerado:
- `dist\MetaXg\MetaXg.exe`

## Execucao do EXE
- O `.env` deve ficar acessivel para o executavel
- ODBC Driver 17 ou 18 precisa existir na maquina
- Exemplo:
```powershell
.\dist\MetaXg\MetaXg.exe --no-email
```

## Artefatos operacionais
O executavel gera e publica artefatos em:
- `P:\ProcessoMetaX\logs`
- `P:\ProcessoMetaX\relatorios`
- `P:\ProcessoMetaX\json`
- `P:\ProcessoMetaX\screenshots`

## Convencao de nomes
- `AAAA-MM-DD__HHhMM__status__relatorio_execucao.txt`
- `AAAA-MM-DD__HHhMM__status__manifest.json`
- `AAAA-MM-DD__HHhMM__status__manifest_parcial.json`
- `AAAA-MM-DD__HHhMM__status__resumo_execucao.md`

## Build de release
```powershell
build_zip.bat -Bump patch
```

Saida padrao:
- `P:\ProcessoMetaX\releases\MetaXg_<versao>.zip`
- `P:\ProcessoMetaX\releases\MetaXg_<versao>.sha256`
- `P:\ProcessoMetaX\releases\latest.json`
