# Release MetaXg

Este documento descreve como executar em modo dev e como gerar/rodar executáveis portáteis.

## 1) Modo dev (Python/venv)
1. Abra o PowerShell na raiz do repo:
   ```powershell
   cd C:\Users\guilherme.batista\OneDrive\MetaXg
   ```
2. (Opcional) Use o script de setup:
   ```powershell
   .\scripts\setup_windows.ps1
   ```
3. Ative o venv e instale deps (se não usou o script):
   ```powershell
   python -m venv .venv
   . .\.venv\Scripts\Activate.ps1
   pip install -r requirements.txt
   python -m playwright install chromium
   ```
4. Rode em modo dev:
   ```powershell
   python main.py
   ```

## 2) Modo executável (PyInstaller)
1. Instale o PyInstaller (se necessário):
   ```powershell
   pip install pyinstaller
   ```
2. Gere os executáveis:
   ```powershell
   pyinstaller build\metaxg.spec --clean
   ```
3. Executáveis gerados em:
   ```
   dist\MetaXg.exe
   dist\MetaXg_debug.exe
   ```

## 3) Execução do EXE
- O executável procura `.env` **na mesma pasta do exe**. Copie seu `.env` para `dist\`.
- Exemplo de execução:
  ```powershell
  .\dist\MetaXg.exe
  ```

### CLI args suportados
```powershell
--txt <caminho>     # ativa modo TXT
--dry-run           # não envia e-mail (gera relatórios)
--no-email          # não envia e-mail
--headless          # tenta rodar sem UI
--log-level INFO|DEBUG|WARN|ERROR
```

Exemplo (modo TXT + sem email):
```powershell
.\dist\MetaXg.exe --txt P:\\GuilhermeCostaProenca\\04_AUTOMACOES\\MetaXg\\inputs\\cadastrar_metax.txt --no-email
```

## 4) Artefatos gerados (evidências)
Gerados no diretório local (ou na pasta do exe):
- `logs/` (logs)
- `relatorios/` (relatórios TXT)
- `json/` (manifest e debug JSON)
- `logs/screenshots/` (screenshots)

Em falha de verificação, são criados:
- `logs/screenshots/verify_fail_<cpf>_<timestamp>__<execid>.png`
- `json/verify_debug_<cpf>_<timestamp>__<execid>.json`

## 5) Modo TXT (fila automática)
- Arquivo padrão: `P:\\GuilhermeCostaProenca\\04_AUTOMACOES\\MetaXg\\inputs\\cadastrar_metax.txt`
- Formato: um nome por linha; ignora vazios e comentários `#`
- O TXT funciona como fila: nomes processados (attempted=True) ou ignorados por cache são removidos.
- Se ficar vazio, o arquivo é apagado.

## 6) Manifest e interpretação
O manifest (`json/manifest_*.json`) é a fonte da verdade.
- `run_status`: CONSISTENT ou INCONSISTENT
- `action_saved`: quantidade de rascunhos salvos
- `verified_success`: sucesso verificado
- `saved_not_verified`: salvo mas não confirmado
- `failed_action`: falha na ação
- `failed_verification`: erro na verificação

Regra: SUCCESS somente quando `verified=true`.

## 7) Playwright browsers
Se o navegador não abrir, instale:
```powershell
python -m playwright install chromium
```

Se estiver em outra máquina sem Python, copie a pasta de browsers do Playwright
(geralmente em %LOCALAPPDATA%\ms-playwright) e defina:
```powershell
$env:PLAYWRIGHT_BROWSERS_PATH="C:\\caminho\\para\\ms-playwright"
```
