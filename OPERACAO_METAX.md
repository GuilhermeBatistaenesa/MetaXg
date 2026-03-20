# Operacao diaria - MetaXg

## 1. Atualizar o repositorio
```powershell
cd Z:\T.I\MetaXg
git pull
```

## 2. Executar o robo
Modo Python:
```powershell
. .\.venv\Scripts\Activate.ps1
python main.py
```

Modo rapido:
```powershell
run_main.bat
```

Modo executavel:
```powershell
.\dist\MetaXg\MetaXg.exe
```

## 3. Acompanhar o CAPTCHA
- O MetaX abre no navegador.
- Quando chegar no captcha, o robo mostra um alerta visual grande e toca som.
- Ele tenta clicar no checkbox e no botao `Validar`.
- Se abrir desafio de imagem, resolva manualmente.

## 4. Onde conferir o resultado
- Logs: `P:\ProcessoMetaX\logs`
- Relatorios: `P:\ProcessoMetaX\relatorios`
- Manifests: `P:\ProcessoMetaX\json`
- Screenshots: `P:\ProcessoMetaX\screenshots`
- Fotos com erro: `P:\ProcessoMetaX\fotos\erros`
- Fotos processadas: `P:\ProcessoMetaX\fotos\processados`

## 5. Como ler os nomes dos arquivos
Exemplo:
`2026-03-19__15h15__inconsistent__manifest.json`

Significa:
- data da execucao
- hora da execucao
- status final da execucao
- tipo do artefato

## 6. Modo TXT
- Arquivo: `P:\ProcessoMetaX\entrada\cadastrar_metax.txt`
- Um nome por linha
- Linhas com `#` sao comentarios
- Se nao existir TXT, o robo roda em modo SQL padrao

## 7. E-mail
- O robo tenta Outlook ou SMTP.
- Os anexos devem sair preferencialmente de `P:\ProcessoMetaX`.
- Pendencias podem seguir com PDF e JPG da foto.

## 8. Evidencias de falha de verificacao
- Screenshot: `P:\ProcessoMetaX\screenshots\verify_fail_<cpf>_<timestamp>__<execid>.png`
- Debug JSON: `P:\ProcessoMetaX\json\verify_debug_<cpf>_<timestamp>__<execid>.json`

## 9. Regra pratica de analise
- `CONSISTENT`: sem falha de acao e sem falha de verificacao
- `INCONSISTENT`: houve erro real, salvo nao verificado ou problema tecnico na execucao
