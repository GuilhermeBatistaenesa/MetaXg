# Runbook Operacional - ProcessoMetaX

## Pre-check
1. Confirmar acesso ao `P:\ProcessoMetaX`
2. Confirmar acesso ao `Z:\T.I\MetaXg`
3. Confirmar SQL, SharePoint e MetaX disponiveis
4. Confirmar `.env` valido

## Execucao
1. Rodar o metodo homologado
2. Acompanhar o captcha quando o alerta visual aparecer
3. Conferir o log da execucao
4. Validar relatorio e manifest ao final

## Pos-execucao
1. Conferir `relatorios`
2. Conferir `json`
3. Conferir `fotos\processados` e `fotos\erros`
4. Conferir envio de e-mail

## Sinais de atencao
- `INCONSISTENT`
- `SAVED_NOT_VERIFIED`
- `FAILED_ACTION`
- `FAILED_VERIFICATION`
- falha de escrita em pasta publica

## Evidencias minimas
- log `.jsonl`
- manifest parcial
- manifest final
- relatorio TXT
- screenshot ou JSON tecnico quando houver falha de verificacao
