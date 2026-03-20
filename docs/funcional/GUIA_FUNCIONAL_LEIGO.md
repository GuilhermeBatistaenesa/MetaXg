# Guia Funcional (Publico Leigo) - ProcessoMetaX

## O que este robo faz
O robo pega os dados do colaborador, busca a foto, entra no portal MetaX e deixa o cadastro como rascunho para continuidade operacional.

## O que entra
- Dados do colaborador vindos do RM
- Foto vinda do SharePoint
- Ou, no modo manual, um TXT em `P:\ProcessoMetaX\entrada`

## O que sai
- Relatorio da execucao
- Manifest com o resultado de cada colaborador
- Fotos classificadas em processados ou erros
- Evidencias tecnicas quando algo falha

## Como saber se deu certo
1. Existe relatorio em `P:\ProcessoMetaX\relatorios`
2. Existe manifest em `P:\ProcessoMetaX\json`
3. O resultado do colaborador aparece como `VERIFIED_SUCCESS`

## Quando voce precisa agir
1. Quando o alerta do captcha aparecer
2. Quando houver pendencia manual em PDF
3. Quando o relatorio indicar `INCONSISTENT`

## O que fazer se der erro
1. Verificar o relatorio
2. Verificar o manifest
3. Verificar a pasta `fotos\erros`
4. Acionar o suporte tecnico com o nome do arquivo da execucao
