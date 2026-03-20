# Documentacao Oficial - MetaXg

## Visao geral
MetaXg e o robo de cadastro de colaboradores no portal MetaX com integracao entre RM Labore, SharePoint, Playwright e camada de evidencias operacionais.

## Ambiente oficial
- Codigo: `Z:\T.I\MetaXg`
- Operacao publica: `P:\ProcessoMetaX`

## Fluxo resumido
1. Buscar colaboradores no SQL dentro da janela configurada
2. Classificar por contrato
3. Baixar foto
4. Abrir sessao no MetaX
5. Passar pelo captcha em modo semiautomatico
6. Preencher cadastro
7. Salvar rascunho
8. Verificar se o CPF apareceu na lista
9. Gerar artefatos
10. Enviar e-mail

## Artefatos publicos
- `P:\ProcessoMetaX\logs`
- `P:\ProcessoMetaX\relatorios`
- `P:\ProcessoMetaX\json`
- `P:\ProcessoMetaX\screenshots`
- `P:\ProcessoMetaX\fotos\processados`
- `P:\ProcessoMetaX\fotos\erros`

## Convencao atual de nomes
- `AAAA-MM-DD__HHhMM__status__relatorio_execucao.txt`
- `AAAA-MM-DD__HHhMM__status__manifest.json`
- `AAAA-MM-DD__HHhMM__status__manifest_parcial.json`
- `AAAA-MM-DD__HHhMM__status__resumo_execucao.md`
- `AAAA-MM-DD__HHhMM__pendencia_contrato_indice_nome_chapa.pdf`

## E-mail corporativo
- Outlook continua suportado
- SMTP esta suportado como fallback principal quando o Outlook COM falhar
- Anexos priorizam caminhos publicos em `P:\ProcessoMetaX`

## Captcha
- Destaque visual na aba do MetaX
- Alerta sonoro
- Clique semiautomatico no checkbox
- Tentativa automatica no botao `Validar`
- Espera manual somente quando o desafio exigir imagem

## Observabilidade
- Log estruturado em `.jsonl`
- Manifest parcial e final
- Relatorio TXT
- Resumo Markdown
- Relatorio JSON
- PDFs de pendencia individual
- Screenshot e JSON tecnico em falha de verificacao

## Regras de leitura do resultado
- `VERIFIED_SUCCESS`: rascunho salvo e CPF confirmado
- `SAVED_NOT_VERIFIED`: portal respondeu sucesso, mas o CPF nao foi confirmado
- `FAILED_ACTION`: falha durante preenchimento ou salvamento
- `FAILED_VERIFICATION`: erro tecnico na etapa de verificacao
- `SKIPPED_ALREADY_EXISTS`: CPF ja estava no cache de rascunhos

## Runner e releases
- Releases publicas em `P:\ProcessoMetaX\releases`
- Fallback opcional via GitHub Releases
- Validacao de integridade por SHA256
