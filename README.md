# RPA MetaXg

Automacao em Python para cadastrar colaboradores no portal MetaX a partir do RM Labore, baixar fotos do SharePoint e gerar evidencias operacionais em pasta publica.

## O que o robo faz
- Consulta colaboradores no SQL Server dentro da janela configurada por `DIAS_RETROATIVOS`.
- Baixa e prepara a foto do colaborador.
- Faz login no MetaX e preenche os dados pessoais, documentos, endereco e dados profissionais.
- Salva o credenciamento como rascunho.
- Verifica se o CPF apareceu na lista de rascunhos.
- Gera relatorios, manifests, logs, PDFs de pendencia e evidencias de erro.
- Envia e-mail por Outlook ou SMTP, com preferencia por anexos em `P:\ProcessoMetaX`.

## Pastas operacionais
- Codigo oficial: `Z:\T.I\MetaXg`
- Base publica: `P:\ProcessoMetaX`
- Entrada manual: `P:\ProcessoMetaX\entrada`
- Fotos em processamento: `P:\ProcessoMetaX\fotos\em_processamento`
- Fotos processadas: `P:\ProcessoMetaX\fotos\processados`
- Fotos com erro: `P:\ProcessoMetaX\fotos\erros`
- Logs: `P:\ProcessoMetaX\logs`
- Relatorios: `P:\ProcessoMetaX\relatorios`
- JSON e manifests: `P:\ProcessoMetaX\json`
- Screenshots: `P:\ProcessoMetaX\screenshots`
- Releases: `P:\ProcessoMetaX\releases`

## Nome dos artefatos
Os arquivos principais agora seguem um nome legivel:
- `2026-03-19__15h15__consistent__relatorio_execucao.txt`
- `2026-03-19__15h15__inconsistent__manifest.json`
- `2026-03-19__15h15__inconsistent__manifest_parcial.json`
- `2026-03-19__15h15__inconsistent__resumo_execucao.md`
- `2026-03-19__15h15__pendencia_eletromecanica_01_nome_12501125.pdf`

## Captcha
O login esta em modo semiautomatico:
- o robo destaca a janela do MetaX
- toca um alerta sonoro
- tenta clicar no checkbox do reCAPTCHA
- tenta clicar em `Validar`
- se abrir desafio de imagem, espera voce assumir manualmente

## E-mail
- O envio pode usar Outlook Desktop quando o COM estiver saudavel.
- Se o Outlook falhar, o robo pode usar SMTP.
- Os anexos priorizam os arquivos da pasta publica `P:\ProcessoMetaX`.
- PDFs e fotos JPG de pendencias podem ser anexados no mesmo envio.

## Como executar
```powershell
python main.py
```

Ou:
```powershell
run_main.bat
```

## Principais argumentos
```powershell
python main.py --no-email
python main.py --dry-run
python main.py --txt "P:\ProcessoMetaX\entrada\cadastrar_metax.txt"
```

## Observacoes importantes
- `SUCCESS` so existe quando a verificacao encontra o CPF nos rascunhos.
- `SAVED_NOT_VERIFIED` significa que o portal respondeu sucesso, mas o CPF nao foi confirmado na lista.
- O relatorio sempre precisa ser interpretado junto com a janela processada pela execucao.

## Documentacao
- `OPERACAO_METAX.md`
- `DOCUMENTACAO_OFICIAL.md`
- `README_Runner.md`
- `docs/`
