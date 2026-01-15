# Manual Técnico - RPA MetaX
**Versão:** 1.0.0
**Data:** 15/01/2026
**Responsável:** Equipe de Automação

---

## 1. Visão Geral da Arquitetura

O projeto adota uma arquitetura modular, separando responsabilidades em arquivos distintos para facilitar a manutenção e escalabilidade. O fluxo de dados segue o padrão ETL (Extract, Transform, Load):

1.  **Extract (Extração)**: O `main.py` aciona a consulta SQL via `pyodbc` para buscar dados dos funcionários no banco RM Labore.
2.  **Enrich (Enriquecimento)**: O script busca e baixa a foto correspondente do funcionário no SharePoint usando `sharepoint.py`.
3.  **Load (Carga/Automação)**: O `rpa_metax.py` utiliza o Playwright para inserir os dados no portal web MetaX.

### Diagrama de Módulos

- **`main.py`**: Controlador principal. Orquestra a ordem de execução e o tratamento de erros macro.
- **`config.py`**: Camada de configuração. Centraliza o acesso a variáveis de ambiente (`.env`).
- **`utils.py`**: Camada de utilitários. Funções puras para tratamento de dados (CPF, PIS, Datas) e imagens.
- **`mappings.py`**: Camada de dados estáticos. Contém dicionários de correlação (DE-PARA) entre códigos do RM e valores dos selects HTML do MetaX.
- **`sharepoint.py`**: Módulo conector. Responsável por autenticar e baixar arquivos do SharePoint.
- **`rpa_metax.py`**: Módulo de automação. Contém a lógica de interação com a interface web.

---

## 2. Detalhamento dos Módulos

### 2.1. `main.py`
Função principal: `main()`.
- Busca a lista de dicionários de funcionários (`buscar_funcionarios_para_cadastro`).
- Dispara o download em lote das fotos (`baixar_fotos_em_lote`).
- Inicia a sessão no navegador (`iniciar_sessao`).
- Itera sobre cada funcionário chamando `cadastrar_funcionario`.
- **Tratamento de Erro:** Se um funcionário falhar, o erro é logado, a página é redirecionada para a home, e o loop continua para o próximo.

### 2.2. `rpa_metax.py`
Foi refatorado para evitar uma "função deus". As responsabilidades foram divididas:
- `navegar_para_cadastro(page)`: Lida com modais impeditivos e navegação de URL.
- `preencher_dados_pessoais(...)`: Preenche inputs simples e selects mapeados.
- `preencher_documentos(...)`: Preenche RG, CTPS, PIS.
- `preencher_endereco(...)`: Lida com a busca de CEP e lógica de fallback (se o CEP não preencher bairro/rua, preenche com dados do banco).
- `anexar_foto(...)`: Redimensiona a imagem em memória (via `utils`) e faz o upload invisível.

### 2.3. `sharepoint.py`
Utiliza a biblioteca `Office365-REST-Python-Client`.
- A estrutura de pastas é dinâmica baseada na data de admissão: `.../ANO/MES/DIA/NOME_DO_FUNCIONARIO/`.

### 2.4. `utils.py`
Contém a lógica complexa de compressão de imagem (`reduzir_foto_para_metax`).
- O MetaX exige imagens < 40KB.
- A função redimensiona para 300x400 e reduz a qualidade JPEG progressivamente até atingir o tamanho alvo.

---

## 3. Mapeamentos (`mappings.py`)

O sistema depende de "mapas" para traduzir o que vem do banco de dados (códigos ou descrições RM) para os IDs (`value`) esperados pelo HTML do MetaX.
- **Importante:** Se um novo cargo for criado no RM ou uma nova opção surgir no MetaX, este arquivo deve ser atualizado.

---

## 4. Tratamento de Erros e Logs

- **Logs**: Utiliza `custom_logger.py` para gerar logs estruturados (JSON Lines) em `logs/execution_log_DATETIME.json`. Isso facilita a ingestão futura por ferramentas como Splunk ou ElasticSearch.
- **Screenshots**: Em caso de erro no salvamento do rascunho, um print da tela é salvo automaticamente em `logs/screenshots/`.

## 5. Manutenção

Para adicionar uma nova regra de negócio:
1.  Verifique se é uma **formatação de dado**: Adicione em `utils.py`.
2.  Verifique se é uma **nova interação no site**: Adicione em `rpa_metax.py`.
3.  Verifique se é uma **alteração de credencial**: Atualize o `.env`.
