# RPA MetaX - Automação de Cadastro de Funcionários

Este projeto é uma automação (RPA) desenvolvida em Python para integrar o banco de dados SQL Server (RM Labore) com o portal web MetaX, além de baixar fotos de funcionários do SharePoint.

## 🚀 Funcionalidades

- **Extração de Dados**: Consulta funcionários admitidos no dia (ou data específica) diretamente do banco de dados SQL Server.
- **Download de Fotos**: Busca e baixa a foto do funcionário no SharePoint corporativo, organizadas por data de admissão.
- **Automação Web**: Realiza o login no portal MetaX, preenche o formulário de credenciamento (Dados Pessoais, Endereço, Documentos, Dados Profissionais) e salva como rascunho.
- **Resiliência**: Tratamento de erros, re-tentativas (retries) e logs estruturados em JSON.
- **Tratamento de Imagens**: Redimensiona e ajusta a qualidade das imagens automaticamente para atender aos requisitos do portal (máx 40KB).

## 🛠️ Tecnologias Utilizadas

- **Python 3.10+**
- **Playwright**: Automação de navegador.
- **PyODBC**: Conexão com banco de dados SQL Server.
- **Office365-REST-Python-Client**: Integração com API do SharePoint.
- **Pillow (PIL)**: Processamento de imagens.
- **Python-Dotenv**: Gerenciamento de variáveis de ambiente.

## ⚙️ Configuração

1.  **Clone o repositório** (ou copie os arquivos para sua máquina).
2.  **Crie o ambiente virtual**:
    ```bash
    python -m venv venv
    source venv/bin/activate  # Linux/Mac
    venv\Scripts\activate     # Windows
    ```
3.  **Instale as dependências**:
    ```bash
    pip install -r requirements.txt
    playwright install chromium
    ```
4.  **Instale o Poppler (para conversão de PDFs)**:
    - **Windows**: Baixe de https://github.com/oschwartz10612/poppler-windows/releases/
    - Extraia o arquivo ZIP e adicione a pasta `bin` ao PATH do sistema
    - Ou instale via Chocolatey: `choco install poppler`
    - **Linux**: `sudo apt-get install poppler-utils` (Ubuntu/Debian) ou `sudo yum install poppler-utils` (CentOS/RHEL)
    - **Mac**: `brew install poppler`
    
    > **Nota**: O código tenta usar PyMuPDF primeiro (mais rápido), e usa pdf2image como fallback. Se PyMuPDF não funcionar no seu ambiente, o Poppler é necessário para pdf2image.
    
    > **Dica**: Execute `python verificar_poppler.py` para verificar se o Poppler está instalado corretamente e obter instruções detalhadas de instalação.

5.  **Configure as Variáveis de Ambiente**:
    - Copie o arquivo `.env.example` para `.env`.
    - Preencha as chaves com suas credenciais reais (SharePoint, Banco de Dados, MetaX).
6.  **Pré-requisito Windows**:
    - ODBC Driver 17/18 precisa estar instalado no Windows alvo.
7.  **E-mail (opcional)**:
    - Requer Outlook Desktop instalado.
    - A biblioteca `pywin32` ? instalada via `requirements.txt`.
    - Se n?o quiser enviar e-mail, deixe `EMAIL_NOTIFICACAO` vazio ou rode com `--no-email`.

## ▶️ Como Executar

Para rodar a automação (irá buscar funcionários admitidos hoje e dias retroativos conforme `.env`):

```bash
python main.py
```

Os logs de execução serão salvos na pasta `logs/`. Screenshots de erros serão salvos em `logs/screenshots/`. Relatórios em `relatorios/`.
Se existir um diretório de rede configurado (`PUBLIC_BASE_DIR`), as fotos ficam em `em processamento/`, `processados/` e `erros/`, e os relatórios/logs são espelhados nos aliases `logs/`, `relatorios/` e `json/`. Caso contrário, usa o diretório do projeto.
Fotos em `processados/` e `erros/` sao organizadas por data (YYYY-MM-DD).

## ❓ Troubleshooting (Resolução de Problemas)

### 1. "Cargo não encontrado"
Se o script acusar que o cargo do RM não existe no MetaX:
- Verifique o log para ver as opções disponíveis listadas.
- Adicione/Corrija o mapeamento em `mappings.py` na variável `MAPA_CARGOS_METAX`.

### 2. Modais bloqueando a tela
O script possui um mecanismo (`fechar_modais_bloqueantes`) que tenta fechar popups automaticamente. Se persistir, verifique se houve mudança no layout do MetaX.

### 3. Timeout ao salvar
Se a internet estiver lenta, o script tenta clicar em "Salvar" novamente e aguarda até 90 segundos.

### 4. Portabilidade (Outros Computadores)
Para rodar em outra máquina:
1. **Instale Python e as dependências** (`requirements.txt`).
2. **Copie o `.env`**: Configure as variáveis, especialmente `PUBLIC_BASE_DIR`, `FOTOS_EM_PROCESSAMENTO_DIR`, `FOTOS_PROCESSADOS_DIR`, `FOTOS_ERROS_DIR` (ou `PASTA_FOTOS` legado) se o caminho de rede não existir na nova máquina.
3. **Driver SQL**: Certifique-se que a máquina tem o "ODBC Driver for SQL Server" instalado.

## 📂 Estrutura do Projeto

- `main.py`: Orquestrador principal.
- `rpa_metax.py`: Lógica de automação web (Playwright).
- `sharepoint.py`: Integração com SharePoint.
- `utils.py`: Funções auxiliares (formatação, tratamento de imagem).
- `config.py`: Carregamento de configurações.
- `mappings.py`: Dicionários de-para (Cargos, Estados, etc.).
- `custom_logger.py`: Configuração de logs.

## 📚 Documentação

A documentação técnica detalhada e versionada encontra-se na pasta `docs/`.

## Teste manual de verificacao (auditoria)
1. Forcar `verificar_cadastro` retornar True -> deve gerar `VERIFIED_SUCCESS`.
2. Forcar `verificar_cadastro` retornar False -> `action_saved` pode ser true, mas `outcome = SAVED_NOT_VERIFIED` e `run_status = INCONSISTENT`.
3. Forcar excecao na verificacao -> `outcome = FAILED_VERIFICATION` e `run_status = INCONSISTENT`.

## Padrao operacional (MetaXg)

### Runner (update e execucao)
- Arquivo: `runner.py`
- Config: `config.json`
- Fluxo: verifica `latest.json` no `network_release_dir` (config.json), valida SHA256, instala em `install_dir\app\current`, executa `MetaXg.exe`.
- Fallback: GitHub Releases (`GuilhermeBatistaenesa/MetaXg`) se `github_repo` estiver configurado.

### Scripts e atalhos
- `run_main.bat`: executa `main.py` (usa `.venv` se existir).
- `run_tests.bat`: executa pytest.
- `build_windows.bat`: gera build onedir via PyInstaller.
- `build_zip.bat`: gera zip + sha256 + latest.json.
Obs: `run_main.bat` cria o venv e instala dependÃªncias automaticamente na primeira execuÃ§Ã£o.

### Releases
- Padrão de artefatos:
  - `MetaXg_<versao>.zip`
  - `MetaXg_<versao>.sha256`
  - `latest.json`
- Diretorio de releases: `network_release_dir` (config.json ou env)
