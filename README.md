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

## ▶️ Como Executar

Para rodar a automação para os funcionários admitidos **hoje**:

```bash
python main.py
```

Os logs de execução serão salvos na pasta `logs/`. Screenshots de erros serão salvos em `logs/screenshots/`.

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
