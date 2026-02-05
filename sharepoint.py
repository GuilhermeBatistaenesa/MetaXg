import os
import unicodedata
from custom_logger import logger


from config import SHAREPOINT_SITE_URL, SHAREPOINT_CLIENT_ID, SHAREPOINT_CLIENT_SECRET

SITE_URL = SHAREPOINT_SITE_URL
CLIENT_ID = SHAREPOINT_CLIENT_ID
CLIENT_SECRET = SHAREPOINT_CLIENT_SECRET

BASE_FOLDER = (
    "/sites/Corporativo/"
    "Documentos Compartilhados/Enesa/"
    "125 - ARAUCO/DP/"
    "01.ADMISSÃO/"
    "04.DOCUMENTAÇÃO MOBILIZAÇÃO"
)

# CONEXÃO
def conectar_sharepoint():
    """Establishing a connection to SharePoint."""
    from office365.sharepoint.client_context import ClientContext
    from office365.runtime.auth.client_credential import ClientCredential

    creds = ClientCredential(CLIENT_ID, CLIENT_SECRET)
    return ClientContext(SITE_URL).with_credentials(creds)

# UTILIDADES
# UTILIDADES
def normalizar(texto: str) -> str:
    """Normalize text by removing accents and converting to uppercase."""
    return unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode("ascii")

def mes_por_extenso(data: any) -> str:
    """Returns the month name (in Portuguese) for a given date object."""
    meses = {
        1: "JANEIRO",
        2: "FEVEREIRO",
        3: "MARCO",
        4: "ABRIL",
        5: "MAIO",
        6: "JUNHO",
        7: "JULHO",
        8: "AGOSTO",
        9: "SETEMBRO",
        10: "OUTUBRO",
        11: "NOVEMBRO",
        12: "DEZEMBRO"
    }
    return meses[data.month]

# BUSCA LOCAL
def _buscar_imagem_local(cpf_limpo: str, pastas_busca: list[str]) -> str | None:
    import glob
    for pasta in pastas_busca:
        if not pasta or not os.path.exists(pasta):
            continue
        padrao = os.path.join(pasta, f"*{cpf_limpo}*.*")
        arquivos_locais = glob.glob(padrao)
        imagens_validas = [f for f in arquivos_locais if f.lower().endswith((".jpg", ".jpeg", ".png"))]
        if imagens_validas:
            return imagens_validas[0]
    return None



# DOWNLOAD DA FOTO
def baixar_foto_funcionario(funcionario: dict, pasta_destino: str, pastas_busca: list[str] | None = None) -> str | None:
    """
    Downloads the employee's photo from SharePoint.
    
    Args:
        funcionario (dict): Dictionary with employee data.
        pasta_destino (str): Directory to save the photo.
        pastas_busca (list[str] | None): Pastas para buscar foto local antes do download.
        
    Returns:
        str | None: Path to the downloaded photo or None if not found.
    """
    ctx = conectar_sharepoint()

    nome = funcionario["NOME"]
    data = funcionario["DATAADMISSAO"]

    ano = str(data.year)
    mes = mes_por_extenso(data)
    dia = f"{data.day:02d}"

    nome_pasta = normalizar(nome.upper())

    cpf_limpo = ''.join(filter(str.isdigit, str(funcionario.get("CPF", ""))))
    pastas_busca = [p for p in (pastas_busca or []) if p]
    if not pastas_busca:
        pastas_busca = [pasta_destino]

    if cpf_limpo:
        foto_local = _buscar_imagem_local(cpf_limpo, pastas_busca)
        if foto_local:
            logger.info(f"Foto j?? existe localmente: {nome}", details={"path": foto_local})
            return foto_local

    caminho_pasta = f"{BASE_FOLDER}/{ano}/{mes}/{dia}/{nome_pasta}"

    try:
        pasta = ctx.web.get_folder_by_server_relative_url(caminho_pasta)
        arquivos = pasta.files
        ctx.load(arquivos)
        ctx.execute_query()
    except Exception:
        logger.warn(f"Pasta n??o encontrada: {caminho_pasta}", details={"path": caminho_pasta})
        return None

    for arq in arquivos:
        nome_arquivo = arq.name.upper()
        if nome_arquivo.startswith("00_FOTO_") and nome_arquivo.endswith((".JPG", ".JPEG", ".PNG")):
            os.makedirs(pasta_destino, exist_ok=True)
            caminho_local = os.path.join(pasta_destino, arq.name)

            with open(caminho_local, "wb") as f:
                arq.download(f).execute_query()

            logger.info(f"Foto baixada: {nome}", details={"file": arq.name, "path": caminho_local})
            return caminho_local

    logger.warn(f"Foto n??o encontrada: {nome}", details={"nome": nome, "pasta": caminho_pasta})
    return None

# PROCESSAMENTO EM LOTE
def baixar_fotos_em_lote(funcionarios: list[dict], pasta_destino: str, pastas_busca: list[str] | None = None) -> dict:
    """
    Retorna:
    {
        CPF: caminho_da_foto | None
    }
    """
    resultados = {}

    for func in funcionarios:
        cpf = func.get("CPF")
        try:
            caminho = baixar_foto_funcionario(func, pasta_destino, pastas_busca=pastas_busca)
            resultados[cpf] = caminho
        except Exception as e:
            logger.error(f"Falha ao baixar foto de {func.get('NOME')}: {e}", details={"error": str(e), "cpf": cpf})
            resultados[cpf] = None

    return resultados
