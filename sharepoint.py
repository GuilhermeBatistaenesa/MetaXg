import os
import unicodedata
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.client_credential import ClientCredential
from custom_logger import logger


SITE_URL = "https://enesaengenharia.sharepoint.com/sites/Corporativo"

CLIENT_ID = "1a1f030d-2177-4f0a-bb58-b9491bb7c8c7"
CLIENT_SECRET = "ExR8Q~~PsBWzR4W50uQ7fXDemgl6MJ~XNcUanaDu"

BASE_FOLDER = (
    "/sites/Corporativo/"
    "Documentos Compartilhados/Enesa/"
    "125 - ARAUCO/DP/"
    "01.ADMISSÃO/"
    "04.DOCUMENTAÇÃO MOBILIZAÇÃO"
)

# CONEXÃO
def conectar_sharepoint():
    creds = ClientCredential(CLIENT_ID, CLIENT_SECRET)
    return ClientContext(SITE_URL).with_credentials(creds)

# UTILIDADES
def normalizar(texto: str) -> str:
    return unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode("ascii")

def mes_por_extenso(data):
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


# DOWNLOAD DA FOTO
def baixar_foto_funcionario(funcionario: dict, pasta_destino: str) -> str | None:
    ctx = conectar_sharepoint()

    nome = funcionario["NOME"]
    data = funcionario["DATAADMISSAO"]

    ano = str(data.year)
    mes = mes_por_extenso(data)
    dia = f"{data.day:02d}"

    nome_pasta = normalizar(nome.upper())

    caminho_pasta = f"{BASE_FOLDER}/{ano}/{mes}/{dia}/{nome_pasta}"

    try:
        pasta = ctx.web.get_folder_by_server_relative_url(caminho_pasta)
        arquivos = pasta.files
        ctx.load(arquivos)
        ctx.execute_query()
    except Exception:
        logger.warn(f"Pasta não encontrada: {caminho_pasta}", details={"path": caminho_pasta})
        return None

    for arq in arquivos:
        if arq.name.upper().startswith("00_FOTO_"):
            os.makedirs(pasta_destino, exist_ok=True)
            caminho_local = os.path.join(pasta_destino, arq.name)

            with open(caminho_local, "wb") as f:
                arq.download(f).execute_query()

            logger.info(f"Foto baixada: {nome}", details={"file": arq.name, "path": caminho_local})
            return caminho_local

    logger.warn(f"Foto não encontrada: {nome}", details={"nome": nome, "pasta": caminho_pasta})
    return None


# PROCESSAMENTO EM LOTE
def baixar_fotos_em_lote(funcionarios: list[dict], pasta_destino: str) -> dict:
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
            caminho = baixar_foto_funcionario(func, pasta_destino)
            resultados[cpf] = caminho
        except Exception as e:
            logger.error(f"Falha ao baixar foto de {func.get('NOME')}: {e}", details={"error": str(e), "cpf": cpf})
            resultados[cpf] = None

    return resultados
