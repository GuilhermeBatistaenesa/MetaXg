import unicodedata
from datetime import date, datetime
import os
import tempfile
from PIL import Image
from custom_logger import logger
from mappings import MAPA_CARGOS_METAX

def reduzir_foto_para_metax(caminho_original, tamanho_max_kb=40):
    try:
        img = Image.open(caminho_original).convert("RGB")
        img = img.resize((300, 400), Image.LANCZOS)

        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".jpg")
        temp_path = temp_file.name
        temp_file.close()

        qualidade = 85

        while qualidade >= 20:
            img.save(
                temp_path,
                format="JPEG",
                quality=qualidade,
                optimize=True,
                progressive=True
            )

            tamanho_kb = os.path.getsize(temp_path) / 1024

            if tamanho_kb <= tamanho_max_kb:
                logger.info(
                    f"Foto pronta para MetaX | {tamanho_kb:.1f} KB",
                    details={"size_kb": tamanho_kb, "quality": qualidade}
                )
                return temp_path

            qualidade -= 5

        logger.warn("Imagem não ficou abaixo de 40KB mesmo após resize")
        return None

    except Exception as e:
        logger.warn(f"Erro ao preparar imagem: {e}", details={"error": str(e)})
        return None


def buscar_foto_por_cpf(pasta_fotos, cpf):
    cpf_numerico = ''.join(filter(str.isdigit, str(cpf)))

    if not os.path.exists(pasta_fotos):
        return None

    for arquivo in os.listdir(pasta_fotos):
        nome = arquivo.lower()
        if cpf_numerico in nome and nome.endswith((".jpg", ".jpeg", ".png")):
            return os.path.join(pasta_fotos, arquivo)

    return None

def ajustar_descricao_cargo(descricao_rm: str) -> str:
    return MAPA_CARGOS_METAX.get(descricao_rm.strip().upper(), descricao_rm.strip().upper())

def formatar_telefone_numerico(telefone):
    if not telefone:
        return ""
    # remove tudo que não for número
    telefone = ''.join(filter(str.isdigit, str(telefone)))
    # valida tamanho mínimo
    if len(telefone) < 10:
        return ""
    return telefone

def formatar_data(data):
    if not data:
        return ""

    if isinstance(data, (date, datetime)):
        return data.strftime("%d/%m/%Y")

    try:
        data_convertida = datetime.strptime(str(data), "%Y-%m-%d")
        return data_convertida.strftime("%d/%m/%Y")
    except ValueError:
        raise ValueError(f"Data de nascimento inválida: {data}")

def normalizar_texto(txt):
    if not txt:
        return ""
    txt = txt.strip().upper()
    txt = unicodedata.normalize('NFD', txt)
    txt = ''.join(c for c in txt if unicodedata.category(c) != 'Mn')
    return txt

def formatar_pis(pis):
    if not pis:
        return ""
    pis = ''.join(filter(str.isdigit, str(pis)))
    if len(pis) < 11:
        pis = pis.zfill(11)
    if len(pis) != 11:
        raise ValueError(f"PIS/PASEP inválido: {pis}")
    return pis

def formatar_cpf(cpf):
    if not cpf:
        return ""
    cpf = ''.join(filter(str.isdigit, str(cpf)))
    if len(cpf) != 11:
        raise ValueError(f"CPF inválido: {cpf}")
    return f"{cpf[:3]}.{cpf[3:6]}.{cpf[6:9]}-{cpf[9:]}"
