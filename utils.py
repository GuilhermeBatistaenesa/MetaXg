import unicodedata
from datetime import date, datetime
import os
import tempfile
import fitz # PyMuPDF
from PIL import Image
from custom_logger import logger
from mappings import MAPA_CARGOS_METAX

def reduzir_foto_para_metax(caminho_original: str, tamanho_max_kb: int = 40) -> str | None:
    """
    Reduz a imagem (ou converte PDF) para o tamanho máximo especificado.
    Suporta .jpg, .png e .pdf.
    """
    try:
        img = None
        
        # 1. Se for PDF, converte para Imagem
        if caminho_original.lower().endswith(".pdf"):
            logger.info(f"Convertendo PDF para imagem: {caminho_original}", details={"path": caminho_original})
            try:
                doc = fitz.open(caminho_original)
                pagina = doc.load_page(0)  # Pega a primeira página
                pix = pagina.get_pixmap()
                img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                doc.close()
            except Exception as e_pdf:
                logger.error(f"Erro ao converter PDF: {e_pdf}", details={"error": str(e_pdf)})
                return None
        else:
            # 2. Se for imagem normal
            img = Image.open(caminho_original).convert("RGB")

        if not img:
            return None

        # 3. Redimensionar para 300x400 (Padrão MetaX)
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


def buscar_foto_por_cpf(pasta_fotos: str, cpf: str) -> str | None:
    """
    Busca uma foto (ou PDF) na pasta especificada.
    """
    cpf_numerico = ''.join(filter(str.isdigit, str(cpf)))

    if not os.path.exists(pasta_fotos):
        return None

    for arquivo in os.listdir(pasta_fotos):
        nome = arquivo.lower()
        # Agora aceitamos .pdf também
        if cpf_numerico in nome and nome.endswith((".jpg", ".jpeg", ".png", ".pdf")):
            return os.path.join(pasta_fotos, arquivo)

    return None

def ajustar_descricao_cargo(descricao_rm: str) -> str:
    """Normaliza e ajusta a descrição do cargo baseando-se no mapa de cargos."""
    return MAPA_CARGOS_METAX.get(descricao_rm.strip().upper(), descricao_rm.strip().upper())

def formatar_telefone_numerico(telefone: str) -> str:
    """Extrai apenas dígitos de um telefone e valida comprimento mínimo."""
    if not telefone:
        return ""
    # remove tudo que não for número
    telefone = ''.join(filter(str.isdigit, str(telefone)))
    # valida tamanho mínimo
    if len(telefone) < 10:
        return ""
    return telefone

def formatar_data(data: str | date | datetime) -> str:
    """
    Formata uma data para o padrão DD/MM/YYYY.
    Aceita string, date ou datetime.
    """
    if not data:
        return ""

    if isinstance(data, (date, datetime)):
        return data.strftime("%d/%m/%Y")

    try:
        data_convertida = datetime.strptime(str(data), "%Y-%m-%d")
        return data_convertida.strftime("%d/%m/%Y")
    except ValueError:
        raise ValueError(f"Data de nascimento inválida: {data}")

def normalizar_texto(txt: str) -> str:
    """Remove acentos e converte para maiúsculas."""
    if not txt:
        return ""
    txt = txt.strip().upper()
    txt = unicodedata.normalize('NFD', txt)
    txt = ''.join(c for c in txt if unicodedata.category(c) != 'Mn')
    return txt

def formatar_pis(pis: str) -> str:
    """Formata o PIS garantindo 11 dígitos numéricos."""
    if not pis:
        return ""
    pis = ''.join(filter(str.isdigit, str(pis)))
    if len(pis) < 11:
        pis = pis.zfill(11)
    if len(pis) != 11:
        raise ValueError(f"PIS/PASEP inválido: {pis}")
    return pis

def formatar_cpf(cpf: str) -> str:
    """Formata o CPF no padrão XXX.XXX.XXX-XX."""
    if not cpf:
        return ""
    cpf = ''.join(filter(str.isdigit, str(cpf)))
    if len(cpf) != 11:
        raise ValueError(f"CPF inválido: {cpf}")
    return f"{cpf[:3]}.{cpf[3:6]}.{cpf[6:9]}-{cpf[9:]}"
