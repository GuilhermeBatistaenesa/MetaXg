from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
import unicodedata
from datetime import date, datetime
import os
from PIL import Image
import tempfile
from custom_logger import logger
from utils import (
    reduzir_foto_para_metax, buscar_foto_por_cpf, ajustar_descricao_cargo,
    formatar_telefone_numerico, formatar_data, normalizar_texto, 
    formatar_pis, formatar_cpf
)
from mappings import (
    MAPA_CARGOS_METAX, MAPA_ESCOLARIDADE, MAPA_ESTADO_CIVIL, 
    MAPA_SEXO, MAPA_ESTADO_NATAL
)


from config import (
    METAX_LOGIN, METAX_PASSWORD, METAX_URL_LOGIN, 
    PASTA_FOTOS, SCREENSHOT_DIR
)

TIMEOUT = 30000
TEMPO_CAPTCHA_MS = 10000 



def anexar_foto(page, caminho_foto: str) -> None:
    """
    Anexa a foto do funcionário no formulário do MetaX.
    Realiza o redimensionamento antes do upload.
    
    Args:
        page (Page): Objeto de página do Playwright.
        caminho_foto (str): Caminho local para a foto original.
    """
    try:
        foto_reduzida = reduzir_foto_para_metax(caminho_foto)

        if not foto_reduzida:
            logger.info("Foto ignorada (não compatível)", details={"path": caminho_foto})
            return

        # Input de arquivo geralmente é hidden, então esperamos apenas estar anexado ao DOM
        page.wait_for_selector("#avatar", state="attached", timeout=TIMEOUT)

        page.set_input_files("#avatar", foto_reduzida)

        page.evaluate("""
            const input = document.querySelector('#avatar');
            input.dispatchEvent(new Event('change', { bubbles: true }));
        """)

        valor = page.locator("#avatar").input_value()
        
        logger.info("Foto anexada ao formulário MetaX", details={"original_path": caminho_foto})

    except Exception as e:
        logger.warn(f"Falha ao anexar foto: {e}", details={"error": str(e)})


def selecionar_cargo_por_descricao(page, descricao_cargo: str) -> bool:
    """
    Tenta selecionar um cargo no combo box buscando pela descrição parcial.
    
    Args:
        page (Page): Página do Playwright.
        descricao_cargo (str): Descrição do cargo para busca.
        
    Returns:
        bool: True se encontrou e selecionou, False caso contrário.
    """
    descricao_cargo = descricao_cargo.strip().upper()

    page.wait_for_selector("#cargo", timeout=30000)
    page.wait_for_function(
        """() => {
            const sel = document.querySelector('#cargo');
            return sel && sel.options && sel.options.length > 1;
        }""",
        timeout=30000
    )

    select = page.locator("#cargo")
    opcoes = select.locator("option")

    for i in range(opcoes.count()):
        texto_opcao = opcoes.nth(i).inner_text().upper()

        if texto_opcao.startswith(descricao_cargo):
            valor = opcoes.nth(i).get_attribute("value")
            page.select_option("#cargo", value=valor)
            page.locator("#cargo").press("Tab")
            logger.info(f"Cargo selecionado: {texto_opcao}", details={"cargo": texto_opcao})
            return True

    return False



# ==============================================================================
# NOVA FUNÇÃO DE INÍCIO DE SESSÃO (Retorna p, browser, page)
# ==============================================================================
def iniciar_sessao():
    """
    Inicia o browser, realiza login e navega até a tela inicial do sistema.
    
    Returns:
        tuple: (playwright_instance, browser_instance, page_instance)
    """
    # Inicia o Playwright, mas NÃO fecha (quem fecha é o main)
    # Obs: sync_playwright() deve ser usado com contexto. 
    # Para ser persistente, chamamos .start() manualmente
    p = sync_playwright().start()
    browser = p.chromium.launch(channel="chrome", headless=False)
    context = browser.new_context(ignore_https_errors=True)
    page = context.new_page()

    try:
        logger.info("Acessando a página de login...")
        page.goto(METAX_URL_LOGIN, timeout=TIMEOUT, wait_until="domcontentloaded")
        # LOGIN
        page.wait_for_selector('#txtLogin', timeout=TIMEOUT)
        page.fill('#txtLogin', METAX_LOGIN)
        page.wait_for_selector('#txtSenha', timeout=TIMEOUT)
        page.fill('#txtSenha', METAX_PASSWORD)        
        logger.info(f"Aguardando CAPTCHA ({TEMPO_CAPTCHA_MS/1000:.0f}s)...")
        page.wait_for_timeout(TEMPO_CAPTCHA_MS)

        page.click('button:has-text("Validar")')

        page.wait_for_selector('#comboContrato', state='visible')

        # Espera as opções carregarem
        page.wait_for_function(
            """() => {
                const sel = document.querySelector('#comboContrato');
                return sel && sel.options.length > 1;
            }""",
            timeout=TIMEOUT
        )

        page.select_option('#comboContrato', value='6578')

        # força o evento de change (importante nesse sistema)
        page.evaluate("""
            const select = document.querySelector('#comboContrato');
            select.dispatchEvent(new Event('change', { bubbles: true }));
        """)

        page.wait_for_selector('button:has-text("Continuar"):not([disabled])', timeout=TIMEOUT)
        page.click('button:has-text("Continuar")')

        page.wait_for_selector('text=Termo de confirmação', timeout=TIMEOUT)
        page.click('text=Li e Aceito os termos de compromisso')

        page.wait_for_selector('text=Termo de confirmação', state="hidden", timeout=TIMEOUT)
        logger.info("Login concluído com sucesso!")
        
        return p, browser, page

    except Exception as e:
        logger.error(f"Erro no login: {e}", details={"error": str(e)})
        browser.close()
        p.stop()
        raise e

# ==============================================================================
# FUNÇÃO DE CADASTRO (Recebe page logada)
# ==============================================================================
def navegar_para_cadastro(page):
    """Navega do menu inicial até a tela de cadastro (Credenciamento)."""
    # Verifica se tem algum modal de erro travando a tela (bootbox)
    try:
        if page.is_visible("div.bootbox.modal"):
            msg_erro = page.locator("div.bootbox-body").inner_text()
            logger.warn(f"Modal detectado antes de iniciar: {msg_erro}", details={"modal": msg_erro})
            # Tenta fechar
            page.click("button.bootbox-close-button, button[data-bb-handler='ok']", timeout=2000)
    except:
        pass

    # Navegação
    page.click('a[href*="Credenciamento"]')
    # Wait for URL or selector
    try:
        page.wait_for_url("**/CredenciamentoLista", timeout=TIMEOUT)
    except:
        pass 

    page.wait_for_selector('a[href="/Credenciamento/Index"]', timeout=TIMEOUT)
    page.click('a[href="/Credenciamento/Index"]')


def preencher_dados_pessoais(page, funcionario: dict) -> None:
    """Preenche a aba de dados pessoais do funcionário."""
    nome = funcionario["NOME"]
    nome_pai = funcionario.get("NOME_PAI", "")
    nome_mae = funcionario.get("NOME_MAE", "")
    cpf = funcionario["CPF"]

    page.wait_for_selector('#nome', timeout=TIMEOUT)
    page.fill('#nome', nome)

    # APELIDO (Obrigatorio)
    apelido = nome.split()[0]
    
    page.wait_for_selector('#apelido', timeout=TIMEOUT)
    page.evaluate("document.querySelector('#apelido').removeAttribute('disabled')")
    page.fill('#apelido', apelido)

    page.wait_for_selector('#nomePai', timeout=TIMEOUT)
    page.fill('#nomePai', nome_pai)

    page.wait_for_selector('#nomeMae', timeout=TIMEOUT)
    page.fill('#nomeMae', nome_mae)

    cpf_formatado = formatar_cpf(cpf)

    page.wait_for_selector('#cpf', timeout=TIMEOUT)
    page.fill('#cpf', cpf_formatado)

    # ESCOLARIDADE
    codigo_rm = funcionario.get("GRAUINSTRUCAO")
    valor_escolaridade = None
    if codigo_rm:
        codigo_rm = str(codigo_rm).strip().upper()
        valor_escolaridade = MAPA_ESCOLARIDADE.get(codigo_rm)

    if valor_escolaridade:
        page.wait_for_selector('#escolaridade', timeout=TIMEOUT)
        page.select_option('#escolaridade', value=valor_escolaridade)
    else:
        logger.warn(f"Escolaridade não mapeada: {codigo_rm}", details={"codigo_rm": codigo_rm})

    # ESTADO CIVIL
    codigo_ec = funcionario.get("ESTADOCIVIL")
    valor_est_civil = None
    if codigo_ec is not None:
        codigo_ec = str(codigo_ec).strip()
        valor_est_civil = MAPA_ESTADO_CIVIL.get(codigo_ec)

    if valor_est_civil:
        page.wait_for_selector('#estCivil', timeout=TIMEOUT)
        page.select_option('#estCivil', value=valor_est_civil)
    else:
        logger.warn(f"Estado civil não mapeado: {codigo_ec}", details={"codigo_ec": codigo_ec})

    # ESTADO DE NASCIMENTO
    estado_natal_rm = funcionario.get("ESTADONATAL")
    valor_estado_natal = None
    if estado_natal_rm:
        estado_natal_rm = str(estado_natal_rm).strip().upper()
        valor_estado_natal = MAPA_ESTADO_NATAL.get(estado_natal_rm)

    if valor_estado_natal:
        page.wait_for_selector('#estNasc', timeout=TIMEOUT)
        page.select_option('#estNasc', value=valor_estado_natal)
    else:
        logger.warn(f"Estado natal não mapeado: {estado_natal_rm}", details={"uf": estado_natal_rm})

    # PIS / PASEP
    pis_rm = funcionario.get("PISPASEP")
    pis_formatado = formatar_pis(pis_rm)

    if pis_formatado:
        page.wait_for_selector('#pisPasep', timeout=TIMEOUT)
        page.fill('#pisPasep', pis_formatado)

    # CIDADE DE NASCIMENTO
    cidade_rm = funcionario.get("NATURALIDADE")
    if cidade_rm:
        cidade_rm_norm = normalizar_texto(cidade_rm)
        page.wait_for_selector('#cidNasc', timeout=TIMEOUT)
        page.wait_for_function(
            """() => {
                const sel = document.querySelector('#cidNasc');
                return sel && sel.options && sel.options.length > 1;
            }""",
            timeout=TIMEOUT
        )
        opcoes = page.locator('#cidNasc option')
        total = opcoes.count()
        selecionado = False
        for i in range(total):
            texto_opcao = opcoes.nth(i).inner_text()
            texto_opcao_norm = normalizar_texto(texto_opcao)
            if texto_opcao_norm == cidade_rm_norm:
                valor = opcoes.nth(i).get_attribute("value")
                page.select_option('#cidNasc', value=valor)
                logger.info(f"Cidade selecionada (exata): {texto_opcao}", details={"cidade": texto_opcao})
                selecionado = True
                break
        if not selecionado:
            for i in range(total):
                texto_opcao = opcoes.nth(i).inner_text()
                texto_opcao_norm = normalizar_texto(texto_opcao)
                if cidade_rm_norm in texto_opcao_norm:
                    valor = opcoes.nth(i).get_attribute("value")
                    page.select_option('#cidNasc', value=valor)
                    logger.info(f"Cidade selecionada (fallback): {texto_opcao}", details={"cidade": texto_opcao, "original": cidade_rm})
                    selecionado = True
                    break
        if not selecionado:
            logger.warn(f"Cidade não encontrada no MetaX: {cidade_rm}", details={"cidade": cidade_rm})
    else:
        logger.warn("NATURALIDADE vazia no RM")

    # DATA DE NASCIMENTO
    data_nasc_rm = funcionario.get("DTNASCIMENTO")
    data_nasc_formatada = formatar_data(data_nasc_rm)

    if data_nasc_formatada:
        page.wait_for_selector('#dtNasc', timeout=TIMEOUT)
        page.fill('#dtNasc', data_nasc_formatada)
    else:
        logger.warn("Data de nascimento vazia")

    # NACIONALIDADE 
    page.wait_for_selector('#nacionalidade', timeout=TIMEOUT)
    page.select_option('#nacionalidade', value='1')

    # SEXO
    sexo_rm = funcionario.get("SEXO")
    valor_sexo = None
    if sexo_rm:
        sexo_rm = str(sexo_rm).strip().upper()
        valor_sexo = MAPA_SEXO.get(sexo_rm)

    if valor_sexo:
        page.wait_for_selector('#sexo', timeout=TIMEOUT)
        page.select_option('#sexo', value=valor_sexo)
    else:
        logger.warn(f"Sexo inválido ou não mapeado: {sexo_rm}", details={"sexo": sexo_rm})

    # EMAIL    
    email = funcionario.get("EMAIL", "")
    if email:
        page.wait_for_selector('#selecaoPadraoEmail', timeout=TIMEOUT)
        page.fill('#selecaoPadraoEmail', email)

    # TELEFONE EMERGENCIAL
    telefone_rm = funcionario.get("TELEFONE1", "")
    telefone_formatado = formatar_telefone_numerico(telefone_rm)

    if telefone_formatado:
        campo_tel = page.locator('#selecaoTelEmergencial')
        campo_tel.click()
        campo_tel.fill("")  
        campo_tel.type(telefone_formatado, delay=50)
    else:
        logger.warn(f"Telefone emergencial inválido ou vazio: {telefone_rm}", details={"fone": telefone_rm})


def preencher_documentos(page, funcionario: dict) -> None:
    """Preenche a aba de documentos (RG, CTPS, etc)."""
    orgamoemissor = funcionario.get("ORGEMISSORIDENT", "")
    numerorg = funcionario.get("CARTIDENTIDADE", "")
    dataemissao = funcionario.get("DTEMISSAOIDENT", "")
    numerocpts = funcionario.get("CARTEIRATRAB", "")
    seriecpts = funcionario.get("SERIECARTTRAB", "")
    datacpts = funcionario.get("DTCARTTRAB", "")

    # ORGAO EMISSOR
    page.wait_for_selector('#orgEmissorRG', timeout=TIMEOUT)
    page.fill('#orgEmissorRG', orgamoemissor)

    # UF DO RG
    uf_rg_rm = funcionario.get("UFCARTIDENT")
    valor_uf_rg = None
    if uf_rg_rm:
        uf_rg_rm = str(uf_rg_rm).strip().upper()
        valor_uf_rg = MAPA_ESTADO_NATAL.get(uf_rg_rm)

    if valor_uf_rg:
        page.wait_for_selector('#ufRG', timeout=TIMEOUT)
        page.select_option('#ufRG', value=valor_uf_rg)
    else:
        logger.warn(f"UF do RG não mapeada ou vazia: {uf_rg_rm}", details={"uf": uf_rg_rm})
    
    # NUMERO RG
    page.wait_for_selector('#numRG', timeout=TIMEOUT)
    page.fill('#numRG', numerorg)

    # EMISSAO RG
    dataemissao = formatar_data(dataemissao)
    if dataemissao:
        page.wait_for_selector('#dtEmissaoRG', timeout=TIMEOUT)
        page.fill('#dtEmissaoRG', dataemissao)
    else:
        logger.warn("Data de emissão do RG vazia")

    # CTPS DIGITAL
    page.wait_for_selector('#cmbCTPSDigital', timeout=TIMEOUT)
    page.check('#cmbCTPSDigital')

    # NUMERO CTPS
    page.wait_for_selector('#numCTPS', timeout=TIMEOUT)
    page.fill('#numCTPS', numerocpts)

    # SERIE CTPS
    page.wait_for_selector('#serieCTPS', timeout=TIMEOUT)
    page.fill('#serieCTPS', seriecpts)

    # ESTADO CTPS
    estadocpts = funcionario.get("UFCARTTRAB")
    valor_estado = None
    if estadocpts:
        estadocpts = str(estadocpts).strip().upper()
        valor_estado = MAPA_ESTADO_NATAL.get(estadocpts)

    if valor_estado:
        page.wait_for_selector('#ufCTPS', timeout=TIMEOUT)
        page.select_option('#ufCTPS', value=valor_estado)
    else:
        logger.warn(f"Estado natal não mapeado: {estadocpts}", details={"uf": estadocpts})
    
    # DATA CTPS
    data_formatada_cpts = formatar_data(datacpts)
    if data_formatada_cpts:
        page.wait_for_selector('#dtCTPS', timeout=TIMEOUT)
        page.fill('#dtCTPS', data_formatada_cpts)
    else:
        logger.warn("Data de nascimento vazia")


def preencher_endereco(page, funcionario: dict) -> None:
    """Preenche endereço e tenta buscar via CEP."""
    cep = funcionario.get("CEP", "")
    endereconumero = funcionario.get("NUMERO", "")

    page.wait_for_selector('a[href="#menu1"]', timeout=TIMEOUT)
    page.click('a[href="#menu1"]')

    def preencher_e_buscar_cep(cep_tentativa):
        # Force remove readonly if present
        page.evaluate("document.querySelector('#CEP').removeAttribute('readonly')")
        page.fill('#CEP', cep_tentativa)
        page.wait_for_selector("#btnPesquisarCep", state="visible")
        page.locator("#btnPesquisarCep").click()
        page.wait_for_timeout(3000) 

    preencher_e_buscar_cep(cep)

    # Verifica se o bairro foi preenchido
    bairro_preenchido = page.input_value("#nomeBairro").strip()

    if not bairro_preenchido:
        logger.warn(f"CEP {cep} não encontrou endereço. Tentando fallback...", details={"cep": cep})
        preencher_e_buscar_cep("79582034")

    # FALLBACK ENDEREÇO
    estado_rm = funcionario.get("ESTADO", "").strip().upper()
    bairro_rm = funcionario.get("BAIRRO", "").strip().upper()
    rua_rm = funcionario.get("RUA", "").strip().upper()

    # ESTADO
    estado_metax = page.locator("#comboEstado")
    estado_valor_atual = estado_metax.input_value().strip()

    if estado_rm in MAPA_ESTADO_NATAL and MAPA_ESTADO_NATAL[estado_rm] != estado_valor_atual:
        page.evaluate("""
            const sel = document.getElementById('comboEstado');
            sel.disabled = false;
        """)
        page.select_option("#comboEstado", value=MAPA_ESTADO_NATAL[estado_rm])
        page.locator("#comboEstado").press("Tab")
        logger.info(f"Estado ajustado via RM: {estado_rm}", details={"uf": estado_rm})

    # BAIRRO
    campo_bairro = page.locator("#nomeBairro")
    campo_bairro.wait_for(state="visible", timeout=TIMEOUT)
    page.evaluate("document.querySelector('#nomeBairro').disabled = false")
    bairro_metax = campo_bairro.input_value().strip()

    if (not bairro_metax) and bairro_rm:
        campo_bairro.click()
        campo_bairro.fill("")
        campo_bairro.type(bairro_rm, delay=50)
        logger.info(f"Bairro preenchido via RM: {bairro_rm}", details={"bairro": bairro_rm})

    # LOGRADOURO
    campo_logradouro = page.locator("#comboLogradouro")
    campo_logradouro.wait_for(state="visible", timeout=TIMEOUT)
    page.evaluate("document.querySelector('#comboLogradouro').disabled = false")
    logradouro_metax = campo_logradouro.input_value().strip()

    if (not logradouro_metax) and rua_rm:
        campo_logradouro.click()
        campo_logradouro.fill("")
        campo_logradouro.type(rua_rm, delay=50)
        logger.info(f"Logradouro preenchido via RM: {rua_rm}", details={"logradouro": rua_rm})

    # NUMERO
    campo_num = page.locator('input#numero.form-control.input')
    campo_num.wait_for(state="visible", timeout=TIMEOUT)
    campo_num.click()
    campo_num.press("Control+A")
    campo_num.press("Backspace")
    campo_num.type(str(endereconumero), delay=80)
    campo_num.press("Tab")

    logger.info(f"Número do endereço preenchido: {endereconumero}", details={"numero": endereconumero})


def preencher_dados_profissionais(page, funcionario: dict) -> bool:
    """Preenche dados profissionais (Cargo, Salário) e seleciona o cargo."""
    dataadmissao = funcionario.get("DATAADMISSAO", "")
    salario = funcionario.get("SALARIO", "")

    page.wait_for_selector('a[href="#menuProfissional"]', timeout=TIMEOUT)
    page.click('a[href="#menuProfissional"]')

    # DATA ADMISSAO
    data_formatada_admissao = formatar_data(dataadmissao)
    if data_formatada_admissao:
        page.wait_for_selector('#dtAdmissao', timeout=TIMEOUT)
        page.fill('#dtAdmissao', data_formatada_admissao)
    else:
        logger.warn("Data de admissao vazia")

    campo = page.locator('#salario')
    campo.wait_for(state="visible", timeout=TIMEOUT)
    campo.click()
    campo.press("Control+A")
    campo.press("Backspace")
    campo.type(str(salario), delay=50)
    campo.press("Tab")
    
    page.wait_for_selector('#horMens', timeout=TIMEOUT)
    page.select_option('#horMens', value='2')

    descricao_rm = funcionario["DESCRICAO_CARGO"].strip().upper()
    logger.info(f"Tentando cargo RM: {descricao_rm}", details={"cargo": descricao_rm})

    # 1. Tentativa
    selecionado = selecionar_cargo_por_descricao(page, descricao_rm)

    # 2. Tentativa
    if not selecionado:
        descricao_ajustada = ajustar_descricao_cargo(descricao_rm)
        if descricao_ajustada != descricao_rm:
            logger.info(f"Tentando cargo ajustado MetaX: {descricao_ajustada}", details={"cargo_ajustado": descricao_ajustada})
            selecionado = selecionar_cargo_por_descricao(page, descricao_ajustada)

    if not selecionado:
        logger.error(f"Cargo não encontrado no MetaX (RM/ajustado): {descricao_rm}", details={"cargo_rm": descricao_rm})
        return False
    
    # FORÇAR BLUR
    page.locator("#cargo").press("Tab")
    page.wait_for_timeout(500)
    return True


def salvar_cadastro(page, cpf: str) -> bool:
    """Clica em salvar rascunho e valida o sucesso da operação."""
    try:
        page.wait_for_function(
            """() => {
                const btn = document.querySelector('#btnSalvarRascunho');
                return btn && !btn.disabled;
            }""",
            timeout=TIMEOUT
        )

        btn_rascunho = page.locator("#btnSalvarRascunho")
        btn_rascunho.scroll_into_view_if_needed()
        page.wait_for_timeout(300)

        btn_rascunho.click()

        max_retries = 10
        for _ in range(max_retries):
            page.wait_for_timeout(1000)
            if "CredenciamentoLista" in page.url:
                logger.info("RASCUNHO SALVO COM SUCESSO NO METAX")
                return True
        
        logger.error(f"Rascunho NÃO foi salvo. URL atual: {page.url}", details={"url": page.url})

        try:
            alertas = page.locator(".alert, .validation-summary-errors, .bootbox-body").all_inner_texts()
            if alertas:
                logger.error(f"Mensagens de erro na tela: {alertas}", details={"alertas": alertas})
        except:
            pass

        page.screenshot(path=os.path.join(SCREENSHOT_DIR, f"erro_salvar_{cpf}.png"))
        return False

    except Exception as e:
        logger.error(f"Falha ao salvar rascunho: {e}", details={"error": str(e)})
        page.screenshot(path=os.path.join(SCREENSHOT_DIR, f"erro_excecao_salvar_{cpf}.png"))
        return False


def cadastrar_funcionario(page, funcionario: dict, caminho_foto: str = None) -> bool:
    """
    Função principal que orquestra todo o cadastro de um funcionário.
    
    Args:
        page: Página Playwright logada.
        funcionario (dict): Dicionário com dados do funcionário.
        caminho_foto (str, optional): Caminho pré-baixado da foto.
        
    Returns:
        bool: True se salvo com sucesso, False caso contrário.
    """
    nome = funcionario["NOME"]
    cpf = funcionario["CPF"]
    obra = funcionario.get("NUMERO_OBRA", "")
    
    logger.info(f"Cadastrando {nome} | CPF {cpf}", details={"nome": nome, "cpf": cpf, "obra": obra})

    navegar_para_cadastro(page)

    # FOTO
    caminho_final = caminho_foto 
    if not caminho_final:
        caminho_final = buscar_foto_por_cpf(PASTA_FOTOS, cpf)

    if caminho_final:
        anexar_foto(page, caminho_final)
    else:
        logger.info(f"Nenhuma foto encontrada para CPF {cpf} – seguindo sem foto", details={"cpf": cpf})

    preencher_dados_pessoais(page, funcionario)
    preencher_documentos(page, funcionario)
    preencher_endereco(page, funcionario)
    
    sucesso_cargo = preencher_dados_profissionais(page, funcionario)
    if not sucesso_cargo:
        return False

    return salvar_cadastro(page, cpf)
