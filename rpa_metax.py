from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
import unicodedata
from datetime import date, datetime
import os
from PIL import Image
import tempfile
from custom_logger import logger


LOGIN = "984.656.002-87"
SENHA = "Enesa@2026*"
URL_LOGIN = "https://portal.metax.ind.br/SegLogin/"
TIMEOUT = 30000
TEMPO_CAPTCHA_MS = 10000 

PASTA_FOTOS = r"P:\MetaX\fotos_funcionarios"

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


def anexar_foto(page, caminho_foto):
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

MAPA_CARGOS_METAX = {
    "AUXILIAR DE SERVICOS GERAIS": "AUXILIAR DE SERVICOS GERAIS",
    "APROPRIADOR": "APROPRIADOR",
    "ENCARREGADO DE SOLDA I": "ENCARREGADO DE SOLDA",
    "ENCARREGADO DE MECANICA I": "ENCARREGADO DE MECANICA",
    "ASSISTENTE DE ADM DE PESSOAL I": "ASSISTENTE ADM PESSOAL",
    "ALMOXARIFE": "ALMOXARIFE",
    "CALDEIREIRO": "CALDEIRO",
    "ENCARREGADO DE MONTAGEM": "ENCARREGADO DE MONTAGEM",
    "AJUDANTE": "AJUDANTE",
    "TECNICO DE SEGURANCA DE TRABALHO III": "TECNICO DE SEGURANCA - INDIRETA",
    "INSPETOR DE SOLDA NIVEL I": "INSPETOR DE SOLDA N1",

    # ===== NOVOS (45 FUNCIONÁRIOS) =====
    "MECANICO MONTADOR": "MECANICO MONTADOR",
    "MONTADOR DE ANDAIME": "MONTADOR DE ANDAIME",
    "PINTOR INDUSTRIAL": "PINTOR INDUSTRIAL",
}

def ajustar_descricao_cargo(descricao_rm: str) -> str:
    return MAPA_CARGOS_METAX.get(descricao_rm.strip().upper(), descricao_rm.strip().upper())

MAPA_ESCOLARIDADE = {
    "1": "1",   # Analfabeto
    "2": "2",   # Fundamental incompleto
    "3": "3",   # Fundamental completo
    "4": "2",   # 6º ao 9º ano → fundamental incompleto
    "5": "3",   # Fundamental completo
    "6": "4",   # Médio incompleto
    "7": "5",   # Médio completo
    "8": "7",   # Superior incompleto
    "9": "8",   # Superior completo
    "A": "9",   # Pós-graduação incompleta
    "B": "10",  # Pós-graduação completa
    "C": "11",  # Mestrado
    "D": "11",  # Mestrado
    "E": "12",  # Doutorado
    "F": "12",  # Doutorado
    "G": "13",  # Outros
    "H": "13",  # Outros
}

MAPA_ESTADO_CIVIL = {
    "S": "1",  # Solteiro
    "E": "2",  # União Estável
    "C": "3",  # Casado
    "P": "4",  # Separado
    "I": "5",  # Divorciado
    "V": "6",  # Viúvo
}

MAPA_SEXO = {
    "F": "1",  # Feminino
    "M": "2",  # Masculino
}

MAPA_ESTADO_NATAL = {
    "AC": "12",
    "AL": "20",
    "AP": "15",
    "AM": "13",
    "BA": "21",
    "CE": "22",
    "DF": "11",
    "ES": "1",
    "GO": "10",
    "MA": "19",
    "MT": "8",
    "MS": "9",
    "MG": "4",
    "PA": "14",
    "PB": "23",
    "PR": "5",
    "PE": "25",
    "PI": "24",
    "RJ": "2",
    "RN": "26",
    "RS": "7",
    "RO": "17",
    "RR": "16",
    "SC": "6",
    "SP": "3",
    "SE": "27",
    "TO": "18",
    "OUTROS": "28"
}

def selecionar_cargo_por_descricao(page, descricao_cargo):
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

    # Se vier como string (ex: '1990-08-25')
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

    # Mantém somente números
    pis = ''.join(filter(str.isdigit, str(pis)))

    # Garante 11 dígitos (com zero à esquerda, se necessário)
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

# ==============================================================================
# NOVA FUNÇÃO DE INÍCIO DE SESSÃO (Retorna p, browser, page)
# ==============================================================================
def iniciar_sessao():
    # Inicia o Playwright, mas NÃO fecha (quem fecha é o main)
    # Obs: sync_playwright() deve ser usado com contexto. 
    # Para ser persistente, chamamos .start() manualmente
    p = sync_playwright().start()
    browser = p.chromium.launch(channel="chrome", headless=False)
    context = browser.new_context(ignore_https_errors=True)
    page = context.new_page()

    try:
        logger.info("Acessando a página de login...")
        page.goto(URL_LOGIN, timeout=TIMEOUT, wait_until="domcontentloaded")

        # LOGIN
        page.wait_for_selector('#txtLogin', timeout=TIMEOUT)
        page.fill('#txtLogin', LOGIN)

        page.wait_for_selector('#txtSenha', timeout=TIMEOUT)
        page.fill('#txtSenha', SENHA)
        
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
def cadastrar_funcionario(page, funcionario, caminho_foto=None):
    nome = funcionario["NOME"]
    cpf = funcionario["CPF"]
    obra = funcionario["NUMERO_OBRA"]
    nome_pai = funcionario.get("NOME_PAI", "")
    nome_mae = funcionario.get("NOME_MAE", "")
    email = funcionario.get("EMAIL", "")
    orgamoemissor = funcionario.get("ORGEMISSORIDENT", "")
    numerorg = funcionario.get("CARTIDENTIDADE", "")
    dataemissao = funcionario.get("DTEMISSAOIDENT", "")
    numerocpts = funcionario.get("CARTEIRATRAB", "")
    seriecpts = funcionario.get("SERIECARTTRAB", "")
    datacpts = funcionario.get("DTCARTTRAB", "")
    cep = funcionario.get("CEP", "")
    endereconumero = funcionario.get("NUMERO", "")
    dataadmissao = funcionario.get("DATAADMISSAO", "")
    salario = funcionario.get("SALARIO", "")

    logger.info(f"Cadastrando {nome} | CPF {cpf}", details={"nome": nome, "cpf": cpf, "obra": obra})

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
        pass # as vezes a URL não muda instantaneamente ou tem query params

    page.wait_for_selector('a[href="/Credenciamento/Index"]', timeout=TIMEOUT)
    page.click('a[href="/Credenciamento/Index"]')

    # ================= FOTO =================
    # Prioriza o caminho vindo do main (download recente)
    # Se não veio, tenta buscar na pasta (fallback)
    caminho_final = caminho_foto 
    if not caminho_final:
        caminho_final = buscar_foto_por_cpf(PASTA_FOTOS, cpf)

    if caminho_final:
        anexar_foto(page, caminho_final)
    else:
        logger.info(f"Nenhuma foto encontrada para CPF {cpf} – seguindo sem foto", details={"cpf": cpf})


    # PREENCHIMENTO – DADOS PESSOAIS

    page.wait_for_selector('#nome', timeout=TIMEOUT)
    page.fill('#nome', nome)

    # APELIDO (Obrigatorio)
    # Usando o primeiro nome como apelido
    apelido = nome.split()[0]
    
    page.wait_for_selector('#apelido', timeout=TIMEOUT)
    # Força habilitação do campo (caso esteja disabled)
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


    # CIDADE DE NASCIMENTO (EXATA → FALLBACK)

    cidade_rm = funcionario.get("NATURALIDADE")

    if cidade_rm:
        cidade_rm_norm = normalizar_texto(cidade_rm)

        # Espera o select existir
        page.wait_for_selector('#cidNasc', timeout=TIMEOUT)

        # Espera as opções carregarem (AJAX)
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

    #orgaoemissor
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
    
    #numero
    page.wait_for_selector('#numRG', timeout=TIMEOUT)
    page.fill('#numRG', numerorg)

    #emissao
    dataemissao = formatar_data(dataemissao)

    if dataemissao:
        page.wait_for_selector('#dtEmissaoRG', timeout=TIMEOUT)
        page.fill('#dtEmissaoRG', dataemissao)
    else:
        logger.warn("Data de emissão do RG vazia")

    #cpts digital
    page.wait_for_selector('#cmbCTPSDigital', timeout=TIMEOUT)
    page.check('#cmbCTPSDigital')

    #numerocpts
    page.wait_for_selector('#numCTPS', timeout=TIMEOUT)
    page.fill('#numCTPS', numerocpts)

    #seriecpts
    page.wait_for_selector('#serieCTPS', timeout=TIMEOUT)
    page.fill('#serieCTPS', seriecpts)

    #estadocpts
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
    
    #datacpts
    data_formatada_cpts = formatar_data(datacpts)

    if data_formatada_cpts:
        page.wait_for_selector('#dtCTPS', timeout=TIMEOUT)
        page.fill('#dtCTPS', data_formatada_cpts)
    else:
        logger.warn("Data de nascimento vazia")

    #endereco pessoal
    page.wait_for_selector('a[href="#menu1"]', timeout=TIMEOUT)
    page.click('a[href="#menu1"]')

    def preencher_e_buscar_cep(cep_tentativa):
        # Force remove readonly if present
        page.evaluate("document.querySelector('#CEP').removeAttribute('readonly')")
        page.fill('#CEP', cep_tentativa)
        page.wait_for_selector("#btnPesquisarCep", state="visible")
        page.locator("#btnPesquisarCep").click()
        # Aumentando espera para garantir que campos carreguem
        page.wait_for_timeout(3000) 

    preencher_e_buscar_cep(cep)

    # Verifica se o bairro foi preenchido (sinal de sucesso)
    bairro_preenchido = page.input_value("#nomeBairro").strip()

    if not bairro_preenchido:
        logger.warn(f"CEP {cep} não encontrou endereço. Tentando fallback...", details={"cep": cep})
        preencher_e_buscar_cep("79582034")
    # ================= FALLBACK ENDEREÇO (ESTADO / BAIRRO / LOGRADOURO) =================

    estado_rm = funcionario.get("ESTADO", "").strip().upper()     # ex: "MA"
    bairro_rm = funcionario.get("BAIRRO", "").strip().upper()
    rua_rm = funcionario.get("RUA", "").strip().upper()

    # ---------- ESTADO ----------
    estado_metax = page.locator("#comboEstado")
    estado_valor_atual = estado_metax.input_value().strip()

    # Se o CEP não trouxe estado correto
    if estado_rm in MAPA_ESTADO_NATAL and MAPA_ESTADO_NATAL[estado_rm] != estado_valor_atual:
        page.evaluate("""
            const sel = document.getElementById('comboEstado');
            sel.disabled = false;
        """)
        page.select_option("#comboEstado", value=MAPA_ESTADO_NATAL[estado_rm])
        page.locator("#comboEstado").press("Tab")
        logger.info(f"Estado ajustado via RM: {estado_rm}", details={"uf": estado_rm})

    # ---------- BAIRRO ----------
    campo_bairro = page.locator("#nomeBairro")
    campo_bairro.wait_for(state="visible", timeout=TIMEOUT)
    
    # Força o campo a ficar habilitado caso o JS do site o bloqueie
    page.evaluate("document.querySelector('#nomeBairro').disabled = false")

    bairro_metax = campo_bairro.input_value().strip()

    if (not bairro_metax) and bairro_rm:
        campo_bairro.click()
        campo_bairro.fill("")
        campo_bairro.type(bairro_rm, delay=50)
        logger.info(f"Bairro preenchido via RM: {bairro_rm}", details={"bairro": bairro_rm})

    # ---------- LOGRADOURO ----------
    campo_logradouro = page.locator("#comboLogradouro")
    campo_logradouro.wait_for(state="visible", timeout=TIMEOUT)

    # Força habilitação
    page.evaluate("document.querySelector('#comboLogradouro').disabled = false")

    logradouro_metax = campo_logradouro.input_value().strip()

    if (not logradouro_metax) and rua_rm:
        campo_logradouro.click()
        campo_logradouro.fill("")
        campo_logradouro.type(rua_rm, delay=50)
        logger.info(f"Logradouro preenchido via RM: {rua_rm}", details={"logradouro": rua_rm})

    # -------- NÚMERO (campo real dentro da aba) --------
    campo_num = page.locator('input#numero.form-control.input')

    campo_num.wait_for(state="visible", timeout=TIMEOUT)

    campo_num.click()
    campo_num.press("Control+A")
    campo_num.press("Backspace")

    campo_num.type(str(endereconumero), delay=80)

    # BLUR obrigatório para o MetaX validar
    campo_num.press("Tab")

    logger.info(f"Número do endereço preenchido corretamente: {endereconumero}", details={"numero": endereconumero})
    #dados colaborador
    page.wait_for_selector('a[href="#menuProfissional"]', timeout=TIMEOUT)
    page.click('a[href="#menuProfissional"]')

    #data admissao
    data_formatada_admissao = formatar_data(dataadmissao)

    if data_formatada_admissao:
        page.wait_for_selector('#dtAdmissao', timeout=TIMEOUT)
        page.fill('#dtAdmissao', data_formatada_admissao)
    else:
        logger.warn("Data de nascimento vazia")

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

    # 1️⃣ tentativa: cargo original do RM
    selecionado = selecionar_cargo_por_descricao(page, descricao_rm)

    # 2️⃣ tentativa: cargo ajustado para MetaX
    if not selecionado:
        descricao_ajustada = ajustar_descricao_cargo(descricao_rm)

        if descricao_ajustada != descricao_rm:
            logger.info(f"Tentando cargo ajustado MetaX: {descricao_ajustada}", details={"cargo_ajustado": descricao_ajustada})
            selecionado = selecionar_cargo_por_descricao(page, descricao_ajustada)

    # ❌ falhou nas duas
    if not selecionado:
        logger.error(f"Cargo não encontrado no MetaX (RM/ajustado): {descricao_rm}", details={"cargo_rm": descricao_rm})
        return False

    
    # FORÇAR BLUR PARA DISPARAR JS DO METAX
    page.locator("#cargo").press("Tab")
    page.wait_for_timeout(500)

    # ================== SALVAR COMO RASCUNHO (CONTROLADO) ==================
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

        btn_rascunho.click()  # <<< SEM force=True

        # aguarda resposta do sistema (Aumentado para 10s)
        # Verifica mudança de URL em loop
        max_retries = 10
        for _ in range(max_retries):
            page.wait_for_timeout(1000)
            if "CredenciamentoLista" in page.url:
                logger.info("RASCUNHO SALVO COM SUCESSO NO METAX")
                return True
        
        # Se chegou aqui, falhou
        logger.error(f"Rascunho NÃO foi salvo. URL atual: {page.url}", details={"url": page.url})

        # Tenta capturar texto de erro na tela (alertas do sistema)
        try:
            alertas = page.locator(".alert, .validation-summary-errors, .bootbox-body").all_inner_texts()
            if alertas:
                logger.error(f"Mensagens de erro na tela: {alertas}", details={"alertas": alertas})
        except:
            pass

        page.screenshot(path=f"erro_salvar_{cpf}.png")
        return False

    except Exception as e:
        logger.error(f"Falha ao salvar rascunho: {e}", details={"error": str(e)})
        page.screenshot(path=f"erro_excecao_salvar_{cpf}.png")
        return False
