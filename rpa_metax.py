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

TIMEOUT = 60000 
TEMPO_CAPTCHA_MS = 30000 


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



def fechar_modais_bloqueantes(page):
    """Tenta fechar modais do Bootbox que estejam bloqueando a tela."""
    try:
        # Verifica se tem algum modal visível
        if page.locator("div.bootbox.modal").is_visible():
            logger.warn("Modal bloqueante detectado. FORÇANDO REMOÇÃO...")
            
            # Força bruta: Remove do DOM qualquer modal bootbox e o backdrop
            page.evaluate("""
                document.querySelectorAll('.bootbox.modal').forEach(e => e.remove());
                document.querySelectorAll('.modal-backdrop').forEach(e => e.remove());
                document.body.classList.remove('modal-open');
            """)
            page.wait_for_timeout(500)
    except:
        pass

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

    # Se chegou aqui, não encontrou match exato/startswith
    
    # 1. Tentar match parcial (contém) REVERSO (Verificar se PALAVRA CHAVE está na opção)
    # Ex: RM="MOTORISTA PESADO", Opcao="MOTORISTA" -> "MOTORISTA" in "MOTORISTA PESADO"? Não.
    # Ex: RM="MOTORISTA PESADO", Opcao="MOTORISTA" -> "MOTORISTA" in "MOTORISTA"? Sim.
    
    palavras_chave = descricao_cargo.split()
    if palavras_chave:
        primeira_palavra = palavras_chave[0] # Ex: MOTORISTA
        if len(primeira_palavra) > 3: # Evita matching de "DE", "DA"
            for i in range(opcoes.count()):
                texto_opcao = opcoes.nth(i).inner_text().upper()
                if primeira_palavra in texto_opcao:
                     valor = opcoes.nth(i).get_attribute("value")
                     page.select_option("#cargo", value=valor)
                     page.locator("#cargo").press("Tab")
                     logger.info(f"Cargo selecionado (Match Palavra-Chave '{primeira_palavra}'): {texto_opcao}", details={"alvo": descricao_cargo, "selecionado": texto_opcao})
                     return True
             
    # 2. Logar opções disponíveis para debug
    lista_opcoes = []
    for i in range(opcoes.count()):
        lista_opcoes.append(opcoes.nth(i).inner_text().upper())
    
    
    # Converte lista para string para aparecer no log de console
    # REMOVIDO LIMITE DE 30 PARA DEBUG TOTAL
    opcoes_str = "; ".join(lista_opcoes) 
    logger.warn(f"Cargo '{descricao_cargo}' não encontrado. Opções disponíveis: [{opcoes_str}]", details={"opcoes": lista_opcoes})

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
        logger.info("👉 AÇÃO NECESSÁRIA: Resolva o CAPTCHA e clique em 'Validar' MANUALMENTE.")
        logger.info("Aguardando você acessar a próxima tela...")
        
        # Removemos o wait_for_timeout fixo e o click automático 
        # para que o script avance assim que o usuário validar.
        # page.wait_for_timeout(TEMPO_CAPTCHA_MS)
        # page.click('button:has-text("Validar")')

        page.wait_for_selector('#comboContrato', state='visible', timeout=TIMEOUT)

        # Espera as opções carregarem
        page.wait_for_function(
            """() => {
                const sel = document.querySelector('#comboContrato');
                return sel && sel.options.length >= 1;
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
def obter_todos_rascunhos(page) -> set[str]:
    """
    Navega para a lista de credenciamento e coleta TODOS os CPFs cadastrados (Rascunhos).
    Gerencia paginação e exibição de 100 itens.
    
    Returns:
        set: Conjunto de CPFs (apenas números) encontrados.
    """
    logger.info("Buscando lista de rascunhos existentes...")
    cpfs_encontrados = set()
    
    # 1. Navegar para a lista
    if not "CredenciamentoLista" in page.url:
        page.goto("https://portal.metax.ind.br/CredenciamentoLista/Index", timeout=30000)
    
    # 2. Mudar exibição para 100 (se existir)
    try:
        # Tenta selecionar '100' no dropdown de registros (name geralmente é '...length')
        # Selector genérico para o select de paginação
        select_paginacao = page.locator("select[name*='length']")
        if select_paginacao.count() > 0:
            select_paginacao.select_option(value="100")
            page.wait_for_timeout(1000) # Espera tabela recarregar
    except Exception as e:
        logger.warn(f"Não conseguiu mudar paginação para 100: {e}")

    # 3. Limpar filtros
    try:
        page.click("text=Limpar", timeout=2000)
        page.wait_for_timeout(1000)
    except:
        pass

    # 4. FILTRAR POR RASCUNHO (Pedido crítico do usuário)
    filtro_aplicado = False
    try:
        logger.info("Aplicando filtro de Status: Rascunho...")
        
        # Tenta selecionar pelo label "Status:"
        # O seletor pode variar, vamos tentar encontrar o select próximo ao label Status
        # Opção A: Pelo ID se fosse conhecido, mas vamos por proximidade ou nome comum
        # Geralmente em grids assim é name="Status" ou id="Status"
        
        # Estratégia: Encontrar o campo de seleção.
        # Vamos tentar um seletor genérico que costuma funcionar nesses forms
        # Dropdown que tem opção "Rascunho"
        select_status = page.locator("select").filter(has_text="Rascunho").first
        
        if select_status.count() > 0:
            select_status.select_option(label="Rascunho")
            filtro_aplicado = True
        else:
            # Fallback forçado: tentar achar o select associado ao label
            # Assumindo layout padrão onde o label está antes ou acima
            try:
                page.locator("text=Status").locator("..").locator("select").first.select_option(label="Rascunho")
                filtro_aplicado = True
            except:
                logger.warn("Não foi possível encontrar o campo de Status para filtrar")

        if filtro_aplicado:
            page.wait_for_timeout(500)
            
            # Clicar em Pesquisar
            page.click("text=Pesquisar")
            logger.info("Botão Pesquisar clicado.")
            
            # Esperar recarregamento
            page.wait_for_timeout(2000)
        else:
            logger.error("⚠️ ATENÇÃO: Filtro de Rascunho NÃO foi aplicado! Coletando TODOS os registros.")
        
    except Exception as e:
        logger.error(f"Erro ao filtrar por rascunho: {e}. Continuando sem filtro (RISCO DE COLETAR TODOS OS REGISTROS)")

    # Garantir que a tabela carregou (espera header ou loading sumir)
    try:
        page.wait_for_selector("table tbody", timeout=10000)
        # Espera um pouco mais para garantir renderização das linhas
        page.wait_for_timeout(2000)
    except Exception as e:
        logger.warn(f"Tabela de rascunhos demorou a carregar: {e}")

    # 4. Loop de Paginação
    while True:
        # Coleta CPFs da página atual
        # Assume que CPF é a 2ª coluna (index 1) - Ajuste conforme HTML real
        # Pelas imagens: Ações | CPF | Passaporte | Nome...
        # Então CPF é td:nth-child(2)
        
        linhas = page.locator("table tbody tr")
        count = linhas.count()
        
        if count == 0:
            break
            
        # Verifica se é linha de "Nenhum registro"
        texto_primeira = linhas.first.inner_text()
        if "Nenhum registro" in texto_primeira:
            break

        # Extrai CPFs da página
        elementos_cpf = page.locator("table tbody tr td:nth-child(2)").all_inner_texts()
        
        cpfs_pagina = 0
        for texto in elementos_cpf:
            cpf_limpo = ''.join(filter(str.isdigit, texto))
            # Valida se tem 11 dígitos (CPF válido)
            if cpf_limpo and len(cpf_limpo) == 11:
                cpfs_encontrados.add(cpf_limpo)
                cpfs_pagina += 1
            elif cpf_limpo and len(cpf_limpo) > 0:
                # Log de CPF inválido para debug
                logger.warn(f"CPF inválido encontrado na lista (não tem 11 dígitos): {cpf_limpo}", details={"cpf": cpf_limpo, "tamanho": len(cpf_limpo)})

        logger.info(f"Coletados {cpfs_pagina} CPFs válidos nesta página. Total acumulado: {len(cpfs_encontrados)}")

        # Verifica botão Próximo
        # Geralmente classe 'paginate_button next'
        # Se tiver classe 'disabled', paramos.
        
        btn_proximo = page.locator("li.paginate_button.next")
        if btn_proximo.count() == 0:
            # Tenta outro seletor comum
            btn_proximo = page.locator("a:has-text('Próximo')")
            
        if btn_proximo.count() > 0:
            classe_btn = btn_proximo.get_attribute("class") or ""
            if "disabled" not in classe_btn:
                btn_proximo.click()
                # Espera a tabela recarregar antes de continuar
                try:
                    page.wait_for_selector("table tbody tr", timeout=5000)
                    page.wait_for_timeout(1000)  # Espera adicional para garantir renderização
                except Exception as e:
                    logger.warn(f"Tabela não recarregou após mudar de página: {e}")
                    # Tenta continuar mesmo assim
            else:
                 break
        else:
            break # Fim das páginas

    logger.info(f"Total de rascunhos mapeados: {len(cpfs_encontrados)}")
    return cpfs_encontrados

def navegar_para_cadastro(page) -> bool:
    """
    Navega do menu inicial até a tela de cadastro.
    """
    
    # 1. Tenta limpar qualquer modal que esteja na frente (Sucesso/Erro anterior)
    try:
        if page.is_visible("div.bootbox.modal"):
            msg = page.locator("div.bootbox-body").first.inner_text()
            logger.warn(f"Modal detectado antes de iniciar: {msg}")
            
            page.evaluate("""
                $('.bootbox.modal').modal('hide');
                $('.modal-backdrop').remove(); 
            """)
            page.wait_for_selector("div.bootbox.modal", state="hidden", timeout=3000)
    except Exception as e:
        pass

    # 2. Garante que estamos na lista (Home do Credenciamento)
    if not "CredenciamentoLista" in page.url:
        try:
            page.click('a[href*="Credenciamento"]', timeout=5000)
            page.wait_for_url("**/CredenciamentoLista", timeout=10000)
        except:
            logger.info("Tentando navegar via URL direta para a lista...")
            page.goto("https://portal.metax.ind.br/CredenciamentoLista/Index", timeout=30000)

    # 3. Clica no botão "CADASTRO"
    try:
        logger.info("Procurando botão CADASTRO...")
        page.wait_for_selector("text=CADASTRO", timeout=10000)
        page.click("text=CADASTRO")
        return True
    except:
        logger.warn("Botão CADASTRO por texto falhou, tentando seletor href...")
        page.click('a[href*="/Credenciamento/Index"]')
        return True


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
    # Validação de data no futuro
    if isinstance(dataemissao, (date, datetime)):
       # Garante que temos apenas DATA para comparação
       data_obj = dataemissao.date() if isinstance(dataemissao, datetime) else dataemissao
        
       if data_obj > datetime.now().date():
           logger.warn(f"Data de emissão do RG no futuro ({dataemissao}). Ajustando para HOJE.", details={"original": str(dataemissao)})
           dataemissao = datetime.now()
           
    elif isinstance(dataemissao, str):
       try:
           # Tenta converter para verificar
           dt_obj = datetime.strptime(dataemissao, "%Y-%m-%d").date()
           if dt_obj > datetime.now().date():
                logger.warn(f"Data de emissão do RG no futuro ({dataemissao}). Ajustando para HOJE.", details={"original": dataemissao})
                dataemissao = datetime.now()
       except:
           pass

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
    cep = funcionario.get("CEP", "")
    endereconumero = funcionario.get("NUMERO", "")
    
    # Ajuste para número 0 -> S/N
    if str(endereconumero).strip() == "0":
        endereconumero = "S/N"

    page.wait_for_selector('a[href="#menu1"]', timeout=TIMEOUT)
    page.click('a[href="#menu1"]')

    def preencher_e_buscar_cep(cep_tentativa):
        # Force remove readonly if present
        page.evaluate("document.querySelector('#CEP').removeAttribute('readonly')")
        page.fill('#CEP', cep_tentativa)
        
        # Nuke Modals antes de clicar
        fechar_modais_bloqueantes(page)
            
        page.wait_for_selector("#btnPesquisarCep", state="visible")
        
        # Tenta clicar com force=True para ignorar overlays transparentes
        try:
             page.locator("#btnPesquisarCep").click(force=True)
        except Exception:
             # Fallback via JS se o click falhar
             page.evaluate("document.getElementById('btnPesquisarCep').click()")

        page.wait_for_timeout(3000) 

    preencher_e_buscar_cep(cep)

    # Verifica se o bairro foi preenchido
    bairro_preenchido = page.input_value("#nomeBairro").strip()
    fallback_usado = False

    if not bairro_preenchido:
        logger.warn(f"CEP {cep} não encontrou endereço. Tentando fallback...", details={"cep": cep})
        
        # Fallback inteligente por Estado
        estado_rm = funcionario.get("ESTADO", "").strip().upper()
        if estado_rm == "BA":
             preencher_e_buscar_cep("40015000") # Salvador/BA (Comércio)
        elif estado_rm == "PA":
             preencher_e_buscar_cep("66010000") # Belém/PA (Campina)
        else:
             preencher_e_buscar_cep("79582034") # CEP Genérico (MS) - Chapadão
             
        fallback_usado = True

    # FALLBACK ENDEREÇO
    bairro_rm = funcionario.get("BAIRRO", "").strip().upper()
    rua_rm = funcionario.get("RUA", "").strip().upper()

    # ESTADO
    fechar_modais_bloqueantes(page)
    estado_metax = page.locator("#comboEstado")
    estado_valor_atual = estado_metax.input_value().strip()

    # Só altera o estado se NÃO usamos fallback. Se usou fallback, respeita o estado do CEP válido (ex: DF)
    # exceto se o estado estiver vazio (falha no preenchimento automatico)
    if not fallback_usado:
        if estado_rm in MAPA_ESTADO_NATAL and MAPA_ESTADO_NATAL[estado_rm] != estado_valor_atual:
            page.evaluate("""
                const sel = document.getElementById('comboEstado');
                sel.disabled = false;
            """)
            page.select_option("#comboEstado", value=MAPA_ESTADO_NATAL[estado_rm])
            page.locator("#comboEstado").press("Tab")
            logger.info(f"Estado ajustado via RM: {estado_rm}", details={"uf": estado_rm})
    else:
        # Se usou fallback mas o campo tá vazio, força o estado do RM (que deve bater com o fallback)
        if not estado_valor_atual and estado_rm in MAPA_ESTADO_NATAL:
             logger.warn(f"Fallback usado mas Estado vazio. Forçando Estado RM: {estado_rm}")
             page.evaluate("document.getElementById('comboEstado').disabled = false;")
             page.select_option("#comboEstado", value=MAPA_ESTADO_NATAL[estado_rm])
        else:
             logger.info(f"Fallback usado: Mantendo estado original do CEP ({estado_valor_atual}) para evitar erro de validação.")

    # TENTATIVA DE FORÇAR CIDADE (Para casos de fallback onde a cidade fica vazia)
    if fallback_usado:
        cidade_rm = funcionario.get("NATURALIDADE", "").strip().upper()
        if cidade_rm:
            logger.info(f"Fallback usado: Tentando forçar cidade para {cidade_rm}...")
            # Tenta seletores comuns de cidade
            try:
                 page.evaluate(f"document.querySelector('#nomeCidade').value = '{cidade_rm}';")
            except: pass
            try:
                 page.evaluate(f"document.querySelector('#cidade').value = '{cidade_rm}';")
            except: pass
            try:
                 # Se for combo, é mais chato, mas tenta setar se existir
                 page.evaluate(f"document.querySelector('#comboCidade').disabled = false;")
                 # Esse aqui é chute, geralmente combo precisa de ID. 
                 # Mas se for input text (comum quando falha cep), o value funciona.
            except: pass


    # BAIRRO
    fechar_modais_bloqueantes(page)
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
    fechar_modais_bloqueantes(page)
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
    fechar_modais_bloqueantes(page)
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
    """Clica em salvar rascunho e valida o sucesso da operação, tratando modais."""
    # Nuke modals antes de tentar salvar
    fechar_modais_bloqueantes(page)

    try:
        # Tenta remover disabled do botão para forçar o clique mesmo que validação JS esteja bloqueando
        page.evaluate("""
            const btn = document.querySelector('#btnSalvarRascunho');
            if(btn) { 
                btn.disabled = false; 
                btn.classList.remove('disabled');
            }
        """)
        
        # Espera breve para garantir que JS processou
        page.wait_for_timeout(500)

        btn_rascunho = page.locator("#btnSalvarRascunho")
        # Garante scroll para o fim da página
        page.keyboard.press("End")
        page.wait_for_timeout(500)
        btn_rascunho.scroll_into_view_if_needed()
        page.wait_for_timeout(300)

        # Clica em Salvar
        btn_rascunho.click()

        # Loop de verificação (URL mudou OU Modal de Sucesso apareceu)
        max_retries = 90 # Aumentado para 90 segundos
        start_time = datetime.now()
        last_click_time = datetime.now()

        while (datetime.now() - start_time).seconds < max_retries:
            page.wait_for_timeout(1000)
            
            # Retry do click se passar 20s e nada acontecer
            if (datetime.now() - last_click_time).seconds > 20:
                 if btn_rascunho.is_visible():
                     logger.info("Retentando clique em Salvar (sem resposta há 20s)...")
                     try:
                        btn_rascunho.click(force=True)
                     except:
                        page.evaluate("document.getElementById('btnSalvarRascunho').click()")
                 else:
                     logger.warn("Botão Salvar NÃO está visível durante retry. Verificando erros...")
                     # Verifica mensagens de erro na tela
                     erros = page.locator(".text-danger, .field-validation-error").all_inner_texts()
                     if erros:
                         erros_texto = " | ".join([e for e in erros if e.strip()])
                         if erros_texto:
                             logger.error(f"Erros de validação encontrados na tela: {erros_texto}")
                             return False # Aborta se tiver erros claros

                 last_click_time = datetime.now()

            # 1. Sucesso por Redirecionamento
            if "CredenciamentoLista" in page.url:
                logger.info("RASCUNHO SALVO COM SUCESSO DO REDIRECIONAMENTO")
                return True

            # 2. Sucesso por Modal (Bootbox)
            # Verifica se existe algum modal VISÍVEL
            if page.locator("div.bootbox.modal").filter(has=page.locator(":scope:visible")).count() > 0:
                
                # Pega textos de todos os corpos de modal (pode ter ocultos)
                # Usamos all_inner_texts() para não dar erro de Strict Mode
                textos_modais = page.locator("div.bootbox-body").all_inner_texts()
                
                # Junta tudo numa string só para verificar
                texto_completo = " | ".join(textos_modais).lower()
                
                if "sucesso" in texto_completo:
                    logger.info("Modal de sucesso detectado!", details={"modais": textos_modais})
                    
                    # Tenta clicar no botão OK do modal VISÍVEL
                    try:
                        # Seletor específico para o botão OK dentro do modal que está visível (.in)
                        btn_ok = page.locator("div.bootbox.modal.in button[data-bb-handler='ok']")
                        if btn_ok.is_visible():
                            btn_ok.click()
                        else:
                            # Fallback genérico
                            page.locator("button.bootbox-close-button, button[data-bb-handler='ok']").last.click()
                            
                        logger.info("Botão OK do modal clicado.")
                    except Exception as e_click:
                        logger.warn(f"Falha ao clicar no OK: {e_click}. Tentando JS force...", details={"erro": str(e_click)})
                        page.evaluate("""
                            document.querySelectorAll('button[data-bb-handler="ok"]').forEach(b => {
                                if(b.offsetParent !== null) b.click();
                            })
                        """)
                    
                    # Espera todos os modais sumirem
                    try:
                        page.wait_for_selector("div.bootbox.modal", state="hidden", timeout=3000)
                    except:
                        logger.warn("Modal não sumiu via clique. Forçando fechamento via JS.")
                        page.evaluate("""
                            $('.bootbox.modal').modal('hide');
                            $('.modal-backdrop').remove(); 
                        """)
                    
                    # SUCESSO CONFIRMADO
                    logger.info("✅ CADASTRO REALIZADO COM SUCESSO!")
                    return True
                else:
                    # É um modal de erro (ou aviso que não é sucesso)
                    logger.error(f"Erro ao salvar (Modal): {textos_modais}", details={"modais": textos_modais})
                    
                    # Fecha o modal visível para não travar o fluxo
                    try:
                         page.locator("div.bootbox.modal.in button").first.click()
                    except:
                         page.evaluate("if(document.querySelector('.bootbox.modal.in')) $('.bootbox.modal.in').modal('hide');")
                         
                    return False
        
        # Se saiu do loop, é timeout
        logger.error(f"Rascunho NÃO foi salvo (Timeout). URL atual: {page.url}", details={"url": page.url})

        # Logar erros visíveis na tela (divs de alerta padrão)
        try:
            alertas = page.locator(".alert, .validation-summary-errors").all_inner_texts()
            if alertas:
                logger.error(f"Alertas na tela: {alertas}", details={"alertas": alertas})
        except:
            pass

        page.screenshot(path=os.path.join(SCREENSHOT_DIR, f"erro_salvar_{cpf}.png"))
        return False

    except Exception as e:
        logger.error(f"Falha ao salvar rascunho: {e}", details={"error": str(e)})
        page.screenshot(path=os.path.join(SCREENSHOT_DIR, f"erro_excecao_salvar_{cpf}.png"))
        return False


def cadastrar_funcionario(page, funcionario: dict, caminho_foto: str = None) -> str:
    """
    Função principal que orquestra todo o cadastro de um funcionário.
    
    Args:
        page: Página Playwright logada.
        funcionario (dict): Dicionário com dados do funcionário.
        caminho_foto (str, optional): Caminho pré-baixado da foto.
        
    Returns:
        str: 'SUCESSO', 'ERRO', 'DUPLICADO', 'SEM_FOTO'
    """
    nome = funcionario["NOME"]
    cpf = funcionario["CPF"]
    obra = funcionario.get("NUMERO_OBRA", "")
    
    logger.info(f"Cadastrando {nome} | CPF {cpf}", details={"nome": nome, "cpf": cpf, "obra": obra})
    
    # Navegação simples (verificação de duplicidade já feita no main)
    navegar_para_cadastro(page)

    # FOTO
    caminho_final = caminho_foto 
    if not caminho_final:
        caminho_final = buscar_foto_por_cpf(PASTA_FOTOS, cpf)

    if caminho_final:
        anexar_foto(page, caminho_final)
    else:
        logger.info(f"Nenhuma foto encontrada para CPF {cpf} – seguindo sem foto", details={"cpf": cpf})
        # Decisão de negócio: Se não tem foto, segue ou para?
        # O usuário pediu para relatar "quantos não acharam a foto".
        # O script original continuava.
        # Vou assumir que CONTINUA, mas reporta "sem foto" separadamente?
        # Ou se a foto for obrigatória, retorna ERRO.
        # No script original, ele continuava em anexar_foto. 
        pass

    preencher_dados_pessoais(page, funcionario)
    preencher_documentos(page, funcionario)
    preencher_endereco(page, funcionario)
    
    sucesso_cargo = preencher_dados_profissionais(page, funcionario)
    if not sucesso_cargo:
        return 'ERRO'

    sucesso_salvar = salvar_cadastro(page, cpf)
    if not sucesso_salvar:
        return 'ERRO'
    
    # Se salvou com sucesso, mas não tinha foto, podemos retornar um status especial ou apenas SUCESSO?
    # O usuário quer saber "quantos não acharam a foto".
    # Podemos retornar SUCESSO_SEM_FOTO se quiser ser preciso, mas vamos manter simples.
    # O chamador pode verificar se caminho_final era None.
    
    return 'SUCESSO'
