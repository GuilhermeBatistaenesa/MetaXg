from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
import unicodedata
from datetime import date, datetime
import os
from PIL import Image
import tempfile
import re
from custom_logger import logger
from utils import (
    reduzir_foto_para_metax, buscar_foto_por_cpf, ajustar_descricao_cargo,
    formatar_telefone_numerico, formatar_data, normalizar_texto, 
    formatar_pis, formatar_cpf
)
from mappings import (
    MAPA_CARGOS_METAX, MAPA_CARGOS_CODFUNCAO_METAX, MAPA_ESCOLARIDADE, MAPA_ESTADO_CIVIL,
    MAPA_SEXO, MAPA_ESTADO_NATAL, MAPA_CARGOS_OVERRIDE_POR_CONTRATO
)


from config import (
    METAX_LOGIN, METAX_PASSWORD, METAX_URL_LOGIN,
    METAX_CONTRATO_MECANICA_VALUE, METAX_CONTRATO_MECANICA_LABEL,
    METAX_CONTRATO_ELETROMECANICA_VALUE, METAX_CONTRATO_ELETROMECANICA_LABEL,
    METAX_CONTRATO_DEFAULT_VALUE, METAX_CONTRATO_DEFAULT_LABEL,
    FOTOS_BUSCA_DIRS
)
from output_manager import OutputManager, KIND_SCREENSHOTS, KIND_JSON

TIMEOUT = 60000 
TEMPO_CAPTCHA_MS = 180000 


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
        modais_visiveis = page.locator("div.bootbox.modal:visible").count()
        if modais_visiveis > 0:
            textos = []
            try:
                textos = [t for t in page.locator("div.bootbox-body").all_inner_texts() if t.strip()]
            except Exception:
                textos = []
            logger.warn("Modal bloqueante detectado. FORCANDO REMOCAO...", details={"modais": textos})

            # Tenta clicar no OK antes de remover
            try:
                btn_ok = page.locator("div.bootbox.modal:visible button[data-bb-handler='ok']")
                if btn_ok.count() > 0:
                    btn_ok.first.click()
            except Exception:
                pass

        # Forca bruta: remove do DOM qualquer modal bootbox e o backdrop
        page.evaluate("""
            document.querySelectorAll('.bootbox.modal').forEach(e => e.remove());
            document.querySelectorAll('.modal-backdrop').forEach(e => e.remove());
            document.body.classList.remove('modal-open');
        """)
        page.wait_for_timeout(300)
    except Exception:
        pass


def selecionar_opcao_select(page, selector: str, value: str | None = None, label: str | None = None) -> bool:
    """
    Tenta selecionar uma opcao de um SELECT usando value ou label.
    Faz fallback por matching normalizado (sem acento / case-insensitive).
    """
    def _try_value(val: str) -> bool:
        try:
            page.select_option(selector, value=val)
            return True
        except Exception:
            return False

    def _try_label(lab: str) -> bool:
        try:
            page.select_option(selector, label=lab)
            return True
        except Exception:
            return False

    if value and _try_value(value):
        return True
    if label and _try_label(label):
        return True

    alvo = label or value
    if not alvo:
        return False

    try:
        ok = page.evaluate(
            """(sel, alvo) => {
                const norm = (s) => (s || '').normalize('NFD').replace(/\\p{Diacritic}/gu, '').toUpperCase().trim();
                const el = document.querySelector(sel);
                if (!el || !el.options) return false;
                const target = norm(alvo);
                let opt = Array.from(el.options).find(o => norm(o.textContent || '') === target);
                if (!opt) {
                    opt = Array.from(el.options).find(o => norm(o.value || '') === target);
                }
                if (!opt) return false;
                el.value = opt.value;
                el.dispatchEvent(new Event('change', { bubbles: true }));
                el.dispatchEvent(new Event('input', { bubbles: true }));
                return true;
            }""",
            selector,
            alvo,
        )
        return bool(ok)
    except Exception:
        return False


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
            matches = []
            for i in range(opcoes.count()):
                texto_opcao = opcoes.nth(i).inner_text().upper()
                if primeira_palavra in texto_opcao:
                    valor = opcoes.nth(i).get_attribute("value")
                    if valor:
                        matches.append((valor, texto_opcao))
            # Evita selecionar cargo errado quando ha mais de uma opcao com mesma palavra-chave.
            if len(matches) == 1:
                valor, texto_opcao = matches[0]
                page.select_option("#cargo", value=valor)
                page.locator("#cargo").press("Tab")
                logger.info(
                    f"Cargo selecionado (Match Palavra-Chave '{primeira_palavra}'): {texto_opcao}",
                    details={"alvo": descricao_cargo, "selecionado": texto_opcao},
                )
                return True
            if len(matches) > 1:
                logger.warn(
                    f"Match por palavra-chave ambiguo para cargo '{descricao_cargo}'.",
                    details={"palavra_chave": primeira_palavra, "matches": [m[1] for m in matches]},
                )
             
    # 2. Logar opções disponíveis para debug
    lista_opcoes = []
    for i in range(opcoes.count()):
        lista_opcoes.append(opcoes.nth(i).inner_text().upper())
    
    
    # Converte lista para string para aparecer no log de console
    # REMOVIDO LIMITE DE 30 PARA DEBUG TOTAL
    opcoes_str = "; ".join(lista_opcoes) 
    logger.warn(f"Cargo '{descricao_cargo}' não encontrado. Opções disponíveis: [{opcoes_str}]", details={"opcoes": lista_opcoes})

    return False


def _aplicar_override_contrato(descricao: str, contrato_chave: str | None) -> str:
    if not contrato_chave:
        return descricao
    override_map = MAPA_CARGOS_OVERRIDE_POR_CONTRATO.get(contrato_chave, {})
    return override_map.get(descricao, descricao)


# ==============================================================================
# NOVA FUNÇÃO DE INÍCIO DE SESSÃO (Retorna p, browser, page)
# ==============================================================================
def _listar_opcoes_contrato(page) -> list[dict]:
    return page.evaluate(
        """() => {
            const sel = document.querySelector('#comboContrato');
            if (!sel) return [];
            return Array.from(sel.options || []).map(o => ({
                value: o.value || "",
                label: (o.textContent || "").trim()
            }));
        }"""
    )


def _selecionar_contrato(page, contrato_value: str | None, contrato_label: str | None):
    opcoes = _listar_opcoes_contrato(page)
    if not opcoes:
        logger.warn("Nenhuma opcao encontrada no combo de contrato.")

    def _norm_label(texto: str | None) -> str:
        if not texto:
            return ""
        return " ".join(texto.upper().split())

    alvo_value = contrato_value or None
    alvo_label = contrato_label or None
    if not (alvo_value or alvo_label):
        if METAX_CONTRATO_DEFAULT_VALUE:
            alvo_value = METAX_CONTRATO_DEFAULT_VALUE
        elif METAX_CONTRATO_DEFAULT_LABEL:
            alvo_label = METAX_CONTRATO_DEFAULT_LABEL
        else:
            logger.warn("Nenhum contrato informado. Mantendo selecao atual do portal.")
            return

    escolhida = None
    if alvo_value:
        for o in opcoes:
            if o.get("value") == alvo_value:
                escolhida = o
                break
    if not escolhida and alvo_label:
        alvo_norm = _norm_label(alvo_label)
        for o in opcoes:
            if _norm_label(o.get("label")) == alvo_norm:
                escolhida = o
                break
    if not escolhida and alvo_label:
        alvo_up = _norm_label(alvo_label)
        for o in opcoes:
            if alvo_up and alvo_up in _norm_label(o.get("label")):
                escolhida = o
                break

    if not escolhida:
        logger.error(
            "Contrato nao encontrado no combo.",
            details={"value": alvo_value, "label": alvo_label, "opcoes": opcoes},
        )
        raise ValueError("Contrato nao encontrado no combo de selecao.")

    page.select_option('#comboContrato', value=escolhida.get("value", ""), timeout=5000)
    logger.info(
        "Contrato selecionado",
        details={"value": escolhida.get("value"), "label": escolhida.get("label")},
    )


def iniciar_sessao(headless: bool = False, contrato_value: str | None = None, contrato_label: str | None = None):
    """
    Inicia o browser, realiza login e navega até a tela inicial do sistema.
    
    Returns:
        tuple: (playwright_instance, browser_instance, page_instance)
    """
    # Inicia o Playwright, mas NÃO fecha (quem fecha é o main)
    # Obs: sync_playwright() deve ser usado com contexto. 
    # Para ser persistente, chamamos .start() manualmente
    p = sync_playwright().start()
    try:
        browser = p.chromium.launch(channel="chrome", headless=headless)
    except Exception as e:
        msg = "Falha ao iniciar navegador. Verifique se os browsers do Playwright estao instalados (python -m playwright install chromium)."
        logger.error(msg, details={"erro": str(e)})
        raise RuntimeError(msg) from e
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
        logger.info("ACAO NECESSARIA: Resolva o CAPTCHA e clique em 'Validar' MANUALMENTE.")
        logger.info("Aguardando você acessar a próxima tela...")
        
        # Removemos o wait_for_timeout fixo e o click automático 
        # para que o script avance assim que o usuário validar.
        # page.wait_for_timeout(TEMPO_CAPTCHA_MS)
        # page.click('button:has-text("Validar")')

        page.wait_for_selector('#comboContrato', state='visible', timeout=TEMPO_CAPTCHA_MS)

        # Espera as opções carregarem
        page.wait_for_function(
            """() => {
                const sel = document.querySelector('#comboContrato');
                return sel && sel.options.length >= 1;
            }""",
            timeout=TEMPO_CAPTCHA_MS
        )

        _selecionar_contrato(page, contrato_value, contrato_label)

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
            logger.error("ATENCAO: Filtro de Rascunho NAO foi aplicado! Coletando TODOS os registros.")
        
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
        if not selecionar_opcao_select(page, '#escolaridade', value=valor_escolaridade, label=valor_escolaridade):
            logger.warn(
                "Escolaridade nao encontrada no combo do MetaX.",
                details={"codigo_rm": codigo_rm, "valor": valor_escolaridade},
            )
            if selecionar_opcao_select(page, '#escolaridade', value="Outros", label="Outros"):
                logger.warn("Fallback de escolaridade aplicado: Outros.", details={"codigo_rm": codigo_rm})
    else:
        logger.warn(f"Escolaridade não mapeada: {codigo_rm}", details={"codigo_rm": codigo_rm})
        page.wait_for_selector('#escolaridade', timeout=TIMEOUT)
        if selecionar_opcao_select(page, '#escolaridade', value="Outros", label="Outros"):
            logger.warn("Fallback de escolaridade aplicado para codigo nao mapeado.", details={"codigo_rm": codigo_rm})

    # ESTADO CIVIL
    codigo_ec = funcionario.get("ESTADOCIVIL")
    valor_est_civil = None
    if codigo_ec is not None:
        codigo_ec = str(codigo_ec).strip()
        valor_est_civil = MAPA_ESTADO_CIVIL.get(codigo_ec)

    if valor_est_civil:
        page.wait_for_selector('#estCivil', timeout=TIMEOUT)
        if not selecionar_opcao_select(page, '#estCivil', value=valor_est_civil, label=valor_est_civil):
            logger.warn(
                "Estado civil nao encontrado no combo do MetaX.",
                details={"codigo_rm": codigo_ec, "valor": valor_est_civil},
            )
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
        logger.info("Aguardando carregamento de cidades...")
        try:
            page.wait_for_function(
                """() => {
                    const sel = document.querySelector('#cidNasc');
                    return sel && sel.options && sel.options.length > 1;
                }""",
                timeout=15000,
            )
        except Exception as e:
            logger.warn(
                "Lista de cidades nao carregou a tempo. Pulando selecao de cidade.",
                details={"cidade": cidade_rm, "error": str(e)},
            )
            cidade_rm = None

        if cidade_rm:
            try:
                opcoes = page.eval_on_selector_all(
                    "#cidNasc option",
                    "els => els.map(e => ({text: (e.textContent || '').trim(), value: e.value}))",
                )
            except Exception as e:
                logger.warn(
                    "Falha ao ler lista de cidades. Pulando selecao.",
                    details={"cidade": cidade_rm, "error": str(e)},
                )
                opcoes = []

            selecionado = False
            for opt in opcoes:
                texto_opcao = opt.get("text", "")
                texto_opcao_norm = normalizar_texto(texto_opcao)
                if texto_opcao_norm == cidade_rm_norm:
                    valor = opt.get("value", "")
                    if valor:
                        page.select_option('#cidNasc', value=valor)
                        logger.info(f"Cidade selecionada (exata): {texto_opcao}", details={"cidade": texto_opcao})
                        selecionado = True
                        break
            if not selecionado:
                for opt in opcoes:
                    texto_opcao = opt.get("text", "")
                    texto_opcao_norm = normalizar_texto(texto_opcao)
                    if cidade_rm_norm in texto_opcao_norm:
                        valor = opt.get("value", "")
                        if valor:
                            page.select_option('#cidNasc', value=valor)
                            logger.info(
                                f"Cidade selecionada (fallback): {texto_opcao}",
                                details={"cidade": texto_opcao, "original": cidade_rm},
                            )
                            selecionado = True
                            break
            if not selecionado:
                logger.warn(f"Cidade nao encontrada no MetaX: {cidade_rm}", details={"cidade": cidade_rm})
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
        if not selecionar_opcao_select(page, '#sexo', value=valor_sexo, label=valor_sexo):
            logger.warn(
                "Sexo nao encontrado no combo do MetaX.",
                details={"codigo_rm": sexo_rm, "valor": valor_sexo},
            )
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
    CEP_FALLBACK = "79582034"
    UF_FALLBACK = "MS"
    CIDADE_FALLBACK = "CHAPADAO DO SUL"
    BAIRRO_FALLBACK = "CENTRO"
    LOGRADOURO_FALLBACK = "RUA SEM NOME"

    cep = funcionario.get("CEP", "")
    endereconumero = funcionario.get("NUMERO", "")
    estado_rm = (funcionario.get("ESTADO") or "").strip().upper()
    if not estado_rm:
        logger.warn("Estado (RM) vazio; usando fallback UF=SP", details={"cpf": funcionario.get("CPF")})
        estado_rm = "SP"
    
    def _normalizar_numero(valor):
        digits = "".join([c for c in str(valor) if c.isdigit()])
        if not digits:
            return "0"
        digits = digits.lstrip("0") or "0"
        return digits

    def _valor_invalido(valor: str) -> bool:
        valor_norm = normalizar_texto(valor)
        if not valor_norm:
            return True
        return valor_norm in {
            "0",
            "SELECIONE",
            "SELECIONE...",
            "SELECIONAR",
            "SELECIONE UM",
            "SELECIONE O BAIRRO",
            "SELECIONE O LOGRADOURO",
        }

    def _set_campo_endereco(selector: str, valor: str, allow_first: bool = True) -> dict:
        try:
            return page.evaluate(
                """(sel, val, allowFirst) => {
                    const norm = (s) => (s || '').toUpperCase().trim();

                    const el = document.querySelector(sel);
                    if (!el) return { ok: false, reason: 'not_found' };

                    el.removeAttribute('readonly');
                    el.removeAttribute('disabled');
                    if (el.readOnly) el.readOnly = false;
                    if (el.disabled) el.disabled = false;

                    const desired = norm(val);

                    const pickFirst = (options) => {
                        const opt = (options || []).find(o => {
                            const v = (o.value || '').trim();
                            return v && v !== '0';
                        });
                        return opt || null;
                    };

                    if ((el.tagName || '').toUpperCase() === 'SELECT') {
                        let opt = null;
                        if (val) {
                            opt = Array.from(el.options || []).find(o => {
                                return norm(o.textContent || '') === desired || norm(o.value || '') === desired;
                            });
                        }
                        if (!opt && allowFirst) opt = pickFirst(Array.from(el.options || []));
                        if (!opt) return { ok: false, reason: 'no_option' };
                        el.value = opt.value;
                        el.dispatchEvent(new Event('change', { bubbles: true }));
                        el.dispatchEvent(new Event('input', { bubbles: true }));
                        return { ok: true, selected: (opt.textContent || opt.value || '').trim() };
                    }

                    const listId = el.getAttribute('list');
                    if (listId) {
                        const list = document.getElementById(listId);
                        if (list) {
                            let opt = null;
                            if (val) {
                                opt = Array.from(list.options || []).find(o => norm(o.value || '') === desired);
                            }
                            if (!opt && allowFirst) opt = (list.options || []).length ? list.options[0] : null;
                            if (opt) {
                                el.value = opt.value || opt.textContent || '';
                                el.dispatchEvent(new Event('input', { bubbles: true }));
                                el.dispatchEvent(new Event('change', { bubbles: true }));
                                return { ok: true, selected: (opt.value || opt.textContent || '').trim() };
                            }
                        }
                    }

                    if (val) {
                        el.value = val;
                        el.dispatchEvent(new Event('input', { bubbles: true }));
                        el.dispatchEvent(new Event('change', { bubbles: true }));
                        return { ok: true, selected: val };
                    }

                    return { ok: false, reason: 'no_value' };
                }""",
                selector,
                valor,
                allow_first,
            )
        except Exception:
            return {"ok": False}

    

    def _snapshot_endereco(stage: str) -> dict:
        try:
            snap = page.evaluate(
                """() => {
                    const info = (el) => {
                        if (!el) return { value: "", tag: "", options: 0, disabled: false, readonly: false };
                        let options = 0;
                        if ((el.tagName || '').toUpperCase() === 'SELECT') {
                            options = (el.options || []).length;
                        } else {
                            const listId = el.getAttribute('list');
                            if (listId) {
                                const list = document.getElementById(listId);
                                if (list && list.options) options = list.options.length;
                            }
                        }
                        return {
                            value: (el.value || '').trim(),
                            tag: (el.tagName || '').toUpperCase(),
                            options,
                            disabled: !!el.disabled,
                            readonly: !!el.readOnly,
                        };
                    };

                    const cidadeEl =
                        document.querySelector('#comboCidade') ||
                        document.querySelector('select[id*="Cidade"], input[id*="Cidade"], select[name*="Cidade"], input[name*="Cidade"]');

                    return {
                        cep: (document.querySelector('#CEP')?.value || '').trim(),
                        bairro: info(document.querySelector('#nomeBairro')),
                        logradouro: info(document.querySelector('#comboLogradouro')),
                        estado: info(document.querySelector('#comboEstado')),
                        cidade: info(cidadeEl),
                        numero: info(document.querySelector('#numero')),
                    };
                }"""
            )
        except Exception:
            snap = {}

        snap["stage"] = stage
        return snap

    def _campo_ok(info: dict) -> bool:
        try:
            valor = (info or {}).get("value", "")
            if not _valor_invalido(valor):
                return True
            return (info or {}).get("options", 0) > 1
        except Exception:
            return False

    def _cep_parece_valido(snap: dict) -> bool:
        return _campo_ok(snap.get("logradouro", {}))

    endereconumero = _normalizar_numero(endereconumero)

    page.wait_for_selector('a[href="#menu1"]', timeout=TIMEOUT)
    page.click('a[href="#menu1"]')

    def preencher_e_buscar_cep(cep_tentativa):
        # Normaliza CEP para 8 digitos (sem hifen)
        cep_digits = "".join([c for c in str(cep_tentativa) if c.isdigit()])
        if len(cep_digits) == 8:
            cep_formatado = f"{cep_digits[:5]}-{cep_digits[5:]}"
            cep_input = cep_digits
        else:
            cep_formatado = str(cep_tentativa)
            cep_input = cep_digits or cep_formatado

        # Forca setar valor via JS (campo pode estar readonly/disabled)
        page.evaluate(
            """(val) => {
                const el = document.querySelector('#CEP');
                if (!el) return;
                el.removeAttribute('readonly');
                el.removeAttribute('disabled');
                el.value = val;
                el.dispatchEvent(new Event('input', { bubbles: true }));
                el.dispatchEvent(new Event('change', { bubbles: true }));
            }""",
            cep_input,
        )

        # Tentativa adicional de digitar (alguns formularios so atualizam com input real)
        try:
            campo = page.locator("#CEP")
            campo.click(force=True)
            campo.press("Control+A")
            campo.press("Backspace")
            campo.type(str(cep_input), delay=20)
        except Exception:
            pass
        
        # Nuke Modals antes de clicar
        fechar_modais_bloqueantes(page)
            
        page.wait_for_selector("#btnPesquisarCep", state="visible")
        
        # Tenta clicar com force=True para ignorar overlays transparentes
        try:
             page.locator("#btnPesquisarCep").click(force=True)
        except Exception:
             # Fallback via JS se o click falhar
             page.evaluate("document.getElementById('btnPesquisarCep').click()")

        # Se o portal retornar modal de erro imediatamente, nao faz espera longa
        try:
            page.wait_for_timeout(800)
            modal = page.locator("div.bootbox.modal:visible")
            if modal.count() > 0:
                textos = []
                try:
                    textos = [t for t in page.locator("div.bootbox-body").all_inner_texts() if t.strip()]
                except Exception:
                    textos = []
                logger.warn("CEP: erro imediato no portal", details={"cep": cep_formatado, "modal": " | ".join(textos)})
                try:
                    btn_ok = page.locator("div.bootbox.modal:visible button[data-bb-handler='ok']")
                    if btn_ok.count() > 0:
                        btn_ok.first.click()
                    else:
                        page.locator("div.bootbox.modal:visible button").first.click()
                except Exception:
                    page.evaluate("if(document.querySelector('.bootbox.modal.in')) $('.bootbox.modal.in').modal('hide');")
                return
        except Exception:
            pass

        page.wait_for_timeout(3000)

        for tentativa in range(2):
            try:
                page.wait_for_function("""() => {
                    const val = (el) => (el && (el.value || '')).trim();
                    const norm = (s) => (s || '').toUpperCase().trim();
                    const valid = (s) => s && s !== '0' && s !== 'SELECIONE' && s !== 'SELECIONE...';

                    const bairro = document.querySelector('#nomeBairro');
                    const logradouro = document.querySelector('#comboLogradouro');

                    const cidadeSel = document.querySelector('#comboCidade') ||
                        document.querySelector('select[id*="Cidade"], select[name*="Cidade"]');

                    const selectOk = (el) => el && (el.tagName || '').toUpperCase() === 'SELECT' && (el.options || []).length > 1;

                    const bairroOk = selectOk(bairro) || (bairro && !bairro.disabled && !bairro.readOnly && valid(norm(val(bairro))));
                    const logOk = selectOk(logradouro) || (logradouro && !logradouro.disabled && !logradouro.readOnly && valid(norm(val(logradouro))));
                    const cidadeOk = cidadeSel && (cidadeSel.options || []).length > 1;

                    return bairroOk || logOk || cidadeOk;
                }""", timeout=20000)
                break
            except Exception as e:
                if tentativa == 0:
                    logger.warn(
                        "CEP: resposta lenta, tentando novamente",
                        details={"cep": cep_formatado, "error": str(e)},
                    )
                    try:
                        page.locator("#btnPesquisarCep").click(force=True)
                    except Exception:
                        page.evaluate("document.getElementById('btnPesquisarCep').click()")
                    page.wait_for_timeout(1500)
                    continue
                logger.warn("CEP: resposta nao carregou a tempo", details={"cep": cep_formatado, "error": str(e)})

    preencher_e_buscar_cep(cep)

    # Verifica se o bairro foi preenchido
    bairro_preenchido = ""
    try:
        bairro_preenchido = page.input_value("#nomeBairro").strip()
    except Exception:
        bairro_preenchido = ""
    fallback_usado = False

    snap_cep = _snapshot_endereco("apos_busca_cep")
    logger.info("Endereco snapshot apos CEP", details=snap_cep)

    # Dados RM para complementar endereco
    bairro_rm = funcionario.get("BAIRRO", "").strip().upper()
    rua_rm = funcionario.get("RUA", "").strip().upper()

    # Se logradouro vier invalido, tenta ajustar via RM antes do fallback
    if not _campo_ok(snap_cep.get("logradouro", {})) and rua_rm:
        result = _set_campo_endereco("#comboLogradouro", rua_rm, allow_first=True)
        if result.get("ok"):
            logger.info("Logradouro ajustado via RM (pre-fallback)", details={"logradouro": result.get("selected", rua_rm)})
        snap_cep = _snapshot_endereco("apos_logradouro_rm")

    if not _campo_ok(snap_cep.get("bairro", {})) and bairro_rm:
        result = _set_campo_endereco("#nomeBairro", bairro_rm, allow_first=True)
        if result.get("ok"):
            logger.info("Bairro ajustado via RM (pre-fallback)", details={"bairro": result.get("selected", bairro_rm)})
        snap_cep = _snapshot_endereco("apos_bairro_rm")

    if not _cep_parece_valido(snap_cep):
        logger.warn(f"CEP {cep} não encontrou endereço. Tentando fallback...", details={"cep": cep})

        # Fallback fixo (orientacao MetaX)
        preencher_e_buscar_cep(CEP_FALLBACK) # CEP Generico (MS) - Chapadao
             
        fallback_usado = True

        snap_fallback = _snapshot_endereco("apos_fallback_cep")
        logger.info("Endereco snapshot apos fallback CEP", details=snap_fallback)

    # FALLBACK ENDERECO

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
        valor_fallback = MAPA_ESTADO_NATAL.get(UF_FALLBACK)
        if valor_fallback and estado_valor_atual != valor_fallback:
            logger.warn(f"Fallback usado: Forcando Estado para {UF_FALLBACK}")
            page.evaluate("document.getElementById('comboEstado').disabled = false;")
            page.select_option("#comboEstado", value=valor_fallback)
        else:
            logger.info(f"Fallback usado: Mantendo estado atual ({estado_valor_atual}).")

    # Se fallback foi usado, manter cidade/bairro/logradouro conforme o CEP valido
    def _forcar_cidade_fallback():
        try:
            ok = page.evaluate("""(cidade) => {
                const alvo = (cidade || "").trim().toUpperCase();
                const candidatos = Array.from(document.querySelectorAll(
                    'select[id*="Cidade"], select[name*="Cidade"], input[id*="Cidade"], input[name*="Cidade"]'
                ));
                let aplicou = false;
                for (const el of candidatos) {
                    if (!el) continue;
                    el.removeAttribute('readonly');
                    el.removeAttribute('disabled');
                    if (el.tagName === "SELECT") {
                        let opt = Array.from(el.options || []).find(o => (o.textContent || "").trim().toUpperCase() === alvo);
                        if (!opt) {
                            opt = Array.from(el.options || []).find(o => (o.value || "").trim() && (o.value || "") !== "0");
                        }
                        if (opt) {
                            el.value = opt.value;
                            el.dispatchEvent(new Event('change', { bubbles: true }));
                            el.dispatchEvent(new Event('input', { bubbles: true }));
                            aplicou = true;
                        }
                    } else {
                        el.value = cidade;
                        el.dispatchEvent(new Event('input', { bubbles: true }));
                        el.dispatchEvent(new Event('change', { bubbles: true }));
                        aplicou = true;
                    }
                }
                return aplicou;
            }""", CIDADE_FALLBACK);
            return bool(ok)
        except Exception:
            return False

    if fallback_usado:
        logger.info("Fallback usado: mantendo endereco do CEP (sem sobrescrever cidade/bairro/logradouro).")
        if not _forcar_cidade_fallback():
            logger.warn("Fallback usado: nao foi possivel forcar cidade.")

    # BAIRRO
    fechar_modais_bloqueantes(page)
    page.wait_for_selector("#nomeBairro", state="visible", timeout=TIMEOUT)
    bairro_metax = ""
    try:
        bairro_metax = page.input_value("#nomeBairro").strip()
    except Exception:
        bairro_metax = ""

    if _valor_invalido(bairro_metax):
        if (not fallback_usado) and bairro_rm:
            result = _set_campo_endereco("#nomeBairro", bairro_rm, allow_first=True)
            if result.get("ok"):
                logger.info(
                    f"Bairro ajustado via RM: {bairro_rm}",
                    details={"bairro": result.get("selected", bairro_rm)},
                )
            else:
                result = _set_campo_endereco("#nomeBairro", "", allow_first=True)
                if result.get("ok"):
                    logger.warn("Bairro ajustado pela primeira opcao disponivel.")
                else:
                    _set_campo_endereco("#nomeBairro", BAIRRO_FALLBACK, allow_first=False)
                    logger.warn(f"Bairro preenchido com padrao '{BAIRRO_FALLBACK}'.")
        else:
            result = _set_campo_endereco("#nomeBairro", "", allow_first=True)
            if result.get("ok"):
                logger.warn("Fallback usado: Bairro selecionado pela primeira opcao disponivel.")
            else:
                _set_campo_endereco("#nomeBairro", BAIRRO_FALLBACK, allow_first=False)
                logger.warn(f"Fallback usado: Bairro preenchido com padrao '{BAIRRO_FALLBACK}'.")

    # LOGRADOURO
    fechar_modais_bloqueantes(page)
    page.wait_for_selector("#comboLogradouro", state="visible", timeout=TIMEOUT)
    logradouro_metax = ""
    try:
        logradouro_metax = page.input_value("#comboLogradouro").strip()
    except Exception:
        logradouro_metax = ""

    if _valor_invalido(logradouro_metax):
        if (not fallback_usado) and rua_rm:
            result = _set_campo_endereco("#comboLogradouro", rua_rm, allow_first=True)
            if result.get("ok"):
                logger.info(
                    f"Logradouro ajustado via RM: {rua_rm}",
                    details={"logradouro": result.get("selected", rua_rm)},
                )
            else:
                result = _set_campo_endereco("#comboLogradouro", "", allow_first=True)
                if result.get("ok"):
                    logger.warn("Logradouro ajustado pela primeira opcao disponivel.")
                else:
                    _set_campo_endereco("#comboLogradouro", LOGRADOURO_FALLBACK, allow_first=False)
                    logger.warn(f"Logradouro preenchido com padrao '{LOGRADOURO_FALLBACK}'.")
        else:
            result = _set_campo_endereco("#comboLogradouro", "", allow_first=True)
            if result.get("ok"):
                logger.warn("Fallback usado: Logradouro selecionado pela primeira opcao disponivel.")
            else:
                _set_campo_endereco("#comboLogradouro", LOGRADOURO_FALLBACK, allow_first=False)
                logger.warn(f"Fallback usado: Logradouro preenchido com padrao '{LOGRADOURO_FALLBACK}'.")

    if fallback_usado and endereconumero == "0":
        endereconumero = "1"
        logger.warn("Fallback usado: Numero ajustado para 1 (campo nao aceita S/N).")

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

    snap_final = _snapshot_endereco("final_endereco")
    logger.info("Endereco snapshot final", details=snap_final)


def preencher_dados_profissionais(page, funcionario: dict, contrato_chave: str | None = None) -> bool:
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
    cod_funcao = (funcionario.get("CODFUNCAO") or "").strip().upper()

    descricao_base = descricao_rm
    if cod_funcao:
        descricao_cod = MAPA_CARGOS_CODFUNCAO_METAX.get(cod_funcao)
        if descricao_cod:
            descricao_base = descricao_cod
            logger.info("Tentando cargo por CODFUNCAO", details={"cod_funcao": cod_funcao, "cargo_metax": descricao_cod})
        else:
            logger.info("CODFUNCAO sem mapeamento; usando descricao RM", details={"cod_funcao": cod_funcao, "cargo_rm": descricao_rm})
    else:
        logger.info(f"Tentando cargo RM: {descricao_rm}", details={"cargo": descricao_rm})

    # Monta candidatos de forma ordenada com override por contrato para lidar
    # com diferencas entre catalogos de Mecanica e Eletromecanica.
    candidatos = []

    def _push_candidato(valor: str | None):
        if not valor:
            return
        v = valor.strip().upper()
        if not v:
            return
        v = _aplicar_override_contrato(v, contrato_chave)
        if v not in candidatos:
            candidatos.append(v)

    _push_candidato(descricao_base)
    _push_candidato(descricao_rm)
    _push_candidato(ajustar_descricao_cargo(descricao_rm))
    if descricao_rm in MAPA_CARGOS_METAX:
        _push_candidato(MAPA_CARGOS_METAX[descricao_rm])

    selecionado = False
    for idx, candidato in enumerate(candidatos):
        if idx == 0:
            logger.info("Tentando cargo principal", details={"contrato": contrato_chave, "cargo": candidato})
        else:
            logger.info("Tentando cargo fallback", details={"contrato": contrato_chave, "cargo": candidato})
        if selecionar_cargo_por_descricao(page, candidato):
            selecionado = True
            break

    if not selecionado:
        logger.error(
            "Cargo nao encontrado no MetaX para o contrato atual.",
            details={
                "cargo_rm": descricao_rm,
                "cod_funcao": cod_funcao,
                "contrato": contrato_chave,
                "candidatos_testados": candidatos,
            },
        )
        logger.error(
            f"Cargo nao encontrado no MetaX (RM/ajustado): {descricao_rm}",
            details={"cargo_rm": descricao_rm, "cod_funcao": cod_funcao, "contrato": contrato_chave},
        )
        return False

    # FORÇAR BLUR
    page.locator("#cargo").press("Tab")
    page.wait_for_timeout(500)
    return True


def salvar_cadastro(page, cpf: str, output_manager: OutputManager) -> dict:
    """Clica em salvar rascunho e retorna o resultado factual do salvamento (sem declarar sucesso final)."""
    fechar_modais_bloqueantes(page)

    try:
        page.evaluate("""
            const btn = document.querySelector('#btnSalvarRascunho');
            if(btn) { 
                btn.disabled = false; 
                btn.classList.remove('disabled');
            }
        """)

        page.wait_for_timeout(500)

        btn_rascunho = page.locator("#btnSalvarRascunho")
        page.keyboard.press("End")
        page.wait_for_timeout(500)
        btn_rascunho.scroll_into_view_if_needed()
        page.wait_for_timeout(300)
        btn_rascunho.click()

        max_retries = 90
        start_time = datetime.now()
        last_click_time = datetime.now()

        while (datetime.now() - start_time).seconds < max_retries:
            page.wait_for_timeout(1000)

            if (datetime.now() - last_click_time).seconds > 20:
                if btn_rascunho.is_visible():
                    logger.info("Retentando clique em Salvar (sem resposta ha 20s)...")
                    try:
                        btn_rascunho.click(force=True)
                    except Exception:
                        page.evaluate("document.getElementById('btnSalvarRascunho').click()")
                else:
                    logger.warn("Botao Salvar NAO esta visivel durante retry. Verificando erros...")
                    erros = page.locator(".text-danger, .field-validation-error").all_inner_texts()
                    if erros:
                        erros_texto = " | ".join([e for e in erros if e.strip()])
                        if erros_texto:
                            logger.error(f"Erros de validacao encontrados na tela: {erros_texto}")
                            return {"attempted": True, "saved": False, "error": "Erros de validacao na tela."}

                last_click_time = datetime.now()

            if "CredenciamentoLista" in page.url:
                logger.info("Rascunho salvo (confirmacao por redirecionamento).")
                return {"attempted": True, "saved": True, "error": "", "detail": "confirmado_por_redirecionamento"}

            if page.locator("div.bootbox.modal").filter(has=page.locator(":scope:visible")).count() > 0:
                textos_modais = page.locator("div.bootbox-body").all_inner_texts()
                texto_completo = " | ".join(textos_modais).lower()

                if "sucesso" in texto_completo:
                    logger.info("Modal de confirmacao detectado.", details={"modais": textos_modais})

                    try:
                        btn_ok = page.locator("div.bootbox.modal.in button[data-bb-handler='ok']")
                        if btn_ok.is_visible():
                            btn_ok.click()
                        else:
                            page.locator("button.bootbox-close-button, button[data-bb-handler='ok']").last.click()

                        logger.info("Botao OK do modal clicado.")
                    except Exception as e_click:
                        logger.warn(f"Falha ao clicar no OK: {e_click}. Tentando JS force...", details={"erro": str(e_click)})
                        page.evaluate("""
                            document.querySelectorAll('button[data-bb-handler="ok"]').forEach(b => {
                                if(b.offsetParent !== null) b.click();
                            })
                        """)

                    try:
                        page.wait_for_selector("div.bootbox.modal", state="hidden", timeout=3000)
                    except Exception:
                        logger.warn("Modal nao sumiu via clique. Forcando fechamento via JS.")
                        page.evaluate("""
                            $('.bootbox.modal').modal('hide');
                            $('.modal-backdrop').remove(); 
                        """)

                    return {"attempted": True, "saved": True, "error": "", "detail": "confirmado_por_modal"}
                else:
                    logger.error(f"Erro ao salvar (Modal): {textos_modais}", details={"modais": textos_modais})

                    try:
                        page.locator("div.bootbox.modal.in button").first.click()
                    except Exception:
                        page.evaluate("if(document.querySelector('.bootbox.modal.in')) $('.bootbox.modal.in').modal('hide');")

                    return {"attempted": True, "saved": False, "error": "Erro ao salvar (modal).", "detail": ""}

        logger.error(f"Rascunho NAO foi salvo (Timeout). URL atual: {page.url}", details={"url": page.url})

        try:
            alertas = page.locator(".alert, .validation-summary-errors").all_inner_texts()
            if alertas:
                logger.error(f"Alertas na tela: {alertas}", details={"alertas": alertas})
        except Exception:
            pass

        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        filename = f"erro_salvar_{cpf}_{timestamp}__{output_manager.execution_id}.png"
        data = page.screenshot()
        output_manager.save_screenshot_bytes(filename, data)
        return {"attempted": True, "saved": False, "error": "Timeout ao salvar rascunho.", "detail": ""}

    except Exception as e:
        logger.error(f"Falha ao salvar rascunho: {e}", details={"error": str(e)})
        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        filename = f"erro_excecao_salvar_{cpf}_{timestamp}__{output_manager.execution_id}.png"
        data = page.screenshot()
        output_manager.save_screenshot_bytes(filename, data)
        return {"attempted": True, "saved": False, "error": str(e), "detail": ""}


def marcar_sem_foto_quando_disponivel(page) -> bool:
    """
    Tenta marcar opcao de "sem foto"/equivalente, caso exista no formulario.
    Retorna True quando algum controle foi marcado.
    """
    try:
        marcado = page.evaluate("""
            () => {
                const textosAlvo = ["sem foto", "nao possui foto", "foto nao", "sem imagem"];

                const norm = (s) => (s || "")
                    .normalize("NFD")
                    .replace(/[\\u0300-\\u036f]/g, "")
                    .toLowerCase();

                const labels = Array.from(document.querySelectorAll("label"));
                for (const label of labels) {
                    const texto = norm(label.textContent || "");
                    if (!textosAlvo.some(t => texto.includes(t))) continue;

                    const forId = label.getAttribute("for");
                    let target = forId ? document.getElementById(forId) : null;
                    if (!target) target = label.querySelector("input[type='checkbox'], input[type='radio']");
                    if (!target) continue;

                    if (target.type === "checkbox" || target.type === "radio") {
                        target.checked = true;
                    }
                    target.dispatchEvent(new Event("input", { bubbles: true }));
                    target.dispatchEvent(new Event("change", { bubbles: true }));
                    label.click();
                    return true;
                }

                const candidatos = Array.from(
                    document.querySelectorAll("input[type='checkbox'], input[type='radio']")
                );
                for (const input of candidatos) {
                    const id = norm(input.id || "");
                    const name = norm(input.name || "");
                    if (
                        id.includes("semfoto") || id.includes("sem_foto") || id.includes("nofoto") || id.includes("no_photo") ||
                        name.includes("semfoto") || name.includes("sem_foto") || name.includes("nofoto") || name.includes("no_photo")
                    ) {
                        input.checked = true;
                        input.dispatchEvent(new Event("input", { bubbles: true }));
                        input.dispatchEvent(new Event("change", { bubbles: true }));
                        return true;
                    }
                }

                return false;
            }
        """)
        return bool(marcado)
    except Exception as e:
        logger.warn("Falha ao tentar marcar opcao sem foto.", details={"error": str(e)})
        return False


def cadastrar_funcionario(
    page,
    funcionario: dict,
    output_manager: OutputManager,
    caminho_foto: str = None,
    contrato_chave: str | None = None,
) -> dict:
    """
    Funcao principal que orquestra todo o cadastro de um funcionario.

    Returns:
        dict: {"attempted": bool, "saved": bool, "error": str|None, "no_photo": bool}
    """
    nome = funcionario["NOME"]
    cpf = funcionario["CPF"]
    obra = funcionario.get("NUMERO_OBRA", "")

    logger.info(f"Cadastrando {nome} | CPF {cpf}", details={"nome": nome, "cpf": cpf, "obra": obra})

    # Navegacao simples (verificacao de duplicidade ja feita no main)
    navegar_para_cadastro(page)

    caminho_final = caminho_foto
    if not caminho_final:
        caminho_final = buscar_foto_por_cpf(FOTOS_BUSCA_DIRS, cpf)

    no_photo = False
    if caminho_final:
        anexar_foto(page, caminho_final)
    else:
        no_photo = True
        logger.info(f"Nenhuma foto encontrada para CPF {cpf} - seguindo sem foto", details={"cpf": cpf})

    preencher_dados_pessoais(page, funcionario)
    preencher_documentos(page, funcionario)
    preencher_endereco(page, funcionario)

    sucesso_cargo = preencher_dados_profissionais(page, funcionario, contrato_chave=contrato_chave)
    if not sucesso_cargo:
        return {"attempted": True, "saved": False, "no_photo": no_photo, "error": "Cargo nao encontrado no MetaX.", "detail": ""}

    resultado_salvar = salvar_cadastro(page, cpf, output_manager)
    if no_photo and not resultado_salvar.get("saved"):
        marcou_sem_foto = marcar_sem_foto_quando_disponivel(page)
        if marcou_sem_foto:
            logger.info(
                "Opcao sem foto marcada. Retentando salvar rascunho.",
                details={"cpf": cpf},
            )
            resultado_salvar = salvar_cadastro(page, cpf, output_manager)

    if not resultado_salvar.get("saved"):
        return {
            "attempted": True,
            "saved": False,
            "no_photo": no_photo,
            "error": resultado_salvar.get("error") or "Falha ao salvar rascunho.",
            "detail": resultado_salvar.get("detail") or "",
        }

    return {"attempted": True, "saved": True, "no_photo": no_photo, "error": "", "detail": resultado_salvar.get("detail") or ""}


def verificar_cadastro(page, funcionario: dict, output_manager: OutputManager, max_paginas: int = 3) -> tuple[bool, str]:
    """
    Verifica se o CPF aparece na lista de rascunhos apos o salvamento.
    Retorna (True, msg) se encontrado; caso contrario (False, msg).
    """
    try:
        cpf = funcionario.get("CPF")
        cpf_limpo = "".join(filter(str.isdigit, str(cpf)))
        if not cpf_limpo:
            return False, "CPF invalido para verificacao."

        if "CredenciamentoLista" not in page.url:
            page.goto("https://portal.metax.ind.br/CredenciamentoLista/Index", timeout=30000)

        # Aplica filtro de Status: Rascunho (mesma estrategia do fluxo principal)
        filtro_aplicado = False
        try:
            select_status = page.locator("select").filter(has_text="Rascunho").first
            if select_status.count() > 0:
                select_status.select_option(label="Rascunho")
                filtro_aplicado = True
            else:
                page.locator("text=Status").locator("..").locator("select").first.select_option(label="Rascunho")
                filtro_aplicado = True

            if filtro_aplicado:
                page.wait_for_timeout(500)
                page.click("text=Pesquisar")
                page.wait_for_timeout(1500)
        except Exception as e:
            logger.warn(f"Falha ao aplicar filtro de rascunho: {e}")

        search_usado = False
        # Se existir campo de busca (DataTables), usa para filtrar pelo CPF
        try:
            search_input = page.locator("input[type='search']").first
            if search_input.count() > 0 and search_input.is_visible():
                search_usado = True
                search_input.fill(cpf_limpo)
                page.wait_for_timeout(1000)
        except Exception:
            # Sem busca ou falha no campo, segue para varredura
            pass

        linhas_texto = []
        # Varredura paginada limitada
        pagina_atual = 0
        while pagina_atual < max_paginas:
            try:
                page.wait_for_selector("table tbody", timeout=10000)
            except Exception:
                pass

            linhas = page.locator("table tbody tr")
            for i in range(linhas.count()):
                texto = linhas.nth(i).inner_text()
                linhas_texto.append(texto)
                encontrados = re.findall(r"\\b\\d{11}\\b", texto)
                if cpf_limpo in encontrados:
                    return True, "CPF encontrado na lista de rascunhos."

            btn_proximo = page.locator("li.paginate_button.next")
            if btn_proximo.count() > 0:
                classe_btn = btn_proximo.get_attribute("class") or ""
                if "disabled" in classe_btn:
                    break
                btn_proximo.click()
                page.wait_for_timeout(1000)
                pagina_atual += 1
            else:
                break

        detalhe = f"cpf nao encontrado apos {pagina_atual + 1} tentativas"
        _registrar_evidencia_verificacao(
            page=page,
            output_manager=output_manager,
            cpf=cpf_limpo,
            filtro_aplicado=filtro_aplicado,
            search_usado=search_usado,
            linhas_texto=linhas_texto,
        )
        return False, detalhe
    except Exception as e:
        try:
            _registrar_evidencia_verificacao(
                page=page,
                output_manager=output_manager,
                cpf="".join(filter(str.isdigit, str(funcionario.get("CPF")))),
                filtro_aplicado=False,
                search_usado=False,
                linhas_texto=[],
            )
        except Exception:
            pass
        return False, f"Erro na verificacao: {e}"


def _registrar_evidencia_verificacao(
    page,
    output_manager: OutputManager,
    cpf: str,
    filtro_aplicado: bool,
    search_usado: bool,
    linhas_texto: list[str],
):
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    filename = f"verify_fail_{cpf}_{timestamp}__{output_manager.execution_id}.png"
    try:
        data = page.screenshot(timeout=60000)
        output_manager.save_screenshot_bytes(filename, data)
    except Exception as e:
        logger.warn("Falha ao gerar screenshot de verificacao", details={"erro": str(e)})

    debug = {
        "url": page.url,
        "total_linhas": len(linhas_texto),
        "linhas_sample": linhas_texto[:5],
        "filtro_aplicado": filtro_aplicado,
        "search_usado": search_usado,
    }
    output_manager.write_json(KIND_JSON, f"verify_debug_{cpf}_{timestamp}__{output_manager.execution_id}.json", debug)
