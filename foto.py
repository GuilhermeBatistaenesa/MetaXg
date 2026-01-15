from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
import os

# ================== CONFIGURAÇÕES ==================
LOGIN = "984.656.002-87"
SENHA = "Enesa@2026*"
URL_LOGIN = "https://portal.metax.ind.br/SegLogin/"
TIMEOUT = 30000
TEMPO_CAPTCHA_MS = 10000

PASTA_FOTOS = r"P:\MetaX\fotos_funcionarios"


# ================== BUSCAR FOTO PELO CPF ==================
def buscar_foto_por_cpf(pasta_fotos, cpf):
    cpf_numerico = ''.join(filter(str.isdigit, cpf))

    if not os.path.exists(pasta_fotos):
        return None

    for arquivo in os.listdir(pasta_fotos):
        nome = arquivo.lower()
        if cpf_numerico in nome and nome.endswith((".jpg", ".jpeg", ".png")):
            return os.path.join(pasta_fotos, arquivo)

    return None


# ================== ANEXAR FOTO ==================
def anexar_foto(page, caminho_foto):
    page.wait_for_selector("#avatar", state="attached", timeout=TIMEOUT)

    page.set_input_files("#avatar", caminho_foto)

    # força evento JS (MetaX depende disso)
    page.evaluate("""
        const input = document.querySelector('#avatar');
        input.dispatchEvent(new Event('change', { bubbles: true }));
    """)

    print(f"[OK] Foto anexada: {os.path.basename(caminho_foto)}")


# ================== EXECUÇÃO PRINCIPAL ==================
def executar_metax(funcionario):
    cpf = funcionario["CPF"]

    caminho_foto = buscar_foto_por_cpf(PASTA_FOTOS, cpf)

    with sync_playwright() as p:
        browser = p.chromium.launch(channel="chrome", headless=False)
        context = browser.new_context()
        page = context.new_page()

        try:
            # ---------- LOGIN ----------
            print("[INFO] Acessando MetaX")
            page.goto(URL_LOGIN, wait_until="domcontentloaded", timeout=TIMEOUT)

            page.fill("#txtLogin", LOGIN)
            page.fill("#txtSenha", SENHA)

            print(f"[INFO] Resolva o CAPTCHA ({TEMPO_CAPTCHA_MS//1000}s)")
            page.wait_for_timeout(TEMPO_CAPTCHA_MS)

            page.click('button:has-text("Validar")')

            # ---------- CONTRATO ----------
            page.wait_for_selector("#comboContrato", state="visible", timeout=TIMEOUT)
            page.select_option("#comboContrato", value="6578")

            page.evaluate("""
                document.querySelector('#comboContrato')
                    .dispatchEvent(new Event('change', { bubbles: true }));
            """)

            page.click('button:has-text("Continuar")')

            # ---------- TERMOS ----------
            page.wait_for_selector('text=Termo de confirmação', timeout=TIMEOUT)
            page.click('text=Li e Aceito os termos de compromisso')
            page.wait_for_selector('text=Termo de confirmação', state="hidden")

            print("[OK] Login concluído")

            # ---------- CREDENCIAMENTO ----------
            page.click('a[href*="Credenciamento"]')
            page.wait_for_url("**/CredenciamentoLista", timeout=TIMEOUT)

            page.click('a[href="/Credenciamento/Index"]')
            page.wait_for_load_state("domcontentloaded")

            print("[INFO] Tela de cadastro aberta")

            # ---------- FOTO (CONDICIONAL) ----------
            if caminho_foto:
                anexar_foto(page, caminho_foto)
            else:
                print(f"[INFO] Nenhuma foto encontrada para CPF {cpf} – seguindo sem foto")

            # ---------- DEBUG ----------
            valor = page.locator("#avatar").input_value()
            print("[DEBUG] Input avatar:", valor)

            # ---------- EXEMPLO DE CAMPO ----------
            page.fill("#nome", funcionario.get("NOME", "TESTE SEM NOME"))

            # ---------- SALVAR RASCUNHO ----------
            page.wait_for_function("""
                () => {
                    const btn = document.querySelector('#btnSalvarRascunho');
                    return btn && !btn.disabled;
                }
            """, timeout=TIMEOUT)

            page.locator("#btnSalvarRascunho").click(force=True)
            print("[OK] Rascunho salvo")

            page.wait_for_timeout(5000)

        except PlaywrightTimeoutError:
            print("[ERRO] Timeout no MetaX")
            page.screenshot(path="erro_timeout.png")

        except Exception as e:
            print("[ERRO] Erro inesperado:", e)
            page.screenshot(path="erro_inesperado.png")

        finally:
            browser.close()


# ================== MAIN ==================
if __name__ == "__main__":
    funcionario_exemplo = {
        "NOME": "VITOR ANTONIO ALVES BORGES",
        "CPF": "12193578630"
    }

    executar_metax(funcionario_exemplo)
