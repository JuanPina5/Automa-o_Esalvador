import re
from playwright.sync_api import TimeoutError

class ESalvadorAutomacao:
    def __init__(self, page, logger, extractor_cls, config):
        self.page = page
        self.logger = logger
        self.Extractor = extractor_cls
        self.config = config

    def esperar_login(self):
        self.logger.section("LOGIN DO eSALVADOR")
        self.logger.info("Faça o login manualmente no navegador.")
        self.logger.info("Depois de entrar no sistema, volte para o terminal e pressione ENTER.")
        input("\nPressione ENTER para continuar...")
        self.logger.success("Login finalizado.")

    def ir_para_consulta(self):
        try:
            self.page.goto(self.config.URL_CONSULTA, wait_until="domcontentloaded", timeout=self.config.TIMEOUT_DEFAULT)
            # Aguarda até que um campo de texto (ou não-oculto) fique visível
            self.page.wait_for_selector("input:not([type='hidden'])", state="visible", timeout=15000)
            # Opcional: aguardar estabilidade de rede
            try:
                self.page.wait_for_load_state("networkidle", timeout=5000)
            except TimeoutError:
                pass
            return True
        except Exception as e:
            self.logger.error(f"Erro ao acessar página de consulta: {e}")
            return False

    def pesquisar_processo(self, processo, ano):
        try:
            # Encontrar e clicar botão Limpar se existir (para resetar o formulário)
            botao_limpar = self.page.get_by_role("button", name=re.compile(r"^\s*LIMPAR\s*$", re.IGNORECASE)).first
            if botao_limpar.count() > 0 and botao_limpar.is_visible():
                botao_limpar.click()
                self.page.wait_for_timeout(500)

            # Preencher formulário
            inputs = self.page.locator("input:not([type='hidden'])").all()
            visiveis = []
            for inp in inputs:
                if inp.is_visible():
                    visiveis.append(inp)

            campo_proc, campo_ano = None, None
            
            for inp in visiveis:
                try:
                    html = inp.evaluate("el => el.outerHTML").lower()
                    if "interno" in html or "processo" in html:
                        if not campo_proc: campo_proc = inp
                    elif "ano" in html:
                        if not campo_ano: campo_ano = inp
                except Exception:
                    pass

            # Fallback absoluto: 1º input visível = Processo, 2º input visível = Ano
            if not campo_proc and len(visiveis) >= 1:
                self.logger.debug("Usando fallback de posição para campo Processo")
                campo_proc = visiveis[0]
            if not campo_ano and len(visiveis) >= 2:
                self.logger.debug("Usando fallback de posição para campo Ano")
                campo_ano = visiveis[1]

            if not campo_proc:
                self.logger.error("Campo de processo não encontrado (nenhum input visível).")
                return False

            campo_proc.fill(str(processo))
            if campo_ano and ano:
                campo_ano.fill(str(ano))

            # Buscar
            botao_buscar = self.page.get_by_role("button", name=re.compile(r"^\s*BUSCAR\s*$", re.IGNORECASE)).first
            if not botao_buscar.count() > 0:
                for b in self.page.locator("button").all():
                    if b.is_visible() and b.inner_text().strip().upper() == "BUSCAR":
                        botao_buscar = b
                        break
            
            if botao_buscar:
                try:
                    with self.page.expect_response(lambda response: "consulta" in response.url or response.status == 200, timeout=15000):
                        botao_buscar.click()
                except TimeoutError:
                    botao_buscar.click()
                
                self.page.wait_for_load_state("networkidle", timeout=10000)
                self.logger.success("Processo pesquisado")
                return True
            else:
                self.logger.error("Botão BUSCAR não encontrado.")
                return False

        except Exception as e:
            self.logger.error(f"Erro ao pesquisar processo: {e}")
            return False

    def clicar_lupa_acao(self, processo, ano):
        try:
            self.page.wait_for_selector("table", timeout=10000)
            linhas = self.page.locator("table tr").all()
            
            for linha in linhas:
                if not linha.is_visible(): continue
                texto = linha.inner_text().lower()
                if str(processo) in texto and (not ano or str(ano) in texto):
                    celulas = linha.locator("td").all()
                    if not celulas: continue
                    ultima_celula = celulas[-1]
                    
                    lupas = ultima_celula.locator("a, button, i[class*='search'], [title*='detalhar' i], [title*='visualizar' i]").all()
                    for lupa in lupas:
                        if lupa.is_visible():
                            pai = lupa.locator("xpath=ancestor::a[1]")
                            click_target = pai if pai.count() > 0 else lupa
                            
                            # Remove target='_blank' do elemento ou do seu pai para que não abra nova aba
                            try:
                                click_target.evaluate("""el => {
                                    if (el.tagName === 'A') el.removeAttribute('target');
                                    let p = el.closest('a');
                                    if (p) p.removeAttribute('target');
                                }""")
                            except Exception:
                                pass
                                
                            try:
                                click_target.click()
                            except Exception:
                                lupa.click()
                                
                            self.page.wait_for_load_state("domcontentloaded", timeout=10000)
                            self.logger.success("Lupa de ação encontrada e clicada")
                            return True
            
            self.logger.error("Processo não encontrado na tabela de resultados.")
            return False
        except Exception as e:
            self.logger.error(f"Erro ao clicar na lupa de ação: {e}")
            return False

    def validar_pagina_detalhe(self, processo, ano):
        try:
            self.page.wait_for_selector("body", timeout=10000)
            texto = self.page.inner_text("body").lower()
            
            padrao_proc = re.compile(r"(?<!\d)" + re.escape(str(processo)) + r"(?!\d)")
            if not padrao_proc.search(texto):
                self.logger.error("Processo não bate na validação da página.")
                return False
                
            if ano:
                padrao_ano = re.compile(r"(?<!\d)" + re.escape(str(ano)) + r"(?!\d)")
                if not padrao_ano.search(texto):
                    self.logger.error("Ano não bate na validação da página.")
                    return False
            
            self.logger.success("Página de detalhes aberta e processo validado")
            return True
        except Exception as e:
            self.logger.error(f"Erro ao validar página de detalhe: {e}")
            return False

    def processar_extravel(self, processo, ano):
        ext = self.Extractor(self.page, self.logger, self.config.DEBUG, self.config.DEBUG_DIR)
        return ext.extract_localizacao(processo, ano)
