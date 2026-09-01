import os
import re

class Extractor:
    def __init__(self, page, logger, debug=False, debug_dir=""):
        self.page = page
        self.logger = logger
        self.debug = debug
        self.debug_dir = debug_dir

    def extract_localizacao(self, processo, ano):
        unidade = None
        orgao = None
        dias = None
        
        try:
            campo_dias = self.page.get_by_text("Dias na Unidade", exact=True).first
            
            if campo_dias.count() > 0:
                self.logger.debug("Texto 'Dias na Unidade' encontrado, buscando container.")
                container = campo_dias
                
                found_container = False
                for _ in range(10):
                    try:
                        container = container.locator("xpath=..")
                        if container.count() == 0: break
                        
                        texto_container = container.inner_text()
                        if "Unidade" in texto_container and "Dias na Unidade" in texto_container:
                            found_container = True
                            break
                    except Exception:
                        break
                        
                if found_container:
                    try:
                        filhos = container.locator("xpath=./*").all()
                        textos = [f.inner_text().strip() for f in filhos if f.is_visible()]
                        
                        for texto in textos:
                            norm = texto.lower()
                            if norm.startswith("dias na unidade"):
                                valor = texto[len("Dias na Unidade"):].strip()
                                dias = self._parse_dias(valor)
                            elif norm.startswith("unidade") and "dias" not in norm:
                                valor = texto[len("Unidade"):].strip()
                                unidade = self._validate_unidade(valor)
                            elif norm.startswith("órgão") or norm.startswith("orgao"):
                                # Pega o valor após a palavra Órgão
                                start_len = len("órgão") if norm.startswith("órgão") else len("orgao")
                                valor = texto[start_len:].strip()
                                # Remove possíveis caracteres residuais do layout
                                if valor.startswith(":"): valor = valor[1:].strip()
                                if valor.startswith("-"): valor = valor[1:].strip()
                                orgao = self._validate_unidade(valor)
                                
                    except Exception as e:
                        self.logger.debug(f"Erro ao processar filhos: {e}")
                
            if not unidade or dias is None:
                self.logger.debug("Tentando fallback via 'Localização Atual'")
                loc_atual = self.page.get_by_text("Localização Atual", exact=True).first
                if loc_atual.count() > 0:
                    container = loc_atual
                    for _ in range(10):
                        try:
                            container = container.locator("xpath=..")
                            texto_container = container.inner_text()
                            if "Unidade" in texto_container and "Dias na Unidade" in texto_container:
                                linhas = [l.strip() for l in texto_container.splitlines() if l.strip()]
                                for i, linha in enumerate(linhas):
                                    norm = linha.lower()
                                    if norm == "unidade" and i + 1 < len(linhas):
                                        val = linhas[i+1]
                                        if not self._is_label(val) and not unidade:
                                            unidade = self._validate_unidade(val)
                                    elif (norm == "órgão" or norm == "orgao") and i + 1 < len(linhas):
                                        val = linhas[i+1]
                                        if not self._is_label(val) and not orgao:
                                            orgao = self._validate_unidade(val)
                                    elif norm == "dias na unidade" and i + 1 < len(linhas):
                                        val = linhas[i+1]
                                        if dias is None:
                                            dias = self._parse_dias(val)
                                break
                        except Exception:
                            pass
            
            if not unidade or dias is None:
                self._dump_debug(processo, ano)
                return None
                
            # Compor a string final da unidade com o órgão
            unidade_final = unidade
            if orgao:
                orgao_upper = orgao.upper()
                # Somente anexa o órgão se NÃO for a SMS e não for igual à unidade
                if orgao_upper != "SMS" and "SMS -" not in orgao_upper and orgao_upper != unidade.upper():
                    unidade_final = f"{orgao}/{unidade}"
                
            self.logger.info("\nLOCALIZAÇÃO ATUAL")
            self.logger.separator()
            self.logger.info(f"Unidade: {unidade_final}")
            self.logger.info(f"Dias na Unidade: {dias}")
            self.logger.separator()
            
            return {"unidade": unidade_final, "dias": dias}
            
        except Exception as e:
            self.logger.error(f"Erro durante a extração: {e}")
            self._dump_debug(processo, ano)
            return None

    def _is_label(self, text):
        lower_text = str(text).lower()
        labels = ["unidade", "dias na unidade", "órgão", "selecione", "-", "localização atual"]
        return any(lower_text == l or lower_text.startswith("selecione") for l in labels)

    def _validate_unidade(self, unidade):
        if not unidade:
            return None
        u = str(unidade).strip()
        if not u or self._is_label(u) or u.isdigit() or len(u) < 2:
            return None
        return u

    def _parse_dias(self, texto):
        if not texto:
            return None
        numeros = re.findall(r"\d+", str(texto))
        if numeros:
            return int(numeros[-1])
        return None

    def _dump_debug(self, processo, ano):
        if self.debug:
            try:
                safe_name = f"processo_{processo}_{ano}".replace("/", "_").replace("\\", "_")
                html_path = os.path.join(self.debug_dir, f"{safe_name}.html")
                png_path = os.path.join(self.debug_dir, f"{safe_name}.png")
                
                with open(html_path, "w", encoding="utf-8") as f:
                    f.write(self.page.content())
                self.page.screenshot(path=png_path)
                self.logger.warning(f"Dump de debug salvo em ./debug/{safe_name}")
            except Exception as e:
                self.logger.debug(f"Erro ao salvar dump de debug: {e}")
