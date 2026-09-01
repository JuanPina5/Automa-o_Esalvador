import time
from playwright.sync_api import sync_playwright

import config
from logger import Logger
from excel_manager import ExcelManager
from extractor import Extractor
from esalvador import ESalvadorAutomacao


def processar_planilha(excel, automacao, logger):
    """Processa todos os processos de uma planilha já configurada.
    Retorna o dicionário de stats ao final."""
    processos = excel.get_processes()
    total_processos = len(processos)
    logger.info(f"Processos encontrados: {total_processos}")

    if total_processos == 0:
        logger.warning("Nenhum processo válido encontrado. Pulando esta planilha.")
        return {"atualizados": 0, "corretos": 0, "erros": 0, "sem_dados": 0, "total": 0}

    stats = {"atualizados": 0, "corretos": 0, "erros": 0, "sem_dados": 0, "total": total_processos}

    for i, proc_data in enumerate(processos, 1):
        processo = proc_data["processo"]
        ano      = proc_data["ano"]
        linha    = proc_data["row"]

        logger.step(processo, ano, i, total_processos)

        if not automacao.ir_para_consulta():
            stats["erros"] += 1
            continue

        if not automacao.pesquisar_processo(processo, ano):
            stats["erros"] += 1
            continue

        if not automacao.clicar_lupa_acao(processo, ano):
            stats["erros"] += 1
            continue

        if not automacao.validar_pagina_detalhe(processo, ano):
            logger.warning("Página não validada. Pulando para evitar gravação indevida.")
            stats["erros"] += 1
            continue

        dados = automacao.processar_extravel(processo, ano)

        if dados:
            mudou = excel.update_process(linha, dados.get("unidade"), dados.get("dias"))
            if mudou:
                stats["atualizados"] += 1
            else:
                stats["corretos"] += 1

            if not excel.save():
                logger.error("Interrompendo devido a falha no salvamento.")
                stats["erros"] += 1
                break
        else:
            stats["sem_dados"] += 1

        # Evita sobrecarregar o site rapidamente
        time.sleep(1)

    return stats


def main():
    logger = Logger(debug_mode=config.DEBUG)
    logger.section("AUTOMAÇÃO eSALVADOR -> EXCEL")

    # ── Preparar planilha GERAL ──────────────────────────────────
    logger.section("VERIFICANDO PLANILHA GERAL")
    excel_geral = ExcelManager(
        config.EXCEL_ORIGINAL_PATH,
        config.EXCEL_UPDATED_PATH,
        logger
    )
    geral_ok = excel_geral.setup()

    import os
    pagamento_existe = os.path.exists(config.PAGAMENTO_ORIGINAL_PATH)

    if not geral_ok and not pagamento_existe:
        logger.error("Nenhuma planilha pôde ser carregada ou encontrada. Finalizando.")
        return

    # ── Navegador (sessão única / login único) ──────────────────
    with sync_playwright() as p:
        logger.section("INICIANDO NAVEGADOR")
        try:
            browser = p.chromium.launch(headless=config.HEADLESS)
            page    = browser.new_page(viewport={"width": 1536, "height": 864})
            page.goto(config.URL_LOGIN, wait_until="domcontentloaded", timeout=config.TIMEOUT_DEFAULT)
        except Exception as e:
            logger.error(f"Não foi possível abrir o eSalvador: {e}")
            return

        automacao = ESalvadorAutomacao(page, logger, Extractor, config)
        automacao.esperar_login()

        if not automacao.ir_para_consulta():
            logger.error("Falha ao abrir a página de consulta.")
            browser.close()
            return

        # ── Processar GERAL ──────────────────────────────────────
        stats_geral = {"atualizados": 0, "corretos": 0, "erros": 0, "sem_dados": 0, "total": 0}
        if geral_ok:
            logger.section(f"PLANILHA: {config.EXCEL_FILE_NAME}")
            stats_geral = processar_planilha(excel_geral, automacao, logger)

        # ── Preparar e Processar PAGAMENTO ───────────────────────
        stats_pagamento = {"atualizados": 0, "corretos": 0, "erros": 0, "sem_dados": 0, "total": 0}
        pagamento_ok = False

        if pagamento_existe:
            logger.section("PREPARANDO PRÓXIMA PLANILHA")
            logger.info("Aguardando estabilização e retornando à consulta...")
            time.sleep(3)
            if not automacao.ir_para_consulta():
                logger.error("Falha ao retornar à consulta. Encerrando antes da PAGAMENTO.")
                browser.close()
                return

            logger.section("VERIFICANDO PLANILHA PAGAMENTO")
            excel_pagamento = ExcelManager(
                config.PAGAMENTO_ORIGINAL_PATH,
                config.PAGAMENTO_UPDATED_PATH,
                logger
            )
            pagamento_ok = excel_pagamento.setup()

            if pagamento_ok:
                logger.section(f"PLANILHA: {config.PAGAMENTO_FILE_NAME}")
                stats_pagamento = processar_planilha(excel_pagamento, automacao, logger)
            else:
                logger.error("Falha ao configurar a planilha PAGAMENTO. Pulando.")

        browser.close()

    # ── Relatório final ──────────────────────────────────────────
    logger.section("PROCESSAMENTO FINALIZADO")

    if geral_ok:
        logger.info(f"[ {config.EXCEL_FILE_NAME} ]")
        logger.info(f"  Total de processos : {stats_geral['total']}")
        logger.info(f"  Atualizados        : {stats_geral['atualizados']}")
        logger.info(f"  Já estavam corretos: {stats_geral['corretos']}")
        logger.info(f"  Sem dados          : {stats_geral['sem_dados']}")
        logger.info(f"  Erros              : {stats_geral['erros']}")
        logger.info(f"  Saída              : {config.EXCEL_OUTPUT_NAME}\n")

    if pagamento_ok:
        logger.info(f"[ {config.PAGAMENTO_FILE_NAME} ]")
        logger.info(f"  Total de processos : {stats_pagamento['total']}")
        logger.info(f"  Atualizados        : {stats_pagamento['atualizados']}")
        logger.info(f"  Já estavam corretos: {stats_pagamento['corretos']}")
        logger.info(f"  Sem dados          : {stats_pagamento['sem_dados']}")
        logger.info(f"  Erros              : {stats_pagamento['erros']}")
        logger.info(f"  Saída              : {config.PAGAMENTO_OUTPUT_NAME}\n")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nProcesso interrompido pelo usuário.")
    except Exception as e:
        import traceback
        print("\n[ERRO FATAL]")
        traceback.print_exc()

    input("\nPressione ENTER para fechar...")

