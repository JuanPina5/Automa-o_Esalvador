import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DEBUG_DIR = os.path.join(BASE_DIR, "debug")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")

EXCEL_FILE_NAME = "GERAL.xlsx"
EXCEL_OUTPUT_NAME = "GERAL_ATUALIZADO.xlsx"

EXCEL_ORIGINAL_PATH = os.path.join(BASE_DIR, EXCEL_FILE_NAME)
EXCEL_UPDATED_PATH = os.path.join(OUTPUT_DIR, EXCEL_OUTPUT_NAME)

PAGAMENTO_FILE_NAME = "PAGAMENTO.xlsx"
PAGAMENTO_OUTPUT_NAME = "PAGAMENTO_ATUALIZADO.xlsx"

PAGAMENTO_ORIGINAL_PATH = os.path.join(BASE_DIR, PAGAMENTO_FILE_NAME)
PAGAMENTO_UPDATED_PATH = os.path.join(OUTPUT_DIR, PAGAMENTO_OUTPUT_NAME)

URL_LOGIN = "https://esalvador.salvador.ba.gov.br"
URL_CONSULTA = "https://esalvador.salvador.ba.gov.br/consulta"

DEBUG = True
HEADLESS = False
TIMEOUT_DEFAULT = 60000
