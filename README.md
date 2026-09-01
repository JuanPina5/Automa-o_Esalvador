# Automação eSalvador -> Excel

Um sistema completo, construído do zero, para automatizar a extração das informações de `Unidade` e `Dias na Unidade` de processos no portal eSalvador, e atualizar automaticamente uma planilha Excel.

## Funcionalidades
- Não cria linhas novas nem histórico de movimentação (apenas sobrescreve o valor atual).
- Identifica automaticamente os cabeçalhos (`PROCESSO`, `ANO`, `LOCALIZAÇÃO ATUAL`, `TRAMITAÇÃO SETOR (DIAS)`).
- Interação autônoma baseada em identificadores DOM do Playwright, e não apenas XPath estáticos ou regex.
- Retries e fallback configurados.
- Salvamento automático de cópia (`GERAL_ATUALIZADO.xlsx`) a cada processo concluído com sucesso.
- Dump inteligente de HTML e Screenshots em caso de erro na extração (se o DEBUG estiver ativado no `config.py`).

## Estrutura do Projeto
```
esalvador_automacao/
│
├── config.py           # Configurações globais (urls, tempos, debug)
├── logger.py           # Estilização do output de terminal
├── excel_manager.py    # Tratamento e alteração de Excel (openpyxl)
├── extractor.py        # Inteligência de raspagem do HTML do processo
├── esalvador.py        # Motor de navegação (Playwright wrapper)
├── main.py             # Orquestrador da automação
├── requirements.txt    # Dependências de biblioteca
│
├── GERAL.xlsx          # (O arquivo deve ser colocado aqui)
│
├── output/             # Diretório onde sairá a planilha atualizada
│   └── GERAL_ATUALIZADO.xlsx
│
└── debug/              # Diretório onde dumps serão salvos
```

## Como Instalar

1. Certifique-se de que possui o [Python](https://www.python.org/downloads/) 3.8+ instalado.
2. Abra o terminal nesta pasta (`esalvador_automacao`).
3. Instale as bibliotecas necessárias usando o `pip`:
   ```bash
   pip install -r requirements.txt
   ```
4. Instale os navegadores baseados em Chromium do Playwright:
   ```bash
   playwright install chromium
   ```

## Como Usar

1. Coloque a sua planilha original com o nome **`GERAL.xlsx`** dentro desta mesma pasta `esalvador_automacao`.
2. Dê um duplo-clique no arquivo `main.py` ou abra o terminal e rode:
   ```bash
   python main.py
   ```
3. O navegador Chromium irá abrir no site eSalvador.
4. **Faça o Login Manualmente.** 
5. Volte na janela preta (terminal) e aperte **ENTER**.
6. Acompanhe a automação atualizar todos os processos do seu arquivo, com cópias sendo geradas para a pasta `output/GERAL_ATUALIZADO.xlsx`.

## Modo Diagnóstico (Debug)
Dentro do arquivo `config.py` existe a variável:
```python
DEBUG = True
```
Quando configurada como `True`, a automação irá fornecer mais detalhes no console. Caso a página apresente um erro ou ela falhe em extrair alguma unidade ou dias na página final, um arquivo `.html` e um `.png` (screenshot) serão salvos dentro da pasta `debug/`.
Caso queira desligar e tornar o terminal mais limpo, altere para `False`.

# Entra na pasta onde estão seus códigos
cd C:\caminho\da\sua\pasta\com\os\codigos

# Inicia git
git init

# Adiciona todos os arquivos
git add .

# Primeiro commit
git commit -m "Inicial: automação eSalvador e Diário Oficial"

# Conecta ao GitHub (cola a URL que copiou)
git remote add origin https://github.com/SEU_USER/automacao-salvador.git

# Envia para GitHub
git branch -M main
git push -u origin main
