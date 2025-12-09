
# ======================== config.py ========================

class ConfigEmail:
    ACCOUNT_NAME = None

class ConfigGA:
    URL = "https://ga.flashcourier.com.br/logs"
    
    CAMPO_CONSULTAS = "//a[contains(text(), 'CONSULTAS')]"
    CAMPO_ARQUIVOS = "//a[contains(text(), 'ARQUIVOS PROCESSADOS')]"
    CAMPO_PESQUISA = "input[placeholder*='Pesquisar']"
    BOTAO_EXCEL = "//button[contains(text(), 'EXCEL')]"
    
    DOWNLOAD_PATH = "./downloads"

class ConfigArquivos:
    OUTPUT_EMAILS = "emails_{data}.xlsx"
    OUTPUT_GA = "ga_relatorios_{data}.xlsx"
    NOME_ARQUIVO_GA = "Arquivos Processados.xlsx"