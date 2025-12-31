# ======================== ga.py (atualizado) ========================

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
import pandas as pd
import time
import logging
import os
from pathlib import Path
from dotenv import load_dotenv

# Carrega variáveis de ambiente do arquivo .env
load_dotenv()

logger = logging.getLogger(__name__)

class ExtratorGA:
    
    def __init__(self, url: str, download_path: str, email: str = None, senha: str = None):
        self.url = url
        # Se email/senha não forem passados, pega do .env
        self.email = email or os.getenv('GA_EMAIL')
        self.senha = senha or os.getenv('GA_SENHA')
        self.download_path = download_path
        self.driver = None
        self.wait = None
        self.timestamp_inicio = None
        self.arquivos_processados = []
        
        # Valida se as credenciais foram carregadas
        if not self.email or not self.senha:
            raise ValueError("Credenciais não encontradas. Defina GA_EMAIL e GA_SENHA no arquivo .env")
    
    def inicializar_driver(self) -> bool:
        try:
            os.makedirs(self.download_path, exist_ok=True)
            
            chrome_options = Options()
            chrome_options.add_argument("--headless")
            chrome_options.add_argument("--no-sandbox")
            chrome_options.add_argument("--disable-dev-shm-usage")
            
            prefs = {
                "download.default_directory": self.download_path,
                "download.prompt_for_download": False,
            }
            chrome_options.add_experimental_option("prefs", prefs)
            
            self.driver = webdriver.Chrome(options=chrome_options)
            self.wait = WebDriverWait(self.driver, 15)
            
            logger.info("✓ Driver Chrome iniciado")
            return True
        
        except Exception as e:
            logger.error(f"✗ Erro ao inicializar driver: {e}")
            return False
    
    def fazer_login(self) -> bool:
        try:
            logger.info("Acessando GA...")
            self.driver.get(self.url)
            time.sleep(2)
            
            self.timestamp_inicio = time.time()
            
            usuario_box = self.driver.find_element(By.NAME, "email")
            usuario_box.send_keys(self.email)
            
            senha_box = self.driver.find_element(By.NAME, "password")
            senha_box.send_keys(self.senha)
            
            login_button = self.driver.find_element(By.XPATH, '//*[@id="login"]/section/form/div[3]/button')
            login_button.click()
            
            time.sleep(3)
            logger.info("✓ Login realizado com sucesso")
            return True
        
        except Exception as e:
            logger.error(f"✗ Erro ao fazer login: {e}")
            return False
    
    def extrair_relatorio_cliente(self, cliente: str) -> dict:
        try:
            logger.info(f"Extraindo relatório para: {cliente}")
            
            # NOVO: Detecta ALELO-KIT e busca com filtro correto
            termo_busca = cliente
            is_alelo_kit = (cliente.upper() == "ALELO-KIT")
            is_alelo_normal = ("ALELO" in cliente.upper() and not is_alelo_kit)
            
            if is_alelo_kit:
                termo_busca = "ELO-RE"
                logger.info(f"🎯 ALELO-KIT detectado. Buscando por: {termo_busca} (filtro: COM _KIT)")
            elif is_alelo_normal:
                termo_busca = "ELO-RE"
                logger.info(f"🎯 ALELO normal detectado. Buscando por: {termo_busca} (filtro: SEM _KIT)")
            
            campo_pesquisa = self.wait.until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "input[aria-controls='dataTableBuilder']"))
            )
            campo_pesquisa.clear()
            campo_pesquisa.send_keys(termo_busca)
            
            logger.info(f"Aguardando 5 segundos...")
            time.sleep(5)
            
            botao_excel = self.wait.until(
                EC.element_to_be_clickable((By.ID, "spreadsheet"))
            )
            botao_excel.click()
            
            logger.info("Download iniciado...")
            time.sleep(7)
            
            resultado = self._processar_arquivo_excel(cliente, is_alelo_kit, is_alelo_normal)
            
            return resultado
        
        except Exception as e:
            logger.error(f"✗ Erro ao extrair relatório para {cliente}: {e}")
            return {'total': 0}
    
    def _processar_arquivo_excel(self, cliente: str, is_alelo_kit: bool, is_alelo_normal: bool) -> dict:
        try:
            arquivo = self._obter_arquivo_recente()
            
            if not arquivo:
                logger.warning("Nenhum arquivo foi identificado")
                return {'total': 0}
            
            downloads_path = str(Path.home() / "Downloads")
            arquivo_path = os.path.join(downloads_path, arquivo)
            
            if not os.path.exists(arquivo_path):
                logger.warning(f"Arquivo não encontrado: {arquivo_path}")
                return {'total': 0}
            
            logger.info(f"Processando arquivo: {arquivo}")
            df = pd.read_excel(arquivo_path)
            logger.info(f"Arquivo carregado com {len(df)} linhas e {df.shape[1]} colunas")
            
            if df.shape[1] >= 7:
                coluna_c = df.iloc[:, 2]
                coluna_d = df.iloc[:, 3]
                coluna_e = df.iloc[:, 4]
                coluna_g = df.iloc[:, 6]
                
                filtro_base = (coluna_g.astype(str).str.upper() == "ENTREGUE") & (~coluna_d.astype(str).str.contains(".SD1", case=False, na=False))
                
                if is_alelo_kit:
                    # ALELO-KIT: Filtra COM _KIT
                    filtro_kit = filtro_base & (coluna_c.astype(str).str.contains("_KIT", case=False, na=False))
                    total = int(coluna_e[filtro_kit].sum())
                    
                    logger.info(f"✅ ALELO-KIT (com _KIT): {total}")
                    self.arquivos_processados.append(arquivo)
                    
                    return {'total': total}
                
                elif is_alelo_normal:
                    # ALELO normal: Filtra SEM _KIT
                    filtro_alelo = filtro_base & (~coluna_c.astype(str).str.contains("_KIT", case=False, na=False))
                    total = int(coluna_e[filtro_alelo].sum())
                    
                    logger.info(f"✅ ALELO Normal (sem _KIT): {total}")
                    self.arquivos_processados.append(arquivo)
                    
                    return {'total': total}
                
                else:
                    # Outros clientes: Sem filtro especial
                    total = int(coluna_e[filtro_base].sum())
                    
                    logger.info(f"Total somado da coluna E: {total}")
                    self.arquivos_processados.append(arquivo)
                    
                    return {'total': total}
            else:
                logger.warning("Arquivo não possui coluna G")
                return {'total': 0}
        
        except Exception as e:
            logger.error(f"✗ Erro ao processar Excel: {e}")
            return {'total': 0}
    
    def _obter_arquivo_recente(self):
        try:
            downloads_path = str(Path.home() / "Downloads")
            
            time.sleep(2)
            
            arquivos_xlsx = [f for f in os.listdir(downloads_path) if f.endswith('.xlsx') and not f.startswith('~')]
            
            if not arquivos_xlsx:
                logger.warning("Nenhum arquivo .xlsx encontrado em Downloads")
                return None
            
            arquivos_novos = []
            agora = time.time()
            
            for arquivo in arquivos_xlsx:
                caminho_completo = os.path.join(downloads_path, arquivo)
                tempo_modificacao = os.path.getmtime(caminho_completo)
                
                if tempo_modificacao > self.timestamp_inicio and (agora - tempo_modificacao) < 30:
                    arquivos_novos.append(arquivo)
                    logger.info(f"Arquivo candidato: {arquivo} (modificado há {int(agora - tempo_modificacao)}s)")
            
            if not arquivos_novos:
                logger.warning("Nenhum arquivo novo encontrado em Downloads")
                logger.info(f"Timestamp início: {self.timestamp_inicio}, Agora: {agora}")
                return None
            
            arquivo_mais_recente = max(
                arquivos_novos,
                key=lambda f: os.path.getmtime(os.path.join(downloads_path, f))
            )
            
            logger.info(f"Arquivo selecionado: {arquivo_mais_recente}")
            return arquivo_mais_recente
        
        except Exception as e:
            logger.error(f"✗ Erro ao identificar arquivo: {e}")
            return None
    
    def fechar(self):
        try:
            if self.driver:
                self.driver.quit()
                logger.info("✓ Driver fechado")
        except Exception as e:
            logger.error(f"✗ Erro ao fechar driver: {e}")