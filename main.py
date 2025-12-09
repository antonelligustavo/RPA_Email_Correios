# ======================== main.py ========================

import logging
from datetime import datetime
import time
import os
from dotenv import load_dotenv

script_dir = os.path.dirname(os.path.abspath(__file__))
env_path = os.path.join(script_dir, '.env')

print(f"🔍 Procurando arquivo .env em: {env_path}")

if os.path.exists(env_path):
    print(f"✅ Arquivo .env encontrado!")
    load_dotenv(env_path)
else:
    print(f"⚠️ Arquivo .env NÃO encontrado em: {env_path}")
    print(f"🔍 Tentando carregar do diretório atual...")
    load_dotenv()

from config import ConfigEmail, ConfigGA, ConfigArquivos
from emails import ColetorEmails
from ga import ExtratorGA
from planilhas import GerenciadorPlanilhas
from respostas import RespostorEmails

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

def main():
    
    logger.info("="*60)
    logger.info("INICIANDO PROCESSO DE VALIDAÇÃO CORREIOS")
    logger.info("="*60)
    
    logger.info("\n[ETAPA 1] Coletando e-mails do Outlook...")
    
    coletor = ColetorEmails(nome_pasta="Processamento Correios")
    
    if not coletor.conectar():
        logger.error("Falha ao conectar. Abortando.")
        return
    
    emails = coletor.buscar_emails_do_dia()
    
    if not emails:
        logger.warning("Nenhum e-mail encontrado!")
        return
    
    logger.info(f"✓ {len(emails)} e-mail(s) coletado(s)")
    
    arquivo_emails = GerenciadorPlanilhas.salvar_emails(
        emails,
        ConfigArquivos.OUTPUT_EMAILS
    )
    
    clientes = [e["Cliente"] for e in emails]
    logger.info(f"Clientes encontrados: {', '.join(clientes)}")
    
    logger.info("\n[ETAPA 2] Extraindo relatórios do GA...")
    
    # ALTERAÇÃO: Agora o ExtratorGA busca email e senha automaticamente do .env
    extrator = ExtratorGA(
        url=ConfigGA.URL,
        download_path=ConfigGA.DOWNLOAD_PATH
    )
    
    if not extrator.inicializar_driver():
        logger.error("Falha ao inicializar Selenium. Abortando.")
        return
    
    try:
        if not extrator.fazer_login():
            logger.error("Falha no login do GA. Abortando.")
            return
        
        resultados_ga = {}
        
        for cliente in clientes:
            resultado = extrator.extrair_relatorio_cliente(cliente)
            
            resultados_ga[cliente] = resultado.get('total', 0)
            
            time.sleep(2)
        
        arquivo_ga = GerenciadorPlanilhas.salvar_relatorios_ga(
            resultados_ga,
            ConfigArquivos.OUTPUT_GA
        )
        
    finally:
        extrator.fechar()
    
    logger.info("\n[ETAPA 3] Realizando validação cruzada...")
    
    dados_validacao = GerenciadorPlanilhas.gerar_dados_validacao(
        emails,
        resultados_ga
    )
    
    arquivo_validacao = GerenciadorPlanilhas.salvar_validacao(
        dados_validacao,
        "validacao_{data}.xlsx"
    )
    
    logger.info("\n📤 Enviando relatório para o Teams...")
    GerenciadorPlanilhas.enviar_para_teams(dados_validacao)
    
    logger.info("\n[ETAPA 4] Respondendo e-mails automaticamente...")
    
    responsor = RespostorEmails(nome_pasta="Processamento Correios")
    
    if responsor.conectar():
        responsor.responder_emails(dados_validacao)
    else:
        logger.warning("Não foi possível responder e-mails")
    
    logger.info("\n" + "="*60)
    logger.info("PROCESSO FINALIZADO COM SUCESSO!")
    logger.info("="*60)
    logger.info(f"Arquivo de E-mails: {arquivo_emails}")
    logger.info(f"Arquivo de GA: {arquivo_ga}")
    logger.info(f"Arquivo de Validação: {arquivo_validacao}")
    logger.info("="*60)

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        logger.info("\n⚠ Processo interrompido pelo usuário")
    except Exception as e:
        logger.error(f"✗ Erro não tratado: {e}", exc_info=True)