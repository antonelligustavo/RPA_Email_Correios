# ======================== emails.py ========================

import win32com.client
from datetime import datetime, timedelta
import re
from typing import List, Dict
import logging
import unicodedata

logger = logging.getLogger(__name__)

def normalizar_texto(texto: str) -> str:
    nfd = unicodedata.normalize('NFD', texto)
    sem_acentos = ''.join(char for char in nfd if unicodedata.category(char) != 'Mn')
    return sem_acentos.upper()

def contem_validacao(texto: str) -> bool:
    """
    Verifica se o texto contém variações de "VALIDAÇÃO" mesmo com erros de digitação.
    Aceita: VALIDAÇÃO, VALIDACAO, VALDAÇÃO, VADAÇÃO, VALIDAÇAO, etc.
    """
    texto_normalizado = normalizar_texto(texto)
    
    # Lista de variações comuns de erro
    variacoes = [
        "VALIDACAO",    # Correto sem acento
        "VALIDAÇÃO",    # Correto com acento (normalizado vira VALIDACAO)
        "VALDACAO",     # Faltando I
        "VALDAÇÃO",     # Faltando I com acento
        "VADACAO",      # Faltando LI
        "VADAÇÃO",      # Faltando LI com acento
        "VALIDACÃO",    # Ã no lugar errado
        "VALIDAÇAO",    # Ç sem til
        "VALI DACAO",   # Com espaço
        "VALIDA CAO",   # Com espaço
    ]
    
    # Verifica se alguma variação está no texto
    for variacao in variacoes:
        variacao_norm = normalizar_texto(variacao)
        if variacao_norm in texto_normalizado:
            return True
    
    # Busca mais genérica: palavras que começam com VAL e terminam com CAO
    # Isso pega variações como VALDCAO, VALIDCAO, etc.
    import re
    if re.search(r'VAL[DI]*[DA]*C[AÃ]*O', texto_normalizado):
        return True
    
    return False

class ColetorEmails:
    
    def __init__(self, nome_pasta: str = "Processamento Correios"):
        self.outlook = None
        self.inbox = None
        self.nome_pasta = nome_pasta
    
    def conectar(self) -> bool:
        try:
            self.outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = self.outlook.GetNamespace("MAPI")
            
            self.inbox = self._obter_pasta(namespace, self.nome_pasta)
            
            if self.inbox is None:
                logger.error(f"Pasta '{self.nome_pasta}' não encontrada. Usando Inbox padrão.")
                self.inbox = namespace.GetDefaultFolder(6)
            
            logger.info(f"✓ Conectado ao Outlook na pasta: {self.nome_pasta}")
            return True
        except Exception as e:
            logger.error(f"✗ Erro ao conectar ao Outlook: {e}")
            return False
    
    def _obter_pasta(self, namespace, nome_pasta: str):
        try:
            inbox_padrao = namespace.GetDefaultFolder(6)
            
            for pasta in inbox_padrao.Folders:
                if pasta.Name.lower() == nome_pasta.lower():
                    logger.info(f"✓ Pasta '{nome_pasta}' encontrada!")
                    return pasta
            
            for conta in namespace.Folders:
                for pasta in conta.Folders:
                    if pasta.Name.lower() == nome_pasta.lower():
                        logger.info(f"✓ Pasta '{nome_pasta}' encontrada na raiz!")
                        return pasta
            
            return None
        
        except Exception as e:
            logger.error(f"✗ Erro ao buscar pasta: {e}")
            return None
    
    def buscar_emails_do_dia(self) -> List[Dict]:
        try:
            emails_dados = []
            agora = datetime.now()
            
            for item in self.inbox.Items:
                try:
                    if not hasattr(item, 'Subject'):
                        continue
                    
                    subject = item.Subject
                    
                    # Usa a nova função que aceita variações de VALIDAÇÃO
                    if not contem_validacao(subject):
                        continue
                    
                    try:
                        received_time = item.ReceivedTime
                        if received_time.date() != agora.date():
                            continue
                    except:
                        continue
                    
                    email_info = self._extrair_dados_email(item, agora)
                    
                    if email_info:
                        emails_dados.append(email_info)
                
                except Exception as e:
                    logger.warning(f"Erro ao processar item: {e}")
                    continue
            
            logger.info(f"E-mails com VALIDAÇÃO encontrados: {len(emails_dados)}")
            return emails_dados
        
        except Exception as e:
            logger.error(f"✗ Erro ao buscar e-mails: {e}")
            return []
    
    def _extrair_dados_email(self, item, agora) -> Dict:
        try:
            subject = item.Subject
            
            # Usa a nova função que aceita variações de VALIDAÇÃO
            if not contem_validacao(subject):
                return None
            
            corpo = self._extrair_corpo_email(item)
            
            if "ALELO" in subject.upper():
                cliente = self._extrair_cliente_subject(subject)
            else:
                cliente = self._extrair_cliente(corpo)
            
            # EXTRAI AMBOS: SOMA e TOTAL informado
            total_soma = self._extrair_total_somando_contratos(corpo)
            total_informado = self._extrair_total(corpo)
            
            if not cliente:
                logger.warning(f"Cliente não encontrado. Subject: {subject}")
                return None
            
            logger.info(f"📧 Cliente: {cliente}")
            logger.info(f"   🧮 SOMA dos contratos: {total_soma}")
            logger.info(f"   📝 TOTAL informado: {total_informado}")
            
            return {
                "Cliente": cliente,
                "Total_Soma": total_soma,
                "Total_Informado": total_informado,
                "Subject": subject
            }
        
        except Exception as e:
            logger.error(f"✗ Erro ao extrair dados do e-mail: {e}")
            return None
    
    def _extrair_cliente_subject(self, subject: str) -> str:
        try:
            subject_limpo = normalizar_texto(subject)
            subject_limpo = subject_limpo.replace("VALIDACAO", "").replace("CORREIOS", "").strip()
            
            if subject_limpo.startswith("-"):
                subject_limpo = subject_limpo[1:].strip()
            
            partes = subject_limpo.split(" - ")
            
            if len(partes) > 0:
                cliente = partes[0].strip()
                return cliente
            
            return ""
        
        except Exception as e:
            logger.error(f"✗ Erro ao extrair cliente do subject: {e}")
            return ""
    
    def _extrair_corpo_email(self, item) -> str:
        try:
            if hasattr(item, 'Body'):
                return item.Body
            
            if hasattr(item, 'HTMLBody'):
                return item.HTMLBody
            
            return ""
        
        except Exception as e:
            logger.error(f"✗ Erro ao extrair corpo: {e}")
            return ""
    
    def _extrair_cliente(self, corpo: str) -> str:
        try:
            linhas = corpo.split('\n')
            
            for linha in linhas:
                if re.match(r'^\d{8,}', linha):
                    match = re.search(r'\b([A-Z][A-Z0-9]*[-_][A-Z0-9_]+)\b', linha)
                    if match:
                        return match.group(1)
            
            return ""
        
        except Exception as e:
            logger.error(f"✗ Erro ao extrair cliente: {e}")
            return ""
    
    def _extrair_total_somando_contratos(self, corpo: str) -> int:
        """
        Soma os valores individuais das linhas de contrato.
        """
        try:
            total_somado = 0
            linhas = corpo.split('\n')
            
            padrao_contrato = r'^\d{8,}\s+.*?\s+([A-Z0-9_-]+)\s+(\d+)\s*$'
            
            for linha in linhas:
                linha_limpa = linha.strip()
                
                if not linha_limpa or 'TOTAL' in linha_limpa.upper():
                    continue
                
                match = re.search(padrao_contrato, linha_limpa)
                
                if match:
                    valor = int(match.group(2))
                    total_somado += valor
                    logger.debug(f"  ➕ Linha: {linha_limpa[:50]}... | Valor: {valor}")
            
            return total_somado
        
        except Exception as e:
            logger.error(f"✗ Erro ao somar contratos: {e}")
            return 0
    
    def _extrair_total(self, corpo: str) -> int:
        """
        Extrai o TOTAL informado pelo usuário no e-mail.
        """
        try:
            match = re.search(r'TOTAL[\s:]+(\d+)', corpo, re.IGNORECASE)
            if match:
                return int(match.group(1))
            return 0
        
        except Exception as e:
            logger.error(f"✗ Erro ao extrair total: {e}")
            return 0