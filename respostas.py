# ======================== respostas.py (atualizado) ========================

import win32com.client
import logging
from datetime import datetime
import unicodedata
import re

logger = logging.getLogger(__name__)

# ============= FUNÇÕES AUXILIARES =============

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
    import re
    if re.search(r'VAL[DI]*[DA]*C[AÃ]*O', texto_normalizado):
        return True
    
    return False

def contem_kit(texto: str) -> bool:
    """
    Verifica se o texto contém "KIT" com variações de separadores.
    Aceita: KIT, _KIT, -KIT, KIT_, -KIT-, _KIT_, etc.
    """
    texto_normalizado = normalizar_texto(texto)
    
    # Remove espaços e caracteres especiais ao redor de KIT
    # Procura por KIT com qualquer separador antes/depois ou sozinho
    padrao = r'[_\-\s]*KIT[_\-\s]*'
    
    if re.search(padrao, texto_normalizado):
        return True
    
    return False

def identificar_tipo_alelo(subject: str, corpo: str = "") -> str:
    """
    Identifica se é ALELO-KIT, ALELO normal ou outro cliente.
    Retorna: "ALELO-KIT", "ALELO", ou None
    """
    subject_norm = normalizar_texto(subject)
    corpo_norm = normalizar_texto(corpo) if corpo else ""
    
    # Verifica se tem ALELO no subject ou corpo
    tem_alelo = "ALELO" in subject_norm or "ALELO" in corpo_norm
    
    if not tem_alelo:
        return None
    
    # Verifica se tem KIT usando a função flexível
    tem_kit = contem_kit(subject)
    
    if tem_kit:
        logger.info(f"🎯 ALELO-KIT detectado no título: {subject}")
        return "ALELO-KIT"
    else:
        logger.info(f"🎯 ALELO normal detectado no título: {subject}")
        return "ALELO"

class RespostorEmails:
    
    def __init__(self, nome_pasta: str = "Processamento Correios", nome_pasta_processados: str = "Correios Processados"):
        self.outlook = None
        self.inbox = None
        self.pasta_processados = None
        self.nome_pasta = nome_pasta
        self.nome_pasta_processados = nome_pasta_processados
    
    def conectar(self) -> bool:
        try:
            self.outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = self.outlook.GetNamespace("MAPI")
            
            self.inbox = self._obter_pasta(namespace, self.nome_pasta)
            
            if self.inbox is None:
                logger.error(f"Pasta '{self.nome_pasta}' não encontrada. Usando Inbox padrão.")
                self.inbox = namespace.GetDefaultFolder(6)
            
            self.pasta_processados = self._obter_ou_criar_pasta(namespace, self.nome_pasta_processados)
            
            if self.pasta_processados is None:
                logger.warning(f"Não foi possível criar pasta '{self.nome_pasta_processados}'. E-mails não serão movidos.")
            
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
    
    def _obter_ou_criar_pasta(self, namespace, nome_pasta: str):
        try:
            pasta = self._obter_pasta(namespace, nome_pasta)
            
            if pasta:
                return pasta
            
            inbox_padrao = namespace.GetDefaultFolder(6)
            nova_pasta = inbox_padrao.Folders.Add(nome_pasta)
            logger.info(f"✓ Pasta '{nome_pasta}' criada com sucesso!")
            return nova_pasta
        
        except Exception as e:
            logger.error(f"✗ Erro ao criar pasta: {e}")
            return None
    
    def responder_emails(self, dados_validacao: list):
        try:
            emails_respondidos = 0
            emails_ignorados = 0
            agora = datetime.now()
            
            for validacao in dados_validacao:
                cliente = validacao["Cliente"]
                status = validacao["Status"]
                
                # NOVA LÓGICA: Só processa e-mails com status OK
                if status != "✓ OK":
                    logger.info(f"⚠️ Cliente {cliente} com DIVERGÊNCIA - e-mail NÃO será respondido")
                    emails_ignorados += 1
                    continue
                
                email_encontrado = False
                
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
                        
                        if self._ja_foi_respondido(item):
                            logger.info(f"E-mail do cliente {cliente} já foi respondido. Ignorando.")
                            continue
                        
                        # NOVA LÓGICA: Verifica se o e-mail corresponde ao cliente
                        corpo = item.Body if hasattr(item, 'Body') else ""
                        
                        # Identifica o tipo de ALELO do e-mail
                        tipo_alelo_email = identificar_tipo_alelo(subject, corpo)
                        
                        # Verifica se o cliente da validação é ALELO
                        cliente_upper = cliente.upper()
                        
                        # Match para ALELO-KIT
                        if cliente_upper == "ALELO-KIT":
                            if tipo_alelo_email == "ALELO-KIT":
                                logger.info(f"✓ E-mail ALELO-KIT encontrado para validação: {cliente}")
                            else:
                                continue
                        
                        # Match para ALELO normal
                        elif "ALELO" in cliente_upper and "KIT" not in cliente_upper:
                            if tipo_alelo_email == "ALELO":
                                logger.info(f"✓ E-mail ALELO normal encontrado para validação: {cliente}")
                            else:
                                continue
                        
                        # Match para outros clientes
                        else:
                            if cliente_upper not in subject.upper() and cliente_upper not in corpo.upper():
                                continue
                        
                        logger.info(f"📧 E-mail encontrado para cliente: {cliente}")
                        
                        # Só envia resposta se for OK (sempre será neste ponto)
                        self._enviar_resposta_ok(item, validacao)
                        
                        self._mover_email(item, cliente)
                        
                        emails_respondidos += 1
                        email_encontrado = True
                        break
                    
                    except Exception as e:
                        logger.warning(f"Erro ao processar e-mail para {cliente}: {e}")
                        continue
                
                if not email_encontrado:
                    logger.warning(f"⚠️ E-mail não encontrado para cliente: {cliente}")
            
            logger.info(f"✓ {emails_respondidos} e-mail(s) respondido(s) com sucesso")
            logger.info(f"⚠️ {emails_ignorados} e-mail(s) com divergência (não respondidos)")
        
        except Exception as e:
            logger.error(f"✗ Erro ao responder e-mails: {e}")
    
    def _mover_email(self, item, cliente: str):
        try:
            if self.pasta_processados is None:
                logger.warning(f"Pasta de processados não disponível. E-mail de {cliente} não foi movido.")
                return
            
            item.Move(self.pasta_processados)
            logger.info(f"✓ E-mail de {cliente} movido para '{self.nome_pasta_processados}'")
        
        except Exception as e:
            logger.error(f"✗ Erro ao mover e-mail de {cliente}: {e}")
    
    def _ja_foi_respondido(self, item) -> bool:
        try:
            try:
                if hasattr(item, 'Replied') and item.Replied:
                    logger.info("E-mail já tem flag de respondido")
                    return True
            except:
                pass
            
            try:
                namespace = self.outlook.GetNamespace("MAPI")
                sent_items = namespace.GetDefaultFolder(5)
                
                conversation_id = item.ConversationID if hasattr(item, 'ConversationID') else None
                
                if conversation_id:
                    for sent_item in sent_items.Items:
                        try:
                            if hasattr(sent_item, 'ConversationID') and sent_item.ConversationID == conversation_id:
                                if hasattr(sent_item, 'SentOn') and hasattr(item, 'ReceivedTime'):
                                    if sent_item.SentOn > item.ReceivedTime:
                                        logger.info(f"Encontrada resposta anterior na conversa")
                                        return True
                        except:
                            continue
            except Exception as e:
                logger.warning(f"Erro ao verificar Sent Items: {e}")
            
            return False
        
        except Exception as e:
            logger.warning(f"Erro ao verificar se já foi respondido: {e}")
            return False
    
    def _enviar_resposta_ok(self, item_original, resultado):
        try:
            reply = item_original.ReplyAll()
            
            cliente = resultado["Cliente"]
            total_exibicao = resultado["Total_Exibicao"]  # Frontend usa este valor
            total_ga = resultado["Total_GA"]
            metodo = resultado["Metodo_Validacao"]
            
            # Monta corpo baseado no método de validação
            if "SOMA" in metodo:
                # TOTAL estava errado, mas SOMA validou
                corpo = f"""Bom dia,

Validação concluída com SUCESSO para o cliente {cliente}.

Detalhes:
- Total Email: {total_exibicao}
- Total GA: {total_ga}
- Status: ✓ OK

A validação foi processada corretamente.

Att."""
            else:
                # TOTAL estava correto
                corpo = f"""Bom dia,

Validação concluída com SUCESSO para o cliente {cliente}.

Detalhes:
- Total Email: {total_exibicao}
- Total GA: {total_ga}
- Status: ✓ OK

A validação foi processada corretamente.

Att."""
            
            reply.Body = corpo
            reply.Send()
            
            logger.info(f"✓ Resposta OK enviada para {cliente}")
        
        except Exception as e:
            logger.error(f"✗ Erro ao enviar resposta OK: {e}")