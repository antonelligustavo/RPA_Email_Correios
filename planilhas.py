# ======================== planilhas.py ========================

import pandas as pd
from datetime import datetime
import logging
import os
from typing import List, Dict
import requests

logger = logging.getLogger(__name__)

class GerenciadorPlanilhas:
    
    @staticmethod
    def salvar_emails(emails_dados: List[Dict], arquivo_template: str) -> str:
        try:
            pasta_saida = "resultados"
            os.makedirs(pasta_saida, exist_ok=True)
            
            data = datetime.now().strftime('%Y%m%d')
            arquivo = os.path.join(pasta_saida, arquivo_template.format(data=data))
            
            df = pd.DataFrame(emails_dados)
            df = df[["Cliente", "Total_Soma", "Total_Informado", "Subject"]]
            
            df.to_excel(arquivo, index=False, sheet_name="E-mails")
            logger.info(f"✓ Planilha de e-mails salva: {arquivo}")
            logger.info(f"  Total de linhas: {len(df)}")
            
            return arquivo
        
        except Exception as e:
            logger.error(f"✗ Erro ao salvar planilha de e-mails: {e}")
            return None
    
    @staticmethod
    def salvar_relatorios_ga(resultados_ga: Dict, arquivo_template: str) -> str:
        try:
            pasta_saida = "resultados"
            os.makedirs(pasta_saida, exist_ok=True)
            
            data = datetime.now().strftime('%Y%m%d')
            arquivo = os.path.join(pasta_saida, arquivo_template.format(data=data))
            
            df = pd.DataFrame(
                list(resultados_ga.items()),
                columns=["Cliente", "Total GA (Entregue)"]
            )
            
            df.to_excel(arquivo, index=False, sheet_name="Relatórios GA")
            logger.info(f"✓ Planilha de GA salva: {arquivo}")
            logger.info(f"  Total de clientes: {len(df)}")
            
            return arquivo
        
        except Exception as e:
            logger.error(f"✗ Erro ao salvar planilha de GA: {e}")
            return None
    
    @staticmethod
    def gerar_dados_validacao(emails_dados: List[Dict], resultados_ga: Dict) -> List[Dict]:
        try:
            emails_por_cliente = {}
            for item in emails_dados:
                cliente = item["Cliente"]
                emails_por_cliente[cliente] = {
                    "soma": item["Total_Soma"],
                    "informado": item["Total_Informado"]
                }
            
            dados_validacao = []
            
            for cliente in emails_por_cliente.keys():
                total_soma = emails_por_cliente[cliente]["soma"]
                total_informado = emails_por_cliente[cliente]["informado"]
                total_ga = resultados_ga.get(cliente, 0)
                
                logger.info(f"🔍 Validando {cliente}:")
                logger.info(f"   📊 SOMA: {total_soma} | INFORMADO: {total_informado} | GA: {total_ga}")
                
                # VALIDAÇÃO DUPLA (BACKEND): Se SOMA OU INFORMADO bater com GA = OK
                soma_ok = (total_soma == total_ga)
                informado_ok = (total_informado == total_ga)
                
                # Define qual valor será mostrado no FRONTEND
                if informado_ok:
                    # Se TOTAL informado está correto, usa ele
                    valor_exibicao = total_informado
                    metodo_validacao = "TOTAL"
                    status = "✓ OK"
                    logger.info(f"   ✅ Validado por TOTAL informado")
                elif soma_ok:
                    # Se TOTAL está errado mas SOMA está certa, usa SOMA
                    valor_exibicao = total_soma
                    metodo_validacao = "SOMA (TOTAL divergente)"
                    status = "✓ OK"
                    logger.info(f"   ✅ Validado por SOMA (TOTAL estava errado)")
                else:
                    # Ambos divergem, mantém TOTAL
                    valor_exibicao = total_informado
                    metodo_validacao = "Nenhum"
                    status = "✗ DIVERGÊNCIA"
                    logger.warning(f"   ❌ Nenhum valor bateu com GA!")
                
                dados_validacao.append({
                    "Cliente": cliente,
                    "Total_Soma": total_soma,  # Mantém no backend para logs
                    "Total_Informado": total_informado,  # Mantém no backend para logs
                    "Total_Exibicao": valor_exibicao,  # Valor mostrado no frontend
                    "Total_GA": total_ga,
                    "Metodo_Validacao": metodo_validacao,
                    "Status": status
                })
            
            return dados_validacao
        
        except Exception as e:
            logger.error(f"✗ Erro ao gerar dados de validação: {e}")
            return []
    
    @staticmethod
    def salvar_validacao(dados_validacao: List[Dict], arquivo_template: str) -> str:
        try:
            pasta_saida = "resultados"
            os.makedirs(pasta_saida, exist_ok=True)
            
            data = datetime.now().strftime('%Y%m%d')
            arquivo = os.path.join(pasta_saida, arquivo_template.format(data=data))
            
            # Salva todas as colunas no Excel (incluindo backend)
            df = pd.DataFrame(dados_validacao)
            df.to_excel(arquivo, index=False, sheet_name="Validação")
            
            logger.info(f"✓ Planilha de validação salva: {arquivo}")
            logger.info(f"  Total de clientes: {len(df)}")
            
            ok_count = (df["Status"] == "✓ OK").sum()
            divergencia_count = (df["Status"] == "✗ DIVERGÊNCIA").sum()
            
            logger.info(f"  ✓ OK: {ok_count}")
            logger.info(f"  ✗ DIVERGÊNCIA: {divergencia_count}")
            
            return arquivo
        
        except Exception as e:
            logger.error(f"✗ Erro ao salvar planilha de validação: {e}")
            return None
    
    @staticmethod
    def enviar_para_teams(dados_validacao: List[Dict]) -> bool:
        try:
            teams_webhook_url = os.getenv('TEAMS_WEBHOOK_URL')
            
            if not teams_webhook_url:
                logger.warning("TEAMS_WEBHOOK_URL não configurado no .env")
                return False
            
            timestamp = datetime.now().strftime('%d/%m/%Y %H:%M:%S')
            
            total_clientes = len(dados_validacao)
            total_ok = sum(1 for item in dados_validacao if "OK" in item["Status"])
            total_divergencias = sum(1 for item in dados_validacao if "DIVERGÊNCIA" in item["Status"])
            
            facts_adaptive = []
            
            for item in dados_validacao:
                cliente = item["Cliente"]
                total_exibicao = item["Total_Exibicao"]  # Frontend usa este valor
                total_ga = item["Total_GA"]
                status = item["Status"]
                metodo = item["Metodo_Validacao"]
                
                if "OK" in status:
                    icone_status = "✅"
                    # Mostra TOTAL ou SOMA dependendo de qual validou
                    if "SOMA" in metodo:
                        valor_texto = f"Email: {total_exibicao} (SOMA corrigida) | GA: {total_ga}"
                    else:
                        valor_texto = f"Email: {total_exibicao} | GA: {total_ga}"
                else:
                    icone_status = "❌"
                    # Em divergência, sempre mostra TOTAL
                    valor_texto = f"Email: {total_exibicao} | GA: {total_ga} ⚠️"
                
                facts_adaptive.append({
                    "title": f"{icone_status} {cliente}",
                    "value": valor_texto
                })
            
            if total_divergencias > 0:
                container_style = "attention"
                status_geral = "⚠️ DIVERGÊNCIAS DETECTADAS"
            else:
                container_style = "good"
                status_geral = "✅ Todas Validações OK"
            
            adaptive_payload = {
                "type": "message",
                "attachments": [{
                    "contentType": "application/vnd.microsoft.card.adaptive",
                    "contentUrl": None,
                    "content": {
                        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                        "type": "AdaptiveCard",
                        "version": "1.4",
                        "body": [
                            {
                                "type": "TextBlock",
                                "weight": "Bolder",
                                "size": "Medium",
                                "text": f"📊 Validação Correios - {status_geral}"
                            },
                            {
                                "type": "TextBlock",
                                "isSubtle": True,
                                "wrap": True,
                                "spacing": "None",
                                "text": f"**Execução:** {timestamp}"
                            },
                            {
                                "type": "Container",
                                "style": container_style,
                                "items": [
                                    {"type": "FactSet", "facts": facts_adaptive}
                                ]
                            },
                            {
                                "type": "TextBlock",
                                "wrap": True,
                                "text": f"**Total de clientes:** {total_clientes} | **✅ OK:** {total_ok} | **❌ Divergências:** {total_divergencias}"
                            }
                        ]
                    }
                }]
            }
            
            response = requests.post(teams_webhook_url, json=adaptive_payload, timeout=15)
            
            if response.status_code == 202:
                logger.info("✅ Relatório enviado para o Teams com sucesso!")
                return True
            else:
                logger.error(f"❌ Erro ao enviar para o Teams: {response.status_code}")
                logger.error(f"Resposta: {response.text}")
                return False
        
        except Exception as e:
            logger.error(f"❌ Erro ao enviar mensagem para o Teams: {e}")
            return False