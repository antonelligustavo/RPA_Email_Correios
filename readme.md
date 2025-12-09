# Sistema de Validação Correios

Sistema automatizado para validação de entregas dos Correios, integrando dados do Outlook e do sistema GA (Flash Courier).

## 📋 Funcionalidades

- **Coleta automática de e-mails** do Outlook com validações de entregas
- **Extração de relatórios** do sistema GA via Selenium
- **Validação cruzada** entre dados informados e registros do GA
- **Respostas automáticas** aos e-mails com resultado da validação
- **Notificações no Microsoft Teams** com resumo das validações
- **Movimentação automática** de e-mails processados para pasta específica

## 🔧 Requisitos

### Sistema Operacional
- Windows (necessário para integração com Outlook via pywin32)

### Software
- Python 3.8 ou superior
- Microsoft Outlook instalado e configurado
- Google Chrome (para Selenium)
- ChromeDriver compatível com sua versão do Chrome

### Credenciais
- Acesso ao sistema GA (Flash Courier)
- Webhook URL do Microsoft Teams (opcional)

## 📦 Instalação

1. Clone o repositório:
```bash
git clone <url-do-repositorio>
cd <nome-do-diretorio>
```

2. Crie um ambiente virtual (recomendado):
```bash
python -m venv venv
```

3. Ative o ambiente virtual:
```bash
# Windows
venv\Scripts\activate
```

4. Instale as dependências:
```bash
pip install -r requirements.txt
```

5. Configure o arquivo `.env` na raiz do projeto:
```env
# Credenciais do GA (Flash Courier)
GA_EMAIL=seu_email@exemplo.com
GA_SENHA=sua_senha

# Webhook do Microsoft Teams (opcional)
TEAMS_WEBHOOK_URL=https://outlook.office.com/webhook/...
```

## 📁 Estrutura do Projeto

```
.
├── config.py           # Configurações gerais (URLs, caminhos, XPaths)
├── emails.py           # Coleta e processamento de e-mails do Outlook
├── ga.py              # Extração de dados do sistema GA via Selenium
├── planilhas.py       # Geração e salvamento de planilhas Excel
├── respostas.py       # Envio automático de respostas aos e-mails
├── main.py            # Orquestrador principal do sistema
├── .env               # Variáveis de ambiente (não versionado)
├── requirements.txt   # Dependências Python
└── README.md         # Documentação
```

## 🚀 Como Usar

### Execução Básica

Execute o script principal:
```bash
python main.py
```

### Fluxo de Execução

O sistema executa automaticamente as seguintes etapas:

1. **Coleta de E-mails**: Busca e-mails do dia atual na pasta "Processamento Correios" do Outlook que contenham variações de "VALIDAÇÃO" no assunto

2. **Extração de Dados**: 
   - Extrai nome do cliente
   - Calcula soma dos contratos individuais
   - Obtém total informado pelo usuário

3. **Consulta ao GA**: 
   - Faz login automaticamente no sistema GA
   - Busca relatórios para cada cliente
   - Baixa e processa planilhas Excel

4. **Validação Cruzada**:
   - Compara total informado vs. total do GA
   - Se divergir, compara soma calculada vs. total do GA
   - Gera status: ✓ OK ou ✗ DIVERGÊNCIA

5. **Geração de Relatórios**:
   - Cria planilhas Excel na pasta `resultados/`
   - Envia notificação ao Microsoft Teams

6. **Respostas Automáticas**:
   - Responde cada e-mail com resultado da validação
   - Move e-mails para pasta "Correios Processados"

## 📊 Planilhas Geradas

O sistema gera três planilhas na pasta `resultados/`:

- **emails_YYYYMMDD.xlsx**: Dados extraídos dos e-mails
- **ga_relatorios_YYYYMMDD.xlsx**: Totais obtidos do GA
- **validacao_YYYYMMDD.xlsx**: Resultado da validação cruzada

## 🎯 Casos de Uso Especiais

### Cliente ALELO
O sistema possui tratamento especial para o cliente ALELO:
- Busca por "ELO-RE" no sistema GA
- Filtra entregas sem "_KIT" para ALELO normal
- Filtra entregas com "_KIT" para ALELO-KIT

### Variações de "VALIDAÇÃO"
O sistema aceita diversas variações no assunto do e-mail:
- VALIDAÇÃO, VALIDACAO
- VALDAÇÃO, VADAÇÃO
- VALIDACÃO, VALIDAÇAO
- Com ou sem espaços

## ⚙️ Configurações Avançadas

### Pasta do Outlook
Por padrão, o sistema busca e-mails na pasta "Processamento Correios". Para alterar:

```python
# Em main.py
coletor = ColetorEmails(nome_pasta="Sua Pasta Customizada")
```

### Pasta de Processados
E-mails respondidos são movidos para "Correios Processados". Para alterar:

```python
# Em main.py
responsor = RespostorEmails(
    nome_pasta="Processamento Correios",
    nome_pasta_processados="Sua Pasta de Processados"
)
```

### Download Path
Por padrão, arquivos são baixados em `./downloads`. Para alterar:

```python
# Em config.py
class ConfigGA:
    DOWNLOAD_PATH = "C:/seu/caminho/customizado"
```

## 🔍 Logs

O sistema gera logs detalhados no console com informações sobre:
- Conexões estabelecidas
- E-mails processados
- Consultas ao GA
- Validações realizadas
- Erros encontrados

Nível de log: `INFO` (pode ser alterado em `main.py`)

## ⚠️ Solução de Problemas

### Erro ao conectar ao Outlook
- Verifique se o Outlook está instalado e configurado
- Execute o Python com permissões de administrador

### ChromeDriver incompatível
- Baixe o ChromeDriver compatível com sua versão do Chrome
- Adicione ao PATH do sistema

### Arquivo .env não encontrado
- Certifique-se de criar o arquivo `.env` na raiz do projeto
- Verifique se as variáveis `GA_EMAIL` e `GA_SENHA` estão definidas

### Pasta do Outlook não encontrada
- O sistema usará a Inbox padrão se não encontrar a pasta especificada
- Crie manualmente a pasta "Processamento Correios" no Outlook

## 📝 Licença

Este projeto é de uso interno.

## 🤝 Contribuindo

Para contribuir com melhorias:
1. Faça um fork do projeto
2. Crie uma branch para sua feature
3. Commit suas mudanças
4. Push para a branch
5. Abra um Pull Request

## 📧 Suporte

Para dúvidas ou problemas, entre em contato com a equipe de desenvolvimento.