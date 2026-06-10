import os
import gspread
from google.oauth2.service_account import Credentials
from dotenv import load_dotenv

# Carrega as variáveis de ambiente do arquivo .env na raiz do projeto
load_dotenv()

# Escopo de permissão necessário para leitura e escrita em planilhas do Google
SCOPE = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive"
]

def obter_cliente_gspread():
    """
    Autentica e retorna um cliente do gspread usando as credenciais do Google.
    
    Retorna:
        gspread.Client: Cliente autenticado do gspread.
        
    Exceções:
        FileNotFoundError: Se o arquivo de credenciais não for encontrado.
        Exception: Se ocorrer um erro durante a autenticação.
    """
    # Obtém o caminho do arquivo de credenciais a partir das variáveis de ambiente
    # Padrão: 'credentials.json' na raiz do projeto
    credenciais_path = os.getenv("GOOGLE_CREDENTIALS_PATH", "credentials.json")
    
    # Verifica se o arquivo de credenciais existe no caminho especificado
    if not os.path.exists(credenciais_path):
        raise FileNotFoundError(
            f"Arquivo de credenciais do Google não encontrado em: {credenciais_path}\n"
            "Por favor, verifique se a variável GOOGLE_CREDENTIALS_PATH no arquivo .env "
            "está apontando para o local correto do seu arquivo de credenciais (ex: service_account.json)."
        )
    
    try:
        # Carrega as credenciais a partir do arquivo JSON de conta de serviço
        creds = Credentials.from_service_account_file(credenciais_path, scopes=SCOPE)
        
        # Autentica e retorna o cliente do gspread
        client = gspread.authorize(creds)
        return client
        
    except Exception as e:
        raise Exception(f"Erro ao autenticar com o Google Drive: {str(e)}")

def obter_planilha(titulo_ou_url):
    """
    Abre e retorna uma planilha específica do Google Sheets.
    
    Args:
        titulo_ou_url (str): O título exato da planilha ou a URL completa da planilha.
        
    Returns:
        gspread.Spreadsheet: Objeto da planilha aberta.
    """
    try:
        client = obter_cliente_gspread()
        
        # Tenta abrir pela URL primeiro (método mais seguro e preciso)
        if titulo_ou_url.startswith("http"):
            planilha = client.open_by_url(titulo_ou_url)
        else:
            # Caso contrário, tenta abrir pelo título exato
            planilha = client.open(titulo_ou_url)
            
        return planilha
        
    except gspread.SpreadsheetNotFound:
        raise Exception(
            f"Planilha não encontrada: '{titulo_ou_url}'.\n"
            "1. Verifique se o título está exatamente igual ao da planilha no Google Drive, "
            "ou prefira usar a URL completa da planilha.\n"
            "2. Certifique-se de que o e-mail da 'Conta de Serviço' (presente no seu arquivo JSON de credenciais) "
            "foi adicionado como 'Leitor' (ou 'Editor') na planilha do Google Drive."
        )
    except Exception as e:
        raise Exception(f"Erro ao acessar a planilha: {str(e)}")