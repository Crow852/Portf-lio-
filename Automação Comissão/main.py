import json
from googleapiclient.discovery import build
from google.oauth2 import service_account

# --- CONFIGURAÇÃO ---
SERVICE_ACCOUNT_FILE = 'automacao-comissoes-459718.json'  # Caminho do seu .json
SCOPES = ['https://www.googleapis.com/auth/drive']
PASTA_PRINCIPAL_ID = '1yYN7e7oeLMrXYmM4G0Hc-m-l4I8-uuR1'  # ID da pasta principal no Drive

# --- AUTENTICAÇÃO COM GOOGLE DRIVE ---
credentials = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES)
service = build('drive', 'v3', credentials=credentials)

# --- LISTAR SUBPASTAS ---
def listar_subpastas(pasta_id):
    query = f"'{pasta_id}' in parents and mimeType = 'application/vnd.google-apps.folder' and trashed = false"
    resultado = service.files().list(
        q=query,
        spaces='drive',
        fields='files(id, name)'
    ).execute()

    arquivos = resultado.get('files', [])
    pastas = {arquivo['name']: arquivo['id'] for arquivo in arquivos}
    return pastas

# --- EXECUÇÃO ---
pastas_licenciados = listar_subpastas(PASTA_PRINCIPAL_ID)

# --- SALVAR COMO JSON ---
with open('pastas_licenciados.json', 'w', encoding='utf-8') as f:
    json.dump(pastas_licenciados, f, ensure_ascii=False, indent=4)

print("✅ Arquivo pastas_licenciados.json salvo com sucesso!")
