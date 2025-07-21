import json
from googleapiclient.discovery import build
from google.oauth2 import service_account

# Caminho do seu arquivo de credenciais
SERVICE_ACCOUNT_FILE = 'automacao-comissoes-459718.json'
SCOPES = ['https://www.googleapis.com/auth/drive.metadata.readonly']

# ID da pasta-mãe que contém as pastas dos licenciados
PASTA_MAE_ID = '1yYN7e7oeLMrXYmM4G0Hc-m-l4I8-uuR1'  # <-- Substitua aqui

# Autenticação
creds = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES
)
service = build('drive', 'v3', credentials=creds)

# Consulta para pegar todas as subpastas dentro da pasta-mãe
query = (
    f"'{PASTA_MAE_ID}' in parents and mimeType = 'application/vnd.google-apps.folder' and trashed = false"
)

try:
    folders = []
    page_token = None

    while True:
        response = service.files().list(
            q=query,
            spaces='drive',
            fields='nextPageToken, files(id, name)',
            pageToken=page_token
        ).execute()

        folders.extend(response.get('files', []))
        page_token = response.get('nextPageToken', None)
        if page_token is None:
            break

    pastas_licenciados = {folder['name'].strip(): folder['id'] for folder in folders}

    with open('pastas_licenciados.json', 'w', encoding='utf-8') as f:
        json.dump(pastas_licenciados, f, ensure_ascii=False, indent=2)

    print(f"✅ JSON criado com {len(pastas_licenciados)} pastas.")
except Exception as e:
    print(f"❌ Erro ao buscar pastas: {e}")
