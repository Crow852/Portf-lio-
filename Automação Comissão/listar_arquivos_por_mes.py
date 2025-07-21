import json
import os
from googleapiclient.discovery import build
from google.oauth2 import service_account

SERVICE_ACCOUNT_FILE = 'automacao-comissoes-459718.json'
SCOPES = ['https://www.googleapis.com/auth/drive.readonly']

creds = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES
)
service = build('drive', 'v3', credentials=creds)

with open('pastas_licenciados.json', 'r', encoding='utf-8') as f:
    pastas_licenciados = json.load(f)

MES_BUSCADO = "06-2025"
arquivos_por_licenciado = {}

for licenciado, pasta_id in pastas_licenciados.items():
    try:
        # Buscar TODAS as subpastas com o nome do mês
        query = (
            f"'{pasta_id}' in parents and name = '{MES_BUSCADO}' "
            f"and mimeType = 'application/vnd.google-apps.folder' and trashed = false"
        )
        resultado = service.files().list(
            q=query,
            fields="files(id, name, createdTime)"
        ).execute()
        subpastas = resultado.get('files', [])

        if subpastas:
            # Ordenar por data de criação (mais recente primeiro)
            subpastas.sort(key=lambda x: x['createdTime'], reverse=True)
            subpasta_id = subpastas[0]['id']

            # Buscar os arquivos dentro da subpasta mais recente
            query_arquivos = f"'{subpasta_id}' in parents and trashed = false"
            arquivos = service.files().list(
                q=query_arquivos,
                fields="files(id, name)"
            ).execute().get('files', [])

            arquivos_por_licenciado[licenciado] = [
                {'id': arquivo['id'], 'name': arquivo['name']} for arquivo in arquivos
            ]
        else:
            arquivos_por_licenciado[licenciado] = []

    except Exception as e:
        print(f"Erro ao processar {licenciado}: {e}")
        arquivos_por_licenciado[licenciado] = []

# Salvar os arquivos encontrados
with open('arquivos_por_licenciado.json', 'w', encoding='utf-8') as f:
    json.dump(arquivos_por_licenciado, f, ensure_ascii=False, indent=2)

print("Arquivos por licenciado salvos em arquivos_por_licenciado.json.")
