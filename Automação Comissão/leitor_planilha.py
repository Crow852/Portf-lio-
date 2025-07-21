from google.oauth2 import service_account
from googleapiclient.discovery import build
import json
import logging
from datetime import datetime

# Setup do log
logging.basicConfig(filename='log_automacao.txt', level=logging.INFO, 
                    format='[%(asctime)s] [%(levelname)s] leitor_planilha - %(message)s', datefmt='%Y-%m-%d %H:%M:%S')

try:
    SERVICE_ACCOUNT_FILE = 'automacao-comissoes-459718.json'
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

    credentials = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES
    )

    SHEET_ID = '1cbfl0ReEWGxZ1UvbxpWEvdF51VfMWlQT9VMDz8_5CGU'
    ABA = 'Página1'

    service = build('sheets', 'v4', credentials=credentials)
    sheet = service.spreadsheets()

    # Note que aqui ajustei o range para incluir a coluna D (Mensagem)
    result = sheet.values().get(
        spreadsheetId=SHEET_ID,
        range=f"{ABA}!A2:D"
    ).execute()

    values = result.get('values', [])

    dados_licenciados = {}

    for row in values:
        if len(row) >= 3:
            nome = row[0].strip().upper()
            comissao = row[1].strip()
            email = row[2].strip()
            mensagem = row[3].strip() if len(row) >= 4 else ""
            dados_licenciados[nome] = {
                "email": email,
                "comissao": comissao,
                "mensagem": mensagem
            }

    with open('dados_licenciados.json', 'w', encoding='utf-8') as f:
        json.dump(dados_licenciados, f, indent=4, ensure_ascii=False)

    print("✅ Dados da planilha salvos em dados_licenciados.json")
    logging.info("Dados da planilha salvos com sucesso. Total: %d registros.", len(dados_licenciados))

except Exception as e:
    print(f"❌ Erro ao ler planilha: {e}")
    logging.error("Erro ao ler a planilha: %s", str(e))
    raise
