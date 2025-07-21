# envio_emails.py

import os
import json
import base64
import pickle
import mimetypes
import time

from googleapiclient.discovery import build
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders

from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# Configuração
SCOPES = ['https://www.googleapis.com/auth/gmail.send', 'https://www.googleapis.com/auth/drive.readonly']

# ======== 🔐 AUTENTICAÇÃO OAUTH2 (COM RENOVAÇÃO DE TOKEN) ========
credentials = None
if os.path.exists('token.pickle'):
    with open('token.pickle', 'rb') as token:
        credentials = pickle.load(token)

try:
    if not credentials or not credentials.valid:
        if credentials and credentials.expired and credentials.refresh_token:
            credentials.refresh(Request())
        else:
            raise Exception("Token inválido ou corrompido.")
except Exception as e:
    print(f"⚠️ Token inválido ou expirado. Reautenticando... ({e})")
    flow = InstalledAppFlow.from_client_secrets_file('credentials_oauth.json', SCOPES)
    credentials = flow.run_local_server(port=0)

    with open('token.pickle', 'wb') as token:
        pickle.dump(credentials, token)

# Inicializa os serviços da Google API
drive_service = build('drive', 'v3', credentials=credentials)
gmail_service = build('gmail', 'v1', credentials=credentials)

# ========= FUNÇÕES =========

def baixar_arquivo(file_id, nome_arquivo):
    request = drive_service.files().get_media(fileId=file_id)
    with open(nome_arquivo, 'wb') as f:
        f.write(request.execute())

def enviar_email(destinatarios, assunto, corpo, anexos=[]):
    mensagem = MIMEMultipart()
    mensagem['to'] = ', '.join(destinatarios)
    mensagem['subject'] = assunto
    mensagem.attach(MIMEText(corpo, 'html'))


    for anexo in anexos:
        tipo_mime, _ = mimetypes.guess_type(anexo)
        if tipo_mime is None:
            tipo_mime = 'application/octet-stream'
        tipo_principal, subtipo = tipo_mime.split('/', 1)

        with open(anexo, 'rb') as f:
            mime = MIMEBase(tipo_principal, subtipo)
            mime.set_payload(f.read())
            encoders.encode_base64(mime)
            mime.add_header('Content-Disposition', f'attachment; filename="{os.path.basename(anexo)}"')
            mensagem.attach(mime)

    raw = base64.urlsafe_b64encode(mensagem.as_bytes()).decode()
    mensagem_body = {'raw': raw}

    gmail_service.users().messages().send(userId='me', body=mensagem_body).execute()

def normalizar_nome(nome):
    return ' '.join(nome.upper().strip().split())

# ========= FUNÇÃO PRINCIPAL =========

def enviar_em_lotes(log_func=print, periodo="01/06/2025 a 30/06/2025", data_nf="20/07/2025", msg_extra_global=""):
    # Carrega os dados
    with open('arquivos_por_licenciado.json', 'r', encoding='utf-8') as f:
        arquivos_por_licenciado = json.load(f)

    with open('dados_licenciados.json', 'r', encoding='utf-8') as f:
        licenciados = json.load(f)

    licenciados_normalizados = {normalizar_nome(nome): dados for nome, dados in licenciados.items()}
    os.makedirs('tmp', exist_ok=True)

    # Parâmetros de controle
    TAMANHO_LOTE = 10
    INTERVALO_ENTRE_EMAILS = 2
    INTERVALO_ENTRE_LOTES = 10

    licenciados_enviados = 0
    total_licenciados = len(arquivos_por_licenciado)
    itens = list(arquivos_por_licenciado.items())

    for i in range(0, total_licenciados, TAMANHO_LOTE):
        lote = itens[i:i + TAMANHO_LOTE]
        log_func(f"🚀 Enviando lote {i // TAMANHO_LOTE + 1} ({len(lote)} e-mails)...")

        for nome_licenciado, arquivos in lote:
            nome_normalizado = normalizar_nome(nome_licenciado)

            if nome_normalizado in licenciados_normalizados:
                dados = licenciados_normalizados[nome_normalizado]
                email_raw = dados['email']
                comissao = dados['comissao']
                mensagem_extra = dados.get('mensagem', '').strip()
                anexos = []

                # Suporta múltiplos e-mails separados por vírgula
                destinatarios = [e.strip() for e in email_raw.split(',') if e.strip()]

                for arquivo in arquivos:
                    nome_drive = arquivo.get('name') or f"arquivo_{arquivo['id']}.pdf"
                    nome_local = os.path.join("tmp", nome_drive)
                    baixar_arquivo(arquivo['id'], nome_local)
                    anexos.append(nome_local)

                # Corpo do e-mail personalizado
                corpo = f"""
<p>Segue anexo o relatório de premiações.</p>

<p>Reforço que o período computado se refere a <b>01/06/2025 a 30/06/2025</b> e que o relatório contempla apenas os faturamentos recebidos.<br>
Gentileza emitir a NF fiscal no valor de <b>{comissao}</b>, até dia <b>20/07/2025</b> para que possamos dar sequência no pagamento, até a data acordada em contrato.<br>
Favor incluir na NF os dados bancários para pagamento.</p>

<p><b>Informação importante:</b><br>
Comunicamos que não procederemos com a compensação do montante relativo à comissão em débitos em atraso. Em caso de inadimplência por parte do parceiro, no momento de pagamento das comissões, o valor devido será retido até que a situação seja regularizada. Ressaltamos que após a regularização, os pagamentos serão disponibilizados nas datas de vencimento estipuladas, ou seja, nos dias 10, 20 e 30 de cada mês.</p>

<p>Qualquer dúvida, estou à disposição.</p>

<p>Atenciosamente,<br>
Fredy Domingues<br>
CPE BH</p>
"""

                # Usa o nome do licenciado no assunto do e-mail
                assunto = f"LICENCIADA - {nome_licenciado} - {periodo}"

                try:
                    enviar_email(destinatarios, assunto, corpo, anexos)
                    log_func(f"📨 E-mail enviado para {', '.join(destinatarios)}")
                    licenciados_enviados += 1
                except Exception as e:
                    log_func(f"❌ Erro ao enviar para {', '.join(destinatarios)}: {e}")

                for anexo in anexos:
                    os.remove(anexo)

                time.sleep(INTERVALO_ENTRE_EMAILS)
            else:
                log_func(f"⚠️ '{nome_licenciado}' não encontrado na planilha (após normalização).")

        log_func(f"✅ Lote {i // TAMANHO_LOTE + 1} finalizado.\n")
        if i + TAMANHO_LOTE < total_licenciados:
            time.sleep(INTERVALO_ENTRE_LOTES)

    log_func(f"\n📬 Envio finalizado: {licenciados_enviados} e-mails enviados com sucesso.")
