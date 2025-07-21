import subprocess
import time
import os
import json
import sys
import logging

# Setup do log
logging.basicConfig(filename='log_automacao.txt', level=logging.INFO,
                    format='[%(asctime)s] [%(levelname)s] executar_tudo - %(message)s', datefmt='%Y-%m-%d %H:%M:%S')

def rodar(script, descricao):
    print(f"\n▶️ {descricao}...")
    logging.info("Iniciando: %s", descricao)
    try:
        subprocess.run(["python", script], check=True)
        print(f"✅ {descricao} finalizado com sucesso.")
        logging.info("Finalizado com sucesso: %s", descricao)
    except subprocess.CalledProcessError as e:
        print(f"❌ Erro ao executar {script}: {e}")
        logging.error("Erro ao executar %s: %s", script, str(e))
        sys.exit(1)

def verificar_json(caminho, descricao):
    print(f"🔍 Verificando {descricao} em '{caminho}'...")
    if not os.path.exists(caminho):
        print(f"❌ Arquivo {caminho} não foi encontrado.")
        logging.error("Arquivo %s não encontrado durante verificação de %s.", caminho, descricao)
        sys.exit(1)

    try:
        with open(caminho, 'r', encoding='utf-8') as f:
            dados = json.load(f)
        if not dados:
            print(f"⚠️ Atenção: O arquivo {caminho} está vazio.")
            logging.warning("Arquivo %s está vazio (%s).", caminho, descricao)
        else:
            print(f"✅ {descricao} carregado com sucesso. {len(dados)} registros encontrados.")
            logging.info("%s carregado com sucesso. %d registros.", descricao, len(dados))
    except json.JSONDecodeError:
        print(f"❌ Erro ao ler o JSON de {caminho}. Verifique o conteúdo.")
        logging.error("Erro ao carregar JSON: %s", caminho)
        sys.exit(1)

if __name__ == "__main__":
    print("🔁 Iniciando execução completa do projeto...\n")
    logging.info("Execução do projeto iniciada.")

    rodar("listar_arquivos_por_mes.py", "1. Listagem de arquivos no Google Drive")
    verificar_json("arquivos_por_licenciado.json", "Arquivos por Licenciado")

    time.sleep(1)

    rodar("leitor_planilha.py", "2. Leitura da planilha de licenciados")
    verificar_json("dados_licenciados.json", "Dados dos Licenciados")

    time.sleep(1)

    print("\n🚀 Pronto! Agora abrindo a interface gráfica para envio dos e-mails.")
    logging.info("Abrindo interface gráfica de envio.")
    subprocess.run(["python", "interface_envio.py"])
