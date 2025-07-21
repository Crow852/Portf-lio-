import subprocess
import os
import sys
import time
import socket
from pathlib import Path

print("🔧 Iniciando o app...")
time.sleep(1)

if getattr(sys, 'frozen', False):
    base_path = Path(sys._MEIPASS)
else:
    base_path = Path(__file__).parent

app_path = base_path / "app" / "app.py"
print(f"📄 Caminho do app.py: {app_path}")
time.sleep(1)

os.environ["STREAMLIT_SERVER_PORT"] = "8501"
os.environ["STREAMLIT_BROWSER_SERVER_PORT"] = "8501"
os.environ["STREAMLIT_BROWSER_GATHER_USAGE_STATS"] = "false"
os.environ["STREAMLIT_HEADLESS"] = "true"
print("🌐 Variáveis de ambiente definidas.")
time.sleep(1)

def is_port_open(host: str, port: int) -> bool:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.settimeout(1)
        return sock.connect_ex((host, port)) == 0

def run_streamlit():
    # Tente localizar o executável streamlit na variável PATH
    streamlit_cmd = "streamlit"
    if getattr(sys, 'frozen', False):
        # Se estiver empacotado, pode precisar do path completo
        possible_path = base_path / "Scripts" / "streamlit.exe"
        if possible_path.exists():
            streamlit_cmd = str(possible_path)
    
    command = [
        streamlit_cmd, "run",
        str(app_path),
        "--server.headless=true",
        "--server.port=8501"
    ]

    print(f"Executando comando: {' '.join(command)}")
    process = subprocess.Popen(command, shell=True)

    for _ in range(20):
        if is_port_open("localhost", 8501):
            subprocess.Popen(["cmd", "/c", "start", "http://localhost:8501"])
            break
        time.sleep(0.5)

    process.wait()

if __name__ == "__main__":
    run_streamlit()
