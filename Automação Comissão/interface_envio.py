import tkinter as tk
from tkinter import scrolledtext
import threading
import envio_emails
import logging
import subprocess  # <- Para rodar outro script

# Setup do log
logging.basicConfig(filename='log_automacao.txt', level=logging.INFO,
                    format='[%(asctime)s] [%(levelname)s] interface_envio - %(message)s', datefmt='%Y-%m-%d %H:%M:%S')

class InterfaceEnvioApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Envio de E-mails")
        self.root.geometry("700x650")

        self.label = tk.Label(root, text="Automação de Envio de Comissões", font=("Arial", 14))
        self.label.pack(pady=10)

        # Entradas para período, data NF e mensagem extra global
        self.periodo_label = tk.Label(root, text="Período (ex: 01/06/2025 a 30/06/2025):")
        self.periodo_label.pack()
        self.periodo_entry = tk.Entry(root, width=40)
        self.periodo_entry.insert(0, "01/06/2025 a 30/06/2025")
        self.periodo_entry.pack(pady=2)

        self.data_nf_label = tk.Label(root, text="Data limite da NF (ex: 20/07/2025):")
        self.data_nf_label.pack()
        self.data_nf_entry = tk.Entry(root, width=40)
        self.data_nf_entry.insert(0, "20/07/2025")
        self.data_nf_entry.pack(pady=2)

        self.msg_extra_label = tk.Label(root, text="Mensagem extra padrão (usada se planilha estiver vazia):")
        self.msg_extra_label.pack()
        self.msg_extra_entry = tk.Entry(root, width=80)
        self.msg_extra_entry.pack(pady=5)

        self.text_area = scrolledtext.ScrolledText(root, wrap=tk.WORD, width=80, height=25, font=("Courier", 10))
        self.text_area.pack(padx=10, pady=10)

        self.botao_enviar = tk.Button(root, text="Iniciar Envio", command=self.iniciar_envio_thread)
        self.botao_enviar.pack(pady=5)

        self.botao_atualizar = tk.Button(root, text="Atualizar Pastas de Licenciados", command=self.atualizar_pastas_thread)
        self.botao_atualizar.pack(pady=5)

    def iniciar_envio_thread(self):
        thread = threading.Thread(target=self.iniciar_envio)
        thread.daemon = True
        thread.start()

    def atualizar_pastas_thread(self):
        thread = threading.Thread(target=self.atualizar_pastas)
        thread.daemon = True
        thread.start()

    def iniciar_envio(self):
        self.log("🟡 Iniciando envio de e-mails...\n", nivel="info")
        try:
            periodo = self.periodo_entry.get().strip()
            data_nf = self.data_nf_entry.get().strip()
            msg_extra = self.msg_extra_entry.get().strip()
            envio_emails.enviar_em_lotes(log_func=self.log, periodo=periodo, data_nf=data_nf, msg_extra_global=msg_extra)
            self.log("✅ Envio concluído com sucesso.", nivel="info")
        except Exception as e:
            self.log(f"❌ Erro geral: {e}", nivel="error")

    def atualizar_pastas(self):
        self.log("🟡 Atualizando lista de pastas de licenciados...\n", nivel="info")
        try:
            resultado = subprocess.run(["python", "atualizar_pastas_licenciados.py"], capture_output=True, text=True)
            self.log(resultado.stdout)
            if resultado.stderr:
                self.log(f"⚠️ Erros durante a execução:\n{resultado.stderr}", nivel="warning")
            else:
                self.log("✅ Lista de pastas atualizada com sucesso.", nivel="info")
        except Exception as e:
            self.log(f"❌ Erro ao atualizar pastas: {e}", nivel="error")

    def log(self, mensagem, nivel="info"):
        print(mensagem)
        if nivel == "info":
            logging.info(mensagem)
        elif nivel == "warning":
            logging.warning(mensagem)
        elif nivel == "error":
            logging.error(mensagem)

        self.root.after(0, lambda: self._log_safe(mensagem))

    def _log_safe(self, mensagem):
        self.text_area.insert(tk.END, mensagem + "\n")
        self.text_area.see(tk.END)

if __name__ == "__main__":
    root = tk.Tk()
    app = InterfaceEnvioApp(root)
    root.mainloop()
