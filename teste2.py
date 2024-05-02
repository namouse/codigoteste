import tkinter as tk
from tkinter import filedialog
import requests
import openpyxl
import time
from tkinter import ttk
import os

class MinhaInterface:
    def __init__(self, janela):
        self.janela = janela
        self.janela.title("Comandos")
        self.progresso = ttk.Progressbar(janela, orient="horizontal", length=220, mode="determinate")
        self.progresso.grid(row=0, column=0, pady=10, columnspan=2)

        self.rotulo_porcentagem = tk.Label(janela, text="0%")
        self.rotulo_porcentagem.grid(row=1, column=0, columnspan=2)

        self.botao_arquivo = tk.Button(janela, text="Escolher Arquivo", command=self.escolher_arquivo)
        self.botao_arquivo.grid(row=2, column=0, pady=5, columnspan=2)

        self.rotulo_nome_arquivo = tk.Label(janela, text="")
        self.rotulo_nome_arquivo.grid(row=3, column=0, columnspan=2, pady=5)

        self.label_tamanho = tk.Label(janela, text="Linhas do Excel:")
        self.label_tamanho.grid(row=4, column=0, pady=1)

        self.entrada_tamanho = tk.Entry(janela)
        self.entrada_tamanho.grid(row=4, column=1, padx=1)

        self.botao_enviar = tk.Button(janela, text="Enviar", command=self.enviar_comandos)
        self.botao_enviar.grid(row=5, column=0, columnspan=2, pady=10)

    
        self.placa_desenvolvido_por = tk.Label(janela, text="Desenvolvido por João Pedro")
        self.placa_desenvolvido_por.grid(row=6, column=0, sticky=tk.SW, pady=(0, 0), padx=(1, 0))

    def escolher_arquivo(self):
        self.arquivo = filedialog.askopenfilename(defaultextension=".xlsx", filetypes=[("Arquivos Excel", "*.xlsx")])
        print("Arquivo selecionado:", self.arquivo)

        nome_arquivo = os.path.basename(self.arquivo)
        self.rotulo_nome_arquivo.config(text=f"Arquivo Selecionado: {nome_arquivo}")

    def enviar_comandos(self):
        tamanho_texto = self.entrada_tamanho.get()
        tamanho = int(tamanho_texto) + 1
        print("Tamanho do arquivo:", tamanho)

        self.progresso["value"] = 0
        self.progresso["maximum"] = tamanho

        workbook = openpyxl.load_workbook(self.arquivo)
        sheet = workbook.active

        coluna1 = [str(cell.value) for cell in sheet['A']]
        coluna2 = [str(cell.value) for cell in sheet['B']]

        linha = 1
        while linha < tamanho:
            valor_coluna1 = coluna1[linha - 1] if linha <= len(coluna1) else None
            valor_coluna2 = coluna2[linha - 1] if linha <= len(coluna2) else None

            if valor_coluna1 is None and valor_coluna2 is None:
                break

            print(f"Valor da Coluna 1: {valor_coluna1}, Valor da Coluna 2: {valor_coluna2}")

            phone_number = valor_coluna2
            message = valor_coluna1

            url = "https://api.smsmarket.com.br/webservice-rest/send-single.php"

            payload = {
                'number': phone_number,
                'content': message,
                'type': '0',
            }

            headers = {
                'Authorization': "Basic c3lzdGVtc2F0OlN5c3RlbXNhdDIwMjRA",
            }

            response = requests.post(url, data=payload, headers=headers)
            print(response.text)
            time.sleep(1)

            porcentagem = int((linha / tamanho) * 100)
            self.progresso["value"] = linha
            self.rotulo_porcentagem.config(text=f"{porcentagem}%")
            self.janela.update_idletasks()

            linha += 1

if __name__ == "__main__":
    
    janela_principal = tk.Tk()
    app = MinhaInterface(janela_principal)

    janela_principal.geometry("320x200")  # Adjusted geometry
    janela_principal.mainloop()