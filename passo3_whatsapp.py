from datetime import datetime
from urllib.parse import quote
import subprocess
import os
import time
import pyautogui

def passo3_whatsapp():
    # Formate a data
    data = datetime.now().strftime('%d/%m/%Y %H:%M:%S')

    # Crie a mensagem
    mensagem = "*DATACEAF*\nDados atualizados: " + data

    # Codifique a mensagem para formato URL
    mensagem_codificada = quote(mensagem)

    # Números de telefone (apenas os dígitos, sem espaços, parênteses ou hífens)
    telefones = ("61993532376", "33991333327")

    # Caminho para o Microsoft Edge
    edge_path = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"

    # Para cada número de telefone, crie o link e envie a mensagem
    for numero_telefone in telefones:
        # Crie o link
        url = f"https://web.whatsapp.com/send?phone=55{numero_telefone}&text={mensagem_codificada}"

        # Abre o Microsoft Edge com a URL especificada
        subprocess.Popen([edge_path, url])

        # Apertar enter (enviar a mensagem)
        time.sleep(10)
        pyautogui.press('enter')
        time.sleep(3)

        # Fecha o Edge
        os.system("taskkill /f /im msedge.exe")



