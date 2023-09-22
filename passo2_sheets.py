import win32com.client
from datetime import datetime
import time
import pyautogui
import subprocess
import os

def passo2_sheets(lista_atualizacao):
    #Abrir o excel
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True
    caminho = r"\\srvnas\sctie\daf\CGMDE\documentos\GCEAF\CGMEDEX\SisCEAF\Bases de Dados\Sistemas\SisCEAF_Atualizador.xlsx"
    workbook = excel.Workbooks.Open(caminho)

    #PARA CADA ITEM DA LISTA ATUALIZACAO, RODAR
    for nome_aba, url in lista_atualizacao:
        # Seleciona a aba e copia os dados
        sheet = workbook.Worksheets(nome_aba)
        sheet.Activate()
        used_range = sheet.UsedRange
        used_range.Select()
        used_range.Copy()

        # Abre o Microsoft Edge com a URL especificada
        edge_path = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
        subprocess.Popen([edge_path, url])

        # Espera o navegador carregar e cola apenas os valores
        time.sleep(10)
        pyautogui.hotkey('ctrl', 'shift', 'v')
        time.sleep(30)

        # Fecha o Edge
        os.system("taskkill /f /im msedge.exe")
