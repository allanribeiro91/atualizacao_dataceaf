import win32com.client
from datetime import datetime
import time

def countdown_timer(seconds):
    for i in range(seconds, 0, -1):
        print(f"Tempo restante: {i} segundos", end='\r')
        time.sleep(1)

def passo1_excel(caminho):
    # Inicie o Excel
    excel = win32com.client.Dispatch("Excel.Application")
    
    # Tornar o Excel visível (opcional)
    excel.Visible = True
    
    # Abra o arquivo desejado
    workbook = excel.Workbooks.Open(caminho)
    
    # Atualize todas as conexões de dados
    for connection in workbook.Connections:
        connection.Refresh()
    
    #Tempo de espera
    countdown_timer(300)

    # Acesse a aba "home"
    home_sheet = workbook.Worksheets("home")

    # Atualize a célula Q30 com a data e hora atual
    current_datetime = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    home_sheet.Range("Q30").Value = current_datetime

    # Salve e feche o arquivo
    workbook.Save()
    workbook.Close()
    
    # Feche o Excel
    excel.Quit()