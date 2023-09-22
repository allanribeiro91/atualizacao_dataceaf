from datetime import datetime
import pyautogui
from passo1_excel import passo1_excel
from passo2_sheets import passo2_sheets
from passo3_whatsapp import passo3_whatsapp


#Início do processo de atualização
print("Início!")
print("Data/hora de início: ", datetime.now().strftime('%d/%m/%Y %H:%M:%S'))

#PASSO 1 - Atualizar os dados no excel (SisCEAF Atualizador)
caminho = r"\\srvnas\sctie\daf\CGMDE\documentos\GCEAF\CGMEDEX\SisCEAF\Bases de Dados\Sistemas\SisCEAF_Atualizador.xlsx"
passo1_excel(caminho)

#PASSO 2 - Levar os dados para o Google Sheets
lista_atualizacao = (
    ("tab1_programacao_med", "https://docs.google.com/spreadsheets/d/1aYa_FL_HZFENtXkReVCOtgaitWZC12AqldCCGbnkdro/edit#gid=0"),
    ("tab2_programacao_uf", "https://docs.google.com/spreadsheets/d/1aYa_FL_HZFENtXkReVCOtgaitWZC12AqldCCGbnkdro/edit#gid=598546521")
    ("tab3_distribuicao", "https://docs.google.com/spreadsheets/d/1aYa_FL_HZFENtXkReVCOtgaitWZC12AqldCCGbnkdro/edit#gid=129111645"),
    ("tab0_abastecimento", "https://docs.google.com/spreadsheets/d/1zvYF2169nzqPvBsAnnI2dDg3d81C8pyCwjU-neoATt0/edit#gid=1971235427"),
    ("tab1_contratos", "https://docs.google.com/spreadsheets/d/1zvYF2169nzqPvBsAnnI2dDg3d81C8pyCwjU-neoATt0/edit#gid=0"),
    ("tab2_contratos_parcelas", "https://docs.google.com/spreadsheets/d/1zvYF2169nzqPvBsAnnI2dDg3d81C8pyCwjU-neoATt0/edit#gid=437130809"),
    ("tab3_estoquesismat", "https://docs.google.com/spreadsheets/d/1zvYF2169nzqPvBsAnnI2dDg3d81C8pyCwjU-neoATt0/edit#gid=1986728444")
)
passo2_sheets(lista_atualizacao)

#PASSO 3 - Encaminhar whatsapp
passo3_whatsapp()

#Finalização do processo de atualização
print("Data/hora de início: ", datetime.now().strftime('%d/%m/%Y %H:%M:%S'))
print("Fim!")


#Supender
pyautogui.hotkey('win', 'l')
