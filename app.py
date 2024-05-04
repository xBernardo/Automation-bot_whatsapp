"""
Automatização de Mensagens no Whatsapp
"""
import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui

webbrowser.open('https://web.whasapp.com/')
sleep(30)  # Pausa de 30s para você se conectar com seu wpp vai QRcode

workbook = openpyxl.load_workbook('teste.xlsx')
pagina_clientes = workbook['Planilha1']

for linha in pagina_clientes.iter_rows(min_row=2):
    nome = linha[0].value # index da planilha
    telefone = linha[1].value
    vencimento = linha[2].value

    mensagem = f'Olá {nome}, Bom dia, caloteiro seu boleto vence dia {vencimento.strftime('%d/%m/%Y')}. Por favor, nos pague seu lindo'

    # Criando os links personalizados do whatsapp
    # Enviar mensagens para cada cliente com base nos dados da planilha
    link_mensagem_wpp = f'https://web.whasapp.com/send?phone={telefone}&text={quote(mensagem)}'
    webbrowser.open(link_mensagem_wpp)
    sleep(10)  # Tempo de 10s para recarregamento da página
    try:
        seta = pyautogui.locateCenterOnScreen('enter.png')
        sleep(5)   # Tempo de 5s para recarregamento da página
        pyautogui.click(seta[0], seta[1])
        sleep(5)
    except:
        print(f'Não foi possivel enviar mensagem para o caloteiro {nome}')
        with open('erros.csv', 'a', newline='', encoding='utf-8') as arquivo:
        # arquivo.csv, no qual vai separar informações com virgula
            arquivo.write(f'{nome}, {telefone}')