import pyautogui, webbrowser, openpyxl
from urllib.parse import quote
from time import sleep

numbers = openpyxl.load_workbook('numeros.xlsx')
page = numbers['Planilha1']

for line in page.iter_rows(min_row=2):
    # nome, telefone
    name = line[0].value
    phone = line[1].value

    msg = (f'Eae {name} Isso é só um teste de um botzinho em python, mas se quiser me enviar um pix, pode mandar no meu '
           f'número <3!!')

    # Tratamento de erro.
    try:
        link_msg_whatsapp = f'https://web.whatsapp.com/send?phone={phone}&text={quote(msg)}'
        webbrowser.open(link_msg_whatsapp)
        sleep(10)
        arrow = pyautogui.locateOnScreen('send.png')
        sleep(5)
        pyautogui.click(arrow[0], arrow[1])
        sleep(5)
        pyautogui.hotkey('ctrl', 'w')
        sleep(5)
    except:
        print(f'Não foi possivel enviar mensagem para {name}')
        with open('erros.csv', 'a', newline='', encoding="utf-8") as file:
            file.write(f'{name}, {phone}')