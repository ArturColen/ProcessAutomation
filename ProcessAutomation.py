import pyautogui
import pyperclip
import time
import pandas as pd

pyautogui.PAUSE = 1

pyautogui.press('win')
pyautogui.write('chrome')
pyautogui.press('enter')

pyperclip.copy('https://drive.google.com/drive/folders/1aS5zKtPLntPs6VsJeIu3ufV0rZKAFYLQ?usp=sharing')
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('enter')
time.sleep(3)

pyautogui.click(x=500, y=400)
pyautogui.click(x=1650, y=200)
pyautogui.click(x=1500, y=645)
time.sleep(4)

spreadsheet = pd.read_excel(r'C:\Users\artur\Downloads\Sales.xlsx')

billing = spreadsheet['Valor Final'].sum()
quantity = spreadsheet['Quantidade'].sum()

pyautogui.hotkey('ctrl', 't')
pyperclip.copy('https://mail.google.com/mail/u/0/#inbox')
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('enter')
time.sleep(3)

pyautogui.click(x=80, y=200)
pyautogui.write('arturbcolen@gmail.com')
pyautogui.press('tab') # Select e-mail
pyautogui.press('tab')

pyperclip.copy('Relatório de vendas')
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('tab')

text = f"""
Prezados,
Segue o relatório de vendas.
Faturamento: R${billing:,.2f}
Quantidade de produtos vendidos: {quantity:,}

Qualquer dúvida estou à disposição.
Att,
Artur Bomtempo
"""

pyperclip.copy(text)
pyautogui.hotkey('ctrl', 'v')

pyautogui.hotkey('ctrl', 'enter')
