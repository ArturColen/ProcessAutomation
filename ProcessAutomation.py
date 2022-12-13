# Import the libraries
import pyautogui
import pyperclip
import time
import pandas as pd

# Control the speed at which each command will be executed
pyautogui.PAUSE = 1

# Open the chrome browser
pyautogui.press('win')
pyautogui.write('chrome')
pyautogui.press('enter')

# Enter the address where the spreadsheet will be picked up
pyperclip.copy('https://drive.google.com/drive/folders/1aS5zKtPLntPs6VsJeIu3ufV0rZKAFYLQ?usp=sharing')
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('enter')
time.sleep(3)

# Download the spreadsheet with the sales data
pyautogui.click(x=500, y=400) # Click on the file
pyautogui.click(x=1650, y=200) # Click the button to display more options for the file
pyautogui.click(x=1500, y=645) # Click the download button
time.sleep(4)

# Import the spreadsheet with data into Python
spreadsheet = pd.read_excel(r'C:\Users\artur\Downloads\Sales.xlsx')

# Calculate the company's indicators from the spreadsheet data
billing = spreadsheet['Valor Final'].sum()
quantity = spreadsheet['Quantidade'].sum()

# Send an e-mail with the company's indicators
pyautogui.hotkey('ctrl', 't')
pyperclip.copy('https://mail.google.com/mail/u/0/#inbox')
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('enter')
time.sleep(3)

pyautogui.click(x=80, y=200) # Click the write e-mail button
pyautogui.write('arturbcolen@gmail.com') # Fill in the e-mail address to which the report will be sent
pyautogui.press('tab') # Select e-mail
pyautogui.press('tab') # Skip to the subject field

pyperclip.copy('Relatório de vendas') # Write the subject of the e-mail
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('tab') # Skip to the body field of the e-mail

text = f"""
Prezados,
Segue o relatório de vendas.
Faturamento: R${billing:,.2f}
Quantidade de produtos vendidos: {quantity:,}

Qualquer dúvida estou à disposição.
Att,
Artur Bomtempo
"""

pyperclip.copy(text) # Write the text that will be sent in the e-mail
pyautogui.hotkey('ctrl', 'v')

pyautogui.hotkey('ctrl', 'enter') # Send the e-mail