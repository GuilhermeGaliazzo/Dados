import pyautogui, time, pyperclip, pandas

email = "guilhermemassari.rodrigues@gmail.com"



pyautogui.PAUSE = 2

read = pandas.read_excel(r"C:\Users\guilh\Desktop\Vendas - Dez.xlsx")
valortotal = read["Valor Final"].sum()
qtdtotal = read["Quantidade"].sum()
mensagem = f"""Boa tarde, 

Valor total vendido esse mês: R${valortotal:,.2f} reais
Quantidade de produtos vendidos esse mês: {qtdtotal} produtos

obrigado.
"""

#Acessar navegador
pyautogui.press("win")
pyautogui.write("Microsoft Edge")
pyautogui.press("enter")

time.sleep(3)
#Acessar gmail
pyautogui.write("https://mail.google.com/mail/u/0/")
pyautogui.press("enter")

#Escrever um novo email
pyautogui.click(x=78, y=175)
pyautogui.write(email)
pyautogui.press("tab")
pyautogui.press("tab")
pyautogui.write("Boa tarde")
pyautogui.press("tab")
pyperclip.copy(mensagem)
pyautogui.hotkey("ctrl", "v")
pyautogui.click(x=1301, y=992)


