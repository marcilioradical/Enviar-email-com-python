import wn32com.client as win32

#  criar a integração com o outlook
outlook =
win32.Dispatch('outlook.application')

#criar um email
email - outlook.CreateItem(0)

faturamento = 1500
qtde_produtos = 10
ticket_medio = faturamento / qtde_produtos

#configurar as informações dom seu e-mail
emeil.To ="destino; destino2"
email.Subject = "E-mail automático do Python"
email.HTMLBody = f"""
<p> Olá Lira, aqui é o codigo Python</p>

<p>O faturamento da loja foi de
R${faturamento}</p>
<p>Vendemos {qtde_produtos} produtos</>
<p>O ticket Médio foi de R${ticket_medio}</p>

<p>Abs,</p>
<p>Código Python</p>
"""

# anexo =
"C://User/joap/Downloads/arquivo.xlsx"
# email.Attachments.Add(anexo)

email.Send()
print("Email Enviado")