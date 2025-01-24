# Baixar biblotecas
import pandas as pd
from datetime import datetime, timedelta
import yagmail
# pip install openpyx
# pip install yagmail
# pip install pandas
# pip install datetime



# Função para enviar email usando yagmail
def enviar_email(destinatario, assunto, mensagem):
    remetente = 'texte5460@gmail.com'
    senha = 'vbxw oifj ptcm qyuk'
    
    yag = yagmail.SMTP(remetente, senha)
    yag.send(to=destinatario, subject=assunto, contents=mensagem)
caminho_arquivo = 'D:\Heitor\Documents\PlanilhaTeste.xlsx' # caminho para a planilha dos clientes
df = pd.read_excel(caminho_arquivo)

# Para a coluna chamada 'DataRetirada'
df['DataRetirada'] = pd.to_datetime(df['DataRetirada'], format='%d/%m/%Y')

# Calcula a data de devolução automática (45 dias após a retirada)
df['DataDevolucao'] = df['DataRetirada'] + pd.DateOffset(days=45)

# Verifica se a data de devolução é menor que a data atual
hoje = datetime.now()
atrasados = df[df['DataDevolucao'] < hoje]

# Se houver devoluções atrasadas, envia um email
for index, row in atrasados.iterrows():
    destinatario = row['EmailCliente']  # Supondo que há uma coluna 'EmailCliente'
    assunto = 'Devolução Pendente'
    dias_atraso = (hoje - row['DataDevolucao']).days
    mensagem = f"Favor devolver o produto retirado em {row['DataRetirada'].strftime('%d/%m/%Y')}, conforme a data de devolução ({row['DataDevolucao'].strftime('%d/%m/%Y')}) já passou. Já se passaram {dias_atraso} dias do prazo de devolução."

    enviar_email(destinatario, assunto, mensagem)

# Exibe o resultado
print(df)
