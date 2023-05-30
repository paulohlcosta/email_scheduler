#!/usr/bin/env python
# coding: utf-8

# In[9]:


# este codigo vai listar os arquivos de uma pasta, por ordem de (data de modificacao)
# e enviar um deles por email


# Importar os módulos necessários
import os
import time
import datetime
import sys
import logging
from datetime import datetime 

# bibliotecas do email
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

# Criar um objeto log com o nível desejado
log = logging.getLogger(__name__)
log.setLevel(logging.ERROR)
# Criar um objeto FileHandler com o nome do arquivo e o modo de escrita
fh = logging.FileHandler("log.txt", mode="w")
# Adicionar o FileHandler ao log
log.addHandler(fh)


print("\nName of Python script:", sys.argv)
timestamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
caminho = 'C:\\Users\\PAULO.COSTA\\Downloads\\'

# Definir as informações do email
email_de = "sender@mail.com"
email_para = "recipient@mail.com"
assunto = "Relatório de notas fiscais"
mensagem = "Segue em anexo relatório gerado em " + timestamp
anexo = "arquivo que vai enviar.extensão"



flag0 = flag1 = flag2 = flag3 = True 
now = datetime.now()
last_hour = now.hour - 1

print(' --- Iniciando Python... ---\n', timestamp)

if len(sys.argv) == 1: 
    flag0 = False    #------------------------------ para executar uma unica vez = nao passar argumentos
else:
    try:
        int(sys.argv[1])
        horario_executar = []
        for x in sys.argv[1:]:
            horario_executar.append(int(x))
        print('\n\nConfiguração para execução segunda a sábado, diariamente às', horario_executar, 'horas.\n\n')
    except:
        flag0 = False

# esta função deverá avaliar se está na hora de executar, retornando True ou False;
# poderá executar até 1 vez por hora
def verifica_agenda():
    global flag1, flag2, flag3, last_hour, now
    print(now.day, now.hour, now.minute, 'iniciando verifica_agenda...', end="")
    if (now.hour != last_hour) and (now.hour in horario_executar): flag1 = True
    else: flag1 = False
    if (now.minute > 30): flag2 = True
    else: flag2 = False
    if (now.weekday != 6): flag3 = True
    else: flag3 = False

    print('    | resultado dos flags:', flag1 * flag2 * flag3)
    print('execução diaria as', horario_executar, '| aguardando proximo ciclo...')
    time.sleep(2)
    return(flag1 * flag2 * flag3)


# In[12]:


# Definir a função lista_arquivos que recebe como parâmetro a pasta a ser avaliada
def lista_arquivos(pasta):
    global file_notas
    
    # Criar uma lista vazia para armazenar os nomes e as datas de modificação dos arquivos
    arquivos = []

    # Percorrer todos os arquivos na pasta
    for arquivo in os.listdir(pasta):
        # Obter o caminho completo do arquivo
        caminho = os.path.join(pasta, arquivo)
        # Verificar se é um arquivo e não uma subpasta
        if os.path.isfile(caminho):
            # Obter a data de modificação do arquivo em segundos desde a época
            data = os.path.getmtime(caminho)
            # Converter a data em um objeto datetime
            data = datetime.fromtimestamp(data)
            # Adicionar o nome e a data do arquivo à lista
            arquivos.append((arquivo, data))

    # Ordenar a lista em ordem decrescente de data de modificação
    arquivos.sort(key=lambda x: x[1], reverse=True)

    file_notas = arquivos[1][0]
    print('Arquivo das notas fiscais:', file_notas)
    
    # Imprimir os 4 primeiros elementos da lista
    for i in range(4):
        print(arquivos[i][0], arquivos[i][1])


# In[14]:


#pasta = "C:\\Users\\TEC6\\Downloads"
#lista_arquivos(pasta)


# In[ ]:


def eemail(anexo):
    global email_de, email_para, assunto, mensagem, flag, last_hour
    
# Criar objeto do email
    msg = MIMEMultipart()
    msg["From"] = email_de
    msg["To"] = email_para
    msg["Subject"] = assunto

# Adicionar corpo do email
    msg.attach(MIMEText(mensagem, "plain"))

# Adicionar anexo
    with open(anexo, "rb") as f:
        arquivo = MIMEApplication(f.read(), _subtype="pdf")
        arquivo.add_header("content-disposition", "attachment", filename=anexo)
        msg.attach(arquivo)
        
# Enviar o email
    print('  Enviando arquivo por email...')
    NOVO = smtplib.SMTP("smtp.server.com", 587)
    try:
        print('start try')
        NOVO.starttls()
        print('tls ok')
        NOVO.login(email_de, "password")
        print('login ok')
        NOVO.sendmail(email_de, email_para, msg.as_string())
        print('   ............................................ email enviado com sucesso!')     
        last_hour = now.hour        
    except smtplib.SMTPResponseException as err:
        error_code = err.smtp_code
        print('ERROR CODE:', error_code)
        error_message = SMTP_ERROR_CODES.get(error_code, f"Unknown error ({error_code})")
        log.error(f"Observed exception while send email to recipient email!,\n "
                  f"Exception: {error_message.format(err.smtp_error)}")
    except Exception as Err:
        log.error(f"Observed exception while send email to recipient email!,\n "
                  f"Exception: {Err}")
        print('::: ocorreu um erro no envio do email, veja o log :::')    
        
    NOVO.quit()


# In[ ]:


while flag0 == True:                         # --------------- loop infinito
    now = datetime.now()
    if verifica_agenda():                    # quando a funcao retorna 1 ou True, executar:
        try:
            lista_arquivos(caminho)
            eemail(caminho + file_notas)
        except:
            pass
    time.sleep(58)

lista_arquivos(caminho)
eemail(caminho + file_notas)
    
print('END OF PYTHON PROGRAM')

