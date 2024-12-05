import csv
import time
import pandas as pd
import smtplib
import email.message

alerts_list = {} #lista para fazer o arquivo xlsx contendo todos os alertas
body = "" #inicia a variavel body vazia



def comparator_file_column(column, alerts): #função de if para comparar as quantidades
    global body #define a varivel body como global para reutilizar depois

    if int(column[2]) < 1: #vê se a quantidade na esteira é menor que 1
        print('\n \033[31m Esteira 1 está com estoque abaixo do crítico \033[0m') #a mensagem na tela
        if 'esteira1' not in alerts: #caso não exista a chave esteira 1 cria ela
            alerts['esteira1'] = []
        alert = {
            'data': column[0],
            'hora': column[1],
            'Status': 'Estoque abaixo do crítico'
        } # cria um dicionario para adicionar ao alerts_list
        alerts['esteira1'].append(alert) #adiciona


        body += """<p>A Esteira 1 está com o estoque abaixo do crítico!</p>""" #incrementa a variavel body um texto em html paara criar o corpo do email

    if int(column[3]) < 1:
        print('\n \033[31m Esteira 2 está com estoque abaixo do crítico \033[0m')
        if 'esteira2' not in alerts:
            alerts['esteira2'] = []
        alert = {
            'data': column[0],
            'hora': column[1],
            'Status': 'Estoque abaixo do crítico'
        }
        alerts['esteira2'].append(alert)


        body += """<p>A Esteira 2 está com o estoque abaixo do crítico!</p>"""

    if int(column[4]) < 1:
        print('\n \033[31m Esteira 3 está com estoque abaixo do crítico \033[0m')
        if 'esteira3' not in alerts:
            alerts['esteira3'] = []
        alert = {
            'data': column[0],
            'hora': column[1],
            'Status': 'Estoque abaixo do crítico'
        }
        alerts['esteira3'].append(alert)


        body += """<p>A Esteira 3 está com o estoque abaixo do crítico!</p>"""



def reader_csv(endpoint, alerts): #leitor do arquivo csv
    with open(endpoint, 'r') as file: #abre o arquivo csv
        reader = csv.reader(file, delimiter=',') # a variavel reader servirá como o nosso leitor para o arquivo
        row = 0 # variavel que será usada para ver em qual linha está
        for column in reader: # loop usado para ler todas as colunas
            if row > 0: # se a linha não for a linha 0
                comparator_file_column(column, alerts) # chama a função comparadora
            row += 1 #adiciona um a linha
            time.sleep(1) # tempo de delay para não sobrecarregar o usuário de informações


def xlsx_creator(alerts): #função que cria o o arquivo excel
    if alerts: #se tiver a lista alerta criará a outra lista all_alerts
        all_alerts = []
        for esteira, alert_list in alerts.items(): #vai percorrer todos os itens e ever se a esteira está nela
            for alert in alert_list: # se a o alert estiver na alert_list
                all_alerts.append({
                    'Esteira': esteira,
                    'Data': alert['data'],
                    'Hora': alert['hora'],
                    'Status': alert['Status']
                })#novo dicionario que servira para criar o excel

        df = pd.DataFrame(all_alerts) #armazena os dados em forma de tabelas
        df.to_excel('alertas.xlsx', index=False) #cria o arquivo
        print('\n Arquivo Excel criado com sucesso!') #notifica o usuário se o documento foi criado



def send_email(body): #função para criar o email
    body_email = body #o corpo do email sera a váriavel global body

    msg = email.message.Message()
    msg['Subject'] = "Alerta de estoque baixo" #assunto do email
    msg['From'] = 'sprocodigo@gmail.com' #email do remetente código, ensira seu email
    msg['To'] = 'martins07.yasmim@gmail.com' #email do destinatario, ensira o email desejado
    password = 'kydp qmkg trqt pxat' #senha do email do rementente
    msg.add_header('Content-Type', 'text/html') #tipo de linguagem que será escrito o corpo do email
    msg.set_payload(body_email) #define o conteudo do corpo do email

    try:
        smt = smtplib.SMTP('smtp.gmail.com', 587 ) #cria conexão com o servidor smtp
        smt.starttls() #cria uma conexão segura com o servidor
        smt.login(msg['From'], password) #autentifica o usuário
        smt.sendmail(msg['From'], [msg['To']], msg.as_string().encode('utf-8')) # que converte a mensagem para um tipo adequado para o envio
        print("E-mail enviado com sucesso!") #aviisa se foi enviado com sucesso
    except Exception as e: #avisa caso haja erro em enviar
        print(f"Erro ao enviar o e-mail: {e}")



csv_file = 'C:/Users/53661815857/Documents/integrador da marcia/Esp8266_Receiver.csv' #caminho do documento
reader_csv(csv_file, alerts_list) #chamando a função reader_csv
xlsx_creator(alerts_list) #chama a função xlsx_creator
send_email(body) #chama a função send_email
