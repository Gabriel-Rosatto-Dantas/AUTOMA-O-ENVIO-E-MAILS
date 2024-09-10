!pip install pretty_html_table

import smtplib
from email.mime.text import MIMEText
from google.colab import auth
import gspread
from google.auth import default
import pandas as pd
from pretty_html_table import build_table
from time import sleep

# Autenticação no Google Drive
auth.authenticate_user()
creds, _ = default()
gc = gspread.authorize(creds)

# Abertura da base de dados no Google Sheets
basedados_atendimento = gc.open('FUP BASES ATENDIMENTO')
basedados_recebimento = gc.open('FUP BASES RECEBIMENTO')

def enviar_email(assunto, corpo, remetente, destinatarios, senha, cc=None):
    msg = MIMEText(corpo, 'html')  # Defina o tipo de conteúdo como HTML
    msg['Subject'] = assunto
    msg['From'] = remetente
    msg['To'] = ', '.join(destinatarios)  # Converta a lista de destinatários em uma única string

    if cc:
        msg['Cc'] = ', '.join(cc)  # Converta a lista de Cc em uma única string
        destinatarios += cc  # Adicione os endereços de Cc à lista de destinatários

    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp_server:
        smtp_server.login(remetente, senha)
        smtp_server.sendmail(remetente, destinatarios, msg.as_string())
        print("E-mail enviado!")

def percorrer_abas_recebimento():
    abas = basedados_recebimento.worksheets()

    # Dicionário que mapeia nomes de abas para endereços de emails
    mapeamento_abas_emails = {
        'AJU': ['E-mail1'],
        'BEL': ['E-mail2'],
        'BVB': ['E-mail3'],
        'BSB': ['E-mail4'],
        'CGR': ['E-mail5'],
        'GRU': ['E-mail6'],
        'CML': ['E-mail7'],
        'CNF': ['E-mail8'],
        'CGH': ['E-mail9'],
        'CWB': ['E-mail10'],
        'FLN': ['E-mail12'],
        'FOR': ['E-mail13'],
        'IGU': ['E-mail14'],
        'GIG': ['E-mail15'],
        'GYN': ['E-mail16'],
        'IOS': ['E-mail17'],
        'IMP': ['E-mail18'],
        'JJG': ['E-mail19'],
        'JPA': ['E-mail20'],
        'JOI': ['E-mail21'],
        'LDB': ['E-mail22'],
        'MCP': ['E-mail23'],
        'MAO': ['E-mail23'],
        'MAB': ['E-mail25'],
        'MGF': ['E-mail26'],
        'NVT': ['E-mail27'],
        'PMW': ['E-mail28'],
        'NAT': ['E-mail29'],
        'POA': ['E-mail30'],
        'BPS': ['E-mail31'],
        'PVH': ['E-mail32'],
        'REC': ['E-mail33'],
        'RAO': ['E-mail34'],
        'RBR': ['E-mail35'],
        'MCZ': ['E-mail36'],
        'SJP': ['E-mail37'],
        'SSA': ['E-mail38'],
        'STM': ['E-mail39'],
        'SDU': ['E-mail40'],
        'MRO': ['E-mail41'],
        'SLZ': ['E-mail42'],
        'THE': ['E-mail43'],
        'UDI': ['E-mail44'],
        'CGB': ['E-mail45'],
        'VCP': ['E-mail46'],
        'VIX': ['E-mail47'],
        'XAP': ['E-mail48'],
        'OPS': ['E-mail49']
    }
    cc = ['felipe.ginicolo@latam.com', 'pedro.alves@latam.com']

    for aba in abas:
        nome_aba = aba.title
        if nome_aba in mapeamento_abas_emails:
            destinatarios = mapeamento_abas_emails[nome_aba]
            dados = aba.get_all_values()
            if len(dados) > 1:  # Verifica se há mais de uma linha de dados
                nome_contato = dados[1][1]  # Nome na primeira coluna (coluna A)
                assunto = f'ITENS PENDENTES DE RECEBIMENTO - {nome_contato}'
                informacoes = pd.DataFrame(dados)
                informacoes.to_csv('informacoes.csv', index=False, header=False)
                # Leia o arquivo CSV com separação adequada
                informacoes_formatada = pd.read_csv('informacoes.csv', sep=',', header=None)
                informacoes_formatada.columns = informacoes_formatada.iloc[0].values
                informacoes_formatada = informacoes_formatada[1:]  # Remove a primeira linha (os números das colunas)
                informacoes_formatada = informacoes_formatada.reset_index(drop=True)  # Remove os números das linhas (índice)
                table_html = build_table(informacoes_formatada, color='blue_light', font_size='12px', font_family='sans-serif')
                corpo = f'''
                <html>
                    <body>
                        Ola Equipe {nome_contato}! Espero que estejam bem !<br>
                        Por favor, poderia nos ajudar com recebimento no SAP as linhas abaixo? Materiais entregues conforme detalhes abaixo.<br><br>
                        {table_html}<br><br>
                        Muito obrigado!!!<br><br>
                        Gabriel Dantas<br>
                    </body>
                </html>
            '''
                enviar_email(assunto, corpo, remetente, destinatarios, senha, cc)
                sleep(20)

def percorrer_abas_atendimento():
    abas = basedados_atendimento.worksheets()

    # Dicionário que mapeia nomes de abas para endereços de emails
    mapeamento_abas_emails = {
        'AJU': ['E-mail1'],
        'BEL': ['E-mail2'],
        'BVB': ['E-mail3'],
        'BSB': ['E-mail4'],
        'CGR': ['E-mail5'],
        'GRU': ['E-mail6'],
        'CML': ['E-mail7'],
        'CNF': ['E-mail8'],
        'CGH': ['E-mail9'],
        'CWB': ['E-mail10'],
        'FLN': ['E-mail12'],
        'FOR': ['E-mail13'],
        'IGU': ['E-mail14'],
        'GIG': ['E-mail15'],
        'GYN': ['E-mail16'],
        'IOS': ['E-mail17'],
        'IMP': ['E-mail18'],
        'JJG': ['E-mail19'],
        'JPA': ['E-mail20'],
        'JOI': ['E-mail21'],
        'LDB': ['E-mail22'],
        'MCP': ['E-mail23'],
        'MAO': ['E-mail23'],
        'MAB': ['E-mail25'],
        'MGF': ['E-mail26'],
        'NVT': ['E-mail27'],
        'PMW': ['E-mail28'],
        'NAT': ['E-mail29'],
        'POA': ['E-mail30'],
        'BPS': ['E-mail31'],
        'PVH': ['E-mail32'],
        'REC': ['E-mail33'],
        'RAO': ['E-mail34'],
        'RBR': ['E-mail35'],
        'MCZ': ['E-mail36'],
        'SJP': ['E-mail37'],
        'SSA': ['E-mail38'],
        'STM': ['E-mail39'],
        'SDU': ['E-mail40'],
        'MRO': ['E-mail41'],
        'SLZ': ['E-mail42'],
        'THE': ['E-mail43'],
        'UDI': ['E-mail44'],
        'CGB': ['E-mail45'],
        'VCP': ['E-mail46'],
        'VIX': ['E-mail47'],
        'XAP': ['E-mail48'],
        'OPS': ['E-mail49']
    }
    cc = ['Email da copia1', 'Email da Copia2']

    for aba in abas:
        nome_aba = aba.title
        if nome_aba in mapeamento_abas_emails:
            destinatarios = mapeamento_abas_emails[nome_aba]
            dados = aba.get_all_values()
            if len(dados) > 1:  # Verifica se há mais de uma linha de dados
                nome_contato = dados[1][0]  # Nome na primeira coluna (coluna A)
                assunto = f'ITENS PENDENTES ENVIO - {nome_contato}'
                informacoes = pd.DataFrame(dados)
                informacoes.to_csv('informacoes_atendimento.csv', index=False, header=False)
                # Leia o arquivo CSV com separação adequada
                informacoes_formatada = pd.read_csv('informacoes_atendimento.csv', sep=',', header=None)
                informacoes_formatada.columns = informacoes_formatada.iloc[0].values
                informacoes_formatada = informacoes_formatada[1:]  # Remove a primeira linha (os números das colunas)
                informacoes_formatada = informacoes_formatada.reset_index(drop=True)  # Remove os números das linhas (índice)
                table_html = build_table(informacoes_formatada, color='red_light', font_size='12px', font_family='sans-serif')
                corpo = f'''
                <html>
                    <body>
                        Ola Equipe {nome_contato}! Espero que estejam bem !<br>
                        Por favor, poderia nos ajudar com finalização do atendimento das requisições abaixo?<br><br>
                        Linhas proximas do vencimento SLA da fase.<br><br>
                        {table_html}<br><br>
                        Muito obrigado!!!<br><br>
                        Gabriel Dantas<br>
                    </body>
                </html>
            '''
                enviar_email(assunto, corpo, remetente, destinatarios, senha, cc)
                sleep(20)

# Configurações
remetente = "Seu e-mail"
senha = "Sua senha"  # Substitua pela sua senha

# Chama a função para percorrer as abas
percorrer_abas_atendimento()
percorrer_abas_recebimento()
