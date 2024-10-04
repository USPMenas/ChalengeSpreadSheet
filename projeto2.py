import os
import smtplib
from email.message import EmailMessage
from imbox import Imbox
import pandas as pd


PATH_DATABASE = 'dataBase/banco_de_dados.xlsx'

# Informações de login do email
EMAIL_ADDRESS = 'mas.serclaro@gmail.com'
EMAIL_PASSWORD = 'dxxb vznv sxff pfmq'  # Use sua senha de app ou credenciais

def update_spreadsheet():
    # Carregar a planilha antiga
    caminho_planilha_antiga = 'dataBase/banco_de_dados.xlsx'  # Substitua pelo caminho correto da planilha antiga
    df_antiga = pd.read_excel(caminho_planilha_antiga)  # Ajuste o skiprows conforme necessário

    # Carregar a planilha nova
    caminho_planilha_nova = 'trashCan/spreedSheet.xlsx'  # Substitua pelo caminho correto da planilha nova
    df_nova = pd.read_excel(caminho_planilha_nova, skiprows=13)

    # Ajuste a seleção de colunas com base nos nomes corretos
    colunas_desejadas = ['Data/hora', 'Local', 'Produto', 'Quantidade', 'Valor (R$)']  # Ajuste os nomes conforme necessário
    df_nova = df_nova[colunas_desejadas]

    # Remover ' un' da coluna 'Quantidade' e converter para inteiro em ambas as planilhas
    df_nova['Quantidade'] = df_nova['Quantidade'].str.replace(' un', '', regex=False)

    # Substituir os valores NaN por zero (ou outro valor padrão)
    df_nova['Quantidade'] = df_nova['Quantidade'].fillna(0).astype(int)

    # Converter a coluna 'Data/hora' para datetime para garantir que a comparação seja feita corretamente
    # df_antiga['Data/hora'] = pd.to_datetime(df_antiga['Data/hora'])
    df_nova['Data/hora'] = pd.to_datetime(df_nova['Data/hora'])

    # Filtrar os dados da nova planilha que não estão na antiga (comparando apenas as datas)
    novos_dados = df_nova[~df_nova['Data/hora'].isin(df_antiga['Data/hora'])]

    # Adicionar os novos dados à planilha antiga
    df_atualizado = pd.concat([novos_dados, df_antiga], ignore_index=True)

    # Salvar a planilha antiga com os novos dados adicionados
    novo_caminho_planilha_antiga = 'dataBase/banco_de_dados.xlsx'
    df_atualizado.to_excel(novo_caminho_planilha_antiga, index=False)

    print(f"Planilha atualizada!'")


# Função para enviar resposta com anexo
def send_reply_with_attachment(to_email, subject, body, attachment_path):
    # Cria uma mensagem de email
    msg = EmailMessage()
    msg['Subject'] = subject
    msg['From'] = EMAIL_ADDRESS
    msg['To'] = to_email
    msg.set_content(body)

    # Lê o arquivo e o anexa ao email
    with open(attachment_path, 'rb') as f:
        file_data = f.read()
        file_name = os.path.basename(attachment_path)
        msg.add_attachment(file_data, maintype='application', subtype='octet-stream', filename=file_name)

    # Conecta ao servidor SMTP e envia o email
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
        smtp.send_message(msg)
        print(f"Resposta enviada para {to_email} com o arquivo '{file_name}' anexado.")

while(True):
    # Conexão com o servidor IMAP do Gmail
    with Imbox('imap.gmail.com',
            username=EMAIL_ADDRESS,
            password=EMAIL_PASSWORD,
            ssl=True,
            starttls=False) as imbox:

        # Pega todas as mensagens com o assunto 'dados@'
        inbox_messages_subject_spreadsheet = imbox.messages(subject='dados@', unread=True)

        # Obtém o diretório atual onde o script está rodando
        current_directory = os.getcwd()

        if inbox_messages_subject_spreadsheet:
            print(f"{len(inbox_messages_subject_spreadsheet)} mensagens encontradas com o assunto 'dados@'")
        else:
            print("Nenhuma mensagem encontrada com o assunto 'dados@'")

        for uid, message in inbox_messages_subject_spreadsheet:

            # Verifica se há anexos na mensagem
            if message.attachments:
                print(f"Anexos encontrados na mensagem com UID: {uid}")
                for attachment in message.attachments:
                    print(f"Analisando o anexo com nome: {attachment['filename']} e tipo: {attachment['content-type']}")
                    
                    # Verifica se o anexo é um arquivo Excel (XLSX)
                    if attachment['content-type'] == 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet':
                        # Pega o nome e o conteúdo do anexo
                        xlsx_name = attachment['filename']
                        xlsx_content = attachment['content']  # O conteúdo em _io.BytesIO

                        # Pega os bytes do conteúdo do anexo
                        xlsx_data = xlsx_content.getvalue()
                        imbox.mark_seen(uid)


                        print(f"Arquivo xlsx '{xlsx_name}' encontrado e armazenado na variável 'xlsx_data'.")

                        # Cria o caminho completo para salvar o arquivo no diretório atual
                        xlsx_path = 'trashCan/spreedSheet.xlsx'

                        # Salva o xlsx no diretório em que o Python está rodando
                        with open(xlsx_path, 'wb') as f:
                            f.write(xlsx_data)
                        print(f"xlsx '{xlsx_name}' salvo no diretório: {xlsx_path}")

                        # Envia o email de resposta com o anexo
                        remetente = message.sent_from[0]['email']  # Obtém o email do remetente original
                        resposta_assunto = f"Re: {message.subject}"
                        resposta_corpo = "Aqui está o arquivo em anexo conforme solicitado."
                        
                        update_spreadsheet()
                        send_reply_with_attachment(remetente, resposta_assunto, resposta_corpo, PATH_DATABASE)
                        # imbox.mark_seen(uid)
            else:
                print("Nenhum anexo encontrado.")
