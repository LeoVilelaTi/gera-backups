import time
import os
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import pyodbc
import datetime
import subprocess


# Passo 1: Chamar o executável de backup do dia
caminho_bat = r"C:\Program Files\Microsoft SQL Server\MSSQL15.SQLEXPRESS\MSSQL\Backup\backup.bat"

if os.path.exists(caminho_bat):
    # Abre o executável
    subprocess.run([caminho_bat])
else:
    print("O executável não foi encontrado.")


# Passo 2: Procura no diretorio necessário o arquivo mais atual do backup
caminho = r"C:\Program Files\Microsoft SQL Server\MSSQL15.SQLEXPRESS\MSSQL\Backup"

if os.path.isdir(caminho):

    lista_arquivos = os.listdir(caminho)

    lista_datas = []

    # Percorre cada item da lista
    for arquivo in lista_arquivos:
        # Descobrir a data do arquivo
        data = os.path.getmtime(f"{caminho}\{arquivo}")
        lista_datas.append((data, arquivo))

    lista_datas.sort(reverse=True)
    ultimo_arquivo = lista_datas[0]


    # Passo 3: Obtem os dados do banco de dados

    def retornar_conexao_sql():
        server = ".\SQLEXPRESS"
        database = "DbUniversoMangas"
        uid = "sa"
        password = "112358"
        string_conexao = 'Driver={SQL Server Native Client 11.0}; Server=' + server + ';Database=' + database + ';UID=' + uid + ';PWD=' + password + ';'
        conexao = pyodbc.connect(string_conexao)
        return conexao


    query = ("""
                select 
                    * 
                from 
                    TB_UNI_VENDA
                where
                    ATIVO = 1 and
                    CONVERT(CHAR(10), INFO_INSERT_DATA_UTC, 103) = CONVERT(CHAR(10),GETDATE(),103)
             """)

    data1 = pd.read_sql(query, retornar_conexao_sql())

    quantItens = data1.QUANTIDADE_MERCADORIAS.sum()
    valorTotal = data1.VALOR_TOTAL.sum()

    query = ("""
                  select 
                      v.INFO_INSERT_DATA_UTC,
                      m.DESCRICAO,
                      m.PRECO_COMPRA,
                      m.PRECO_VENDA,
                      vm.QUANTIDADE,
                      v.TAXA_PAGAMENTO,
                      v.VALOR_DESCONTO
                  from 
                       TB_UNI_VENDA as v
                   inner join 
                       TB_UNI_VENDA_MERCADORIA as vm 
                           on v.ID = vm.ID_VENDA
                   inner join 
                       TB_UNI_MERCADORIA as m 
                           on vm.ID_MERCADORIA = m.ID
                  where
                      v.ATIVO = 1 and
                      CONVERT(CHAR(10), v.INFO_INSERT_DATA_UTC, 103) = CONVERT(CHAR(10), GETDATE(), 103)       
             """)

    data2 = pd.read_sql(query, retornar_conexao_sql())

    # Passo 4: Monta o e-mail e dispara um resumo, com o anexo do backup

    linhas = ""

    for index, linha in data2.iterrows():
        linhas += "<tr>"
        linhas += f"""<td style="width: auto; border: 1px solid #cecece;"> {datetime.date.strftime(linha['INFO_INSERT_DATA_UTC'], "%d/%m/%Y")} </td>"""
        linhas += f"""<td style="width: auto; border: 1px solid #cecece;"> {linha.DESCRICAO} </td>"""
        linhas += f"""<td style="width: auto; border: 1px solid #cecece;"> {linha.PRECO_VENDA} </td>"""
        linhas += f"""<td style="width: auto; border: 1px solid #cecece;"> {linha.QUANTIDADE} </td>"""
        linhas += f"""<td style="width: auto; border: 1px solid #cecece;"> {linha.VALOR_DESCONTO} </td>"""
        linhas += "</tr>"

    body = f"""
                <h4>Boa tarde</h4>
                <div>Resumo das vendas do dia:</div>
                <ul>
                <li>Total de itens vendidos: {quantItens} itens</li>
                <li>Valor total vendido: R$ {valorTotal:,.2f} </li>
                </ul>
                <table style="height: 104px; border: 1px solid #cecece;" width="627">
                <tbody>
                <tr>
                <td style="width: auto; border: 1px solid #cecece;"><strong>Data Hora</strong></td>
                <td style="width: auto; border: 1px solid #cecece;"><strong>Descrição</strong></td>
                <td style="width: auto; border: 1px solid #cecece;"><strong>Preço Venda</strong></td>
                <td style="width: auto; border: 1px solid #cecece;"><strong>Quantidade</strong></td>
                <td style="width: auto; border: 1px solid #cecece;"><strong>Valor Desconto</strong></td>
                </tr>
                {linhas}
                </tbody>
                </table>
                <div>&nbsp;</div>
                <div>Obs.: Em anexo segue o backup do sistema.</div>
                <div>&nbsp;</div>
                <div>Att.</div>
                <div>Equipe de vendas</div>
            """
    


    # Configurações do servidor SMTP
    smtp_server = 'smtp.gmail.com'
    smtp_port = 587  # Porta típica para STARTTLS
    smtp_username = 'laurelios754@gmail.com'
    ##smtp_password = '@utomatico'
    smtp_password = 'cxpy awnv pjrk mhmo'
    destinatario = 'leovilelati@gmail.com'

    # Construa a mensagem de e-mail
    assunto = 'Resumo de vandas / Backup do sistema'
    corpo_mensagem_html = body

    mensagem = MIMEMultipart()
    mensagem['From'] = smtp_username
    mensagem['To'] = destinatario
    mensagem['Subject'] = assunto

    # Adicione corpo do e-mail
    mensagem.attach(MIMEText(corpo_mensagem_html, 'html'))

    # Adicione anexo ao e-mail
    caminho_anexo = f"{caminho}\{ultimo_arquivo[1]}"

    anexo = MIMEBase('application', 'octet-stream')
    anexo.set_payload(open(caminho_anexo, 'rb').read())
    encoders.encode_base64(anexo)
    anexo.add_header('Content-Disposition', f'attachment; filename={caminho_anexo}')

    mensagem.attach(anexo)


    # Conecte-se ao servidor SMTP e envie o e-mail
    try:
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            # Inicie a comunicação com o servidor
            server.starttls()

            # Faça login no servidor
            server.login(smtp_username, smtp_password)

            # Envie o e-mail
            server.sendmail(smtp_username, destinatario, mensagem.as_string())

        print('E-mail com anexo enviado com sucesso!')

    except Exception as e:
        print(f'Erro ao enviar e-mail com anexo: {e}')