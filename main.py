import time
import os
import pandas as pd
import smtplib
import email.mime.multipart
import email.mime.text
import email.mime.application
import pyodbc
import datetime

time.sleep(3)

# Passo 1: Chamar o executável de backup do dia
os.startfile(r"C:\Program Files\Microsoft SQL Server\MSSQL15.SQLEXPRESS\MSSQL\Backup\backup.bat", 'open')

time.sleep(20)

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
        server = "LAPTOP-A80SI4IH\SQLEXPRESS"
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

    FROM = "laurelios754@gmail.com"
    PASSWORD = "@utomatico"

    TO = "leovilelati@gmail.com"

    msg = email.mime.multipart.MIMEMultipart()
    msg['Subject'] = "Resumo de vandas / Backup do sistema"
    msg['From'] = FROM
    msg['To'] = TO

    linhas = ""

    for index, linha in data2.iterrows():
        linhas += "<tr>"
        linhas += f"""<td style="width: auto; border: 1px solid #cecece;"> {datetime.date.strftime(linha['INFO_INSERT_DATA_UTC'], "%d/%m/%Y")} </td>"""
        linhas += f"""<td style="width: auto; border: 1px solid #cecece;"> {linha.DESCRICAO} </td>"""
        linhas += f"""<td style="width: auto; border: 1px solid #cecece;"> {linha.PRECO_VENDA} </td>"""
        linhas += f"""<td style="width: auto; border: 1px solid #cecece;"> {linha.QUANTIDADE} </td>"""
        linhas += f"""<td style="width: auto; border: 1px solid #cecece;"> {linha.VALOR_DESCONTO} </td>"""
        linhas += "</tr>"

    body = email.mime.text.MIMEText(f"""
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
                                    """, 'html')
    msg.attach(body)

    filename = f"{caminho}\{ultimo_arquivo[1]}"
    fp = open(filename, 'rb')
    att = email.mime.application.MIMEApplication(fp.read(), _subtype="bat")
    fp.close()
    att.add_header('Content-Disposition', 'attachment', filename=filename)
    msg.attach(att)

    s = smtplib.SMTP('smtp.gmail.com', 587)
    s.starttls()
    s.login(FROM, PASSWORD)
    s.sendmail(FROM, [TO], msg.as_string())
    s.quit()

    print("Email enviado!")