import win32com.client
import processamento as p
import envioVersao as ev
from datetime import datetime
import re
import win32com.client
from tkinter import messagebox, filedialog, Tk
import os

## FUNÇÕES PARA A PARTE DE ENVIO DE SOLICITAÇÃO
def enviar_email(codCliente, nomeSolicitante, enviar_para,assunto,df):
    #confiugrações de envio
    to = enviar_para
    conteudoBody = [] # Conteudo do body que será preenchido pela função lerCSV()
    subject = assunto

    for index, row in df[(df['Cliente'] == codCliente) & (df['Solicitante'] == nomeSolicitante.upper())].iterrows():
        #texto = f"OS: {row['Numero']}\nSolicitação: {row['Solicitação']}\n"
        r = row['Solicitação'] #define coluna de solicitação para extrair conteudo das tags resumo caso ela tenha

        texto = f"<b>OS:</b> {row['Numero']}<br><b>Solicitação:</b> {row['Solicitação'].replace("-", "").replace("=", "")}<br><br>"
        conteudoBody.append(texto)

    Body ="".join(conteudoBody) # retorna o conteudo do email ja preenchdido
    
    #iniciar outlook
    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0) #cria novo email

    mail.To = to
    mail.Subject = subject
    
    # Hora atual
    agora = datetime.now()
    hora = agora.hour

    # Determinar período do dia para deixar mensagem mais amigavel
    if 5 <= hora < 12:
        periodo = "Bom dia!"
    elif 12 <= hora < 18:
        periodo = "Boa tarde!"
    else:
        periodo = "Boa noite!"

    mail.HTMLBody = f"<b>{periodo}</b><br><b>Segue abaixo solicitações que serão enviadas na próxima versão:</b> <br><br> {Body}" 

    #abrir email
    mail.Display()

def enviarVersao(template):
    to = "yasminaraujo591#gmail.com"
    conteudoBody = [] # Conteudo do body que será preenchido pela função lerCSV()
    subject = "Nova versão Disponivel"

    #iniciar outlook
    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0) #cria novo email

    mail.To = to
    mail.Subject = subject
    mail.HTMLBody = ev.lerdoc(template=template)

    #abrir email
    mail.Display()    


## FUNÇÕES PARA A PARTE DE ENVIO DE VERSÃO
#PEGA O CAMINHO DO TEMPLATE DE EMAIL -- ARQUIVO OTF
def upload():
    root = Tk()
    root.withdraw()
    template = filedialog.askopenfilename(
        title="Selecione o template",
        filetypes=[("Arquivo msg", "*.oft")],
    )
    return template

# ABRE OUTLOOK E ENVIA EMAIL DE ACORDO COM O TAMPLATE INFORMADO BA FUNÇÃO upload()
def enviarVersao(temp, lista_envio):
    to = lista_envio
    oft_path = temp

    if not oft_path or not os.path.exists(oft_path):
        messagebox.showerror("Erro", "Nenhum arquivo .oft selecionado.")
        return

    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItemFromTemplate(oft_path)

        mail.To = to
        mail.Subject = "Nova versão do Ceros Disponível"  # opcional, pode manter do modelo
        mail.Display()  # exibe antes de enviar

    except Exception as e:
        messagebox.showerror("Erro ao carregar modelo", str(e))
        
def enviarVersaoAutomatico(temp, lista_envio):
    to = lista_envio
    oft_path = temp

    if not oft_path or not os.path.exists(oft_path):
        messagebox.showerror("Erro", "Nenhum arquivo .oft selecionado.")
        return

    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItemFromTemplate(oft_path)

        mail.To = to
        mail.Subject = "Nova versão do Ceros Disponível"  # opcional, pode manter do modelo
        mail.Send()  # envia direto sem abrir o outlook

    except Exception as e:
        messagebox.showerror("Erro ao carregar modelo", str(e))