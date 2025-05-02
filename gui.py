import customtkinter as ctk
import email_sender as es
import processamento as p
import envioVersao as ev
import pandas as pd
import tkinter as tk
import traceback
import os 
from time import sleep
from datetime import datetime
from tkinter import ttk, messagebox, filedialog
from pathlib import Path
from PIL import Image


# Configuração do tema do CTk
ctk.set_appearance_mode("light")  # ou "dark"
ctk.set_default_color_theme("blue")

ASSETS_PATH_ICON = Path(__file__).parent / "assets" / "frame0" / "definicoes.ico"
# Criar a janela principal
window = ctk.CTk()
window.geometry("500x600")
window.title("Envio de Versão")
window._set_appearance_mode("light")
window.iconbitmap(ASSETS_PATH_ICON)

# Caixa de entrada para o destinatário
lbl_destinatario = ctk.CTkLabel(window, text="Enviar para:", font=("Roboto", 16))
lbl_destinatario.place(x=15, y=46)

email_val = tk.StringVar()
entry_destinatario = ctk.CTkEntry(window, textvariable=email_val, width=380)
entry_destinatario.place(x=111, y=46)

assunto = tk.StringVar()
assunto.set("Solicitações concluídas")

# Caixa de entrada para o assunto
lbl_assunto = ctk.CTkLabel(window, text="Assunto:", font=("Roboto", 16))
lbl_assunto.place(x=15, y=88)
entry_assunto = ctk.CTkEntry(window, textvariable=assunto,width=380)
entry_assunto.place(x=111, y=84)

# legenda de críticas
status_Titulo = "Ações necessárias:"

status_sem_Template = "Necessário informar o template do email (arquivo .oft)!"
status_com_Template = "Template Carregado com sucesso!"

status_sem_Planilha = "Necessário informar a planilha"
status_com_Planilha = "Planilha Carregada com sucesso!"
lbl_StatusTemplate1 = ctk.CTkLabel(window, text=f"{status_sem_Template}", font=("Arial", 14), text_color="red")
lbl_StatusTemplate2 = ctk.CTkLabel(window, text=f"{status_com_Template}", font=("Arial", 14), text_color="green")
lbl_StatusTitulo = ctk.CTkLabel(window, text=f"{status_Titulo}", font=("Arial", 14, 'bold'))
lbl_StatusPlanilha1 = ctk.CTkLabel(window, text=status_com_Planilha,font=("Arial", 14), text_color="green")
lbl_StatusPlanilha2 = ctk.CTkLabel(window, text=status_sem_Planilha,font=("Arial", 14), text_color="red")


# Variável para RadioButtons
opcao_selecionada = ctk.StringVar(value="EnviarSolicitacao")

#monitorar quando alternar o radio button
def alterou_opcao():
    if opcao_selecionada.get() == "EnviarVersao":
        desabilitarBtnEnvioVersao()
        habilitarBtnSave()
        lbl_StatusTitulo.place(x=15, y=46)
        check_enviaAtutomatico.place(x=350, y=215)
        lbl_StatusTemplate1.place(x=15, y=66)
        lbl_StatusPlanilha2.place(x=15, y=86)
        button_EnviarVersao.place(x=15, y=180)
        button_AbrirTemplate.place(x=135, y=180)
        button_abrirPlanilha.place(x=15, y=215)
        frame_versao.place(x=15, y=250)
        lbl_destinatario.place_forget()
        lbl_assunto.place_forget()
        entry_assunto.place_forget()
        entry_destinatario.place_forget()
        button_AbrirCSV.place_forget() #esconder botão CSV
        button_EnviarEmail.place_forget()
        frame.place_forget()
        assunto.set("Nova versão disponivel")
        
    elif opcao_selecionada.get() == "EnviarSolicitacao":
        desabilitarBtnSave()
        lbl_assunto.place(x=15, y=88)
        entry_assunto.place(x=111, y=84)
        entry_destinatario.place(x=111, y=46)
        button_EnviarEmail.place(x=15, y=180)
        lbl_destinatario.place(x=15, y=46)
        frame.place(x=15, y=220)
        button_AbrirCSV.place(x=135, y=180) #Apresentar botão Docx
        frame_versao.place_forget()
        lbl_StatusTitulo.place_forget()
        lbl_StatusTemplate1.place_forget()
        lbl_StatusTemplate2.place_forget()
        lbl_StatusPlanilha1.place_forget()
        lbl_StatusPlanilha2.place_forget()
        button_AbrirTemplate.place_forget()
        button_abrirPlanilha.place_forget()
        button_EnviarVersao.place_forget()
        check_enviaAtutomatico.place_forget()
        assunto.set("Solicitações concluídas")

# Criando o Radio Button
radio_enviarVersao = ctk.CTkRadioButton(window, text="Enviar versão", variable=opcao_selecionada, value="EnviarVersao", command=alterou_opcao)
radio_enviarVersao.place(x=350, y=118)

radio_enviarSolicitacao = ctk.CTkRadioButton(window, text="Enviar Solicitações", variable=opcao_selecionada, value="EnviarSolicitacao", command=alterou_opcao)
radio_enviarSolicitacao.place(x=350, y=140)

# Botões
# Caminho correto usando pathlib
ASSETS_PATH_JS = Path(__file__).parent / "assets" / "frame0" / "button_2.png"
ASSETS_PATH_CSV = Path(__file__).parent / "assets" / "frame0" / "excel.png"
ASSETS_PATH_DOCX = Path(__file__).parent / "assets" / "frame0" / "word.png"
ASSETS_PATH_OUT = Path(__file__).parent / "assets" / "frame0" / "outlook.png"
ASSETS_PATH_CANC = Path(__file__).parent / "assets" / "frame0" / "cancel.png"
icon_image_js = ctk.CTkImage(light_image=Image.open(ASSETS_PATH_JS), size=(20, 20))
icon_image_csv = ctk.CTkImage(light_image=Image.open(ASSETS_PATH_CSV), size=(20, 20))
icon_image_doc = ctk.CTkImage(light_image=Image.open(ASSETS_PATH_DOCX), size=(20, 20))
icon_image_out = ctk.CTkImage(light_image=Image.open(ASSETS_PATH_OUT), size=(20, 20))
icon_image_canc = ctk.CTkImage(light_image=Image.open(ASSETS_PATH_CANC), size=(20, 20))

dados = [] # lista para preencher os dados do CSV
contatos = [] # lista para preencher os dados do xls contendo os contatos

button_AbrirCSV = ctk.CTkButton(
    window, 
    image=icon_image_csv,
    compound="left",
    text="Abrir planilha", 
    anchor="w",
    width=15, 
    height=24,
    command=lambda: abrirCSV()
)
button_AbrirCSV.place(x=135, y=180)
def abrirCSV():
    global dados
    global df

    #remove csv anterior
    if os.path.exists("arquivo_convertido.csv"):
        os.remove("arquivo_convertido.csv")

    try:
        filename = filedialog.askopenfilename(
        title="Selecione a planilha",
        filetypes=[("Arquivo xlsx", "*.xls")], 
        )
        caminhoDoArquivo = filename
        if not filename:
            messagebox.showinfo("Atenção", "Nenhum arquivo informado")
            habilitarBtnCSV() 
            return # finalizar try se não for informado nenhum arquivo
        # chamar função de conversão para CSV e efetuar leitura da planilha incluindo informações resuimdas na tree view
        else:
            df = p.convertToCSV(caminhoDoArquivo=caminhoDoArquivo)
            qtsSolicitacao = df.groupby(["Cliente","Razão social", "Solicitante"]).size().reset_index(name="qtdOS") #agrupa valores por cod do cliente, razao e solicitante para definir a qtd de solicitações
            emails = df[['Cliente', 'Solicitante', 'Email']].drop_duplicates() #pega os valores de email sem valores duplicados

            qtsSolicitacao = qtsSolicitacao.merge(emails, on=['Cliente', "Solicitante"], how='left') 

            qtsSolicitacao['Email'] = qtsSolicitacao['Email'].fillna("")
            #ordernar dados no grid por razão em ordem ascendente
            qtsSolicitacao = qtsSolicitacao.sort_values(by=["Razão social"])
            # Remove os itens antigos da Treeview
            limparDados()
            # Inclui dados na tree view
            for index, row in qtsSolicitacao.iterrows():
                codcliente = row['Cliente']
                nomeCliente = row['Razão social']
                solicitante = row['Solicitante']
                qtdOS = row['qtdOS']
                email = row['Email']

                dados.append([int(codcliente),nomeCliente,solicitante,qtdOS, email])
            for item in dados:
                item_id = tree.insert("", tk.END, values=("☐", *item), tags=("normal",))
                checkboxes[item_id] = "☐"
                habilitarBtnSave()
                desabilitarRadioEnvioVersao()
        return df       
    except Exception as e:
        # Exibe um erro mais detalhado, incluindo o tipo de erro e a mensagem
        print("Erro ao processar o arquivo:", str(e))
        traceback.print_exc()  # Exibe o rastreamento da exceção
        messagebox.showinfo("Atenção", "não foi possível processar o arquivo")
        # Reabilita o botão após o erro
        habilitarBtnCSV()
# funções de habilitar e desabilitar campos na tela
def habilitarBtnCSV():
        button_AbrirCSV.configure(state="normal")
def desabilitarBtnCSV():
        button_AbrirCSV.configure(state="disabled")
def habilitarBtnSave():
        button_Salvar.configure(state="normal")
def desabilitarBtnSave():
        button_Salvar.configure(state="disabled")
def habilitarRadioEnvioVersao():
        radio_enviarVersao.configure(state="normal")
def desabilitarRadioEnvioVersao():
        radio_enviarVersao.configure(state="disabled")
def habilitarRadioEnvioSolicitacao():
        radio_enviarSolicitacao.configure(state="normal")
def desabilitarRadioSolicitacao():
        radio_enviarSolicitacao.configure(state="disabled")
def desabilitarBtnEnvioVersao():
        button_EnviarVersao.configure(state="disabled")
def habilitarBtnEnvioVersao():
        button_EnviarVersao.configure(state="normal")

def limparDados():
    try:
        dados.clear()
        tree.delete(*tree.get_children()) 
    except Exception as e:
         messagebox.showerror("Erro ao limpar tela", str(e))
#criar novo botão para abrir template do email
def limparContatos():
    try:
        treeVersao.delete(*treeVersao.get_children()) 
    except Exception as e:
         messagebox.showerror("Erro ao limpar tela", str(e))
#criar novo botão para abrir template do email
button_AbrirTemplate = ctk.CTkButton(
    window, 
    image=icon_image_doc,
    compound="left",
    text="Abrir Template", 
    anchor="w",
    width=15, 
    height=24,
    command=lambda: abrirTemplate()
)
def abrirTemplate():
        try:
            global temp #deixar variavel acessivel para enviar o email
            temp = es.upload()
            lbl_StatusTemplate1.place_forget() # esconder mensagem inicial
            if temp == '': # Se o arquivo não for informado manter o status inicial
                lbl_StatusTemplate1.place(x=15, y=66) 
            else:
                lbl_StatusTemplate2.place(x=15, y=66) # exibir confirmação de que o template foi lido
                habilitarBtnEnvioVersao()
            return temp
        except Exception as e:
             messagebox.showerror("Erro ao salvar template", str(e))
    #botão enviar email
button_EnviarEmail = ctk.CTkButton(
    window, 
    image=icon_image_out,
    compound="left",
    text="Enviar e-mail", 
    anchor="w",
    width=15, 
    height=24,
    command=lambda: enviar_email()
)
button_EnviarEmail.place(x=15, y=180)
def enviar_email():
    try:
        global df
        selected_item = tree.selection()
        df = p.lerCSV(caminhoDoArquivo="arquivo_convertido.csv")
        #sempre que que chamar a função, deixar marcado como Enviado
        if not selected_item:
            messagebox.showwarning("Atenção", "Nenhum item selecionado!")
            return 
        for item in tree.get_children():
            tree.item(selected_item[0], tags=("selecionado",))
            linha = tree.item(selected_item[0], "values")
            codCli = int(linha[1])
            solicitante = linha[3]
            colunaEmailSolicitante = linha[5]
            assuntoEmail = entry_assunto.get()
            #caso a coluna email esteja em branco, preencher com o email informado no campo de texto
            if colunaEmailSolicitante == "":
                emailSolicitante = entry_destinatario.get()
            else:
                emailSolicitante = colunaEmailSolicitante
        es.enviar_email(codCliente=codCli, nomeSolicitante=solicitante, enviar_para=emailSolicitante, assunto=assuntoEmail,df=df)
        window.iconify()
    except Exception as e:
         messagebox.showerror("Erro ao abrir Outlook", str(e))

button_abrirPlanilha = ctk.CTkButton(window, 
    text="Abrir planilha de controle de envio",
    image=icon_image_csv,
    compound="left",
    anchor="w",
    width=35, 
    height=25,
    command=lambda: incluir_planilha()                          
)
def incluir_planilha():
     global contatos, plan, planilha_controle

     plan = filedialog.askopenfilename(
         title="Selecione a planilha",
         filetypes=[("Arquivo xlsx", "*.xlsx")],
     ) 
     planilha_controle = plan 
     if not plan:
        messagebox.showerror("Atenção", "Nenhum arquivo informado")
        return
     else:     
        try:
            lbl_StatusPlanilha2.place_forget()
            lbl_StatusPlanilha1.place(x=15, y=86)
            df_controle = p.ler_planilhaDeControle(plan=planilha_controle)
            for i, row in df_controle.iterrows():
                empresa = row['EMPRESA']
                email = row['EMAIL']
                data_envio = row['DATA ENVIO']
                data_atualizacao = row['DATA ATUALIZAÇÃO']

                contatos.append([empresa, email, data_envio, data_atualizacao])
                            
            for item in contatos:
                    if pd.isna(item[2]):
                        item_id = treeVersao.insert("", tk.END, values=("☐", *item), tags=("normal",))
                        checkboxes[item_id] = "☐"
                    elif not pd.isna(item[2]):
                        item_id = treeVersao.insert("", tk.END, values=("☐", *item), tags=("selecionado",))
                        checkboxes[item_id] = "☐"
            desabilitarRadioSolicitacao()
        except Exception as e:
            messagebox.showerror("Erro ao incluir planilha!", str(e))

button_EnviarVersao = ctk.CTkButton(
    window, 
    image=icon_image_out,
    compound="left",
    text="Enviar e-mail", 
    anchor="w",
    width=15, 
    height=24,
    command=lambda: enviarVersao()
)
def enviarVersao():
    lista_empresa = []
    try:
        if check_enviaAtutomatico.get() == 1:
            try:
                popup = ctk.CTkToplevel(window)
                popup.title("Envio automático")
                popup.geometry("300x100")
                popup.resizable(False,False)
                popup.transient(window)
                popup.grab_set()
                label_status = ctk.CTkLabel(popup, text="Iniciando envios...")
                label_status.pack(pady=20)
                                
                for email in lista_envio:
                    sleep(1)
                    label_status.configure(text=f"Enviando para {email}")
                    es.enviarVersaoAutomatico(temp=temp, lista_envio=email)
                    popup.update()
                    sleep(2)
                popup.update()
                popup.destroy()
                messagebox.showinfo(title="Concluído", message="Envio finalizado com sucesso.")
            except Exception as e:
                 messagebox.showerror("Erro ao enviar emails", str(e))
        else:
            for email in lista_envio:
                es.enviarVersao(temp=temp, lista_envio=email)

        # altera os itens selecionados enviado conforme envia os emails (abre o outlook)
        for item_id, status in check.items():
            if status == "☑":
                dia = datetime.now()
                hoje = dia.date().strftime('%d/%m/%Y')

                valores = treeVersao.item(item_id, 'values')
            
                check[item_id] = "☐"
                valores = treeVersao.item(item_id, 'values')
                
                novos_valores = ("☐",) + valores[1:]
                treeVersao.item(item_id, values=novos_valores, tags=('selecionado',))

                valores_atuais = list(treeVersao.item(item_id, 'values'))
                valores_atuais[3] = hoje  
                treeVersao.item(item_id, values=valores_atuais)

                empresa = valores[1]
                lista_empresa.append(empresa)
        p.ajustarDataEnvio(plan=plan, dia=hoje, empresas=lista_empresa)
                
    except Exception as e:
        messagebox.showerror("Erro ao ler template", str(e))
    limpar_lista()

def limpar_lista():
    lista_envio.clear()

#botão desmarcar envio
button_DesmarcarEnvio = ctk.CTkButton(
    window, 
    image=icon_image_canc,
    compound="left",
    text="Desmarcar envio", 
    text_color="#870B0B",
    anchor="w",
    width=15, 
    height=24,
    command=lambda: desmarcar_envio()
)
button_DesmarcarEnvio.place(x=350, y=180)
check_enviaAtutomatico = ctk.CTkCheckBox(window, text="Envio automático")

limpar_empresa = []
def desmarcar_envio():
    if opcao_selecionada.get() == "EnviarSolicitacao":
        selected_item = tree.selection()
        if not selected_item:
            messagebox.showwarning("Atenção", "Nenhum item selecionado!")
            return
        for item in tree.get_children():
            tree.item(selected_item[0], tags=("normal",))

    elif opcao_selecionada.get() == "EnviarVersao":
         for item_id, status in check.items():
            if status == "☑":
                valores = treeVersao.item(item_id, 'values')
            
                check[item_id] = "☐"
                valores = treeVersao.item(item_id, 'values')
                
                novos_valores = ("☐",) + valores[1:]
                treeVersao.item(item_id, values=novos_valores, tags=('normal',))
                limpar_lista()
                
                empresa = valores[1]
                limpar_empresa.append(empresa)
                if empresa in limpar_empresa:
                    valores = list(treeVersao.item(item_id, 'values'))
                    valores[3] = ''  # Atualiza o campo de data (índice 3)
                    treeVersao.item(item_id, values=valores)

                p.limpar_planilha(plan=plan, empresas=limpar_empresa)

                for item_id in treeVersao.get_children():
                    valores = list(treeVersao.item(item_id, 'values'))
                    empresa = valores[1]  # Pegando o nome da empresa (índice 1)

button_Salvar = ctk.CTkButton(
    window,
    image=icon_image_js, 
    fg_color="#FFFFFF",
    text="", 
    command=lambda: salvar(),
    width=24, 
    height=24
)
button_Salvar.place(x=455, y=5)
def salvar():
    nova_planilhaControle = plan
    df = pd.read_excel(nova_planilhaControle)
    pasta_destino = filedialog.askdirectory(title="Salve a planilha")

    if pasta_destino:
        # Monta o caminho completo para salvar
        caminho_completo = os.path.join(pasta_destino, "controle de envio.xlsx")
        df.to_excel(caminho_completo, index=False)
    

# botao limpar tela 
button_LimparTela= ctk.CTkButton(
    window,
    image=icon_image_canc,
    text_color="#000000",
    fg_color="#FFFFFF",
    text="Limpar tela", 
    command=lambda: limparTela(), 
    width=24, 
    height=24
)
button_LimparTela.place(x=325, y=5)
def limparTela():
    if opcao_selecionada.get() == "EnviarVersao":
        limparContatos()
        lbl_StatusPlanilha1.place_forget()
        lbl_StatusPlanilha2.place(x=15, y=86)
        contatos.clear()
        habilitarRadioEnvioSolicitacao()
        #remove xls anterior
    elif opcao_selecionada.get() == "EnviarSolicitacao":
        limparDados()
        habilitarBtnCSV()
        habilitarRadioEnvioVersao()

# Criar um frame vazio para o grid (inicialmente invisível)
frame = ctk.CTkFrame(window, width=480, height=360)
frame.place(x=15, y=220)

#### LAYOUT DO GRID DE ENVIO DE SOLICITAÇÕES
tree = ttk.Treeview(frame, columns=("Check","ID", "Razão", "Solicitante", "QTD_OS", "Email"), show="headings", style="Treeview")
   # Definir os cabeçalhos
tree.heading("Check", text="")
tree.heading("ID", text="ID")
tree.heading("Razão", text="Razão")
tree.heading("Solicitante", text="Solicitante")  
tree.heading("QTD_OS", text="QTD_OS")
tree.heading("Email", text="Email")

#   ajstar largura das colunas
tree.column("Check", width=25, stretch=False)
tree.column("ID", width=30, stretch=False)
tree.column("Razão", width=91,stretch=False)
tree.column("Solicitante", width=91,stretch=False)
tree.column("QTD_OS", width=40,stretch=False)
tree.column("Email", width=200,stretch=False)


#definir preenchimento do chbox
# Dicionário para armazenar o estado dos checkboxes
checkboxes = {}
def toggle_checkbox(event):
    """Permite selecionar apenas um checkbox por vez."""
    global checkboxes, email_val

    # Obtém o item selecionado
    selected_item = tree.identify_row(event.y)
    
    if selected_item:
        # Reseta todos os checkboxes
        for item in checkboxes:
            checkboxes[item] = "☐"
        
        # Marca apenas o item clicado
        checkboxes[selected_item] = "☑"
        
        # Atualiza os valores no Treeview
        for item in checkboxes:
            tree.item(item, values=(checkboxes[item], *tree.item(item, "values")[1:]))    
        
    # preencher campos de texto
    if checkboxes[selected_item] == "☑":
        tree_select = tree.selection()
        val = tree.item(tree_select[0], "values")
        email = val[5] 
        email_val.set(email)  

for item in dados:
    item_id = tree.insert("", tk.END, values=("☐", *item), tags=("normal",))
    checkboxes[item_id] = "☐"

# Bind para capturar cliques e atualizar checkboxes
tree.bind("<ButtonRelease-1>", toggle_checkbox)

# Estilizar o Treeview com CustomTkinter 
style = ttk.Style()
style.configure("Treeview", background="lightgray", foreground="black", font=("Roboto", 10))
style.configure("Treeview.Heading", background="gray", font=("Roboto", 11, "bold"))

# Adicionando tags para cores
tree.tag_configure("normal", background="white")
tree.tag_configure("selecionado", background="green")

# Posicionar o Treeview dentro do frame
tree.place(relx=0.5, rely=0.5, anchor="center", relwidth=0.95, relheight=0.9)


###### LAYOUT DO GRID DE ENVIO DE VERSÃO
# Criar um frame vazio para o grid (inicialmente invisível) para o envio de versão
frame_versao = ctk.CTkFrame(window, width=480, height=320)

treeVersao = ttk.Treeview(frame_versao, columns=("Check","Empresa", "Email","Data Envio", "Data Atualização"), show="headings", style="Treeview")
   # Definir os cabeçalhos
treeVersao.heading("Check", text="")
treeVersao.heading("Empresa", text="Empresa")
treeVersao.heading("Email", text="Email")  
treeVersao.heading("Data Envio", text="Data Envio")
treeVersao.heading("Data Atualização", text="Data Atualização")


#   ajstar largura das colunas
treeVersao.column("Check", width=25, stretch=False)
treeVersao.column("Empresa", width=85, stretch=False)
treeVersao.column("Email", width=85,stretch=False)
treeVersao.column("Data Envio", width=120, stretch=False)
treeVersao.column("Data Atualização", width=140,stretch=False)

#definir preenchimento do chbox
# Dicionário para armazenar o estado dos checkboxes
check = {}
lista_envio = set()
empresas = set()
def checkbox(event):
    global selected_itemVersao, current_values
    selected_itemVersao = treeVersao.focus()
    if selected_itemVersao:
        linha_selecionada = treeVersao.selection()
        current_values = treeVersao.item(selected_itemVersao, "values")
        check_value = check.get(selected_itemVersao, "☐")
        qtd_lista_envio = len(lista_envio)
            
 
        if check_value == "☐":  # Se ainda não está selecionado
                # Marca como selecionado
                check[selected_itemVersao] = "☑"
                email = current_values[2]  
                empresa = current_values[1]
                lista_envio.add(email)
                empresas.add(empresa)

        else:  # Se já está selecionado
            check[selected_itemVersao] = "☐"
            email = current_values[2]
            empresa = current_values[1]
            if email in lista_envio:
                lista_envio.remove(email)
                empresas.remove(empresa)


        # Atualiza o valor no Treeview
        new_values = (check[selected_itemVersao],) + current_values[1:]
        treeVersao.item(selected_itemVersao, values=new_values)
       
             
# Bind para capturar cliques e atualizar checkboxes
treeVersao.bind("<ButtonRelease-1>",checkbox)

# Estilizar o Treeview com CustomTkinter 
style = ttk.Style()
style.configure("Treeview", background="lightgray", foreground="black", font=("Roboto", 10))
style.configure("Treeview.Heading", background="gray", font=("Roboto", 11, "bold"))

# Adicionando tags para cores
treeVersao.tag_configure("normal", background="white")
treeVersao.tag_configure("selecionado", background="green")

# Posicionar o Treeview dentro do frame
treeVersao.place(relx=0.5, rely=0.5, anchor="center", relwidth=0.95, relheight=0.9)


window.resizable(False, False)
window.mainloop()
