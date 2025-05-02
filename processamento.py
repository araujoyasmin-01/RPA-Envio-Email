import re
from io import StringIO
from bs4 import BeautifulSoup
import pandas as pd
import csv
import win32com.client
from tkinter import messagebox, filedialog, Tk
import os


## FUNÇÕES PARA A PARTE DE ENVIO DE SOLICITAÇÕES
def convertToCSV(caminhoDoArquivo):
    with open(caminhoDoArquivo, 'r') as file:
        html_str = file.read()

    soup = BeautifulSoup(html_str, "html.parser")
    table = soup.find("table")

    padrao = re.compile("<resumo>(.*?)<resumo>", re.DOTALL)
    padrao2 = re.compile("RESUMO(.*?)RESUMO", re.DOTALL)

    if not table:
        print("Nenhuma tabela encontrada no HTML.")
        exit()

    # Expressão regex para remover a tag <resumo> (sem fechamento)
    resumo_pattern = re.compile(r"<resumo>.*?", re.DOTALL)

    # Processar os dados da tabela
    data = []
    rows = table.find_all("tr", recursive=False)  # Garante que pegamos apenas os TRs diretos da tabela

    # Identificar a posição da coluna "Solicitação"
    header_row = table.find("tr")  # Assumindo que o primeiro <tr> contém os cabeçalhos
    headers = [th.get_text(strip=True) for th in header_row.find_all("td")]

    try:
        solicitacao_index = headers.index("Solicitação")  # Descobre a posição da coluna "Solicitação"
    except ValueError:
        print("A coluna 'Solicitação' não foi encontrada na tabela.")
        exit()

    # Regex para capturar apenas o conteúdo dentro da tag <resumo>
        resumo_pattern = re.compile(r"<resumo>(.*?)", re.DOTALL)

    # Processar os dados da tabela
    data = [headers]  # Adicionamos os cabeçalhos primeiro
    rows = table.find_all("tr", recursive=False)[1:]  # Pegamos os TRs diretos, exceto o cabeçalho

    for row in rows:
        cells = row.find_all("td", recursive=False)  # Evita pegar TDs aninhados errados
        if not cells:
            continue

        processed_row = []
        for i, cell in enumerate(cells):
            cell_text = str(cell)

            # Aplicar regex apenas na coluna "Solicitação"
            if i == solicitacao_index:
                match = padrao.search(cell_text)
                match2 = padrao2.search(cell_text)
                if match:
                    cell_text = match.group(1).strip()  # Mantém só o conteúdo de <resumo>
                elif match2:
                    cell_text = match2.group(1).strip()  # Mantém só o conteúdo de RESUMO
                else: #se não tiver a tag ou palavra resumo limpar os traços da coluna
                     cell_text = cell.get_text(strip=True).replace("-", "").replace("=", "")  # Remove "-" e "="
            else:
                cell_text = cell.get_text(strip=True)  # Para outras colunas, pega o texto normal

            processed_row.append(cell_text)

        data.append(processed_row)     
        with open("arquivo_convertido.csv", "w", newline="", encoding="utf-8") as csvfile:
                writer = csv.writer(csvfile)
                writer.writerows(data)
        
    return lerCSV(caminhoDoArquivo="arquivo_convertido.csv") #retorna a chamada lendo o CSV gerado

def lerCSV(caminhoDoArquivo):
    #informações do arquivo
    arquivo = caminhoDoArquivo
    csv = pd.read_csv(f"{arquivo}", sep=",", encoding="utf-8").sort_values(by="Razão social", ascending=False)

    df = pd.DataFrame(csv)
    return df

## FUNÇÕES PARA A PARTE DE ENVIO DE VERSÃO
def ler_planilhaDeControle(plan):
    conteudo = []
    caminho_planilha = plan
    df_controle = pd.read_excel(caminho_planilha)
    return df_controle

def ajustarDataEnvio(plan, dia, empresas):
    planilha_controle = plan
    df = pd.read_excel(planilha_controle)
    df.loc[df['EMPRESA'].isin(empresas), 'DATA ENVIO'] = dia
    df.to_excel(planilha_controle, index=False)

def limpar_planilha(plan, empresas):
    planilha_controle = plan
    df_planilha = pd.read_excel(planilha_controle)
    df_planilha.loc[df_planilha['EMPRESA'].isin(empresas), 'DATA ENVIO'] = ''
    df_planilha.to_excel(planilha_controle, index=False)

