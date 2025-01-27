from openpyxl import load_workbook, Workbook
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from utils import data_hora_atual
import os

nome_arquivo = f"Maiores_Arrecadadores_{data_hora_atual}.xlsx"

def capturar_primeiras_seis_colunas(func_navegador):
    try:
        elemento = func_navegador.find_element(By.CSS_SELECTOR, '#ctl00_ContentPlaceHolder1_dvResultado > table.tabelaFormulario')
        texto = elemento.text
    except NoSuchElementException  as e:
        print(f"Erro ao capturar elemento: {e}")
        return {}

    texto = elemento.text

    dados = {}
    for linha in texto.split("\n"):  
        if " : " in linha:  
            chave, valor = linha.split(" : ", 1)  
            dados[chave.strip()] = valor.strip()  

    print(dados)
    return dados  

def salvar_dados_planilha(dados_tabela, nome_arquivo):
    caminho_arquivo = os.path.join(os.path.expanduser("~"), "Documents/Planilha ANM", nome_arquivo)

    if os.path.exists(caminho_arquivo):
        workbook = load_workbook(caminho_arquivo)
    else:
        workbook = Workbook()

    folha = workbook.active
    folha.title = "Arrecadadores"

    if folha.max_row == 1 and folha.cell(row=1, column=1).value != "Arrecadador (Empresa)":
        folha.append(["Arrecadador (Empresa)", "Qtde Títulos", "Operação", "Recolhimento CFEM", "% Recolhimento CFEM"])

    for dado in dados_tabela:
        folha.append([  
            dado["Arrecadador (Empresa)"],
            dado["Qtde Títulos"],
            dado["Operação"],
            dado["Recolhimento CFEM"],
            dado["% Recolhimento CFEM"]
        ])

    workbook.save(caminho_arquivo)
    print(f"Dados salvos na planilha '{caminho_arquivo}' com sucesso!")

def capturar_todos_os_dados(func_navegador, municipio, regiao):
    # Captura os dados da primeira tabela
    elemento = func_navegador.find_element(By.CSS_SELECTOR, '#ctl00_ContentPlaceHolder1_dvResultado > table.tabelaFormulario')
    texto = elemento.text

    dados_gerais = {}
    for linha in texto.split("\n"):
        if " : " in linha:
            chave, valor = linha.split(" : ", 1)
            dados_gerais[chave.strip()] = valor.strip()

    dados_gerais["Município"] = municipio
    dados_gerais["Região"] = regiao
    print("Dados gerais capturados:", dados_gerais)

    # Captura os dados da segunda tabela
    tabela = func_navegador.find_element(By.CSS_SELECTOR, '#ctl00_ContentPlaceHolder1_dvResultado > table.tabelaRelatorio')
    linhas = tabela.find_elements(By.CSS_SELECTOR, "tbody > tr")

    dados_completos = []
    
    if not linhas:  
        print("Tabela vazia. Preenchendo com dados gerais e valores zerados.")
        dados_completos.append({
            "Arrecadador (Empresa)": "",
            "Qtde Títulos": 0,
            "Operação": "",
            "Recolhimento CFEM": 0,
            "% Recolhimento CFEM": 0,
            **dados_gerais  
        })
    else:  
        for linha in linhas:
            colunas = linha.find_elements(By.TAG_NAME, "td")
            if len(colunas) >= 6:
                linha_dados = {
                    "Arrecadador (Empresa)": colunas[1].text,
                    "Qtde Títulos": colunas[2].text,
                    "Operação": colunas[3].text,
                    "Recolhimento CFEM": colunas[4].text,
                    "% Recolhimento CFEM": colunas[5].text,
                    **dados_gerais
                }
                dados_completos.append(linha_dados)
            else:
                print(f"Linha ignorada por não ter colunas suficientes: {linha.text}")

    for dado in dados_completos:
        print(dado)

    return dados_completos

def salvar_dados_completos_planilha(dados_completos, nome_arquivo):
    caminho_arquivo = os.path.join(os.path.expanduser("~"), "Documents/Planilha ANM", nome_arquivo)

    if os.path.exists(caminho_arquivo):
        workbook = load_workbook(caminho_arquivo)
    else:
        workbook = Workbook()

    folha = workbook.active
    folha.title = "Arrecadadores"

    if folha.max_row == 1:
        folha.append([
            "Ano", "Arrecadação por", "Ordenação por", "Substância Agrupadora", "Substância", "Região", "Estado", "Município", "Arrecadador (Empresa)", "Qtde Títulos", "Operação",
            "Recolhimento CFEM", "% Recolhimento CFEM"
        ])

    for dado in dados_completos:
        folha.append([
            dado.get("Ano", ""),
            dado.get("Arrecadação por", ""),
            dado.get("Ordenação por", ""),
            dado.get("Substância Agrupadora", ""),
            dado.get("Substância", ""),
            dado.get("Região", ""),
            dado.get("Estado", ""),
            dado.get("Município", ""),
            dado["Arrecadador (Empresa)"],
            dado["Qtde Títulos"],
            dado["Operação"],
            dado["Recolhimento CFEM"],
            dado["% Recolhimento CFEM"]
            
        ])

    workbook.save(caminho_arquivo)
    print(f"Dados salvos na planilha '{caminho_arquivo}' com sucesso!")