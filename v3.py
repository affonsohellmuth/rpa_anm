from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import pandas as pd

# Função para obter as opções de um dropdown
def obter_opcoes(id):
    select = Select(driver.find_element(By.ID, id))
    return [opcao.text for opcao in select.options[1:]]  # Ignora a primeira opção (geralmente a de "Selecione")

# Função para selecionar um valor no dropdown
def selecionar_combo(id, valor):
    select = Select(driver.find_element(By.ID, id))
    select.select_by_visible_text(valor)

# Inicializando o WebDriver
#driver = webdriver.Chrome()
#driver.maximize_window()
options = webdriver.ChromeOptions()
#options.add_argument("--headless")
options.add_argument("--maximize-window")
options.add_argument("--disable-extensions") 
options.add_argument("--disable-popup-blocking") 
driver = webdriver.Chrome(options=options)


# Acessando a página
url = "https://sistemas.anm.gov.br/arrecadacao/extra/Relatorios/cfem/arrecadadores.aspx"
driver.get(url)

wait = WebDriverWait(driver, 10)

# Aguardar o carregamento da página
time.sleep(2)

# Dados coletados
dados_coletados = []

# Selecionando o ComboBox de "Ano" - selecionar "2025"
ano_combo = Select(driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_nu_Ano"))
ano_combo.select_by_visible_text("2025")

# Selecionando o ComboBox de "Região" - selecionar "Centro Oeste"
regiao_combo = Select(driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_regiao"))
regiao_combo.select_by_visible_text("Centro-Oeste")

# Selecionando o ComboBox de "Estado" - selecionar "Todos os Estados"
estado_combo = Select(driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_Estado"))
estado_combo.select_by_visible_text("Todos os Estados")

# Marcar as opções de comparação e ordenação
wait.until(EC.element_to_be_clickable((By.ID, "ctl00_ContentPlaceHolder1_rdComparacao_5"))).click()
wait.until(EC.element_to_be_clickable((By.ID, "ctl00_ContentPlaceHolder1_rdOrdenacao_0"))).click()

# Coletar subagrupadoras
subs_agrupadoras = obter_opcoes("ctl00_ContentPlaceHolder1_subs_agrupadora")
for sub_agrupadora in subs_agrupadoras:
    selecionar_combo("ctl00_ContentPlaceHolder1_subs_agrupadora", sub_agrupadora)
    print("Subagrupadora selecionada: ", sub_agrupadora)

    # Coletar substâncias
    substancias = obter_opcoes("ctl00_ContentPlaceHolder1_Substancia")
    for substancia in substancias:
        selecionar_combo("ctl00_ContentPlaceHolder1_Substancia", substancia)

        # Clicar no botão "Gera"
        button = driver.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_btnGera"]')

        # Rolar até o botão
        driver.execute_script("arguments[0].scrollIntoView(true);", button)
        button.click()
        
        # Espera a presença da tabela
        try:
            tabelas = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "#ctl00_ContentPlaceHolder1_dvResultado > table.tabelaRelatorio")))

            # Verificar se a tabela está presente
            if tabelas:
                tabela = tabelas[0]  # Pega a tabela
                linhas = tabela.find_elements(By.TAG_NAME, "tr")

                # Processa as linhas da tabela
                for linha in linhas:
                    colunas = linha.find_elements(By.TAG_NAME, "td")
                    if colunas:  # Verifica se a linha tem dados
                        dados = [col.text.strip() for col in colunas]
                        if len(dados) == 6:  # Verifica o número esperado de colunas
                            dados_coletados.append(dados + [sub_agrupadora, substancia])

        except Exception as e:
            print(f"Erro ao processar a tabela: {e}")
            continue

# Se dados foram coletados, organiza e salva no Excel

if dados_coletados:

    dados_coletados = [linha[1:] for linha in dados_coletados] 
    colunas = ["Arrecadador (Empresa)", "Qtde Títulos", "Valor", "Recolhimento CFEM", "% Recolhimento CFEM", "Subagrupadora", "Substância"]

    # Criando um DataFrame com os dados coletados
    df = pd.DataFrame(dados_coletados, columns=colunas)

    # Convertendo valores numéricos e percentuais
    df["Valor"] = df["Valor"].apply(lambda x: float(x.replace(".", "").replace(",", ".")) if isinstance(x, str) else x)
    df["Recolhimento CFEM"] = df["Recolhimento CFEM"].apply(lambda x: float(x.replace(".", "").replace(",", ".")) if isinstance(x, str) else x)
    df["% Recolhimento CFEM"] = df["% Recolhimento CFEM"].apply(lambda x: float(x.replace(",", ".").replace("%", "")) if isinstance(x, str) else x)

    # Salvar os dados no Excel
    df.to_excel("dados_com_subagrupadora_substancia.xlsx", index=False)
    print("Dados salvos em 'dados_com_subagrupadora_substancia.xlsx'.")
else:
    print("Nenhum dado coletado.")

# Fechar o driver após a execução
driver.quit()
