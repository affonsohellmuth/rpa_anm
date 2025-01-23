from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException
import pandas as pd
import time

# Configuração inicial do Selenium
driver = webdriver.Chrome()  # Substitua por seu driver adequado
driver.maximize_window()
url = "https://sistemas.anm.gov.br/arrecadacao/extra/relatorios/cfem/maiores_arrecadadores.aspx"
driver.get(url)
wait = WebDriverWait(driver, 10)

# Definir valores fixos
ano_fixo = "2025"
regiao_fixa = "Centro-Oeste"
estados_validos = ["Distrito Federal", "Goiás", "Mato Grosso", "Mato Grosso do Sul"]

# Selecionar o Ano
ano_combo = Select(wait.until(EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_nu_Ano"))))
ano_combo.select_by_visible_text(ano_fixo)

# Selecionar Região
regiao_combo = Select(wait.until(EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_regiao"))))
regiao_combo.select_by_visible_text(regiao_fixa)

# Iterar por Subs. Agrupadora
subs_agrupadora_combo = Select(wait.until(EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_subs_agrupadora"))))
for subs_agrupadora in subs_agrupadora_combo.options:
    subs_agrupadora_text = subs_agrupadora.text
    # Ignorar a opção "Todas as Agrupadoras"
    if subs_agrupadora_text == "Todas as Agrupadoras":
        continue
    subs_agrupadora_combo.select_by_visible_text(subs_agrupadora_text)
    time.sleep(1)

    # Iterar por Substância
    substancia_combo = Select(wait.until(EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_Substancia"))))
    for substancia in substancia_combo.options:
        substancia_text = substancia.text
        # Ignorar a opção "Todas as Substâncias"
        if substancia_text == "Todas as Substância":
            continue
        substancia_combo.select_by_visible_text(substancia_text)
        time.sleep(1)

        # Iterar por Estados
        estado_combo = Select(wait.until(EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_Estado"))))
        for estado in estado_combo.options:
            try:
                estado_text = estado.text
                # Ignorar a opção "Todos os Estados"
                if estado_text == "Todos os Estados":
                    continue
                
                # Selecionar o estado
                estado_combo.select_by_visible_text(estado_text)
                time.sleep(1)
                
                # Tentar capturar o estado novamente após a interação
                estado_combo = Select(wait.until(EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_Estado"))))
                
            except StaleElementReferenceException:
                # Se o erro acontecer, tenta-se novamente
                print(f"Erro com o estado {estado_text}, tentando novamente...")
                estado_combo = Select(wait.until(EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_Estado"))))
                estado_combo.select_by_visible_text(estado_text)
                time.sleep(1)

            # Iterar por Municípios
            municipio_combo = Select(wait.until(EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_Municipio"))))

            for municipio in municipio_combo.options:
                municipio_text = municipio.text
                # Ignorar as opções "Todas os Municípios" e "***"
                if municipio_text == "Todas os Município" or municipio_text == "***":
                    continue
                try:
                    municipio_combo.select_by_visible_text(municipio_text)
                    time.sleep(1)
                except StaleElementReferenceException:
                    # Se o erro acontecer, tenta-se novamente
                    print("Elemento desatualizado, tentando novamente...")
                    municipio_combo = Select(wait.until(EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_Municipio"))))
                    municipio_text = municipio.text
                    municipio_combo.select_by_visible_text(municipio_text)
                    time.sleep(1)

                arrecadado_por = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_rdComparacao_5")
                ordenado_por = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_rdOrdenacao_0")

                if arrecadado_por.is_selected():
                    pass
                else:
                    arrecadado_por.click()

                time.sleep(2)

                if ordenado_por.is_selected():
                    pass
                else:
                    ordenado_por.click()

                # Clicar em "Gera"
                gera_button = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_btnGera")
                gera_button.click()

                time.sleep(2)

                # Aguardar a tabela carregar
                wait.until(EC.presence_of_element_located((By.CLASS_NAME, "tabelaRelatorio")))

                # Extração da tabela
                tabela = driver.find_element(By.CLASS_NAME, "tabelaRelatorio")
                linhas = tabela.find_elements(By.TAG_NAME, "tr")

                # Verificar se a tabela não está vazia
                if len(linhas) > 2:  # Ignora cabeçalho e rodapé
                    # Processar linhas da tabela
                    dados = []
                    for linha in linhas[2:-1]:  # Ignora cabeçalho e rodapé
                        colunas = linha.find_elements(By.TAG_NAME, "td")
                        dados.append({
                            "Subs. Agrupadora": subs_agrupadora_text,  # Usando o texto salvo no início do loop
                            "Substância": substancia_text,  # Usando o texto salvo no início do loop
                            "Estado": estado_text,  # Usando o texto salvo no início do loop
                            "Município": municipio_text,  # Usando o texto salvo no início do loop
                            "Posição": colunas[0].text.strip(),
                            "Empresa": colunas[1].text.strip(),
                            "Qtde Títulos": colunas[2].text.strip(),
                            "Valor Operação": colunas[3].text.strip(),
                            "Valor Recolhimento CFEM": colunas[4].text.strip(),
                            "% Recolhimento CFEM": colunas[5].text.strip()
                        })
                    # Criação de uma única planilha com os dados extraídos
                    df = pd.DataFrame(dados)

                    # Salva os dados em uma planilha separada por colunas
                    df.to_excel("maiores_arrecadadores_completo.xlsx", index=False)

                    print("Dados extraídos e salvos em 'maiores_arrecadadores_completo.xlsx'.")
                else:
                    print("Tabela vazia, nenhum dado extraído.")

# Encerrar o driver
driver.quit()
