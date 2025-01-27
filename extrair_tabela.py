# from selenium.webdriver.common.by import By


# def capturar_dados_segunda_planilha(func_navegador):
#     # Localiza a tabela
#     tabela = func_navegador.find_element(By.CSS_SELECTOR, '#ctl00_ContentPlaceHolder1_dvResultado > table.tabelaRelatorio')

#     # Localiza as linhas da tabela (exceto o cabeçalho)
#     linhas = tabela.find_elements(By.CSS_SELECTOR, "tbody > tr")

#     # Inicializa uma lista para armazenar os dados
#     dados_tabela = []

#     # Itera sobre as linhas e captura os dados das células
#     for linha in linhas:
#         colunas = linha.find_elements(By.TAG_NAME, "td")  # Captura todas as colunas da linha

#         # Verifica se a linha tem pelo menos 6 colunas
#         if len(colunas) >= 6:
#             # Adiciona os dados da linha como um dicionário
#             dados_tabela.append({
#                 "Arrecadador (Empresa)": colunas[1].text,  # Nome da empresa (coluna 2)
#                 "Qtde Títulos": colunas[2].text,          # Quantidade de títulos (coluna 3)
#                 "Operação": colunas[3].text,              # Valor de operação (coluna 4)
#                 "Recolhimento CFEM": colunas[4].text,     # Valor do recolhimento CFEM (coluna 5)
#                 "% Recolhimento CFEM": colunas[5].text    # Percentual de recolhimento CFEM (coluna 6)
#             })
#         else:
#             # Caso a linha não tenha o número esperado de colunas, registre para depuração
#             print(f"Linha ignorada por não ter colunas suficientes: {linha.text}")

#     # Exibe os dados capturados
#     for dado in dados_tabela:
#         print(dado)
#     return dados_tabela