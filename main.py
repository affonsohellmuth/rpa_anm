from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from seletores import clicar_elemento, selecionar_combo, obter_opcoes
from configs import driver
from utils import data_hora_atual
from salvar_dados import salvar_dados_completos_planilha, capturar_todos_os_dados


def processar():
    ano = '2025'
    regiao = 'Centro-Oeste'
    selecionar_combo('//*[@id="ctl00_ContentPlaceHolder1_nu_Ano"]', ano)
    selecionar_combo('//*[@id="ctl00_ContentPlaceHolder1_regiao"]', regiao)

    arrecadado_por = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_rdComparacao_5")
    ordenado_por = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_rdOrdenacao_0")

    if not arrecadado_por.is_selected():
        arrecadado_por.click()
    
    if not ordenado_por.is_selected():
        ordenado_por.click()

    subs_agrupadoras = obter_opcoes('//*[@id="ctl00_ContentPlaceHolder1_subs_agrupadora"]')
    for sub_agrupadora in subs_agrupadoras:
        if sub_agrupadora == "Todas as Agrupadoras":
            continue
        selecionar_combo('//*[@id="ctl00_ContentPlaceHolder1_subs_agrupadora"]', sub_agrupadora)

        substancias = obter_opcoes('//*[@id="ctl00_ContentPlaceHolder1_Substancia"]')
        for substancia in substancias:
            if substancia == "Todas as Substância":
                continue
            selecionar_combo('//*[@id="ctl00_ContentPlaceHolder1_Substancia"]', substancia)

            estados = obter_opcoes('//*[@id="ctl00_ContentPlaceHolder1_Estado"]')
            for estado in estados:
                if estado == "Todos os Estados":
                    continue
                selecionar_combo('//*[@id="ctl00_ContentPlaceHolder1_Estado"]', estado)

                municipios = obter_opcoes('//*[@id="ctl00_ContentPlaceHolder1_Municipio"]')
                for municipio in municipios:
                    if municipio in ["Todas os Município", "***"]:
                        continue
                    selecionar_combo('//*[@id="ctl00_ContentPlaceHolder1_Municipio"]', municipio)

                    clicar_elemento('//*[@id="ctl00_ContentPlaceHolder1_btnGera"]')

                    dados_completos = capturar_todos_os_dados(func_navegador=driver, municipio=municipio, regiao=regiao)

                    salvar_dados_completos_planilha(dados_completos, nome_arquivo=f"Maiores_Arrecadadores_{data_hora_atual}.xlsx")
                    print(f"Dados do município {municipio} gravados com sucesso.")

# Iniciar o processo
processar()

# Fechar o driver ao final
driver.quit()
