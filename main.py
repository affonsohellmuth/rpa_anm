from selenium.webdriver.common.by import By
from selenium.common.exceptions import (
    NoSuchElementException,
    TimeoutException,
    ElementClickInterceptedException
)
from seletores import clicar_elemento, selecionar_combo, obter_opcoes
from configs import driver
from utils import data_hora_atual
from salvar_dados import salvar_dados_completos_planilha, capturar_todos_os_dados

def processar():
    try:
        ano = '2025'
        regiao = 'Centro-Oeste'

        try:
            selecionar_combo('//*[@id="ctl00_ContentPlaceHolder1_nu_Ano"]', ano)
            selecionar_combo('//*[@id="ctl00_ContentPlaceHolder1_regiao"]', regiao)
        except NoSuchElementException as e:
            print(f"Erro ao selecionar ano ou região: {e}")
            return

        try:
            arrecadado_por = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_rdComparacao_5")
            ordenado_por = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_rdOrdenacao_0")

            if not arrecadado_por.is_selected():
                arrecadado_por.click()
            if not ordenado_por.is_selected():
                ordenado_por.click()
        except NoSuchElementException as e:
            print(f"Erro ao localizar os elementos de ordenação: {e}")
            return

        try:
            subs_agrupadoras = obter_opcoes('//*[@id="ctl00_ContentPlaceHolder1_subs_agrupadora"]')
        except NoSuchElementException as e:
            print(f"Erro ao obter opções de subs agrupadoras: {e}")
            return

        for sub_agrupadora in subs_agrupadoras:
            if sub_agrupadora == "Todas as Agrupadoras":
                continue

            try:
                selecionar_combo('//*[@id="ctl00_ContentPlaceHolder1_subs_agrupadora"]', sub_agrupadora)
            except Exception as e:
                print(f"Erro ao selecionar subs agrupadora '{sub_agrupadora}': {e}. Aplicando refresh...")
                driver.refresh()
                continue

            try:
                substancias = obter_opcoes('//*[@id="ctl00_ContentPlaceHolder1_Substancia"]')
            except NoSuchElementException as e:
                print(f"Erro ao obter opções de substâncias: {e}")
                continue

            for substancia in substancias:
                if substancia == "Todas as Substância":
                    continue

                try:
                    selecionar_combo('//*[@id="ctl00_ContentPlaceHolder1_Substancia"]', substancia)
                except Exception as e:
                    print(f"Erro ao selecionar substância '{substancia}': {e}. Aplicando refresh...")
                    driver.refresh()
                    continue

                try:
                    estados = obter_opcoes('//*[@id="ctl00_ContentPlaceHolder1_Estado"]')
                except NoSuchElementException as e:
                    print(f"Erro ao obter opções de estados: {e}")
                    continue

                for estado in estados:
                    if estado == "Todos os Estados":
                        continue

                    try:
                        selecionar_combo('//*[@id="ctl00_ContentPlaceHolder1_Estado"]', estado)
                    except Exception as e:
                        print(f"Erro ao selecionar estado '{estado}': {e}. Aplicando refresh...")
                        driver.refresh()
                        continue

                    try:
                        municipios = obter_opcoes('//*[@id="ctl00_ContentPlaceHolder1_Municipio"]')
                    except NoSuchElementException as e:
                        print(f"Erro ao obter opções de municípios: {e}")
                        continue

                    for municipio in municipios:
                        if municipio in ["Todas os Município", "***"]:
                            continue

                        try:
                            selecionar_combo('//*[@id="ctl00_ContentPlaceHolder1_Municipio"]', municipio)
                        except Exception as e:
                            print(f"Erro ao selecionar município '{municipio}': {e}. Aplicando refresh...")
                            driver.refresh()
                            continue

                        try:
                            clicar_elemento('//*[@id="ctl00_ContentPlaceHolder1_btnGera"]')
                        except (ElementClickInterceptedException, TimeoutException) as e:
                            print(f"Erro ao clicar no botão 'Gera' para município '{municipio}': {e}. Aplicando refresh...")
                            driver.refresh()
                            continue

                        try:
                            dados_completos = capturar_todos_os_dados(func_navegador=driver, municipio=municipio, regiao=regiao)
                        except Exception as e:
                            print(f"Erro ao capturar dados para município '{municipio}': {e}. Aplicando refresh...")
                            driver.refresh()
                            continue

                        try:
                            salvar_dados_completos_planilha(
                                dados_completos,
                                nome_arquivo=f"Maiores_Arrecadadores_{data_hora_atual}.xlsx",
                            )
                            print(f"Dados do município {municipio} gravados com sucesso.")
                        except Exception as e:
                            print(f"Erro ao salvar dados para município '{municipio}': {e}. Aplicando refresh...")
                            driver.refresh()
                            continue
    except Exception as e:
        print(f"Ocorreu um erro inesperado durante a execução do processo: {e}")
    finally:
        driver.quit()
        print("Driver encerrado com sucesso.")

processar()