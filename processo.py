from selenium.webdriver.common.by import By
from multiprocessing import Pool, Manager
from salvar_dados import capturar_todos_os_dados, salvar_dados_completos_planilha, unificar_planilhas
from seletores import clicar_elemento, selecionar_combo, obter_opcoes
from configs import iniciar_driver
import os
import json
import pandas as pd
from selenium.common.exceptions import (
    TimeoutException,
    ElementClickInterceptedException
)
from utils import data_hora_atual
from controle_execucao import PASTA_SALVAMENTO

# Nome do arquivo de flag para parar a execução
ARQUIVO_PARAR = "parar.flag"

# Função para salvar o progresso em progresso.json
def salvar_progresso(progresso, processo_id):
    """Salva o progresso da execução em um arquivo JSON separado para cada processo."""
    nome_arquivo = os.path.join(PASTA_SALVAMENTO, f"progresso_{processo_id}.json")
    with open(nome_arquivo, "w") as arquivo:
        json.dump(progresso, arquivo, indent=4)
    print(f"✅ Progresso salvo em: {nome_arquivo}")



def processar_sub_agrupadoras(sub_agrupadoras_chunk, processo_id, progresso):
    driver = iniciar_driver(headless=True)
    url = "https://sistemas.anm.gov.br/arrecadacao/extra/relatorios/cfem/maiores_arrecadadores.aspx"
    driver.get(url)
    ano = '2025'
    regiao = 'Centro-Oeste'
    nome_arquivo = f"Maiores_Arrecadadores_{data_hora_atual}_{processo_id}.xlsx"
    planilha_temporaria = f"planilha_temp_{processo_id}.xlsx"

    df_temporario = pd.DataFrame(
        columns=['Ano', 'Subs.Agrupadora', 'Substância', 'Estado', 'Município', 'Arrecadador', 'Qtde Títulos',
                 'Operação', 'Recolhimento CFEM', '% Recolhimento CFEM'])

    try:
        selecionar_combo(driver, '//*[@id="ctl00_ContentPlaceHolder1_nu_Ano"]', ano)
        selecionar_combo(driver, '//*[@id="ctl00_ContentPlaceHolder1_regiao"]', regiao)

        arrecadado_por = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_rdComparacao_5")
        ordenado_por = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_rdOrdenacao_0")

        if not arrecadado_por.is_selected():
            arrecadado_por.click()
        if not ordenado_por.is_selected():
            ordenado_por.click()
    except Exception as e:
        print(f"Erro na configuração inicial: {e}")
        driver.quit()
        return

    for sub_agrupadora in sub_agrupadoras_chunk:
        if os.path.exists(ARQUIVO_PARAR):  # Verifica se a execução deve ser interrompida
            print("⏹️ Execução interrompida pelo usuário.")
            driver.quit()
            return

        # Atualizar progresso
        progresso_atual = progresso if progresso else {}
        progresso_atual['subs_agrupadora'] = sub_agrupadora
        salvar_progresso(progresso_atual, processo_id)

        try:
            selecionar_combo(driver, '//*[@id="ctl00_ContentPlaceHolder1_subs_agrupadora"]', sub_agrupadora)
        except Exception as e:
            print(f"Erro ao selecionar subs_agrupadora '{sub_agrupadora}': {e}")
            driver.refresh()
            continue

        try:
            substancias = obter_opcoes(driver, '//*[@id="ctl00_ContentPlaceHolder1_Substancia"]')
        except Exception as e:
            print(f"Erro ao obter substâncias: {e}")
            continue

        for substancia in substancias:
            if substancia == "Todas as Substância":
                continue

            if os.path.exists(ARQUIVO_PARAR):
                print("⏹️ Execução interrompida pelo usuário.")
                driver.quit()
                return

            # Atualizar progresso
            progresso_atual['substancia'] = substancia
            salvar_progresso(progresso_atual, processo_id)

            try:
                selecionar_combo(driver, '//*[@id="ctl00_ContentPlaceHolder1_Substancia"]', substancia)
            except Exception as e:
                print(f"Erro ao selecionar substância '{substancia}': {e}")
                driver.refresh()
                continue

            try:
                estados = obter_opcoes(driver, '//*[@id="ctl00_ContentPlaceHolder1_Estado"]')
            except Exception as e:
                print(f"Erro ao obter estados: {e}")
                continue

            for estado in estados:
                if estado == "Todos os Estados":
                    continue

                if os.path.exists(ARQUIVO_PARAR):
                    print("⏹️ Execução interrompida pelo usuário.")
                    driver.quit()
                    return

                # Atualizar progresso
                progresso_atual['estado'] = estado
                salvar_progresso(progresso_atual, processo_id)

                try:
                    selecionar_combo(driver, '//*[@id="ctl00_ContentPlaceHolder1_Estado"]', estado)
                except Exception as e:
                    print(f"Erro ao selecionar estado '{estado}': {e}")
                    driver.refresh()
                    continue

                try:
                    municipios = obter_opcoes(driver, '//*[@id="ctl00_ContentPlaceHolder1_Municipio"]')
                except Exception as e:
                    print(f"Erro ao obter municípios: {e}")
                    continue

                for municipio in municipios:
                    if municipio in ["Todas os Município", "***"]:
                        continue

                    if os.path.exists(ARQUIVO_PARAR):
                        print("⏹️ Execução interrompida pelo usuário.")
                        driver.quit()
                        return

                    # Atualizar progresso
                    progresso_atual['municipio'] = municipio
                    salvar_progresso(progresso_atual, processo_id)

                    try:
                        selecionar_combo(driver, '//*[@id="ctl00_ContentPlaceHolder1_Municipio"]', municipio)
                    except Exception as e:
                        print(f"Erro ao selecionar município '{municipio}': {e}. Aplicando refresh...")
                        driver.refresh()
                        continue

                    try:
                        clicar_elemento(driver, '//*[@id="ctl00_ContentPlaceHolder1_btnGera"]')
                    except (ElementClickInterceptedException, TimeoutException) as e:
                        print(f"Erro ao clicar no botão 'Gera' para município '{municipio}': {e}. Aplicando refresh...")
                        driver.refresh()
                        continue


                    try:
                        dados_completos = capturar_todos_os_dados(func_navegador=driver, municipio=municipio,
                                                                  regiao=regiao)
                    except Exception as e:
                        print(f"Erro ao capturar dados para município '{municipio}': {e}. Aplicando refresh...")
                        driver.refresh()
                        continue

                    try:
                        salvar_dados_completos_planilha(
                            dados_completos,
                            planilha_temporaria,
                        )
                        print(f"Dados do município {municipio} gravados com sucesso.")
                    except Exception as e:
                        print(f"Erro ao salvar dados para município '{municipio}': {e}. Aplicando refresh...")
                        driver.refresh()
                        continue
    # Salvando a planilha temporária após todos os processos
    df_temporario.to_excel(planilha_temporaria, index=False)
    driver.quit()

if __name__ == "__main__":
    driver_temp = iniciar_driver()
    driver_temp.get("https://sistemas.anm.gov.br/arrecadacao/extra/relatorios/cfem/maiores_arrecadadores.aspx")
    subs_agrupadoras_total = obter_opcoes(driver_temp, '//*[@id="ctl00_ContentPlaceHolder1_subs_agrupadora"]')
    subs_agrupadoras_total = [s for s in subs_agrupadoras_total if s != "Todas as Agrupadoras"]
    driver_temp.quit()

    num_processos = 4
    chunks = [(subs_agrupadoras_total[i::num_processos], {}, i) for i in range(num_processos)]  # Passando o progresso como dict
    with Manager() as manager:
        progresso = manager.dict()  # Criando o progresso compartilhado entre os processos
        with Pool(num_processos) as pool:
            pool.starmap(processar_sub_agrupadoras, [(chunk[0], chunk[2], progresso) for chunk in chunks])

    unificar_planilhas("Maiores_Arrecadadores_Final.xlsx")
