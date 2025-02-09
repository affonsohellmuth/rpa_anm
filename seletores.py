import time
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC

def clicar_elemento(driver, xpath, tentativas=3):
    for tentativa in range(tentativas):
        try:
            elemento = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, xpath)))
            ActionChains(driver).move_to_element(elemento).click().perform()
            return
        except Exception as e:
            print(f"Erro ao clicar (tentativa {tentativa+1}): {e}")
            driver.refresh() if tentativa == tentativas-1 else time.sleep(2)

def obter_opcoes(driver, xpath, tentativas=3):
    for tentativa in range(tentativas):
        try:
            select = Select(WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, xpath))))
            return [opcao.text for opcao in select.options]
        except Exception as e:
            print(f"Erro ao obter opções (tentativa {tentativa+1}): {e}")
            driver.refresh() if tentativa == tentativas-1 else time.sleep(2)
    return []

def selecionar_combo(driver, xpath, valor, tentativas=3):
    for tentativa in range(tentativas):
        try:
            select = Select(WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, xpath))))
            select.select_by_visible_text(valor)
            return
        except Exception as e:
            print(f"Erro ao selecionar '{valor}' (tentativa {tentativa+1}): {e}")
            driver.refresh() if tentativa == tentativas-1 else time.sleep(2)