import time
import random
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException,
    NoSuchElementException,
    ElementClickInterceptedException,
    ElementNotInteractableException,
)
from utils import esperar_elemento
from configs import driver

def clicar_elemento(xpath, tentativas=3):
    for tentativa in range(tentativas):
        try:
            elemento = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, xpath)))
            actions = ActionChains(driver)
            actions.move_to_element(elemento).perform()
            elemento.click()
            print(f"Botão clicado: {xpath}")
            return  
        except (TimeoutException, ElementClickInterceptedException, ElementNotInteractableException) as e:
            print(f"Erro ao clicar no botão (tentativa {tentativa + 1} de {tentativas}): {e}")
            if tentativa < tentativas - 1:
                time.sleep(random.uniform(1, 3))  
                continue
            driver.refresh()  
            print("Página atualizada após erro ao clicar no botão.")

def obter_opcoes(xpath, tentativas=3):
    for tentativa in range(tentativas):
        try:
            esperar_elemento(xpath)
            select = Select(driver.find_element(By.XPATH, xpath))
            opcoes = [opcao.text for opcao in select.options]
            return opcoes
        except (TimeoutException, NoSuchElementException) as e:
            print(f"Erro ao obter opções (tentativa {tentativa + 1} de {tentativas}): {e}")
            if tentativa < tentativas - 1:
                time.sleep(random.uniform(1, 3))  
                continue
            driver.refresh()  
            print("Página atualizada após erro ao obter opções.")
    return []

def selecionar_combo(xpath, valor, tentativas=3):
    for tentativa in range(tentativas):
        try:
            esperar_elemento(xpath)
            select = Select(driver.find_element(By.XPATH, xpath))
            select.select_by_visible_text(valor)
            print(f"Opção selecionada: {valor}")
            return
        except (TimeoutException, NoSuchElementException, ElementNotInteractableException) as e:
            print(f"Erro ao selecionar opção (tentativa {tentativa + 1} de {tentativas}): {e}")
            if tentativa < tentativas - 1:
                time.sleep(random.uniform(1, 3))  
                continue
            driver.refresh()  
            print("Página atualizada após erro ao selecionar opção.")
