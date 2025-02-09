from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime

def esperar_elemento(driver, xpath, timeout=10):
    WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.XPATH, xpath)))

data_hora_atual = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
nome_arquivo = f"Maiores_Arrecadadores_{data_hora_atual}.xlsx"