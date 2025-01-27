from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from configs import driver
from datetime import datetime

def esperar_elemento(xpath):
    WebDriverWait(driver, 3).until(EC.visibility_of_element_located((By.XPATH, xpath)))

data_hora_atual = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
nome_arquivo = f"Maiores_Arrecadadores_{data_hora_atual}.xlsx"