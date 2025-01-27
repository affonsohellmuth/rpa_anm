from selenium import webdriver

driver = webdriver.Chrome()  
driver.maximize_window()
url = "https://sistemas.anm.gov.br/arrecadacao/extra/relatorios/cfem/maiores_arrecadadores.aspx"
driver.get(url)