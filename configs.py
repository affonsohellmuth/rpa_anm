from selenium.webdriver import Chrome, ChromeOptions

def iniciar_driver(headless=True):
    options = ChromeOptions()
    if headless:
        options.add_argument("--headless")
    options.add_argument("--disable-gpu")
    options.add_experimental_option('excludeSwitches', ['enable-logging'])  # Remove logs do Chrome
    return Chrome(options=options)  # Retorna o driver sem abrir a URL