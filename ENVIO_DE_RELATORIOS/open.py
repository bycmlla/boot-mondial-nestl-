from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
import time

def abrir_whatsapp_web():
    options = Options()
    options.add_argument(r"user-data-dir=C:\\ChromeSelenium")
    options.add_argument("profile-directory=Profile1")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    driver.get("https://web.whatsapp.com")

    input("Pressione Enter para fechar o navegador...")
    time.sleep(300)
    driver.quit()

abrir_whatsapp_web()
