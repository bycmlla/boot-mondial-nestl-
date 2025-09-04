import datetime
import schedule
from selenium import webdriver
from selenium.common import StaleElementReferenceException, TimeoutException
from selenium.webdriver import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import time
import shutil
import tempfile
import logging
import os
from tratar_html import tratar_relatorio
from cred import ENOVA_USUARIO, ENOVA_SENHA, ENOVA_URL, CONTATO_01, CONTATOS_ARRAY

logging.basicConfig(level=logging.ERROR, format='%(asctime)s - %(levelname)s - %(message)s')

def download_relatorio():
    temp_dir = tempfile.mkdtemp()

    options = Options()
    for arg in [
        "--start_maximized",
        "--disable-extensions",
        "--allow-running-insecure-content",
        f"--unsafely-treat-insecure-origin-as-secure={ENOVA_URL}",
        "--disable-gpu",
        "--no-sandbox",
        "--disable-dev-shm-usage"
    ]: options.add_argument(arg)

    options.add_experimental_option("prefs", {
        "download.default_directory": temp_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": False,
    })

    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)

    try:
        driver.get(ENOVA_URL)

        WebDriverWait(driver, 10).until(
            EC.frame_to_be_available_and_switch_to_it((By.ID, "component-1009"))
        )

        username = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.NAME, "Username"))
        )
        username.send_keys(ENOVA_USUARIO)

        password = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.NAME, "Password"))
        )
        password.send_keys(ENOVA_SENHA)

        time.sleep(3)

        connect_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "btnconfirmar"))
        )
        connect_button.click()

        driver.switch_to.default_content()

        try:
            time.sleep(5)
            gestao_fiscal_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//span[text()='5 Gestão Fiscal/Contábil']/ancestor::a[1]"))
            )
            gestao_fiscal_button.click()
        except Exception:
            try:
                gestao_fiscal_button = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable(
                        (By.XPATH, "//span[contains(text(), 'Gestão Fiscal/Contábil')]/ancestor::a[1]"))
                )
                gestao_fiscal_button.click()
            except Exception as e2:
                print(f"Erro ao tentar clicar pelo texto parcial: {e2}")

        menu_fiscal = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, "//span[text()='Fiscal']"))
        )
        actions = ActionChains(driver)
        actions.move_to_element(menu_fiscal).perform()

        elemento_relatorio = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, "//span[text()='Relatório de Conferência Conhecimento']"))
        )
        elemento_relatorio.click()

        time.sleep(3)

        try:
            xpath_relativo_label = "//label[normalize-space(text())='Período']/preceding-sibling::table[contains(@class, 'x-field')][1]//input[@type='text']"
            campo_data = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, xpath_relativo_label))
            )
            campo_data.clear()
            campo_data.send_keys("01/01/2025")
            campo_data.send_keys(Keys.TAB)
        except Exception as e:
            print(f"Erro ao interagir com o campo de data usando XPath relativo ao label: {e}")

        try:
            xpath_botao_consultar = "//b[normalize-space(text())='Consultar']/ancestor::a[contains(@class, 'x-btn')][1]"
            time.sleep(3)
            botao_consultar = WebDriverWait(driver, 50).until(
                EC.element_to_be_clickable((By.XPATH, xpath_botao_consultar))
            )
            botao_consultar.click()
        except Exception as e:
            print(f"Erro ao tentar clicar no botão 'Consultar': {e}")

        WebDriverWait(driver, 120).until(lambda d: len(d.window_handles) > 1)

        aba_original = driver.current_window_handle
        for aba in driver.window_handles:
            if aba != aba_original:
                driver.switch_to.window(aba)
                driver.execute_script("window.open('', '_self', ''); window.close();")
                break
        driver.switch_to.window(aba_original)

        botao_exportar = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, ".x-btn-wrap.x-btn-split.x-btn-split-right"))
        )
        actions = ActionChains(driver)
        actions.move_to_element(botao_exportar).move_by_offset(60, 0).click().perform()

        xpath_opcao_excel = "//div[contains(@class, 'x-menu-item')]//b[normalize-space(text())='Excel']/ancestor::a[contains(@class, 'x-menu-item-link')][1]"

        try:
            opcao_excel = WebDriverWait(driver, 15).until(
                EC.element_to_be_clickable((By.XPATH, xpath_opcao_excel))
            )
            opcao_excel.click()
        except Exception as e_excel:
            print(f"Erro ao clicar na opção Excel: {e_excel}")

        WebDriverWait(driver, 20).until(lambda d: any(fname.endswith(".xls") for fname in os.listdir(temp_dir)))

        caminho_destino = r"\\server\JTDTRANSPORTES2\ANALITCS\ACOMPANHAMENTO DE CARGAS\teste\testte.xlsx"

        tratar_relatorio(temp_dir, caminho_destino)

    except Exception as e:
        print(f"Erro durante a execução: {e}")
    finally:
        time.sleep(3)
        shutil.rmtree(temp_dir, ignore_errors=True)





download_relatorio()








schedule.every().day.at("10:20").do(download_relatorio)

ultimo_log = 0

while True:
    schedule.run_pending()

    agora = time.time()
    if agora - ultimo_log >= 120:
        print(f"[{datetime.datetime.now().strftime('%H:%M:%S')}] \U0001F4C1 Aguardando para ENVIO")
        ultimo_log = agora

    time.sleep(1)