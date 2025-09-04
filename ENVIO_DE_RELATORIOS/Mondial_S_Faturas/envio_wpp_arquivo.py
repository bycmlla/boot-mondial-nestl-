from selenium import webdriver
from selenium.webdriver import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import time
import logging
import os
import pyautogui

logging.basicConfig(level=logging.ERROR, format='%(asctime)s - %(levelname)s - %(message)s')


def send_whatsapp_file(file_path, contact_names):
    print("Entrando na função de envio pelo WhatsApp")
    global driver
    if not os.path.exists(file_path):
        logging.error(f"Arquivo não encontrado: {file_path}. Abortando.")
        return

    options = Options()
    for arg in [
        r"user-data-dir=C:\\ChromeSelenium",
        "profile-directory=Profile1",
        "--remote-debugging-port=0",
        "--disable-gpu",
        "--no-sandbox",
        "--disable-dev-shm-usage"
    ]:
        options.add_argument(arg)

    options.add_experimental_option("prefs", {
        'useAutomationExtension': False,
        'excludeSwitches': ['enable-automation']
    })

    try:
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=options)

        driver.get("https://web.whatsapp.com")

        WebDriverWait(driver, 150).until(
            EC.presence_of_element_located((By.XPATH, '//div[@id="pane-side"]'))
        )

        search_box_xpath = '//div[@contenteditable="true" and @aria-label="Caixa de texto de pesquisa"]'

        for contact_name in contact_names:
            actions = ActionChains(driver)

            try:
                actions.send_keys(Keys.ESCAPE).perform()
                time.sleep(0.5)
                actions.send_keys(Keys.ESCAPE).perform()
                time.sleep(0.5)

                search_box = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.XPATH, search_box_xpath))
                )
                search_box.click()
                time.sleep(0.5)

                search_box.send_keys(Keys.CONTROL + "a")
                time.sleep(0.2)
                search_box.send_keys(Keys.BACKSPACE)
                time.sleep(0.5)

                search_box.send_keys(contact_name)
                time.sleep(1)

                WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, f"//span[@title='{contact_name}']"))
                )

                contact_xpath = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, f"//span[@title='{contact_name}']"))
                )

                contact_xpath.click()
                time.sleep(2)
                contact_xpath.click()

                time.sleep(2)

                print("01")

                plus_button_location = None
                for i in range(5):
                    plus_button_location = pyautogui.locateCenterOnScreen("plusBlack.png", confidence=0.9)
                    if plus_button_location:
                        break
                    time.sleep(1)

                if plus_button_location:
                    pyautogui.moveTo(plus_button_location)
                    pyautogui.click()
                else:
                    logging.error("Botão '+' não encontrado na tela.")
                    raise Exception("Botão '+' não localizado.")

                file_input_xpath = '//input[@type="file"]'
                file_input = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, file_input_xpath))
                )
                file_input.send_keys(file_path)

                send_button_xpath = '//div[@role="button" and @aria-label="Enviar"]'

                send_button = WebDriverWait(driver, 60).until(
                    EC.element_to_be_clickable((By.XPATH, send_button_xpath))
                )
                send_button.click()
                time.sleep(3)

            except Exception as e:
                logging.error(f"Erro ao processar contato '{contact_name}': {e}")
                try:
                    actions.send_keys(Keys.ESCAPE).perform()
                    time.sleep(0.5)
                    actions.send_keys(Keys.ESCAPE).perform()
                except:
                    pass
                continue

    except Exception as e:
        logging.error(f"Erro geral na execução da função: {e}")
    finally:
        if driver:
            time.sleep(10)
            driver.quit()

