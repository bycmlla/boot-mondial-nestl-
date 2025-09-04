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
from envio_wpp_arquivo import send_whatsapp_file
from tratamento_csv import treatment_csv
from Relatorio_Nestle.dwld_relatorio_nestle import download_file_fiscal
from cred import ENOVA_USUARIO, ENOVA_SENHA, ENOVA_URL_3500, CONTATO_01, CONTATOS_ARRAY

logging.basicConfig(level=logging.ERROR, format='%(asctime)s - %(levelname)s - %(message)s')

def access_site():
    temp_dir = tempfile.mkdtemp()

    options = Options()
    for arg in [
        "--start_maximized",
        "--disable-extensions",
        "--allow-running-insecure-content",
        f"--unsafely-treat-insecure-origin-as-secure={ENOVA_URL_3500}",
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
        driver.get(ENOVA_URL_3500)

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

        time.sleep(5)

        try:
            time.sleep(3)
            xpath_relativo = "//label[normalize-space(text())='Lay Out']/preceding-sibling::a[contains(@class, 'x-btn-icon')][1]"
            botao_layout = WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.XPATH, xpath_relativo))
            )
            botao_layout.click()
        except Exception:
            try:
                xpath_relativo_input = "//input[@name='O13DF']/ancestor::table[1]/preceding-sibling::a[contains(@class, 'x-btn-icon')][1]"
                botao_layout = WebDriverWait(driver, 20).until(
                    EC.element_to_be_clickable((By.XPATH, xpath_relativo_input))
                )
                botao_layout.click()
            except Exception as e2:
                print(f"Erro ao tentar clicar no botão usando XPath relativo ao input: {e2}")
        time.sleep(3)
        celula_b = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, "//tr[td/div[text()='Visualiza Notas Fiscais']]"))
        )
        actions = ActionChains(driver)
        actions.double_click(celula_b).perform()

        time.sleep(4)

        xpath_checkboxes_para_desmarcar = "//label[normalize-space(text())='Cancelados']/following-sibling::div[contains(@class,'x-grid')][1]//tr[.//td/div[normalize-space(text())='Sim']]//img[contains(@class, 'x-grid-checkcolumn-checked')]"

        try:
            checkboxes_encontrados = WebDriverWait(driver, 20).until(
                EC.presence_of_all_elements_located((By.XPATH, xpath_checkboxes_para_desmarcar))
            )
            for checkbox_img in checkboxes_encontrados:
                try:
                    checkbox_clicavel = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable(checkbox_img)
                    )
                    checkbox_clicavel.click()
                except (StaleElementReferenceException, TimeoutException):
                    continue
                except Exception:
                    continue
        except TimeoutException:
            pass
        except Exception as e_find:
            print(f"Erro ao procurar por checkboxes na grade 'Cancelados': {e_find}")

        try:
            xpath_botao_grupo_cliente = "//label[normalize-space(text())='Selecionar Grupo Cliente']/preceding-sibling::a[contains(@class, 'x-btn-icon')][1]"
            grupo_cliente_button = WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.XPATH, xpath_botao_grupo_cliente))
            )
            grupo_cliente_button.click()
        except Exception as e:
            print(f"Erro ao tentar clicar no botão 'Selecionar Grupo Cliente': {e}")

        time.sleep(5)
        elemento_selecionar = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, "//td/div[normalize-space(text())='Selecionar']"))
        )
        actions = ActionChains(driver)
        actions.double_click(elemento_selecionar).perform()

        try:
            xpath_botao_grupo_cliente = "//label[normalize-space(text())='Grupo Cliente']/preceding-sibling::a[contains(@class, 'x-btn-icon')][1]"
            botao_grupo_cliente = WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.XPATH, xpath_botao_grupo_cliente))
            )
            botao_grupo_cliente.click()
        except Exception as e:
            print(f"Erro ao tentar clicar no botão 'Grupo Cliente': {e}")

        mondial = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, "//td/div[text()='MONDIAL']"))
        )
        actions = ActionChains(driver)
        actions.double_click(mondial).perform()

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

        WebDriverWait(driver, 60).until(lambda d: len(d.window_handles) > 1)

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

        xpath_opcao_csv = "//div[contains(@class, 'x-menu-item')]//b[normalize-space(text())='CSV']/ancestor::a[contains(@class, 'x-menu-item-link')][1]"

        try:
            opcao_csv = WebDriverWait(driver,15).until(
                EC.element_to_be_clickable((By.XPATH, xpath_opcao_csv))
            )
            opcao_csv.click()
        except Exception as e_csv:
            print(f"Erro ao clicar na opção CSV: {e_csv}")

        WebDriverWait(driver, 20).until(lambda d: any(fname.endswith(".zip") for fname in os.listdir(temp_dir)))

        processed_xlsx_path = treatment_csv(temp_dir)



        if processed_xlsx_path:
            time.sleep(2)
            driver.quit()
            hoje = datetime.datetime.now().weekday()

            contatos_whatsapp = []

            if hoje != 6:
                contatos_whatsapp.append(CONTATO_01)

            if hoje in [0, 2, 5]:
                contatos_whatsapp += CONTATOS_ARRAY

            if contatos_whatsapp:
                send_whatsapp_file(processed_xlsx_path, contatos_whatsapp)
            else:
                logging.info("Hoje é domingo. Nenhum contato definido para envio.")
        else:
            logging.error("Falha no processamento do arquivo CSV. Envio pelo WhatsApp cancelado.")

    except Exception as e:
        print(f"Erro durante a execução: {e}")
    finally:
        time.sleep(3)
        shutil.rmtree(temp_dir, ignore_errors=True)
        print("Baixando Nestlé")
        download_file_fiscal()

# access_site()

schedule.every().day.at("10:20").do(access_site)

ultimo_log = 0

while True:
    schedule.run_pending()

    agora = time.time()
    if agora - ultimo_log >= 120:
        print(f"[{datetime.datetime.now().strftime('%H:%M:%S')}] \U0001F4C1 Aguardando para ENVIO")
        ultimo_log = agora

    time.sleep(1)