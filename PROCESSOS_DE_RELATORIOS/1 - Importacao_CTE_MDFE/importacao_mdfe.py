import os
import time
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from cred import ENOVA_USUARIO, ENOVA_URL, ENOVA_SENHA

def importacao_mdfe(driver, temp_dir):
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

    connect_button = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "btnconfirmar"))
    )
    connect_button.click()

    driver.switch_to.default_content()

    time.sleep(5)
    try:
        gestao_operacional_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//span[text()='3 Gestão Operacional']/ancestor::a[1]"))
        )
        gestao_operacional_button.click()
    except Exception as e:
        print(f"Erro ao tentar clicar em '3 Gestão Operacional': {e}")
        try:
            span_element = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//span[text()='3 Gestão Operacional']"))
            )
            actions = ActionChains(driver)
            actions.move_to_element(span_element).click().perform()
        except Exception as e2:
            print(f"Falha também com ActionChains: {e2}")
    time.sleep(3)

    edi_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//span[text()='EDI']"))
    )
    actions = ActionChains(driver)
    actions.move_to_element(edi_button).perform()

    importar_doc_proprio = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//span[text()='Importar Documentos Próprio']"))
    )
    actions = ActionChains(driver)
    actions.move_to_element(importar_doc_proprio).perform()

    mdfe_item = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((
            By.XPATH,
            "//span[normalize-space()='Manifesto de Documento Fiscal Eletrônico']/ancestor::a[1]"
        ))
    )
    mdfe_item.click()

    driver.execute_script("""
      document.querySelectorAll('.x-mask').forEach(el => el.remove());
      document.querySelectorAll('.x-mask-msg').forEach(el => el.remove());
    """)

    WebDriverWait(driver, 30).until(
        lambda d: "x-disabled" not in d.find_element(By.XPATH, "//a[.//b[text()='Consultar']]").get_attribute("class")
    )

    botao = driver.find_element(By.XPATH, "//a[.//b[text()='Consultar']]")

    driver.execute_script("arguments[0].scrollIntoView(true);", botao)
    time.sleep(1)

    actions = ActionChains(driver)
    actions.move_to_element(botao).pause(0.2).click().perform()

    caminho_arquivo = os.path.join(temp_dir, "LoteXML_02.zip")

    input_upload = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//input[@type='file']"))
    )
    input_upload.send_keys(caminho_arquivo)

    botao_upload = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//span[normalize-space()='Upload']/ancestor::a[1]"))
    )
    botao_upload.click()

    marca_todos = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(
            (By.XPATH, "//span[normalize-space()='Marca Todos']/ancestor::span[contains(@class, 'x-btn-button')]"))
    )
    marca_todos.click()
    time.sleep(1)

    importar = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(
            (By.XPATH, "//a[contains(@class, 'x-btn') and .//span[normalize-space()='Importar']]"))
    )
    importar.click()

    WebDriverWait(driver, 180).until(
        EC.invisibility_of_element_located((By.CSS_SELECTOR, "div.blockUI.blockOverlay"))
    )