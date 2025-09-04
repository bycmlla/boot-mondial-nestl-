import time
import os
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pyautogui
from cred import ENOVA_USUARIO, ENOVA_URL, ENOVA_SENHA

def importacao_enova(driver, temp_dir):
    driver.get(ENOVA_URL)

    driver.maximize_window()

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

    try:
        time.sleep(5)
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

    print("dentro do cte 01")

    importar_doc_proprio = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//span[text()='Importar Documentos Próprio']"))
    )
    actions = ActionChains(driver)
    actions.move_to_element(importar_doc_proprio).perform()
    print("dentro do cte 02")

    cte_eletronico = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//span[text()='Conhecimento de Transporte Eletrônico']"))
    )
    cte_eletronico.click()
    print("dentro do cte 03")

    time.sleep(2)

    carregar = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//a[.//span[contains(text(), 'Carregar Arquivo')]]"))
    )

    carregar.screenshot("loadButton.png")

    carregar.click()

    # carregar = pyautogui.locateOnScreen('carregar_arquivo.png', confidence=0.9)
    # print("dentro do cte 04")
    # if carregar:
    #     centro = pyautogui.center(carregar)
    #     pyautogui.moveTo(centro)
    #     pyautogui.click()
    # else:
    #     print("Botão não encontrado na tela.")


    caminho_arquivo = os.path.join(temp_dir, "LoteXML_01.zip")

    input_upload = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//input[@type='file']"))
    )
    input_upload.send_keys(caminho_arquivo)

    time.sleep(5)

    botao_upload = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//span[normalize-space()='Upload']/ancestor::a[1]"))
    )
    botao_upload.click()

    time.sleep(4)

    marca_todos = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(
            (By.XPATH, "//span[normalize-space()='Marca Todos']/ancestor::span[contains(@class, 'x-btn-button')]"))
    )
    marca_todos.click()
    time.sleep(1)

    local = pyautogui.locateOnScreen('consultar_botao.png', confidence=0.9)
    if local:
        centro = pyautogui.center(local)
        pyautogui.moveTo(centro)
        pyautogui.click()
    else:
        print("Botão não encontrado na tela.")
