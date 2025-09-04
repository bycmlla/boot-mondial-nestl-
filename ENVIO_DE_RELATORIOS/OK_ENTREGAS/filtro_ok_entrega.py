import schedule
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support import expected_conditions as EC
import time
import pyautogui
from datetime import datetime
from cred import OK_ENTREGA_URL, OK_ENTREGA_USUARIO, OK_ENTREGA_SENHA


def filtro_ok_entrega():
    options = Options()
    for arg in [
        "--start_maximized",
        "--disable-extensions",
        "--allow-running-insecure-content",
        f"--unsafely-treat-insecure-origin-as-secure={OK_ENTREGA_URL}",
        "--disable-gpu",
        "--no-sandbox",
        "--disable-dev-shm-usage"
    ]: options.add_argument(arg)

    options.add_experimental_option("prefs", {
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": False,
    })

    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)

    try:
        driver.get(OK_ENTREGA_URL)
        driver.maximize_window()

        email_button = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.NAME, "camp_identificacao"))
        )

        email_button.send_keys(OK_ENTREGA_USUARIO)

        senha = driver.find_element(By.XPATH, '//input[@placeholder="Informe sua senha ..."]')

        senha.send_keys(OK_ENTREGA_SENHA)

        botao_entrar = driver.find_element(By.CLASS_NAME, "loginButton")

        botao_entrar.click()

        print("Entrando no sistema... Aguarde.")
        time.sleep(5)

        hambuguer_menu = WebDriverWait(driver, 10).until((
            EC.presence_of_element_located((By.CLASS_NAME, "hamburger"))
        ))

        hambuguer_menu.click()

        consulta_button = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//a[span[contains(text(), "Consulta")]]'))
        )

        consulta_button.click()
        time.sleep(1)

        exportacao_excel = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//a[text()="Exportação Excel"]'))
        )

        exportacao_excel.click()
        time.sleep(3)

        close_btn = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "icon_close"))
        )
        close_btn.click()

        WebDriverWait(driver, 10).until(
            EC.invisibility_of_element_located((By.ID, "alertModalAviso"))
        )

        time.sleep(3)

        filtro_btn = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "a.nav-link.bell-link"))
        )

        filtro_btn.click()

        intermediario_link = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//a[text()='Intermediário']"))
        )

        intermediario_link.click()

        time.sleep(1)

        area = pyautogui.locateOnScreen('area-elemento.png', confidence=0.8)
        if area:
            print("Área encontrada:", area)

            campo = pyautogui.locateOnScreen('selecione-campo.png', region=area, confidence=0.8)

            if campo:
                print("Campo 'Selecione' encontrado:", campo)
                pyautogui.click(campo)
            else:
                print("Campo 'Selecione' não encontrado.")
        else:
            print("Área do elemento não encontrada.")

        item_emissao = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.XPATH, "//li[@data-value='emissao']"))
        )
        item_emissao.click()

        time.sleep(1)

        area_periodo = pyautogui.locateOnScreen('periodo-area.png', confidence=0.8)
        if area:
            print("Área encontrada:", area_periodo)

            campo_periodo = pyautogui.locateOnScreen('periodo-campo.png', region=area_periodo, confidence=0.8)

            if campo_periodo:
                print("Campo 'Periodo' encontrado:", campo_periodo)
                pyautogui.click(campo_periodo)
            else:
                print("Campo 'Periodo' não encontrado.")
        else:
            print("Área do elemento não encontrada.")

        time.sleep(1)

        data_especifica = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.XPATH, "//li[text()='Data Específica']"))
        )
        data_especifica.click()

        time.sleep(1)
        data_hoje = datetime.today().strftime("%d/%m/%Y")
        intervalo_data = f"01/01/2025 - {data_hoje}"

        campo_data = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "issueSpecificDate"))
        )

        campo_data.clear()
        campo_data.send_keys(intervalo_data)

        try:
            botao_fechar = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, ".iziToast.iziToast-opened .iziToast-close"))
            )
            botao_fechar.click()
        except:
            print("Toast não apareceu ou já sumiu.")

        time.sleep(2)

        botao_filtrar = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "a.filtrar_action.bt_filtrar"))
        )
        botao_filtrar.click()

    except Exception as e:
        print(f"Erro durante a execução: {e}")
    finally:
        time.sleep(3)
        driver.quit()


filtro_ok_entrega()

schedule.every().day.at("07:30").do(filtro_ok_entrega)
schedule.every().day.at("09:30").do(filtro_ok_entrega)
schedule.every().day.at("13:30").do(filtro_ok_entrega)
schedule.every().day.at("17:00").do(filtro_ok_entrega)

ultimo_log = 0

while True:
    schedule.run_pending()

    agora = time.time()
    if agora - ultimo_log >= 120:
        print(f"[{datetime.now().strftime('%H:%M:%S')}] \U0001F4BB Aguardando FILTRO OK ENT.")
        ultimo_log = agora

    time.sleep(1)

