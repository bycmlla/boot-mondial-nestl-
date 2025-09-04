from datetime import datetime, timedelta
import schedule
from selenium import webdriver
from selenium.webdriver import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support import expected_conditions as EC
import time
import tempfile
from cred import FSIS_LOGIN, FSIS_PASSWORD, FSIS_URL
from pathlib import Path
import zipfile


def access_fsis():
    temp_dir = tempfile.mkdtemp()

    options = Options()

    for arg in [
        "--disable-extensions",
        "--window_size=1920,1080",
        "--disable_extensions",
        "--allow-running-insecure-content",
        f"--unsafely-treat-insecure-origin-as-secure={FSIS_URL}",
        "--disable-gpu",
        "--no-sandbox",
        "--disable-dev-shm-usage",
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36",
        "--disable-features=AutofillServerCommunication,AutofillEnableAccountWalletStorage,AutofillEnablePaymentsIdentityDetection,PasswordManagerEnabled,PasswordManagerRedesign,PasswordCheck,SafeBrowsingEnhancedProtection,SafeBrowsingRealTimeUrlLookupEnterprise,SafeBrowsingRealTimeLookup,SafeBrowsingUrlLookupService"
    ]: options.add_argument(arg)

    options.add_experimental_option("prefs", {
        "download.default_directory": temp_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": False,
        "profile.password_manager_enabled": False,
        "credentials_enable_service": False,
        "profile.password_manager_leak_detection": False,
        "useAutomationExtension": False,
    })

    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)

    try:
        driver.get(FSIS_URL)
        driver.maximize_window()

        time.sleep(3)

        username = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.NAME, "usuario"))
        )

        username.send_keys(FSIS_LOGIN)

        password = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.NAME, "senha"))
        )

        password.send_keys(FSIS_PASSWORD)

        login_button = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//input[@type="submit" and @value="Entrar"]'))
        )
        login_button.click()

        print("29293")

        time.sleep(20)

        hoje = datetime.today()
        ontem = hoje - timedelta(days=1)
        data_ontem = ontem.strftime("%d/%m/%Y")
        print(data_ontem)

        dropdown = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH,
                                            "//div[contains(@class, 'titulo-busca') and text()='Período']/following::div[contains(@class, 'but-ads')][1]")
                                           )
        )
        dropdown.click()

        input_data_inicial = driver.find_element(By.XPATH,'//label[contains(text(), "Data Inicial")]/following::input[@type="text" and contains(@class, "form-control")][1]')

        input_data_inicial.click()
        input_data_inicial.send_keys(Keys.CONTROL, 'a')
        input_data_inicial.send_keys(Keys.BACKSPACE)
        input_data_inicial.send_keys(data_ontem + Keys.ENTER)

        input_data2 = driver.find_element(By.XPATH, '//input[@id="data2"]')

        input_data2.click()
        input_data2.send_keys(Keys.CONTROL, 'a')
        input_data2.send_keys(Keys.BACKSPACE)
        input_data2.send_keys(data_ontem + Keys.ENTER)

        time.sleep(4)

        confirmar_button = driver.find_element(By.XPATH, '//input[@type="button" and @value="CONFIRMAR"]')
        confirmar_button.click()

        time.sleep(5)

        selecionar_todas = driver.find_element(By.XPATH,
                                               '//div[@class="btn btn-default" and text()="Selecionar todas"]')
        selecionar_todas.click()

        download_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(
                (By.XPATH, '//div[contains(@class, "btn-success") and .//span[text()="Download"]]'))
        )
        download_button.click()

        btn_xml = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(
                (By.XPATH, '//button[contains(@class, "btn-img") and .//span[text()="Apenas XMLs"]]')
            )
        )
        btn_xml.click()

        zip_path = None
        for _ in range(30):
            matches = list(Path(temp_dir).glob("*.zip"))
            if matches:
                zip_path = matches[0]
                break
            time.sleep(1)

        if not zip_path:
            print("ZIP não encontrado.")
            return

        destino = Path(r"\\server\JTDTRANSPORTES2\ANALITCS\ACOMPANHAMENTO DE CARGAS\xml")
        destino.mkdir(parents=True, exist_ok=True)

        with zipfile.ZipFile(zip_path, "r") as zf:
            zf.extractall(path=destino)

        print(f"Arquivos extraídos para {destino}")


        time.sleep(5)

    except Exception:
        print("getouit")
    finally:
        driver.quit()

# access_fsis()

schedule.every().day.at("09:05").do(access_fsis)

ultimo_log = 0

while True:
    schedule.run_pending()

    agora = time.time()
    if agora - ultimo_log >= 120:
        print(f"[{datetime.now().strftime('%H:%M:%S')}] \U0001F4E6 Aguardando XML")
        ultimo_log = agora

    time.sleep(1)
