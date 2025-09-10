import tempfile
import time
import os
import pyautogui
import schedule
from selenium import webdriver
from datetime import datetime, timedelta
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support import expected_conditions as EC
from importacao_cte import importacao_enova
from importacao_mdfe import importacao_mdfe
from cred import MULTICTE_URL, MULTICTE_SENHA, MULTICTE_USUARIO

def aguardar_download(download_dir, nome_arquivo, timeout=30):
    caminho_final = os.path.join(download_dir, nome_arquivo)
    caminho_temp = caminho_final + ".crdownload"

    tempo_inicial = time.time()
    while time.time() - tempo_inicial < timeout:
        if os.path.exists(caminho_final) and not os.path.exists(caminho_temp):
            try:
                with open(caminho_final, "rb"):
                    return caminho_final
            except PermissionError:
                pass
        time.sleep(0.5)

    raise TimeoutError(f"Download do arquivo {nome_arquivo} não finalizou em tempo hábil.")

def verificar_importacao_cte(driver, timeout=500):
    inicio = time.time()
    while time.time() - inicio < timeout:
        elementos = driver.find_elements(By.XPATH, "//div[contains(text(), 'C:\\Enova\\cache\\enovaerp')]")
        if len(elementos) == 0:
            return True
        time.sleep(2)
    return False

def verificar_importacao_mdfe(driver, timeout=500):
    mensagens_esperadas = [
        "IMPORTADO COM SUCESSO",
        "ARQUIVO CORROMPIDO",
        "MDF-E JA CADASTRADO"
    ]
    inicio = time.time()
    while time.time() - inicio < timeout:
        for mensagem in mensagens_esperadas:
            elementos = driver.find_elements(By.XPATH, f"//div[contains(text(), '{mensagem}')]")
            if elementos:
                print(f"[DEBUG] Encontrou mensagem '{mensagem}' para MDF-e")
                return True
        print("[DEBUG] Nenhuma mensagem encontrada ainda para MDF-e, aguardando...")
        time.sleep(2)
    return False


def download_cte_zip():
    temp_dir = tempfile.mktemp()
    options = Options()

    for arg in [
        "--disable-extensions",
        "--window_size=1920,1080",
        "--disable_extensions",
        "--allow-running-insecure-content",
        f"--unsafely-treat-insecure-origin-as-secure={MULTICTE_URL}",
        "--disable-gpu",
        "--no-sandbox",
        "--disable-dev-shm-usage",
        "--safebrowsing-disable-download-protection",
        "--disable-features=SafeBrowsingSafeDownload",
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36",
        "--disable-features=PasswordManagerEnabled,PasswordCheck,AutofillServerCommunication,AutofillEnableAccountWalletStorage,AutofillEnablePaymentsIdentityDetection"
    ]:
        options.add_argument(arg)

    options.add_experimental_option("prefs", {
        "download.default_directory": temp_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": False,
        "profile.password_manager_enabled": False,
        "credentials_enable_service": False,
        "profile.password_manager_leak_detection": False,
        "useAutomationExtension": False,
        "download_restrictions": 0,
        "download.open_pdf_in_system_reader": False,
        "profile.default_content_settings.popups": 0,
        "safebrowsing.disable_download_protection": True
    })

    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    try:
        driver.get(MULTICTE_URL)
        driver.maximize_window()

        usuario = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.NAME, "Usuario"))
        )
        usuario.screenshot("Abriu_funcao.png")
        usuario.send_keys(MULTICTE_USUARIO)

        senha = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.NAME, "Senha"))
        )

        senha.send_keys(MULTICTE_SENHA)

        acessar_botao = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CLASS_NAME, "btn.btn-info.btn-block.btn-lg.waves-effect.waves-themed"))
        )

        acessar_botao.click()
        time.sleep(5)

        WebDriverWait(driver, 15).until(
            EC.invisibility_of_element_located((By.CSS_SELECTOR, "div.blockUI.blockOverlay"))
        )

        relatorios_menu = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//a[span[text()='Relatórios']]"))
        )

        relatorios_menu.click()

        consulta_cte = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//a[span[text()='Consulta de CT-e']]"))
        )
        consulta_cte.click()

        input_data = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located(
                (By.CSS_SELECTOR, "input.form-control[data-bind*='DataEmissaoInicial.val']")
            )
        )
        data_ontem_str = (datetime.now() - timedelta(days=1)).strftime("%d/%m/%Y")

        driver.execute_script("""
            arguments[0].value = arguments[1];
            arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
            arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
            arguments[0].dispatchEvent(new KeyboardEvent('keydown', { bubbles: true }));
            arguments[0].dispatchEvent(new KeyboardEvent('keyup', { bubbles: true }));
        """, input_data, data_ontem_str)

        time.sleep(5)
        print("01")

        botao_pesquisar = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "button.btn.btn-primary[data-bind*='Pesquisar.eventClick']"))
        )
        botao_pesquisar.click()

        print("02")

        time.sleep(5)

        botao_download = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(
                (By.CSS_SELECTOR, "button.btn.btn-success.dropdown-toggle[data-bs-toggle='dropdown']"))
        )
        botao_download.click()

        time.sleep(1)

        botao_baixar_xml_01 = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//a[span[text()='Baixar Lote de XML']]"))
        )
        botao_baixar_xml_01.click()

        time.sleep(5)

        caminho_01 = aguardar_download(temp_dir, "LoteXML.zip")
        novo_nome_01 = os.path.join(temp_dir, "LoteXML_01.zip")
        os.rename(caminho_01, novo_nome_01)
        print(f"Arquivo 01 baixado e renomeado para: {novo_nome_01}")

        time.sleep(2)
        consulta_mdfe = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//a[span[text()='Consulta de MDF-e']]"))
        )
        consulta_mdfe.click()

        time.sleep(2)

        input_data_mdfe = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located(
                (By.CSS_SELECTOR, "input.form-control[data-bind*='DataEmissaoInicial.val']")
            )
        )
        input_data_mdfe.clear()
        input_data_mdfe.send_keys(data_ontem_str)

        botao_pesquisar_mdfe = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "button.btn.btn-primary[data-bind*='Pesquisar.eventClick']"))
        )
        botao_pesquisar_mdfe.click()

        time.sleep(2)
        botao_download02 = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(
                (By.CSS_SELECTOR, "button.btn.btn-success.dropdown-toggle[data-bs-toggle='dropdown']"))
        )
        botao_download02.click()

        botao_baixar_xml_02 = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//a[span[text()='Baixar Lote de XML']]"))
        )
        botao_baixar_xml_02.click()

        time.sleep(2)

        caminho_02 = aguardar_download(temp_dir, "LoteXML.zip")
        novo_nome_02 = os.path.join(temp_dir, "LoteXML_02.zip")
        os.rename(caminho_02, novo_nome_02)
        print(f"Arquivo 01 baixado e renomeado para: {novo_nome_02}")

        importacao_enova(driver, temp_dir)

        if verificar_importacao_cte(driver, timeout=120):
            importacao_mdfe(driver, temp_dir)

            if verificar_importacao_mdfe(driver, timeout=180):
                print("[DEBUG] MDF-e importado com sucesso.")
            else:
                print("[DEBUG] Falha ao importar MDF-e.")
        else:
            print("[DEBUG] CTe não foi importado corretamente. Ignorando MDF-e.")

        return driver, temp_dir

    except Exception as e:
        print(f"Erro {e}")
        return driver, temp_dir
    finally:
        try:
            driver.quit()
            print("[DEBUG] Driver encerrado com sucesso.")
        except Exception as e:
            print(f"[DEBUG] Erro ao tentar fechar o driver: {e}")

download_cte_zip()

schedule.every().day.at("07:00").do(download_cte_zip)
schedule.every().day.at("10:00").do(download_cte_zip)


ultimo_log = 0

while True:
    schedule.run_pending()

    agora = time.time()
    if agora - ultimo_log >= 120:
        print(f"[{datetime.now().strftime('%H:%M:%S')}] \U0001F4E7 Aguardando IMPORT/EXPORT")
        ultimo_log = agora

    time.sleep(1)
