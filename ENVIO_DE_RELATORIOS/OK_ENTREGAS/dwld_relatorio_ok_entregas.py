import schedule
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support import expected_conditions as EC
import time
import tempfile
import os
import datetime
from tratamento_xlsx import tratar_e_formatar_arquivo
from Mondial_S_Faturas.envio_wpp_arquivo import send_whatsapp_file
from cred import OK_ENTREGA_URL, OK_ENTREGA_USUARIO, OK_ENTREGA_SENHA

def esperar_download_concluir(diretorio, timeout=60):
    for _ in range(timeout):
        arquivos = [f for f in os.listdir(diretorio) if f.endswith(".xlsx")]
        if arquivos:
            return arquivos[0]
        time.sleep(1)
    return None

def segundo_acesso_():
    temp_dir = tempfile.mkdtemp()

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
        "download.default_directory": temp_dir,
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
        time.sleep(5)

        close_btn = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "icon_close"))
        )
        close_btn.click()

        WebDriverWait(driver, 10).until(
            EC.invisibility_of_element_located((By.ID, "alertModalAviso"))
        )

        max_tentativas = 3
        tentativas = 0
        baixado = False

        while tentativas < max_tentativas and not baixado:
            try:
                print("01 - Aguardando tabela aparecer")
                tabela = WebDriverWait(driver, 15).until(
                    EC.presence_of_element_located((By.ID, "listaSolicitacoes"))
                )
                print("02 - Tabela localizada")

                primeira_linha = tabela.find_element(By.CSS_SELECTOR, "tbody tr:first-child")
                status_div = primeira_linha.find_element(By.CSS_SELECTOR, "td:nth-child(2) > div > span")
                print("03 - Status localizado")

                status = status_div.text.strip()
                print(f"Tentativa {tentativas + 1} - Status:", status)
                print("04 - Avaliando status")

                if status.lower() == "executado":
                    try:
                        print("05 - Aguardando link clicável")
                        link_element = WebDriverWait(primeira_linha, 10).until(
                            EC.element_to_be_clickable((By.CSS_SELECTOR, "td:nth-child(7) a"))
                        )
                        href = link_element.get_attribute("href")
                        print("Link encontrado:", href)

                        try:
                            link_element.click()
                        except:
                            driver.execute_script("arguments[0].click();", link_element)

                        print("Cliquei para baixar o arquivo.")

                        nome_arquivo = esperar_download_concluir(temp_dir, timeout=60)
                        if nome_arquivo:
                            print("Arquivo baixado:", nome_arquivo)

                            try:
                                tratar_e_formatar_arquivo(temp_dir, nome_arquivo)
                                print("Arquivo tratado e salvo com sucesso.")

                                caminho_final_temp = os.path.join(temp_dir, "ACOMPANHAMENTO OK ENTREGAS.xlsx")
                                contatos = ["Yasmin", "RELATÓRIO OK ENTREGA - JTD TRANSPORTES"]
                                send_whatsapp_file(caminho_final_temp, contatos)

                                baixado = True

                            except Exception as e:
                                print("Erro ao tratar e salvar o arquivo:", e)
                        else:
                            print("O download não foi concluído dentro do tempo limite.")
                    except Exception as e:
                        print("Erro ao tentar clicar no link de download:", e)
                else:
                    tentativas += 1
                    if tentativas < max_tentativas:
                        print("Ainda não está 'Executado'. Aguardando 10 minutos para tentar novamente...")
                        time.sleep(600)

                        driver.refresh()
                        time.sleep(5)

                        try:
                            ok_btn = WebDriverWait(driver, 10).until(
                                EC.element_to_be_clickable((By.XPATH, '//button[normalize-space(text())="Ok"]'))
                            )
                            ok_btn.click()

                            WebDriverWait(driver, 10).until(
                                EC.invisibility_of_element_located((By.ID, "alertModalAviso"))
                            )
                            print("Modal de aviso fechado com sucesso.")
                        except Exception as e:
                            print("Modal de aviso não apareceu ou já estava fechado:", e)

                        driver.execute_script("window.scrollBy(0, 800)")
                    else:
                        print("Número máximo de tentativas atingido. Encerrando.")
            except Exception as e:
                print("Erro geral no laço:", e)
                tentativas += 1
                time.sleep(5)

    except Exception as e:
        print(f"Erro durante a execução: {e}")
    finally:
        time.sleep(5)
        driver.quit()
        print("Driver encerrado.")


segundo_acesso_()

schedule.every().day.at("07:55").do(segundo_acesso_)

ultimo_log = 0

while True:
    schedule.run_pending()

    agora = time.time()
    if agora - ultimo_log >= 120:
        print(f"[{datetime.datetime.now().strftime('%H:%M:%S')}] \U0001F4E5 Aguardando DOWNLOAD")
        ultimo_log = agora

    time.sleep(1)

