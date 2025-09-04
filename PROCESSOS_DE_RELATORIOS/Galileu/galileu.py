import tempfile
import time
import pyautogui
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from cred import GALILEU_SENHA, GALILEU_URL, GALILEU_USUARIO

def abrir_site_em_chrome_com_selenium():
    temp_dir = tempfile.mkdtemp()
    options = Options()
    for arg in [
        r"user-data-dir=C:\\ChromeSelenium",
        "profile-directory=Profile1",
        f"--unsafely-treat-insecure-origin-as-secure={GALILEU_URL}",
        "--start-maximized",
        "--disable-extensions",
        "--disable-gpu",
        "--no-sandbox",
        "--disable-dev-shm-usage",
        "--disable-features=ClipboardCodeExecution"
    ]:
        options.add_argument(arg)

    prefs = {
        "profile.content_settings.exceptions.clipboard": {
            "https://nestle-br.galileulog.com.br:443,*": {
                "setting": 1,
                "last_modified": int(time.time() * 1000),
                "expiration": 0
            }
        }
    }
    options.add_experimental_option("prefs", prefs)

    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    driver.get(GALILEU_URL)
    return driver


def clicar_img(imagem, confidence=0.8, region=None, delay=0.3, clicks=1):
    pos = pyautogui.locateCenterOnScreen(imagem, confidence=confidence, region=region)
    if pos:
        time.sleep(delay)
        pyautogui.click(pos, clicks=clicks)
        return True
    return False


driver = abrir_site_em_chrome_com_selenium()
time.sleep(10)

if clicar_img('usuario_field.png'):
    pyautogui.write(GALILEU_USUARIO, interval=0.1)
    pyautogui.press('tab')
    pyautogui.write(GALILEU_SENHA, interval=0.1)
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('enter')

time.sleep(5)

clicar_img('gerenciamento.png')
time.sleep(8)

if clicar_img('pesquisar.png'):
    time.sleep(1)
    clicar_img('pesquisar.png')

time.sleep(5)

if clicar_img('status_filtro.png'):
    time.sleep(1)
    clicar_img('status_filtro.png')
time.sleep(5)

menu_filtro = pyautogui.locateOnScreen('menu_filtro.png', confidence=0.8)
if menu_filtro:
    x, y, w, h = menu_filtro
    regiao_filtro = (x, y, w, h)

    if not clicar_img('em_viagem.png', region=regiao_filtro):
        print("Filtro 'EM VIAGEM' não encontrado dentro do menu.")

time.sleep(1)

if clicar_img('concluir_botao.png'):
    time.sleep(0.5)
    clicar_img('concluir_botao.png')

time.sleep(5)

exportar_area = pyautogui.locateOnScreen('exportar_area.png', confidence=0.6)
print("entrando aqui")
if exportar_area:
    x, y, w, h = exportar_area
    regiao_filtro_area = (x, y, w, h)

    pos = pyautogui.locateCenterOnScreen('exportar_botao_02.png', confidence=0.8, region=regiao_filtro_area)
    if pos:
        print("Posição encontrada:", pos)
        pyautogui.moveTo(pos)
        time.sleep(1)
        print("Mouse agora está em:", pyautogui.position())
        pyautogui.click()
    else:
        print("Botão EXTRair não encontrado dentro da área EXPORTAR!")
else:
    print("Área EXPORTAR não localizada!")

time.sleep(1)

confirmar_botao = pyautogui.locateOnScreen("confirmar_botao.png", confidence=0.8)
if confirmar_botao:
    pyautogui.click(confirmar_botao, clicks=2)

time.sleep(20)
