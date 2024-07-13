from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, StaleElementReferenceException, TimeoutException, ElementClickInterceptedException
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
import time
import openpyxl
from datetime import datetime, timedelta
import pyautogui as pa

caminho_planilha = r'V:\\Automatização Union\\BACKOFFICE\\Planilhas - backoffice\\domicilio.xlsx'
pasta_download = r'G:\\Drives compartilhados\\Bienio_Corrente_2024\\2024\\EXTRATO - EFISCO\\07 - JULHO'  # Defina seu diretório de download aqui

try:
    workbook = openpyxl.load_workbook(caminho_planilha)
    extrato_sheet = workbook['Plan1']
    time.sleep(5)
except FileNotFoundError:
    print(f"Arquivo não encontrado: {caminho_planilha}")
    exit(1)

# Configurações do Chrome para depuração e download automático
options = webdriver.ChromeOptions()
options.add_argument("--disable-extensions")
options.add_argument("--disable-popup-blocking")
options.add_argument("--start-maximized")
options.add_experimental_option('prefs', {
    "download.default_directory": pasta_download,  # Muda o diretório de download
    "download.prompt_for_download": False,  # Desativa a pergunta de download
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
})

# Inicia o serviço do ChromeDriver
service = Service(ChromeDriverManager().install())
try:
    navegador = webdriver.Chrome(service=service, options=options)
    # Acessa as configurações de download do Chrome
    navegador.get("https://efisco.sefaz.pe.gov.br/sfi_com_sca/PRMontarMenuAcesso")
    wait = WebDriverWait(navegador, 10)
    
    # Espera o botão de certificado estar clicável
    certificado = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="btt_certificado"]/span')))
    certificado.click()

    # Tempo para selecionar o certificado digital
    time.sleep(20)

    # Espera o menu tributário estar clicável e visível
    tributario = wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="nav_topo"]/li[1]/a')))
    navegador.execute_script("arguments[0].scrollIntoView(true);", tributario)
    time.sleep(2)  # Espera 2 segundos após rolar para garantir que o elemento está visível
    tributario.click()

    # Espera o link de pagamento estar clicável e visível
    pgto = wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="100001"]')))
    navegador.execute_script("arguments[0].scrollIntoView(true);", pgto)
    time.sleep(2)  # Espera 2 segundos após rolar para garantir que o elemento está visível
    pgto.click()

    # Espera o link do extrato estar clicável e visível
    extrato = wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="fmw_id_sidebar_142031"]/span')))
    navegador.execute_script("arguments[0].scrollIntoView(true);", extrato)
    time.sleep(2)  # Espera 2 segundos após rolar para garantir que o elemento está visível
    extrato.click()

    for linha in extrato_sheet.iter_rows(min_row=4):
        # Espera o campo radical CNPJ estar clicável e visível
        radical_cnpj = wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="NuRadicalCNPJ"]')))
        navegador.execute_script("arguments[0].scrollIntoView(true);", radical_cnpj)
        time.sleep(2)  # Espera 2 segundos após rolar para garantir que o elemento está visível
        radical_cnpj.click()
        radical_cnpj.send_keys(linha[4].value)

        # Espera o botão localizar estar clicável e visível
        localizar = wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="btt_localizar"]')))
        navegador.execute_script("arguments[0].scrollIntoView(true);", localizar)
        time.sleep(2)  # Espera 2 segundos após rolar para garantir que o elemento está visível
        localizar.click()

        pa.PAUSE = 5
        # Espera a página carregar os resultados
        time.sleep(5)

        # Simula o comando Ctrl + P para abrir a janela de impressão
        pa.hotkey('ctrl', 'p')
        time.sleep(5)

        # Adiciona um tempo de espera extra antes de começar a digitar
        pa.PAUSE = 3.5

        # Escreve o nome do arquivo e diretório
        pa.write(str(linha[3].value))
        time.sleep(3.5)
        pa.click(x=754, y=226)
        time.sleep(3.5)
        pa.write(pasta_download)
        time.sleep(3.5)

        # Enviar backspace 8 vezes
        radical_cnpj.send_keys(Keys.BACK_SPACE * 8)

        # Simula o Enter para confirmar a impressão
        pa.press('enter')
        time.sleep(3.5)

        pa.press('enter')
        time.sleep(3.5)

        # Espera um tempo para garantir que o print foi salvo
        time.sleep(10)

except ElementClickInterceptedException as e:
    print(f"Erro ao clicar no elemento: {e}")
except Exception as e:
    print(f"Erro ao iniciar o navegador: {e}")
finally:
    if navegador:
        navegador.quit()
