import os
import time
import pyautogui
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Costantes
PASTA_15 = ''
MASK = ''
SETA_CARREGAR = ''
DOWNLOAD_PATH = r''
RELATORIO_PRO = r''
RELATORIO_PARA = r''
BUTTON_SUBSTITUICAO = r''
IMG_DOWN = r''


# Acessando SharePoint
def moves_to_sharepoint(driver):
    # Entrando no SharePoint
    driver.get('')
    print("Entrando no SharePoint...")
    # Carregar o site
    time.sleep(2)

    # indo para a Pasta 15
    Pasta_15 = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, PASTA_15)))
    ActionChains(driver).move_to_element(Pasta_15).click().perform()
    print("Pasta 15 acessada com sucesso")

    # tempo
    time.sleep(2)

    # Entrando na pasta da Máscara
    # test_mask = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH,'')))
    test_mask = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,MASK)))
    ActionChains(driver).move_to_element(test_mask).perform()
    driver.execute_script("arguments[0].scrollIntoView(true);", test_mask)
    test_mask.click()
    print("Pasta Mascará acessada com sucesso")

    time.sleep(2)

    # Sobre até aparecer o botão de Carregar
    scroll = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH,'')))
    driver.execute_script("arguments[0].scrollIntoView()",scroll)
    ActionChains(driver).scroll_to_element(scroll).perform()
    time.sleep(2)

    # Clicando no botão de Carregar, para enviar os arquivos
    load_files = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH,SETA_CARREGAR)))
    ActionChains(driver).move_to_element(load_files).click().perform()
    print("Botão clicado com sucesso")

    # Tempinho
    time.sleep(2)

    # Clicando no botão que irá permitir selecionar e enviar os arquivos
    ficheiro = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH,'')))
    ActionChains(driver).move_to_element(ficheiro).click().perform()
    print("Botão clicado com sucesso")

def imagens_process():
    # Função para verificar a existência de uma imagem
    def verificar_imagem(caminho_imagem):
        if not os.path.exists(caminho_imagem):
            print(f"O arquivo {caminho_imagem} nâo foi encontrado")
            return False
        return True


    # Verificar se todos os arquivos de imagem existem
    imagens = [IMG_DOWN, RELATORIO_PARA, RELATORIO_PRO, BUTTON_SUBSTITUICAO]
    for imagem in imagens:
        if not verificar_imagem(imagem):
            exit()

    time.sleep(2)

    # Tentando encontrar o botão de Downloads
    location_img = pyautogui.locateCenterOnScreen(IMG_DOWN, confidence=0.7)
    if location_img is not None:
        pyautogui.click(location_img)
        print("Pasta encontrada e clicada.")

        time.sleep(2)
    else:
        print("Donwload não encontrado.")
        exit()


    # Tentando encontrar e clicar no primerio arquivo
    location_pro = pyautogui.locateCenterOnScreen(RELATORIO_PRO, confidence=0.8)
    if location_pro is not None:
        pyautogui.click(location_pro)
        print("Primeiro arquivo clicado.")
        time.sleep(1)

        # Segurar a tecla Ctrl
        pyautogui.keyDown('ctrl')


        # Tentando encontrar e clicar no segundo arquivo
        location_para = pyautogui.locateCenterOnScreen(RELATORIO_PARA, confidence=0.8)
        if location_para is not None:
            pyautogui.click(location_para)
            print("Segundo arquivo clicado")
            time.sleep(2)

            # Soltar o Ctrl
            pyautogui.keyUp('ctrl')

            # Precionar o botão Enter
            pyautogui.press('enter')
            print("Arquivos enviados")
        else:
            print("Segundo arquivo não encontrado")
    else:
        print("Primeiro arquivo não encontrado")


# Função para começar o loop de procurar o botão
def wait_for_button(image_path, timeout=30, interval=5):
    start_time = time.time()
    while True:
        try:
            # Tenta localizar o botão na tela
            print("Tentando localizar o botão na tela...")
            location_subs = pyautogui.locateCenterOnScreen(image_path, confidence=0.7)
            if location_subs:
                print("Botão encontrado", location_subs)
                return location_subs
        except pyautogui.ImageNotFoundException:
            print("Imagem do botão não encontrada.")

        # Verifica se o tempo atingiu o timeout
        elapsed_time = time.time() - start_time
        if elapsed_time > timeout:
            print("Tempo limite excedido.")
            return None
        
        # Espera um pouco antes de tentar novamente
        print(f"Aguardando {interval} segundos antes de tentar novamente...")
        time.sleep(interval)


def delete_files():
    download_folder = os.path.expanduser(DOWNLOAD_PATH)
    files_to_delete = ['RELATÓRIO_PARADAS.csv', 'RELATÓRIO_PRODUÇÃO.csv']

    for file_name in files_to_delete:
        file_path = os.path.join(download_folder, file_name)
        try:
            os.remove(file_path)
            print(f"Arquivo {file_name} deletado com sucesso.")
        except FileNotFoundError:
            print(f"Arquivo {file_name} não encontrado")
        except Exception as e:
            print(f"Ocorreu um erro ao tentar deletar {file_name}: {e}")