import os
import time
from datetime import datetime
from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.edge.options import Options as EdgeOptions
from webdriver_manager.microsoft import EdgeChromiumDriverManager # type: ignore
from selenium.common.exceptions import NoSuchElementException , TimeoutException


# Constantes
USERNAME =  ''
PASSWORD = ''
INDUSTRIA = ''
SAP_MII = ''
FIRST_DAY_XPATH = ''
DATE_INPUT_XPATH = ''
EXECUTE_BUTTON_ID = ''
DOWNLOAD_PATH = r''

def login_to_sgr(driver):
    # Entrando no SGR
    driver.get('')
    time.sleep(5)
    # Login
    driver.find_element(By.NAME,'username').send_keys(USERNAME)
    driver.find_element(By.NAME,'password').send_keys(PASSWORD)
    driver.find_element(By.NAME, 'sSubmit').click()
    print("Login realizado com sucesso!")
    time.sleep(10)

# Navegando na Produção e Paradas
# Produção
def navigate_to_production_mii_oee(driver):
    time.sleep(5)
    # Ir até a Indústria 
    industria_menu = driver.find_element(By.XPATH, INDUSTRIA)
    ActionChains(driver).move_to_element(industria_menu).click().perform()
    time.sleep(1)

    # Ir até o SAP MII
    sap_mii_menu = driver.find_element(By.XPATH, SAP_MII)
    ActionChains(driver).move_to_element(sap_mii_menu).click().perform()
    time.sleep(1)

    # Produção MII-OEE
    producao_menu = driver.find_element(By.XPATH,'')
    ActionChains(driver).move_to_element(producao_menu).click().perform()
    
    print("Navegação até a Produção concluida")
    
    time.sleep(15)


# Paradas
def navigate_to_stops_mii_oee(driver):
    # Ir até a Indústria
    industria_menu = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, INDUSTRIA)))
    ActionChains(driver).move_to_element(industria_menu).click().perform()
    time.sleep(2)

    # Ie até o SAP MII
    sap_mii_menu = driver.find_element(By.XPATH,SAP_MII)
    ActionChains(driver).move_to_element(sap_mii_menu).click().perform()
    time.sleep(2)

    # Paradas MII-OEE
    paradas_menu = driver.find_element(By.XPATH,'')
    ActionChains(driver).move_to_element(paradas_menu).click().perform()
    
    print("Navegação até a Paradas concluida")

    time.sleep(15)

# Selecionando Unidade, Data de Início e Data atual
def select_date_range_production(driver):
    driver.find_element(By.XPATH,'').click()
    time.sleep(1)

    WebDriverWait(driver, 5).until(
        EC.visibility_of_element_located((By.XPATH, ''))
    ).find_element(By.ID, "item_element_11").click()

    
    data_atual = datetime.now().strftime("01/%m/%Y")
    date_input = driver.find_element(By.XPATH, FIRST_DAY_XPATH)
    date_input.clear()
    date_input.send_keys(data_atual)

    
    time.sleep(2)

    data_atual = datetime.now().strftime("%d/%m/%Y")
    date_input = driver.find_element(By.XPATH, DATE_INPUT_XPATH)
    date_input.clear()
    date_input.send_keys(data_atual)

    time.sleep(1)

    # Elemento
    driver.find_element(By.XPATH,'').click()
    time.sleep(2)

    # Localizando elemento "TODOS"
    WebDriverWait(driver, 5).until(
        EC.visibility_of_element_located((By.XPATH, ''))
    ).find_element(By.ID, 'item_element_0').click()

    # Recolhendo dropdown
    driver.find_element(By.XPATH,'').click()
    time.sleep(2)

    print("Unidade,Data de início e Data atual selecionados com sucesso!")
    
    time.sleep(2)
    driver.find_element(By.ID, EXECUTE_BUTTON_ID).click()
    print("Botão de executar selecionado com sucesso!")
    time.sleep(20)

# Selecionando Unidade, Data de Início e Data atual
def select_date_range_stops(driver):
    driver.find_element(By.XPATH,'').click()
    time.sleep(1)

    WebDriverWait(driver, 5).until(
        EC.visibility_of_element_located((By.XPATH, ''))
    ).find_element(By.ID, "item_element_11").click()

    
    data_atual = datetime.now().strftime("01/%m/%Y")
    date_input = driver.find_element(By.XPATH, '')
    date_input.clear()
    date_input.send_keys(data_atual)

    
    time.sleep(2)

    data_atual = datetime.now().strftime("%d/%m/%Y")
    date_input = driver.find_element(By.XPATH, '')
    date_input.clear()
    date_input.send_keys(data_atual)

    print("Unidade,Data de início e Data atual selecionados com sucesso!")
    
    time.sleep(2)
    driver.find_element(By.ID, EXECUTE_BUTTON_ID).click()
    print("Botão de executar selecionado com sucesso!")
    time.sleep(20)

# Tela de Consultas
def navigate_to_consults(driver):
    # Navegar para a tela de Consultas
    driver.find_element(By.XPATH,'').click()
    time.sleep(2)

    driver.find_element(By.XPATH,'').click()
    time.sleep(2)
    print("Navegação para Consultas realizada com sucesso!")


# Funções para o encontrar o Download e fazer a Renomeação
def find_downloaded_file(download_path, codigo):
    files = [os.path.join(download_path, f) for f in os.listdir(download_path) if codigo in f]
    if files:
        return max(files, key=os.path.getctime)
    else:
        return None
    

def rename_file(old_name, new_name):
    try:
        os.rename(old_name, new_name)
        print(f"Arquivo renomeado de '{old_name}' para '{new_name}' com sucesso!")
    except Exception as e:
        print(f"Erro ao renomear o arquivo: {e}")


# Espera downloads estar concluído
def wait_for_downloads(download_dir, timeout=300, check_interval=30):
  
    start_time = time.time()
    
    while True:
        # Lista de arquivos na pasta de downloads
        files = os.listdir(download_dir)
        
        # Verifica se há arquivos de download em andamento (ex: arquivos com extensão .crdownload, .part)
        downloading_files = [f for f in files if f.endswith('.crdownload') or f.endswith('.part')]
        
        if not downloading_files:
            print("Todos os downloads foram concluídos.")
            return True
        
        # Verifica se o tempo limite foi atingido
        elapsed_time = time.time() - start_time
        if elapsed_time > timeout:
            print("Tempo limite atingido. Alguns downloads ainda não foram concluídos.")
            return False
        
        # Espera antes de verificar novamente
        time.sleep(check_interval)


# Produção
# Função Paradas
def producao(driver, max_retries=3, current_retry=0):
    # Navegar para a página de Produção, selecionar data e abrir consultas
    navigate_to_production_mii_oee(driver)
    select_date_range_production(driver)
    navigate_to_consults(driver)

    wait = WebDriverWait(driver, 80)
    codigo_encontrado = None

    try:
        # Data e hora atuais
        data_atual = datetime.now()

        # Aguardar a tabela carregar e localizar todas as linhas (tr) dentro de tbody
        row = wait.until(EC.presence_of_element_located((By.XPATH, "")))
        linhas = row.find_elements(By.XPATH, "")

        # Primeira Verificação: Procurar uma linha com status 'À Iniciar'
        for i, linha in enumerate(linhas, start=1):
            descricao_curta = linha.find_element(By.XPATH, "").text
            status = linha.find_element(By.XPATH, "").text
            
            # Se encontrar 'À Iniciar' com a descrição correta, salva o código e sai do loop
            if descricao_curta == '' and status == '' or status == "":
                codigo_encontrado = linha.find_element(By.XPATH, "").text
                print(f"Encontramos na linha {i} com status ''. Código da linha: {codigo_encontrado}")
                break

        # Se o código foi encontrado na primeira verificação, tenta fazer o download
        if codigo_encontrado:
            while current_retry < max_retries:
                try:
                    bnt = wait.until(EC.element_to_be_clickable((By.ID, f'download_{codigo_encontrado}')))
                    time.sleep(2)
                    print("Botão de download encontrado, fazendo o download...")
                    bnt.click()
                    return codigo_encontrado  # Finaliza a função após o download
                except TimeoutException:
                    print(f"O botão de download 'download_{codigo_encontrado}' não foi encontrado. Tentando novamente...")

                # Verifica se existe um botão de erro para reiniciar o processo
                try:
                    bnt_erro = wait.until(EC.element_to_be_clickable((By.ID, f'alert_{codigo_encontrado}')))
                    print(f"Botão de erro encontrado. Refazendo a operação para {descricao_curta}")
                    return producao(driver, max_retries, current_retry + 1)  # Reinicia a função aumentando o contador de tentativas
                except TimeoutException:
                    pass  # Continua tentando até atingir max_retries

            print("Número máximo de tentativas alcançado. Parando a operação.")
            return None

        # Segunda Verificação: Caso não tenha encontrado 'À Iniciar', verifica a data e hora
        for i, linha in enumerate(linhas, start=1):
            descricao_curta = linha.find_element(By.XPATH, "").text
            data_texto = linha.find_element(By.XPATH, "").text

            # Converte a string de data para datetime
            try:
                datas = datetime.strptime(data_texto, "%d/%m/%Y %H:%M:%S")
            except ValueError:
                print(f"Formato de data inválido na linha {i}: {data_texto}")
                continue

            # Se encontrar uma linha com a data e hora atuais e a descrição correta, salva o código
            if descricao_curta == '' and datas.date() == data_atual.date() and datas.hour == data_atual.hour:
                codigo_encontrado = linha.find_element(By.XPATH, "").text
                print(f"Encontramos na linha {i} com a data/hora atuais. Código da linha: {codigo_encontrado}")

                # Tentar fazer o download
                for attempt in range(max_retries):
                    try:
                        bnt = wait.until(EC.element_to_be_clickable((By.ID, f'download_{codigo_encontrado}')))
                        time.sleep(2)
                        print("Botão encontrado, fazendo o download...")
                        bnt.click()
                        return codigo_encontrado  # Finaliza a função após o download
                    except TimeoutException:
                        print(f"O botão de download 'download_{codigo_encontrado}' não foi encontrado. Tentando novamente...")
                        time.sleep(1)

                    # Verifica se existe um botão de erro
                    try:
                        bnt_erro = wait.until(EC.element_to_be_clickable((By.ID, f'alert_{codigo_encontrado}')))
                        print(f"Botão de erro encontrado. Tentativa de refazer a operação {descricao_curta}")
                        if current_retry < max_retries:
                            return producao(driver, max_retries, current_retry + 1)  # Reinicia a função com nova tentativa
                    except TimeoutException:
                        pass

            print(f"Não encontramos o item desejado na linha {i}. Indo para a próxima linha.")
        
        print("Nenhum item foi encontrado com os critérios estabelecidos.")
        return None

    except Exception as e:
        print(f"Ocorreu um erro: {e}")
        return None


# Paradas
def paradas(driver, max_retries=3, current_retry=0):
    # Navegar para a página de Produção, selecionar data e abrir consultas
    navigate_to_stops_mii_oee(driver)
    select_date_range_stops(driver)
    navigate_to_consults(driver)

    wait = WebDriverWait(driver, 80)
    codigo_encontrado = None

    try:
        # Data e hora atuais
        data_atual = datetime.now()

        # Aguardar a tabela carregar e localizar todas as linhas (tr) dentro de tbody
        row = wait.until(EC.presence_of_element_located((By.XPATH, "")))
        linhas = row.find_elements(By.XPATH, "")

        # Primeira Verificação: Procurar uma linha com status 'À Iniciar'
        for i, linha in enumerate(linhas, start=1):
            descricao_curta = linha.find_element(By.XPATH, "").text
            status = linha.find_element(By.XPATH, "").text
            
            # Se encontrar 'À Iniciar' com a descrição correta, salva o código e sai do loop
            if descricao_curta == '' and status == '':
                codigo_encontrado = linha.find_element(By.XPATH, "").text
                print(f"Encontramos na linha {i} com status ''. Código da linha: {codigo_encontrado}")
                break

        # Se o código foi encontrado na primeira verificação, tenta fazer o download
        if codigo_encontrado:
            while current_retry < max_retries:
                try:
                    bnt = wait.until(EC.element_to_be_clickable((By.ID, f'download_{codigo_encontrado}')))
                    time.sleep(2)
                    print("Botão de download encontrado, fazendo o download...")
                    bnt.click()
                    return codigo_encontrado  # Finaliza a função após o download
                except TimeoutException:
                    print(f"O botão de download 'download_{codigo_encontrado}' não foi encontrado. Tentando novamente...")

                # Verifica se existe um botão de erro para reiniciar o processo
                try:
                    bnt_erro = wait.until(EC.element_to_be_clickable((By.ID, f'alert_{codigo_encontrado}')))
                    print(f"Botão de erro encontrado. Refazendo a operação para {descricao_curta}")
                    return paradas(driver, max_retries, current_retry + 1)  # Reinicia a função aumentando o contador de tentativas
                except TimeoutException:
                    pass  # Continua tentando até atingir max_retries

            print("Número máximo de tentativas alcançado. Parando a operação.")
            return None

        # Segunda Verificação: Caso não tenha encontrado 'À Iniciar', verifica a data e hora
        for i, linha in enumerate(linhas, start=1):
            descricao_curta = linha.find_element(By.XPATH, "").text
            data_texto = linha.find_element(By.XPATH, "").text

            # Converte a string de data para datetime
            try:
                datas = datetime.strptime(data_texto, "%d/%m/%Y %H:%M:%S")
            except ValueError:
                print(f"Formato de data inválido na linha {i}: {data_texto}")
                continue

            # Se encontrar uma linha com a data e hora atuais e a descrição correta, salva o código
            if descricao_curta == '' and datas.date() == data_atual.date() and datas.hour == data_atual.hour:
                codigo_encontrado = linha.find_element(By.XPATH, ".//td[1]").text
                print(f"Encontramos na linha {i} com a data/hora atuais. Código da linha: {codigo_encontrado}")

                # Tentar fazer o download
                for attempt in range(max_retries):
                    try:
                        bnt = wait.until(EC.element_to_be_clickable((By.ID, f'download_{codigo_encontrado}')))
                        time.sleep(2)
                        print("Botão encontrado, fazendo o download...")
                        bnt.click()
                        return codigo_encontrado  # Finaliza a função após o download
                    except TimeoutException:
                        print(f"O botão de download 'download_{codigo_encontrado}' não foi encontrado. Tentando novamente...")
                        time.sleep(1)

                    # Verifica se existe um botão de erro
                    try:
                        bnt_erro = wait.until(EC.element_to_be_clickable((By.ID, f'alert_{codigo_encontrado}')))
                        print(f"Botão de erro encontrado. Tentativa de refazer a operação {descricao_curta}")
                        if current_retry < max_retries:
                            return paradas(driver, max_retries, current_retry + 1)  # Reinicia a função com nova tentativa
                    except TimeoutException:
                        pass

            print(f"Não encontramos o item desejado na linha {i}. Indo para a próxima linha.")
        
        print("Nenhum item foi encontrado com os critérios estabelecidos.")
        return None

    except Exception as e:
        print(f"Ocorreu um erro: {e}")
        return None

# Corpo
def body(driver):
    try: 
        login_to_sgr(driver)

        wait = WebDriverWait(driver, 80)
        while True:
            # Paradas
            codpara = paradas(driver)
           
            time.sleep(2)

            # Produção
            codprod = producao(driver)

            # Espera os downloads serem feitos
            downloads_concluidos = wait_for_downloads(DOWNLOAD_PATH)
            if downloads_concluidos:
                print("Downloads concluídos, continue com o próximo passo.")
            else:
                print("Alguns donwloads não foram concluídos dentro do tempo limite")

            time.sleep(10)
            # Renomeação dos arquivos
            old_file_path1 = find_downloaded_file(DOWNLOAD_PATH, codprod)
            if old_file_path1:
                new_file_path1 = os.path.join(DOWNLOAD_PATH, f'.csv')
                rename_file(old_file_path1, new_file_path1)
            time.sleep(2)

            old_file_path2 = find_downloaded_file(DOWNLOAD_PATH, codpara)
            if old_file_path2:
                new_file_path2 = os.path.join(DOWNLOAD_PATH, f'.csv')
                rename_file(old_file_path2, new_file_path2)
                
            time.sleep(2)
            break

        
    except Exception as e:
        print(f"Erro durante a execução: {e}")    
    
    finally:
        driver.quit()
        print("------------------------------SGR finalizada!------------------------------")