# Import´s
import os
import re
import time
import win32com
import win32gui
import pyautogui
import win32com.client
import win32com.client as win32



# Constantes
# Macro
TRANSFERIR = r''    
POP_TRANSF = r''
OK = r''
OK_TRASNSF = r''
CP1 = r''
MC1 = r''
MS1 = r''
# Atualizando e Salvando
def refreshing_save():
    excel = win32com.client.Dispatch("Excel.Application")
    time.sleep(2)
    workbook = excel.ActiveWorkbook

    # Atualiza
    workbook.RefreshAll()
    print("Atualizando...")

    # Aguarda a conclusão da atualização assícronas
    excel.CalculateUntilAsyncQueriesDone()

    # Salva o Excel atual
    workbook.Save()

    print(f"Atualização concluída e feito o save do arquivo {workbook.Name}")


def save():
    excel = win32com.client.Dispatch("Excel.Application")
    time.sleep(2)
    workbook = excel.ActiveWorkbook

    # Salva o Excel atual
    workbook.Save()

    print(f"Atualização concluída e feito o save do arquivo {workbook.Name}")



# Minimizando Excel
def mini_excel():
    # Conectando ao Excel
    excel = win32com.client.Dispatch("Excel.Application")

    # Obtem a janela principal do excel
    hwnd = excel.Hwnd

    # Minimizar a tela
    win32gui.ShowWindow(hwnd, 6)

    print("A janela do Excel foi minimizada")

# Formula que fecha todos os Excels
def close_existing_excel_instances():
    try:
        excel_app = win32com.client.Dispatch("Excel.Application")
        for wb in excel_app.Workbooks:
            wb.Close(SaveChanges=True)
        excel_app.Quit()
    except Exception as e:
        print(f"Ocorreu um erro: {e}")


# MRC
def mascara_cpp(file_path, max_attempts = 3):
    def get_excel_app():
        try:
            excel = win32com.client.GetActiveObject("Excel.Application")
            print("Usando uma instância existente do Excel.")
        except Exception as e:
            print(f"Erro ao conectar a uma instância existente: {e}")
            try:
                excel = win32com.client.Dispatch("Excel.Application")
                print("Criando uma nova instância do Excel.")
            except Exception as e:
                print(f"Erro ao criar uma nova instância do Excel: {e}")
                return None
        return excel

    if not os.path.isfile(file_path):
        print(f"O arquivo '{file_path}' não foi encontrado.")
        return
    
    excel = get_excel_app()
    if excel is None:
        print("Não foi possével inicializar o Excel.")
        return
    time.sleep(2)

    try:
        excel.Visible = True
    except Exception as e:
        print(f"Erro ao definir a propriedade Visible: {e}")

    attempt = 0
    workbook = None
    
    while attempt < max_attempts:
        try:
            workbook = excel.Workbooks.Open(file_path)
            print(f"Arquivo '{workbook.Name}' aberto com sucesso na tentativa {attempt + 1}.")
            # Atualiza
            workbook.RefreshAll()
            print("Atualizando...")

            # Aguarda a conclusão da atualização assícronas
            excel.CalculateUntilAsyncQueriesDone()

            # Salva o Excel atual
            workbook.Save()

            print(f"Atualização concluída e feito o save do arquivo {workbook.Name}")
            time.sleep(2)

            hwnd = excel.Hwnd

            # Minimizar a tela
            win32gui.ShowWindow(hwnd, 6)   

            
            break

        except Exception as e:
            print(f"Erro ao abrir o arquivo Excel na tentativa {attempt + 1}: {e}")

            attempt += 1
            time.sleep(2)


    if workbook:
        print(f"Operação fincalizada com sucesso no arquivo '{workbook.Name}'")
    else:
        print("Falha ao abrir o arquivo após várias tentativas")



# MACRO
def macro(file_path, max_attempts=3):
    def get_excel_app():
        try:
            excel = win32com.client.GetActiveObject("Excel.Application")
            print("Usando uma instância existente do Excel.")
        except Exception as e:
            print(f"Erro ao conectar a uma instância existente: {e}")
            try:
                excel = win32com.client.Dispatch("Excel.Application")
                print("Criando uma nova instância do Excel.")
            except Exception as e:
                print(f"Erro ao criar uma nova instância do Excel: {e}")
                return None
        return excel

    def close_workbook_if_open(workbook):
        if workbook is not None:
            try:
                workbook.Close(SaveChanges=False)
                print(f"Workbook '{workbook.Name}' fechado.")
            except Exception as e:
                print(f"Erro ao fechar o workbook: {e}")

    if not os.path.isfile(file_path):
        print(f"O arquivo '{file_path}' não foi encontrado.")
        return
    
    excel = get_excel_app()
    if excel is None:
        print("Não foi possível inicializar o Excel.")
        return
    time.sleep(2)

    try:
        excel.Visible = True
    except Exception as e:
        print(f"Erro ao definir a propriedade Visible: {e}")

    attempt = 0
    workbook = None
    
    while attempt < max_attempts:
        try:
            workbook = excel.Workbooks.Open(file_path)
            print(f"Arquivo '{workbook.Name}' aberto com sucesso na tentativa {attempt + 1}.")
            
            # Acessa a aba "Config"
            try:
                sheet = workbook.Sheets("Config")
                sheet.Activate()
                print("Aba 'Config' ativada com sucesso.")
            except Exception as e:
                print(f"Erro ao acessar a aba 'Config': {e}")
                close_workbook_if_open(workbook)
                workbook = None
                attempt += 1
                time.sleep(2)
                continue  # Recomeça o loop após fechar o workbook
            
            tranferir()
            pop_transf_ok()
            workbook.Save()
            close_existing_excel_instances()
            time.sleep(2)
            break

        except Exception as e:
            print(f"Erro ao abrir o arquivo Excel na tentativa {attempt + 1}: {e}")
            close_workbook_if_open(workbook)
            workbook = None
            attempt += 1
            time.sleep(2)


# Transferindo dados
def tranferir():
    bnt_tra = pyautogui.locateCenterOnScreen(TRANSFERIR, confidence=0.9)
    if bnt_tra is not None:
        pyautogui.doubleClick(bnt_tra)
        print("Botão clicado com sucesso.")
        time.sleep(1)

    else:
        print("Botão não clicado")


# Espera o pop-up aparecer
def esperando_pop_trans(image_path, timeout=30, interval=5):
    start_time = time.time()
    while True:
        try:
            # Tenta localizar o botão na tela
            print("Tentando localizar o botão na tela...")
            location_pop = pyautogui.locateCenterOnScreen(image_path, confidence=0.7)
            if location_pop:
                print("Botão encontrado", location_pop)
                return location_pop
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
        

# Vai até o pop-up e clica no botão 'ok'
# teste 
def pop_transf_ok():
    # Fórmula que faz loop até encontrar o pop-up de transferência
    location_pop = esperando_pop_trans(POP_TRANSF, timeout=3600, interval=20)

    if location_pop:
            print("Pop-up de transferência apareceu")
            pyautogui.moveTo(location_pop)
            time.sleep(2)
            
    else:
        print("Pop-up não apareceu na tela")

    time.sleep(1)

    # Clicar no botão
    location_ok = pyautogui.locateCenterOnScreen(OK, confidence=0.8)
    if location_ok is not None:
        pyautogui.click(location_ok)
        print("Botão clicado com sucesso.")
    else:
        print("Botão não localizado.")

    time.sleep(5)

def mascara_cpp(file_path, max_attempts = 3):
    def get_excel_app():
        try:
            excel = win32com.client.GetActiveObject("Excel.Application")
            print("Usando uma instância existente do Excel.")
        except Exception as e:
            print(f"Erro ao conectar a uma instância existente: {e}")
            try:
                excel = win32com.client.Dispatch("Excel.Application")
                print("Criando uma nova instância do Excel.")
            except Exception as e:
                print(f"Erro ao criar uma nova instância do Excel: {e}")
                return None
        return excel

    if not os.path.isfile(file_path):
        print(f"O arquivo '{file_path}' não foi encontrado.")
        return
    
    excel = get_excel_app()
    if excel is None:
        print("Não foi possével inicializar o Excel.")
        return
    time.sleep(2)

    try:
        excel.Visible = True
    except Exception as e:
        print(f"Erro ao definir a propriedade Visible: {e}")

    attempt = 0
    workbook = None
    
    while attempt < max_attempts:
        try:
            workbook = excel.Workbooks.Open(file_path)
            print(f"Arquivo '{workbook.Name}' aberto com sucesso na tentativa {attempt + 1}.")
            # Atualiza
            workbook.RefreshAll()
            print("Atualizando...")

            # Aguarda a conclusão da atualização assícronas
            excel.CalculateUntilAsyncQueriesDone()

            # Salva o Excel atual
            workbook.Save()

            print(f"Atualização concluída e feito o save do arquivo {workbook.Name}")
            time.sleep(2)

            hwnd = excel.Hwnd

            # Minimizar a tela
            win32gui.ShowWindow(hwnd, 6)   

            
            break

        except Exception as e:
            print(f"Erro ao abrir o arquivo Excel na tentativa {attempt + 1}: {e}")

            attempt += 1
            time.sleep(2)


    if workbook:
        print(f"Operação fincalizada com sucesso no arquivo '{workbook.Name}'")
    else:
        print("Falha ao abrir o arquivo após várias tentativas")



def macro(file_path, max_attempts=3):
    def get_excel_app():
        try:
            excel = win32com.client.GetActiveObject("Excel.Application")
            print("Usando uma instância existente do Excel.")
        except Exception as e:
            print(f"Erro ao conectar a uma instância existente: {e}")
            try:
                excel = win32com.client.Dispatch("Excel.Application")
                print("Criando uma nova instância do Excel.")
            except Exception as e:
                print(f"Erro ao criar uma nova instância do Excel: {e}")
                return None
        return excel

    if not os.path.isfile(file_path):
        print(f"O arquivo '{file_path}' não foi encontrado.")
        return
    
    excel = get_excel_app()
    if excel is None:
        print("Não foi possível inicializar o Excel.")
        return
    time.sleep(2)

    try:
        excel.Visible = True
    except Exception as e:
        print(f"Erro ao definir a propriedade Visible: {e}")

    attempt = 0
    workbook = None
    
    while attempt < max_attempts:
        try:
            workbook = excel.Workbooks.Open(file_path)
            print(f"Arquivo '{workbook.Name}' aberto com sucesso na tentativa {attempt + 1}.")
            
            # Acessa a aba "Config"
            try:
                sheet = workbook.Sheets("Config")
                sheet.Activate()
                print("Aba 'Config' ativada com sucesso.")
            except Exception as e:
                print(f"Erro ao acessar a aba 'Config': {e}")
                break  # Encerra as tentativas se não conseguir acessar a aba
            
            tranferir()
            pop_transf_ok()
            workbook.Save()
            close_existing_excel_instances()
            time.sleep(2)
            break

        except Exception as e:
            print(f"Erro ao abrir o arquivo Excel na tentativa {attempt + 1}: {e}")
            
            attempt += 1


def mmm_4():
    mascara_cpp(MS1)
    mascara_cpp(CP1)
    macro(MC1)