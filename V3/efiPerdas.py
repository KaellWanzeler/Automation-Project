import time
import psutil
import win32gui
import win32com.client
  

def refreshing_save():
    excel = win32com.client.Dispatch("Excel.Application")
    workbook = excel.ActiveWorkbook

    workbook.RefreshAll()
    print("Atualizando...")

    # Aguarda a conclusão de atualizações assícronas 
    excel.CalculateUntilAsyncQueriesDone()

    # Salva o arquivo após a atualização
    workbook.Save()
    

    print(f'Atualização concluída com sucesso para o arquivo {workbook.Name}')
    time.sleep(10)

    # Desativar os alertas de confirmação
    excel.DisplayAlerts = False

    # Marca o workbook como salvo
    workbook.Saved = True

    time.sleep(2)

    # Fecha o Excel sem solicitar confirmação
    excel.Quit()
# !!! NÃO ESTÁ SENDO USADO !!!
def mini_excel():
    excel = win32com.client.Dispatch("Excel.Application")

    hwnd = excel.Hwnd

    win32gui.ShowWindow(hwnd, 6) 

# fechar os Excel´s
def fechar_instancias_excel():
    # Itera por todos os processos em execução
    for processo in psutil.process_iter(['pid', 'name']):
        # Verifica se o nome do processo é 'EXCEL.EXE'
        if processo.info['name'] == 'EXCEL.EXE':
            # Encerra o processo do Excel
            processo.terminate()

    print("Todas as instâncias do Excel foram fechadas.")



def EfiPerdasMes():
    def get_excel_app():
        # Tenta conectar a uma instância do Excel
        try:
            excel = win32com.client.GetActiveObject("Excel.Application")
            print("Usando uma instância existente do Excel")
        except Exception as e:
            print(f"Erro ao conectar a uma instância existente do Excel: {e}")
            try:     
                # Se não houver onstância existente, cria uma nova
                excel = win32com.client.Dispatch("Excel.Application")
                print("Criando uma nova instância do Excel")
            except Exception as e:
                print(f"Erro ao criar uma nova instância do Excel: {e}")
                return None
        return excel
        
    # Caminho para o arquivo Excel
    file_path = r''

    # Inicializa ou conecta a uma instância do Excel
    excel = get_excel_app()
    if excel is None:
        print(f"Não foi possvel inicializar o Excel.")
        return
    time.sleep(2)
    try:
        excel.Visible = True
    except Exception as e:
        print(f"Erro ao definir a propridade Visible: {e}")

    try:
        # Abre o arquivo Excel
        workbook = excel.Workbooks.Open(file_path)
        print(f"Arquivo '{workbook.Name}' aberto com sucesso.")
    except Exception as e:
        print(f"Erro ao abirr o arquivo Excel: {e}")
        return
    
    time.sleep(5)


def EfiPerdasAno():
    def get_excel_app():
        # Tenta conectar a uma intância do Excel
        try:
            excel = win32com.client.GetActiveObject("Excel.Application")
            print("Usando uma instânica existente do Excel.")
        except Exception as e:
            print(f"Erro ao conectar a uma instância existente do Excel: {e}")
            try:
                # Se não houver instância existente, cria uma nova
                excel = win32com.client.Dispatch("Excel.Application")
                print("Criando uma nova instância do Excel.")
            except Exception as e:
                print(f"Erro aocriar uma nva instância do Excel: {e}")
                return None
        return excel
    
    # Caminho para o arquivo Excel
    file_path = r''

    # Inicializa ou conecta a uma instância do Excel
    excel = get_excel_app()
    if excel is None:
        print(f"Não foi possivel inicializar o Excel.")
        return
    time.sleep(2)

    try:
        excel.Visible = True
    except Exception as e:
        print(f"Erro ao definir a propriedade Visible: {e}")

    try:
        # Abre o arquivo Excel
        workbook = excel.Workbooks.Open(file_path)
        print(f"Arquivo '{workbook.Name}' aberto com sucesso.")
    except Exception as e:
        print(f"Erro ao abrir o arquivo Excel: {e}")
        return

    time.sleep(5)
    

def EfiPerdas():
    def get_excel_app():
        try:
            excel = win32com.client.GetActiveObject("Excel.Application")
            print("Usando uma instãncia existente")
        except Exception as e:
            try:
                # Se não houver instância existente, cria uma nova 
                excel = win32com.client.Dispatch("Excel.Application")
                print("Criando uma noca instância do Excel.")
            except Exception as e:
                print(f"Erro ao criar uma noca instância do Excel: {e}")
                return None
        return excel
    
    # Caiminho para o arquivo Excel
    file_path = r''

    # Iniciando ou conecta a uma instância do Excel
    excel = get_excel_app()
    if excel is None:
        print(f"Não foi possivel inicializar o Excel.")
        return
    time.sleep(2)
    try:
        excel.Visible = True
    except Exception as e:
        print(f"Erro ao definir a propriedade Visible: {e}")
    
    try:
        # Abra o arquivo Excel
        workbook = excel.Workbooks.Open(file_path)
        print(f"Arquivo '{workbook.Name}' aberto com sucesso.")
    except Exception as e:
        print(f"Ero ao abrir o arquivo Excel: {e}")
        return
    
    time.sleep(5)

def EfiPerdasMain():
    EfiPerdasMes()
    time.sleep(2)
    refreshing_save()
    time.sleep(2)
    print("Efiperdas feito com sucesso.")
    
    time.sleep(10)

    EfiPerdasAno()
    time.sleep(2)
    refreshing_save()
    time.sleep(2)
    print("EfiPerdasAno feito com sucesso.")

    time.sleep(10)

    EfiPerdas()
    time.sleep(2)
    refreshing_save()
    time.sleep(2)
    print("------------------------------EfiPerdas feito com sucesso.------------------------------")