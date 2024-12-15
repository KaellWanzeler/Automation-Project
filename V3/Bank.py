# Import´s
import time
import psutil
import win32com
import win32gui
import win32com.client

# Ataulização e Save
def refreshing_save():
    excel = win32com.client.Dispatch("Excel.Application")
    workbook = excel.ActiveWorkbook

    # Atualiza
    workbook.RefreshAll()
    print("Atualizando...")

    # Aguarda a conclusão da atualização assícronas
    excel.CalculateUntilAsyncQueriesDone()

    # Salva o Excel atual
    workbook.Save()

    print(f"Atualização concluída e feito o save do arquivo {workbook.Name}")
    time.sleep(60)

    # Desativa os alertas de confirmação
    excel.DisplayAlerts = False

    # Marca o workbook como salvo
    workbook.Saved = True
    
    # Fecha o Excel sem solicitar confirmação
    excel.Quit()

# Minimizando o Excel
def mini_excel():
    excel = win32com.client.Dispatch("Excel.Application")

    hwnd = excel.Hwnd

    win32gui.ShowWindow(hwnd, 6)


# Formula que fecha todos os Excels

def close_existing_excel_instances():
    try:
        excel_app = win32com.client.Dispatch("Excel.Application")
        for wb in excel_app.Workbooks:
            wb.Close(SaveChanges=True)
        excel_app.Quit()
    except Exception as e:
        print(f"Ocorreu um erro: {e}")


# Mês do Banco
def Bank_Mes():
    def get_excel_app():
        # Tenta conectar a uma instância do Excel
        try:
            excel = win32com.client.GetActiveObject("Excel.Application")
            print("Usando uma instância existente do Excel")
        except Exception as e:
            print(f"Erro ao conectar a uma instância existente ao Excel: {e}")
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
        print(f"Erro ao abrir o arquivo Exel: {e}")
        return

    time.sleep(5)


# Banco
def Bank():
    def get_excel_app():
        # Tenta conectar a uma instância do Excel
        try:
            excel = win32com.client.GetActiveObject("Excel.Application")
            print("Utilizando uma intância existente do Excel")
        except Exception as e:
            print(f"Erro ao conectar a uma instância existente do Excel: {e}")
            try:
                # Se não houver instância existente, cria uma nova
                excel = win32com.client.Dispatch("Excel.Application")
                print("Criando uma nova instância do Excel")
            except Exception as e:
                print(f"Err ao criar uma nova instancia do Excel: {e}")
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


def Bank_Main():
    Bank_Mes()
    time.sleep(2)
    refreshing_save()
    time.sleep(2)
    print("Primeiro arquivo feito com sucesso")

    Bank()
    time.sleep(2)
    refreshing_save()
    time.sleep(2)
    close_existing_excel_instances()
    print("------------------------------Aruivo feito com sucesso!------------------------------")