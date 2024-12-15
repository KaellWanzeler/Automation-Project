# Import´s
import time
import win32com
import win32com.client
import win32com.client as win32

# Atualizando e Salvando
def refreshing_save():
    excel = win32com.client.Dispatch("Excel.Application")
    workbook = excel.ActiveWorkbook

    # Atualiza
    workbook.RefreshAll()
    print("Atualizando...")

     # Aguarda a conclusão da atualização assícronas
    while True:
        try:
            # Verifica se a atualização assíncrona foi concluída
            excel.CalculateUntilAsyncQueriesDone()
            print("Consultas assíncronas concluídas.")
            break
        except Exception as e:
            print(f"Aguardando a conclusão da atualização: {e}")
            time.sleep(5)

    # Salva o Excel atual
    workbook.Save()

    print(f"Atualização concluída e feito o save do arquivo {workbook.Name}")

def recalculo():
    excel = win32com.client.Dispatch("Excel.Application")
    workbook = excel.ActiveWorkbook
    
    sheet = excel.ActiveSheet

    sheet.Calculate()

    time.sleep(5)

def info():
    excel = win32.Dispatch("Excel.Application")
    workbook = excel.ActiveWorkbook
    sheet = workbook.Sheets('INFO')

    sheet.Activate()


# Formula que fecha todos os Excels
def close_existing_excel_instances():
    try:
        excel_app = win32com.client.Dispatch("Excel.Application")
        for wb in excel_app.Workbooks:
            wb.Close(SaveChanges=True)
        excel_app.Quit()
    except Exception as e:
        print(f"Ocorreu um erro: {e}")

# Const
ARQUIVO1 = r''
ARQUIVO2 = r''
ARQUIVO3 = r''
ARQUIVO4 = r''
ARQUIVO5 = r''

# MRC
def mrc(setor):
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

    file_path = setor

    excel = get_excel_app()
    if excel is None:
        print("Não foi possível inicializar o Excel.")
        return

    try:
        excel.Visible = True
    except Exception as e:
        print(f"Erro ao definir a propriedade Visible: {e}")

    try:
        workbook = excel.Workbooks.Open(file_path)
        print(f"Arquivo '{workbook.Name}' aberto com sucesso.")
    except Exception as e:
        print(f"Erro ao abrir o arquivo Excel: {e}")
        return
    

    try:
        sheet = workbook.Sheets('HSMO')
        print("Aba HSMO selecionada com sucesso!")
    except Exception as e:
        print(f"Ocorreu um erro tentando mudar de aba: {e}")
        return
    
    try:
        sheet.Activate()
    except Exception as e:
        print(f"Erro ao ativar")
    
    time.sleep(5)

def atualizar():
    mrc(ARQUIVO1)
    recalculo()
    time.sleep(10)

    mrc(ARQUIVO2)
    recalculo()
    time.sleep(10)
    
    mrc(ARQUIVO3)
    recalculo()
    time.sleep(10)
    
    mrc(ARQUIVO4)
    recalculo()
    time.sleep(10)
    
    mrc(ARQUIVO5)
    recalculo()
    time.sleep(10)
    

    close_existing_excel_instances()