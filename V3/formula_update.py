# Import´s
import re
import time
import win32com
import win32gui
import pyautogui
import win32com.client
import win32com.client as win32

def save():
    excel = win32com.client.Dispatch("Excel.Application")
    time.sleep(2)
    workbook = excel.ActiveWorkbook

    # Salva o Excel atual
    workbook.Save()

    print(f"Atualização concluída e feito o save do arquivo {workbook.Name}")

    time.sleep(2)

    workbook.Save()

    print(f"{workbook.Name} salvo com sucesso!")

    time.sleep(2)

    excel.Quit()

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
ARQUIVO1 =r''
ARQUIVO2 =r''
ARQUIVO3 =r''
ARQUIVO4 =r''
ARQUIVO5 =r''

def aba():
    # Conectando ao Excel que já está aberto
    excel = win32.Dispatch('Excel.Application')
    time.sleep(5)
    # Seleciona a planilha que deseja trabalhar (mude o nome da planilha conforme necessário)
    workbook = excel.ActiveWorkbook
    sheet = workbook.Sheets('RELATÓRIO_PRODUÇÃO')  # Nome da aba

    sheet.Activate()
    print("Aba de 'RELATÓRIO_PRODUÇÃO selecionada")


def retirando_filtros():
    # Conectar ao Excel
    excel = win32.Dispatch("Excel.Application")
    time.sleep(2)
    # Conectar ao arquivo Excel já aberto
    workbook = excel.ActiveWorkbook
    sheet = workbook.ActiveSheet

    # Definir a célula que você deseja capturar (Exemplo: A1)
    cell = sheet.Range("C4")

    if sheet.AutoFilterMode:
        sheet.AutoFilterMode = False

    for table in sheet.ListObjects:
        if table.AutoFilter:
            table.AutoFilter.ShowAllData()

    print("Filtros retirados") 


def move_C4():
    excel = win32.Dispatch("Excel.Application")
    time.sleep(2)
    excel.Visible = True

    workbook = excel.ActiveWorkbook
    sheet = workbook.ActiveSheet

    target_cell = sheet.Range('C4')
    target_cell.Select()

    print("Movido para C4")

    time.sleep(2)

def move_M5():
    excel = win32.Dispatch("Excel.Application")
    time.sleep(2)
    excel.Visible = True

    workbook = excel.ActiveWorkbook
    sheet = workbook.ActiveSheet

    target_cell = sheet.Range('M5')
    target_cell.Select()

    print("Movido para M5")

def move_X5():
    excel = win32.Dispatch("Excel.Application")
    time.sleep(2)
    excel.Visible = True

    workbook = excel.ActiveWorkbook
    sheet = workbook.ActiveSheet

    target_cell = sheet.Range('X5')
    target_cell.Select()

    print("Movido para X5")

def classificacao():
    # Conectar ao Excel
    excel = win32.Dispatch("Excel.Application")
    time.sleep(2)
    # Conectar ao arquivo Excel já aberto
    workbook = excel.ActiveWorkbook
    sheet = workbook.ActiveSheet
    
    DATA = r'c:\Users\wpt28798\Pictures\cv2\Data.png'
    location = pyautogui.locateCenterOnScreen(DATA, confidence=0.7)

    if location:
        pyautogui.click(location)
        print("Celula selecionada")
    else:
        print("Imagem não encontrada")
    # Simular o atalho para "Classificar de A a Z" 
    # No Exel, o atalho para ordenar é de Alt + S + A + L, então usaremos esse atalho
    pyautogui.hotkey('alt', 's') # Abre o menu de Dados
    time.sleep(0.5) # Aguarda o menu abrir
    pyautogui.press('a') # Selecionar o botão de Classificação de A a z
    time.sleep(0.5)
    pyautogui.press('l')

    print("Classificação aplicada")


def formula_M5():
    # Conectar ao Excel
    excel = win32.Dispatch("Excel.Application")
    time.sleep(2)
    # Conectar ao arquivo Excel já aberto
    workbook = excel.ActiveWorkbook
    sheet = workbook.ActiveSheet

    # Localizar a célula com a fórmula (por exemplo, A1)
    cell = sheet.Range('M5') # Substitua com o endereço da célula desejada

    # Atualizar a fórmula
    original_formula = cell.Formula
    new_formula = re.sub(r'RELATÓRIO_PRODUÇÃO!AL\d+', 'RELATÓRIO_PRODUÇÃO!AL5', original_formula)
    cell.Formula = new_formula

    # Encontrar a última linha com dados na coluna 
    last_row = sheet.Cells(sheet.Rows.Count, cell.Column).End(-4162).Row

    # Preencher a fórmula para o restante da coluna
    range_to_fill = sheet.Range(cell, sheet.Cells(last_row, cell.Column))
    range_to_fill.FillDown()

    print("Formula atualizada")


def formula_X5():
    # Conectar ao Excel
    excel = win32.Dispatch("Excel.Application")
    time.sleep(2)
    # Conectar ao arquivo Excel já aberto
    workbook = excel.ActiveWorkbook
    sheet = workbook.ActiveSheet

    # Localizar a célula com a fórmula (por exemplo, A1)
    cell = sheet.Range('X5') # Substitua com o endereço da célula desejada

    # Atualizar a fórmula
    original_formula = cell.Formula
    new_formula = re.sub(r'RELATÓRIO_PRODUÇÃO!AL\d+', 'RELATÓRIO_PRODUÇÃO!AL5', original_formula)
    cell.Formula = new_formula

    # Encontrar a última linha com dados na coluna 
    last_row = sheet.Cells(sheet.Rows.Count, cell.Column).End(-4162).Row

    # Preencher a fórmula para o restante da coluna
    range_to_fill = sheet.Range(cell, sheet.Cells(last_row, cell.Column))
    range_to_fill.FillDown()

    print("Formula atualizada")


def formula_2_X5():
    # Conectar ao Excel
    excel = win32.Dispatch("Excel.Application")
    time.sleep(2)

    # Conectar ao arquivo Excel já aberto
    workbook = excel.ActiveWorkbook
    sheet = workbook.ActiveSheet

    # Localizar a célula com a fórmula (por exemplo, A1)
    cell = sheet.Range('X5') # Substitua com o endereço da célula desejada

    # Atualizar a fórmula
    if isinstance(cell.Formula, str):
        # Substituir qualquer referência á coluna $AE seguida de um número por $AE5
        nova_formula = re.sub(r'\$AE\d+', '$AE5', cell.Formula)
        cell.Formula = nova_formula

    # Encontrar a última linha com dados na coluna 
    last_row = sheet.Cells(sheet.Rows.Count, cell.Column).End(-4162).Row

    # Preencher a fórmula para o restante da coluna
    range_to_fill = sheet.Range(cell, sheet.Cells(last_row, cell.Column))
    range_to_fill.FillDown()

    print("Formula atualizada")


def form():
    aba() # Ir para aba de RELATÓRIO DE PRODUÇÃO
    time.sleep(2)
    retirando_filtros() # Retira os filtros
    time.sleep(2)
    move_C4() # Se move para a celula C4
    time.sleep(2)
    classificacao() # Clica na celula, e aperta algumas teclas para ordenar 
    time.sleep(2)
    move_M5() # Se move para a celula M5
    time.sleep(2)
    formula_M5() # Muda a formula que veio
    time.sleep(2)
    move_X5() # Se move para a celula X5
    time.sleep(2)
    formula_X5() # Muda a formula que veio
    time.sleep(2)
    formula_2_X5() # Muda a fomrula que veio
    time.sleep(2)


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
    time.sleep(2)

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
    
    time.sleep(5)
    
    form()


def update():
    mrc(ARQUIVO1)
    time.sleep(5)
    save()
    time.sleep(5)
    mrc(ARQUIVO2)
    time.sleep(5)
    save()
    time.sleep(5)
    mrc(ARQUIVO3)
    time.sleep(5)
    save()
    time.sleep(5)
    mrc(ARQUIVO4)
    time.sleep(5)
    save()
    time.sleep(5)
    mrc(ARQUIVO5)
    time.sleep(5)
    save()
    time.sleep(5)