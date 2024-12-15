import time
import datetime
import win32com.client
import traceback
from datetime import date
from datetime import datetime
import win32com.client as win32

# Const
ARQUIVO1 = r''
ARQUIVO2 = r''
def save_and_close():
    excel = win32com.client.Dispatch("Excel.Application")
    workbook = excel.ActiveWorkbook

    # Salva o Excel atual
    workbook.Save()

    print(f"Salvando o arquivo {workbook.Name}")

    time.sleep(2)

    workbook.Save()

    print(f"{workbook.Name} salvo com sucesso!")
    workbook.Close(SaveChanges=True)
    time.sleep(2)
def close():
    excel = win32com.client.Dispatch("Excel.Application")
    workbook = excel.ActiveWorkbook

    workbook.Save()

    print(f"Salvando o arquivo {workbook.Name}")
    time.sleep(2)

    print(f"{workbook.Name} salvo com sucesso!")
    excel.Quit()

def cpp_biscoito():
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

    file_path = ARQUIVO1

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
        # Seleciona a aba "Config"
        worsheet = workbook.Sheets("HORAS_DISPONÍVEIS")
        print("Aba 'HORAS_DISPONÍVEIS' selecionada.")
    except Exception as e:
        print(f"Erro ao mudar de aba: {e}") 
        return

    try:   
        worsheet.Activate()
    except Exception as e:
        print(f"Erro ao ativar")


    time.sleep(5)

def cpp_bolos():
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

    file_path = ARQUIVO2

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
        # Seleciona a aba "Config"
        worsheet = workbook.Sheets("HORAS_DISPONÍVEIS")
        print("Aba 'HORAS_DISPONÍVEIS' selecionada.")
    except Exception as e:
        print(f"Erro ao mudar de aba: {e}") 
        return

    try:   
        worsheet.Activate()
    except Exception as e:
        print(f"Erro ao ativar")


    time.sleep(5)


def hsmo_23():
    # Configurações do Excel
    excel = win32.Dispatch("Excel.Application")
    workbook = excel.ActiveWorkbook
    sheet1 = workbook.ActiveSheet

    # Definir o padrão HSMO
    hsmo_padrao = [
        ["", ""],
        ["", ""],
        ["22:40", "24:00"]
    ]

    # Encontrar o número total de linhas com base na coluna DATA
    last_row = sheet1.Cells(sheet1.Rows.Count, 2).End(-4162).Row  # Coluna 2 para DATA

    # Obter a data de hoje
    data_hoje = date.today()

    # Variáveis para controlar se as linhas LB02 e LB03 já foram preenchidas
    lb02_preenchido = False
    lb03_preenchido = False
    lbl02_preenchido = False

    # Iterar pelas linhas da planilha
    for row_idx in range(4, last_row + 1):
        linha_cell = sheet1.Cells(row_idx, 5)  # Coluna LINHA (índice 5)
        data_cell = sheet1.Cells(row_idx, 2)  # Coluna DATA (índice 2)

        # Verifica se a DATA corresponde à data de hoje
        if isinstance(data_cell.Value, datetime):
            data_celula = data_cell.Value.date()  # Obtém a parte da data
            if data_celula == data_hoje:
                # Preencher apenas na linha LB02 se ainda não tiver sido preenchida
                if linha_cell.Value == "LB02" and not lb02_preenchido:
                    sheet1.Range(sheet1.Cells(row_idx + 1, 10), sheet1.Cells(row_idx + 3, 11)).Value = hsmo_padrao
                    lb02_preenchido = True
                    print(f"HSMO aplicado na linha LB02 para a data {data_hoje}.")

                # Preencher apenas na linha LB03 se ainda não tiver sido preenchida
                elif linha_cell.Value == "LB03" and not lb03_preenchido:
                    sheet1.Range(sheet1.Cells(row_idx + 1, 10), sheet1.Cells(row_idx + 3, 11)).Value = hsmo_padrao
                    lb03_preenchido = True
                    print(f"HSMO aplicado na linha LB03 para a data {data_hoje}.")
                
                # Preencher a linha LBL02 se não tiver preenchida
                elif linha_cell.Value == "LBL02" and not lbl02_preenchido:
                    sheet1.Range(sheet1.Cells(row_idx + 1, 10), sheet1.Cells(row_idx + 3, 11)).Value = hsmo_padrao
                    lbl02_preenchido = True
                    print(f"HSMO aplicado na linha LBL02 para a data {data_hoje}")


        # Se ambas as linhas LB02 e LB03 já foram preenchidas, podemos sair do loop
        if lb02_preenchido and lb03_preenchido and lbl02_preenchido:
            break

    print("HSMO aplicado com sucesso.")

def loop():
    print(f"\n---> Iniciando função 'Fechamento de turnos' ({time.strftime("%H:%M:%S")}) ===")
    max_tentativas = 3

    def tentando(func, desc, *args):
        tentativas = 0
        while tentativas < max_tentativas:
            try:
                func(*args)
                print(f"---> {desc} concluída com sucesso ({time.strftime("%H:%M:%S")})\n{'-'*40}")
                return
            except Exception as e:
                tentativas += 1 
                print(f"[ERRO] Tentativa {tentativas}/{max_tentativas} falhou em '{desc}' ({time.strftime("%H:%M:%S")}): {str(e)}")
                traceback.print_exc()
                time.sleep(2)
        print(f"[FALHA] Após {max_tentativas} tentativas, '{desc}' não foi concluída.\an{'-'*40}")

    tentando(cpp_biscoito, "ARQUIVO1 EXCEL")
    tentando(hsmo_23, "FINALIZANDO T")
    tentando(save_and_close, "SALVANDO E FECHANDO")
    
    print("=== ARQUIVO1 finalizada com sucesso!")

    tentando(cpp_bolos, "ARQUIVO2 EXCEL")
    tentando(hsmo_23, "FINALIZANDO T")
    tentando(close, "FECHANDO EXCEL")

    print("=== ARQUIVO2 finalizado com sucesso!")