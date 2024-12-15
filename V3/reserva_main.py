import os
os.chdir(r'')

# Import´s
import time
import ctypes
import datetime
import importlib
import pyautogui
from datetime import datetime
from selenium import webdriver

# Arquvios
import sgr
import shapoint
import renomeacao
import efiPerdas 
import relatorios
import formula_update
import Bank
import move_bank_efi
import hsmo_vt
import cps

# Funções
from shapoint import moves_to_sharepoint, imagens_process, wait_for_button, delete_files
import V3.excel_1 as excel_1, V3.excel_2 as excel_2, V3.excel_3 as excel_3, V3.excel_5 as excel_5, V3.excel_4 as excel_4
from V3.excel_1 import mmm_1
from V3.excel_2 import mmm_2
from V3.excel_3 import mmm_3 
from V3.excel_4 import mmm_4 
from V3.excel_5 import mmm_5 
from renomeacao import renomear_e_mover_arquivos, ARQUIVOS, PASTA_DESTINO
from move_bank_efi import movendo_efi_bank
from efiPerdas import EfiPerdasMain
from Bank import Bank_Main
from relatorios import enviando_rela
from formula_update import update
from hsmo_vt import loop
from cpps import movendo_cpps

importlib.reload(sgr)
importlib.reload(shapoint)
importlib.reload(excel_1)
importlib.reload(excel_2)
importlib.reload(excel_3)
importlib.reload(excel_4)
importlib.reload(excel_5)
importlib.reload(renomeacao)
importlib.reload(efiPerdas)
importlib.reload(Bank)
importlib.reload(move_bank_efi)
importlib.reload(hsmo_vt)
importlib.reload(cps)

def prevent_sleep_windows():
    ctypes.windll.kernel32.SetThreadExecutionState(0x80000002)


def main():
    driver = webdriver.Edge()
    try:
        # SGR
        sgr.body(driver)
        time.sleep(10)

        hora_fim = datetime.now().strftime("%H:%M")
        print(f"-------------São exatamente...{hora_fim}-------------")

        # Enviando os relatórios
        enviando_rela()
        time.sleep(10)

        # Deletar arquivos
        delete_files()
        time.sleep(5)
        
        mmm_1()
        print("------------------------------Arquivo1 finalizada!------------------------------")
        time.sleep(5)
        
        hora_fim = datetime.now().strftime("%H:%M")
        print(f"-------------São exatamente...{hora_fim}-------------")

        mmm_2()
        print("------------------------------Arquivo2 finalizada!------------------------------")
        time.sleep(5)
        
        hora_fim = datetime.now().strftime("%H:%M")
        print(f"-------------São exatamente...{hora_fim}-------------")
        
        mmm_3()
        print("------------------------------Arquivo3 finalizada!------------------------------")
        time.sleep(5)
        
        hora_fim = datetime.now().strftime("%H:%M")
        print(f"-------------São exatamente...{hora_fim}-------------")
        
        mmm_4()
        print("------------------------------Arquivo4 finalizada!------------------------------")
        time.sleep(5)
        
        hora_fim = datetime.now().strftime("%H:%M")
        print(f"-------------São exatamente...{hora_fim}-------------")
        
        mmm_5()
        print("------------------------------Arquivo5 finalizada!------------------------------")
        time.sleep(5)
        
        hora_fim = datetime.now().strftime("%H:%M")
        print(f"-------------São exatamente...{hora_fim}-------------")

        # Formula que atualiza as formulas da CPPs
        update()
        time.sleep(2)

        hora_fim = datetime.now().strftime("%H:%M")
        print(f"-------------São exatamente...{hora_fim}-------------")

        renomear_e_mover_arquivos(ARQUIVOS, PASTA_DESTINO)

        Bank_Main()
        print("------------------------------Banco finalizada!------------------------------")
        time.sleep(5)

        hora_fim = datetime.now().strftime("%H:%M")
        print(f"-------------São exatamente...{hora_fim}-------------")

        EfiPerdasMain()
        print("------------------------------Eficiência e Perdas finalizada!------------------------------")
        time.sleep(5)
        
        hora_fim = datetime.now().strftime("%H:%M")
        print(f"-------------São exatamente...{hora_fim}-------------")

        movendo_efi_bank()
        print("------------------------------Arquivos movidos com sucesso!------------------------------")
        time.sleep(5)



    except Exception as e:
        print(f"Ocorreu um erro: {e}")
        return False

    return True

time.sleep(1)

# Dizer que horas está rodando o código
hora_inicio = datetime.now()

# Loop que faz o código rodar a cada 1hora
def job():
    print(f"Iniciando job...")
    success = main()
    if success:
        print("Tarefa concluída com sucesso")
    else:
        print("Erro na execução da tarefa. Tentando novamente em uma hora.")
        
        
def tarefa():
    job()

def job_diario():
    print("São exatamente 23:58 e começará a rodar o código especial")
    loop()
    
def job_semanal():
    print("=== É sábado e estamos enviando as CPPs")
    movendo_cpps()

if __name__ == '__main__':
    prevent_sleep_windows()
    print("O job será executado a cada 2 hora.")

    # Lista de horários específicos no formato HH:MM
    horarios = ["00:00", "02:00", "04:00", "06:00", "08:00", "10:00", "12:00", "14:00", "16:04", "18:00", "20:00", "22:00"]
    executado_hoje = {horario: False for horario in horarios}

    job_diario_executado = False
    job_semanal_executado = False
    
    while True:
        agora = datetime.now()
        hora_atual = agora.strftime("%H:%M")

        if agora.weekday() == 5 and hora_atual == "11:50" and not job_semanal_executado:
            print(f"=== Executando JOB_SEMANAL às {hora_atual} ===")
            job_semanal()
            job_semanal_executado = True

        elif hora_atual == "23:58" and not job_diario_executado:
            print(f"== Executadndo JOB_DIÁRIO às {hora_atual} ===")
            job_diario()
            job_diario_executado = True

        elif hora_atual in horarios and not executado_hoje[hora_atual]:
            hora_inicio = datetime.now().strftime("%H:%M")
            dia_atual = datetime.now()
            dia_atual.strftime("%d/%m/%Y")
            print(f"-------------São exatamente...{hora_inicio}, {dia_atual}-------------") 
            tarefa()
            hora_fim = datetime.now().strftime("%H:%M")
            print(f"-------------São exatamente...{hora_fim}-------------")
            executado_hoje[hora_atual] = True
        
        # Resetar o status de execução à meia-nite
        if agora.strftime("%H:%M") == "00:00":
            executado_hoje = {horario: False for horario in horarios}
            job_diario_executado = False
            job_semanal_executado = False

        time.sleep(1)
