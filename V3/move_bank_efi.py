import os
import shutil

# Constantes
ARQUIVOS_EFI = r''
ARQUIVOS_BANK = r''

PASTA_DESTINO_EFI = r''
PASTA_DESTINO_BANK = r''

# Função
def mover_arquivos(arquivos, pasta_destino):
    if os.path.exists(arquivos):
        if os.path.exists(pasta_destino):
            shutil.copy2(arquivos, pasta_destino)
            print(f"{arquivos} foi copiado para {pasta_destino}")
        else:
            print(f"Erro: A pasta de destino {pasta_destino} não foi encontrada.")
    else:
        print(f"Erro: O arquivo {arquivos} não foi encontrado.")

def movendo_efi_bank():
    print(f"Arquivo EFI: {ARQUIVOS_EFI}")
    print(f"Pasta de destino EFI: {PASTA_DESTINO_EFI}")
    mover_arquivos(ARQUIVOS_EFI, PASTA_DESTINO_EFI)
    
    print(f"Arquivo BANK: {ARQUIVOS_BANK}")
    print(f"Pasta de destino BANK: {PASTA_DESTINO_BANK}")
    mover_arquivos(ARQUIVOS_BANK, PASTA_DESTINO_BANK)