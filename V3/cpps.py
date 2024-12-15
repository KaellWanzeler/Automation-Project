import os
import shutil

# Constantes
CAMINHO = r''

ARQUIVO1 = rf'{CAMINHO}'
ARQUIVO2 = rf'{CAMINHO}'
ARQUIVO3 = rf'{CAMINHO}'
ARQUIVO4 = rf'{CAMINHO}'
ARQUIVO5 = rf'{CAMINHO}'

PASTA_ARQUIVO1 = rf'{CAMINHO}'
PASTA_ARQUIVO2 = rf'{CAMINHO}'
PASTA_ARQUIVO3 = rf'{CAMINHO}'
PASTA_ARQUIVO4 = rf'{CAMINHO}'
PASTA_ARQUIVO5 = rf'{CAMINHO}'
# Função
def mover_arquivos(arquivos, pasta_destino):
    if os.path.exists(arquivos):
        shutil.copy2(arquivos, pasta_destino)
        print(f"{arquivos} foi copiado para {pasta_destino}")
    else:
        print(f"Arquivo não encontrado: {arquivos}")
def movendo_cpps():
    mover_arquivos(ARQUIVO1, PASTA_ARQUIVO1)
    mover_arquivos(ARQUIVO2, PASTA_ARQUIVO2)
    mover_arquivos(ARQUIVO3, PASTA_ARQUIVO3)
    mover_arquivos(ARQUIVO4, PASTA_ARQUIVO4)
    mover_arquivos(ARQUIVO5, PASTA_ARQUIVO5)