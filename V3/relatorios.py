# Import´s
import os
import shutil

# Constantes
CAMINHO = r''
RELATORIO_PARA = rf'{CAMINHO}'
RELATORIO_PROD = rf'{CAMINHO}'

PATH_RELATORIOS = rf''

DOWNLOAD_PATH =rf'{CAMINHO}'

# Função que move para o arquivo que fica os Relatórios no SharePoint
def mover_arquivos(arquivos, pasta_destino):
    if os.path.exists(arquivos):
        if os.path.exists(pasta_destino):
            shutil.copy2(arquivos, pasta_destino)
            print(f"{arquivos} foi copiado para {pasta_destino}")
        else:
            print(f"Erro: A pasta de destino {pasta_destino} não foi encontrada.")
    else:
        print(f"Erro: O arquivo {arquivos} não foi encontrado.")

# Faz a transferência dos relatórios
def enviando_rela():
    mover_arquivos(RELATORIO_PARA,PATH_RELATORIOS)
    mover_arquivos(RELATORIO_PROD,PATH_RELATORIOS)

def delete_files():
    download_folder = os.path.expanduser(DOWNLOAD_PATH)
    files_to_delete = ['RELATÓRIO_PARADAS.csv', 'RELATÓRIO_PRODUÇÃO.csv']

    for file_name in files_to_delete:
        file_path = os.path.join(download_folder, file_name)
        try:
            os.remove(file_path)
            print(f"Arquivo {file_name} deletado com sucesso.")
        except FileNotFoundError:
            print(f"Arquivo {file_name} não encontrado")
        except Exception as e:
            print(f"Ocorreu um erro ao tentar deletar {file_name}: {e}")