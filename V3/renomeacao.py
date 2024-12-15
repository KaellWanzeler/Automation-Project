import os
import shutil

# Constantes
ARQUIVOS = [
    r'',
    r'',
    r'',
    r'',
    r''
]

PASTA_DESTINO = r''

# Função
def renomear_e_mover_arquivos(arquivos, pasta_destino):
    for arquivo in arquivos:
        # Extrai o nome do arquivo sem extensão 
        nome_arquivo = os.path.basename(arquivo)
        # Remve a parte "_AGOSTO_204" do nome do arquivo
        novo_nome = '_'.join(nome_arquivo.split('_')[:2]) + os.path.splitext(arquivo)[1]

        # Cria o caminho completo para o novo arquivo na pasta de destino
        novo_caminho = os.path.join(pasta_destino, novo_nome)

        # Move e substitui o arquivo na pasta de destino 
        shutil.copy2(arquivo,novo_caminho)
        print(f"{nome_arquivo} foi copiado para {novo_caminho} e renomeado para {novo_nome}")