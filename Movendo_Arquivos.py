from datetime import datetime
from pathlib import Path
from shutil import move, rmtree, copy
import os 
import time 

path_base = r'CAMINHO DA PASTA DOWNLOADS' 
destino = Path(r'CAMINHO DA PASTA DE DESTINO') # No meu caso a pasta es ta alocada na Rede (outro servidor)

#Precisei excluir antes todos os arquivos que contem a palavra chave que a func glob pega
def remover_arquivo_antigo(destino):
    arquivo_antigo = None
    for file in destino.glob("NOME DO ARQUIVO*.xlsx"): # Ecluindo todos os arquivos que contem o nome
        if file.is_file():
            arquivo_antigo = file
            break
    if arquivo_antigo is not None:
        os.remove(arquivo_antigo)
        print(f"ARQUIVO REMOVIDO COM SUCESSO: {arquivo_antigo}")
    print('\n')
remover_arquivo_antigo(destino)
time.sleep(1)

#Copiando o arquivo que acabou de chegar na pasta Download
def copiar_arquivo(origem, destino) -> bool:
    try:
        copy(origem, destino) 
        return True

    except Exception as e:
        print(e)
        return False

#Criando uma pasta com o nome do arquivo para poder separar cada arquivo
def criar_pasta(nome_pasta):
    try:
        path_pasta = Path(nome_pasta)
        path_pasta.mkdir(parents=True, exist_ok=True)
        return True

    except Exception as e:
        print(f'TIVEMOS UM ERRO AO CRIAR A PASTA, SEGUE O ERRO: {e}\n')        
        return False

# Depois de copiar o arquivo e criar a pasta, vamos chamar as funções no try
try:
    date_today = datetime.today().strftime('%d-%m-%Y')
    for i in range(0, 99):
        nome_arquivo = Path(rf"{path_base}/NOME_DO_ARQUIVO__{i}_{date_today}.xlsx")
        if nome_arquivo.is_file():        
            print(f'ARQUIVO ENCONTRADO: {nome_arquivo}') 
            print('\n')
            if copiar_arquivo(origem=nome_arquivo, destino=destino): 
                print('ARQUIVO COPIADO COM SUCESSO. UHUUU!')
                print('\n')
                # Criar pasta com o nome do arquivo na pasta de downloads
                nome_pasta = Path(rf"{path_base}/{nome_arquivo.stem}")
                criar_pasta(nome_pasta)
                
                # Mover arquivo para a pasta criada
                destino_arquivo = nome_pasta / nome_arquivo.name
                move(nome_arquivo, destino_arquivo)
                print(f'ACABEI DE ENVIAR O ARQUIVO PARA: {destino_arquivo}\n')                
            else:
                print('ALGO DE ERRADO NÃO ESTA CERTO...')
            break 
        
except Exception as e:
    print(f'ERRO NA TENTATIVA DE EXECUÇÃO: {e}')
