 # Imports de bibliotecas necessárias para fazermos o processo de complilamento e mover os arquivos
from datetime import datetime
import pandas as pd
from datetime import datetime
from pathlib import Path
from shutil import move, rmtree, copy
from os.path import getmtime
import time   
 
# Criando uma Func
def Compilar_Arquivos(): 
    # Criando laço para reparar bugs 
    try:
        date_today = datetime.today() 
        #hoje = date_today.strftime('%d-%m-%Y' + ' às ' + '%H:%M')   
        data_hoje = date_today.strftime('%d %m %Y')
        # Acrescentando a coluna, compilando o arquivo e enviando para a pasta na rede    
        downloads = Path(r'PASTA DE ORIGEM')
        path_destino = Path(str('ARQUIVO DESEJADO ' + data_hoje + '.xlsx'))
        seu_arquivo = downloads / path_destino
        rede = Path(r'PASTA DE DESTINO') / path_destino

        # Pegando o nome dos arquivos e adicionando na variavel glob
        glob = '(ARQUIVO DESEJADO)*.xlsx'

        # Cria uma pasta em Download para poder separar os arquivos que você precisa fazer o compilamento
        seu_arquivo.mkdir(exist_ok=True)
        for arquivo in downloads.glob(glob): 
            move(arquivo, seu_arquivo)
        # Adicionando o nome das colunas do arquivo fica mais fácil para o for abaixo poder localizar
        colunas = [ 'ADICIONE O NOME DOS CABEÇALHOS DO SEU ARQUIVO' ]

        # Criando o Dataframe com o nome das colunas para receber os dados dos arquivos
        dataframe_final = pd.DataFrame(columns=colunas)

        # Iniciando o for para percorrer todos os arquivos
        for i, arquivo in enumerate(sorted(seu_arquivo.glob(glob), key=getmtime)):
            hora_modificacao = arquivo.stat().st_mtime

            # Adicionando uma coluna de data e hora no arquivo, necessariamente adicone essa linha apenas se for necessário
            # No meu arquivo preciso adicionar. Mas vai do seu gosto
            hora_formatada = datetime.fromtimestamp(hora_modificacao).strftime('%d/%m/%Y' + ' às ' + '%H:%M')
            dataframe_arquivo = pd.read_excel(arquivo)
            
            # Detalhe para poder printar os arquivos que estão sendo compilados
            dataframe_final = pd.concat([dataframe_final, dataframe_arquivo])
            dataframe_final['Filial'] = dataframe_final['Filial'].astype(str)
            print('\n')
            print(f'TRANTANDO O {i+1}º ARQUIVO')
            print(f'{dataframe_final}\n')            

        time.sleep(1)
        # Transformando o dataframe em formato em excel
        dataframe_final.to_excel(seu_arquivo / f'NOME DO ARQUIVO FINAL {data_hoje}.xlsx')
        # Movendo o arquivo para a pasta de destino
        if rede.exists(): 
            rmtree(rede, ignore_errors=True)
        arquivo_compilado = seu_arquivo / f'NOME DO ARQUIVO FINAL {data_hoje}.xlsx'

        # Caso já tenha este arquivo com o mesmo nome, ele será substituido
        copy(arquivo_compilado, rede)
        time.sleep(1)
        print('TRATAMENTO FINALIZADO... ')
        print('\n')
        print(f'ARQUIVO >>> {arquivo_compilado}  <<< ENVIADO PARA A PASTA >>> {rede}')
        time.sleep(8)
    except Exception as e:
        print(f'ERRO NA TENTATIVA DE EXECUÇÃO: {e}')       
        
if __name__ == '__main__':
    Compilar_Arquivos()
