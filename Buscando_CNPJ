import pandas as pd
import requests
import time

# Essa é uma ideia que meu atual chefe faz no próprio excel, resolvi trazer para um robo fazer isso automaticamente e sem precisar ficar olhando terminar o processo dentro do próprio excel.
# Sem falar que no excel ele busca um por um. Aqui eu fiz uma interação para buscar todas as linhas já informadas na planilha



# Ler a planilha de CNPJs (supondo que esteja no formato .xlsx) Pegando cada cnpj que esta na planilha e criando uma api e trazendo esses dados para um novo arquivo em excel

df_cnpjs = pd.read_excel(r'CAMINHO DO ARQUIVO', dtype={'CNPJ': str}) #formatando os numeros da coluna CNPJ para str para facilitar nos proximos passos a concatenação

# Criar uma lista vazia para armazenar as informações retirada na api
lista_informacoes = []

# Definir o tamanho do lote (número de CNPJs a serem processados em cada lote) Como há uma limitação de dados e tempo de retorno atribui um limite de lotes de pesquisa
tamanho_lote = 50

# Inicializar a variável de contagem de lotes
lotes_concluidos = 0

# Iterar sobre os CNPJs em lotes onde pega cada cnpj da planilha 
print(f'Entrando na API e buscando os registros. Buscando {lotes_concluidos + 1}º Lote.')
for i in range(0, len(df_cnpjs), tamanho_lote):
    # Obter os CNPJs do lote atual
    cnpjs_lote = df_cnpjs['CNPJ'][i:i+tamanho_lote]
    
    # Iterar sobre os CNPJs do lote atual
    for cnpj in cnpjs_lote:
        # Converter o CNPJ para string
        cnpj = str(cnpj)
        
        # Construir a URL para obter as informações do CNPJ -- por isso que transformei o cnpj em str para fazer essa concatenação
        site = "https://receitaws.com.br/v1/cnpj/"
        url = site + cnpj
        
        # Fazer a requisição para obter as informações
        response = requests.get(url)
        if response.status_code == 200:
            # Extrair as informações do JSON
            dados = response.json()
            
            # Criar um dicionário com as informações encontradas dentro da api (json)
            informacoes = {
                "abertura": dados.get('abertura', ''),
                "situacao": dados.get('situacao', ''),
                "tipo": dados.get('tipo', ''),
                "nome": dados.get('nome', ''),
                "fantasia": dados.get('fantasia', ''),
                "porte": dados.get('porte', ''),
                "natureza_juridica": dados.get('natureza_juridica', ''),
                "logradouro": dados.get('logradouro', ''),
                "numero": dados.get('numero', ''),
                "municipio": dados.get('municipio', ''),
                "bairro": dados.get('bairro', ''),
                "uf": dados.get('uf', ''),
                "cep": dados.get('cep', ''),
                "email": dados.get('email', ''),
                "telefone": dados.get('telefone', ''),
                "data_situacao": dados.get('data_situacao', ''),
                "motivo_situacao": dados.get('motivo_situacao', ''),
                "ultima_atualizacao": dados.get('ultima_atualizacao', ''),
                "status": dados.get('status', ''),
                "complemento": dados.get('complemento', ''),
                "efr": dados.get('efr', ''),
                "situacao_especial": dados.get('situacao_especial', ''),
                "data_situacao_especial": dados.get('data_situacao_especial', ''),
                "atividade_principal": '',
                "atividade_principal_code": '',
                "atividades_secundarias": '',
                "atividades_secundarias_code": '',
                "capital_social": dados.get('capital_social', '')
            }
            
            # Verificar se existem atividades principais no CNPJ
            if 'atividade_principal' in dados and isinstance(dados['atividade_principal'], list):
                atividades_principais = []
                atividades_principais_code = []
                for atividade in dados['atividade_principal']:
                    atividades_principais.append(atividade.get('text', ''))
                    atividades_principais_code.append(atividade.get('code', ''))
                informacoes['atividade_principal'] = ', '.join(atividades_principais)
                informacoes['atividade_principal_code'] = ', '.join(atividades_principais_code)
            
            # Verificar se existem atividades secundárias no CNPJ
            if 'atividades_secundarias' in dados and isinstance(dados['atividades_secundarias'], list):
                atividades_secundarias = []
                atividades_secundarias_code = []
                for atividade in dados['atividades_secundarias']:
                    atividades_secundarias.append(atividade.get('text', ''))
                    atividades_secundarias_code.append(atividade.get('code', ''))
                informacoes['atividades_secundarias'] = ', '.join(atividades_secundarias)
                informacoes['atividades_secundarias_code'] = ', '.join(atividades_secundarias_code)
            
            # Adicionar as informações à lista
            lista_informacoes.append(informacoes)
        time.sleep(50)
    # Atualizar o número de lotes concluídos
    lotes_concluidos += 1

    # Exibir a mensagem de conclusão do lote atual
    print(f'{lotes_concluidos}º Lote concluído com sucesso!')

    # Aguardar 10 minutos antes de processar o próximo lote
    time.sleep(600)
    #Nesta parte de times eu adicionei um tempo muito grande para que não haja uma sobrecarga de consultas e quebre o robo

# Criar um DataFrame a partir da lista de informações que retiramos da api dentro de cada lote buscado
df_resultado = pd.DataFrame(lista_informacoes)

# Salvar o DataFrame de resultado em um arquivo Excel
df_resultado.to_excel('Resultados.xlsx', index=False)
