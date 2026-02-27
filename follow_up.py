import pandas as pd

def enviar_mensagens(df, token, layout, requests):

    df = df.loc[df['Aprovadores'] != 'ERROR']

    url = 'https://slack.com/api/chat.postMessage'
    
    headers = {
        'Authorization': f'Bearer {token}',
        'Content-Type': 'application/json',
    }

    for index, row in df.iterrows():

        dados_do_fup = {
        'ID_Aprovador' : row['ID Slack Aprovador'],
        'doc': str(row['Documento']).split('.')[0],
        'fornecedor': row['Nome do fornecedor'],
        'centro_custo': row['Centro de Custo'],
        'tipo_documento': row['Tipo Documento'],
        'ID_comprador': row['ID Slack Comprador'],
    }
        mensagem_final = layout.substitute(dados_do_fup)

        data = {
        'channel': 'U072798C68L',
        'text': mensagem_final,
        }

        requests.post(url=url, headers=headers, json=data)

def limpeza_dados(pd):

    df = pd.read_excel(
        'relatorio_aprovadores.xlsx', engine= 'calamine'
    )

    colunas = ['Empresa', 'Nome do fornecedor', 'Tipo Documento', 'Documento','Status Tarefa', 'Centro de Custo', 'Vencimento', 'Aprovadores', 'Criado por']
    
    # 1. Primeiro, transformamos a string "Letícia, Raimundo" em uma lista ['Letícia', 'Raimundo']
    df['Aprovadores'] = df['Aprovadores'].str.split(',')

    # 2. O explode transforma cada item da lista em uma nova linha
    df = df.explode('Aprovadores')

    # 3. Importante: Como geralmente tem um espaço após a vírgula, limpamos os espaços extras
    df['Aprovadores'] = df['Aprovadores'].str.strip()

    # Resetar o index é opcional, mas recomendado para evitar índices duplicados
    df = df.reset_index(drop=True)

    df = df[colunas]
    
    df['Vencimento'] = pd.to_datetime(df['Vencimento'], errors='coerce')
    
    df.sort_values('Documento', inplace=True)

    filtro = (df['Status Tarefa'] == 'Ag. Aprovação') & ((df['Tipo Documento'] == 'Contrato de Compras') | (df['Tipo Documento'] == 'Pedido de Compras'))
    df = df.loc[filtro].reset_index(drop=True)
    
    # Garante que Documento seja lido como string limpa
    df['Documento'] = df['Documento'].astype(str).str.split('.').str[0]
    
    return df

def buscar_id_slack(df):
    # Carrega a base de referência (mantendo o nome original da coluna de busca)
    usuarios = pd.read_excel('users_sap.xlsx')[['Nome', 'ID Slack']]

    # 1. Merge para o Aprovador
    # Usamos left_on (coluna no df principal) e right_on (coluna na planilha de usuários)
    df = pd.merge(df, usuarios, left_on='Aprovadores', right_on='Nome', how='left')
    
    # Renomeia o ID Slack que acabou de chegar e dropa a coluna 'Nome' que vem de brinde no merge
    df.rename(columns={'ID Slack': 'ID Slack Aprovador'}, inplace=True)
    df.drop(columns=['Nome'], inplace=True)

    # 2. Merge para o Comprador (Criado por)
    df = pd.merge(df, usuarios, left_on='Criado por', right_on='Nome', how='left')
    
    # Renomeia para o nome que você quer e dropa o 'Nome' extra de novo
    df.rename(columns={'ID Slack': 'ID Slack Comprador'}, inplace=True)
    df.drop(columns=['Nome'], inplace=True)

    return df
