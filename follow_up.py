import pandas as pd

def enviar_mensagens(df, token, layout, requests):

    url = 'https://slack.com/api/chat.postMessage'
    
    headers = {
        'Authorization': f'Bearer {token}',
        'Content-Type': 'application/json',
    }

    for index, row in df.iterrows():

        dados_do_fup = {
        'ID_Slack' : row['ID Slack'],
        'doc': str(row['Documento']).split('.')[0],
        'fornecedor': row['Nome do fornecedor'],
        'centro_custo': row['Centro de Custo'],
        'tipo_documento': row['Tipo Documento'],
    }
        mensagem_final = layout.substitute(dados_do_fup)

        data = {
        'channel': 'U072798C68L',
        'text': mensagem_final,
        }

        requests.post(url=url, headers=headers, json=data)
        
        print(mensagem_final)

def limpeza_dados(pd):

    df = pd.read_excel(
        'relatorio_aprovadores.xlsx', engine= 'calamine'
    )
    colunas = ['Empresa', 'Nome do fornecedor', 'Tipo Documento', 'Documento','Status Tarefa', 'Centro de Custo', 'Vencimento', 'Aprovadores']
    df = df[colunas]
    df['Vencimento'] = pd.to_datetime(df['Vencimento'], errors='coerce')
    
    df.sort_values('Documento', inplace=True)

    filtro = (df['Status Tarefa'] == 'Ag. Aprovação') & (df['Tipo Documento'] == 'Contrato de Compras')
    df = df.loc[filtro].reset_index(drop=True)
    
    # Garante que Documento seja lido como string limpa
    df['Documento'] = df['Documento'].astype(str).str.split('.').str[0]
    
    return df

def buscar_id_slack(df):

    aprovadores = pd.read_excel('Trabalhadores (2).xlsx')

    aprovadores = aprovadores[['ID do usuário','Nome', 'ID Slack']]

    # A forma correta de renomear colunas
    aprovadores.rename(columns={'Nome': 'Aprovadores'}, inplace=True)

    df = pd.merge(df, aprovadores, on='Aprovadores', how='left')

    return df
