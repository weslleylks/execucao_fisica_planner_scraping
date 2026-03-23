#%%
import requests
import msal
import json
import pandas as pd
import numpy as np
import re

# --- CONFIGURAÇÕES ---
# CLIENT_ID = 'SEU_CLIENT_ID_AQUI'
# CLIENT_SECRET = 'SEU_CLIENT_SECRET_AQUI'
# TENANT_ID = 'SEU_TENANT_ID_AQUI'
ACCESS_TOKEN = 'Bearer eyJ0eXAiOiJKV1QiLCJub25jZSI6InZUN01GMzVvQzN1cHppUFJISUhHZ3F3QjNaMjBEbnBFeDhiRmdRcnVuNW8iLCJhbGciOiJSUzI1NiIsIng1dCI6IlFaZ045SHFOa0dORU00R2VLY3pEMDJQY1Z2NCIsImtpZCI6IlFaZ045SHFOa0dORU00R2VLY3pEMDJQY1Z2NCJ9.eyJhdWQiOiJodHRwczovL2dyYXBoLm1pY3Jvc29mdC5jb20iLCJpc3MiOiJodHRwczovL3N0cy53aW5kb3dzLm5ldC9lZjY0ZDdkMC1mNmRkLTQyOGMtYmM5OC00OWVmMTgxMWU5YjUvIiwiaWF0IjoxNzc0MjY0ODM1LCJuYmYiOjE3NzQyNjQ4MzUsImV4cCI6MTc3NDM1MTUzNSwiYWNjdCI6MCwiYWNyIjoiMSIsImFjcnMiOlsicDEiXSwiYWlvIjoiQVhRQWkvOGJBQUFBNC9CdjRjUDBjUVFuMWtrMjl0bWxTRXVWanRTMWU3NlpnQW90UnBnUk1UWm96bHBYaGd5SXZuQ2h5RXI0NUw0dFNhZU56TkdyNXNVT0hTODR6OC9KeUJGZnNFWXFUQkQ1QllSOENXcDB0SUgvMUsrd29ub3BJMUxoVWl6c2JXUm5BWnVINkNvVndUNXllSk5hK09uajNnPT0iLCJhbXIiOlsicHdkIiwicnNhIiwibWZhIl0sImFwcF9kaXNwbGF5bmFtZSI6Ik1pY3Jvc29mdCBQbGFubmVyIENsaWVudCIsImFwcGlkIjoiNzVmMzE3OTctMzdjOS00OThlLThkYzktNTNjMTZhMzZhZmNhIiwiYXBwaWRhY3IiOiIwIiwiY2Fwb2xpZHNfbGF0ZWJpbmQiOlsiYmQyZWZkNDEtNWM2Zi00MGQ5LTk3ZjItYjY4NjNjNzMzNmIyIl0sImRldmljZWlkIjoiNTk2Y2UxZmEtOGZlYi00ZTI0LTlkZGItNjAzYzZkMzc0OGM2IiwiZmFtaWx5X25hbWUiOiJTaWx2YSIsImdpdmVuX25hbWUiOiJXZXNsbGV5IiwiaWR0eXAiOiJ1c2VyIiwiaXBhZGRyIjoiMjAxLjMxLjE5My4xNTAiLCJuYW1lIjoiV2VzbGxleSBMdWlzIFNpbHZhIiwib2lkIjoiNzJkOTVjNzgtOWIyNy00NWVkLWIzNjYtYzdmYmVkMTc0YmUzIiwib25wcmVtX3NpZCI6IlMtMS01LTIxLTc2MTk5NTM3NS0xNDkyMjg4ODcxLTIyMDI1NjkzMTktODk1ODkiLCJwbGF0ZiI6IjMiLCJwdWlkIjoiMTAwMzIwMDUyN0IzOENDNCIsInJoIjoiMS5BVFFBME5kazc5MzJqRUs4bUVudkdCSHB0UU1BQUFBQUFBQUF3QUFBQUFBQUFBQTBBSE0wQUEuIiwic2NwIjoiQ2FsZW5kYXJzLlJlYWRCYXNpYyBDaGFubmVsTWVtYmVyLlJlYWQuQWxsIENoYXQuUmVhZEJhc2ljIERpcmVjdG9yeS5SZWFkLkFsbCBlbWFpbCBGaWxlcy5SZWFkV3JpdGUuQWxsIEZpbGVTdG9yYWdlQ29udGFpbmVyLlNlbGVjdGVkIEdyb3VwLlJlYWRXcml0ZS5BbGwgR3JvdXBNZW1iZXIuUmVhZFdyaXRlLkFsbCBJbmZvcm1hdGlvblByb3RlY3Rpb25Qb2xpY3kuUmVhZCBvcGVuaWQgT3JnYW5pemF0aW9uLlJlYWQuQWxsIHByb2ZpbGUgU2Vuc2l0aXZpdHlMYWJlbC5SZWFkIFRhc2tzLlJlYWRXcml0ZSBVbmlmaWVkR3JvdXBNZW1iZXIuUmVhZC5Bc0d1ZXN0IFVzZXIuUmVhZC5BbGwgVXNlci5SZWFkQmFzaWMuQWxsIiwic2lkIjoiMDA5YWMxNDktMDYwMi00YzMyLWIwNzktNTYxNWQ3ZGEyN2E0Iiwic2lnbmluX3N0YXRlIjpbImR2Y19tbmdkIiwiZHZjX2NtcCIsImR2Y19kbWpkIiwiaW5rbm93bm50d2siLCJrbXNpIl0sInN1YiI6IjBubWpFcG5FeHBkYmRXX1dfSlZEOUQtajBnUzh3SldsT25MVGlFYmtCMFEiLCJ0ZW5hbnRfcmVnaW9uX3Njb3BlIjoiU0EiLCJ0aWQiOiJlZjY0ZDdkMC1mNmRkLTQyOGMtYmM5OC00OWVmMTgxMWU5YjUiLCJ1bmlxdWVfbmFtZSI6ImJwNTY5MDk0QGJwLm9yZy5iciIsInVwbiI6ImJwNTY5MDk0QGJwLm9yZy5iciIsInV0aSI6Iml4aFpWRHhmbjA2X2RjV24ybUpFQUEiLCJ2ZXIiOiIxLjAiLCJ3aWRzIjpbImI3OWZiZjRkLTNlZjktNDY4OS04MTQzLTc2YjE5NGU4NTUwOSJdLCJ4bXNfYWNkIjoxNjkwODY0MjkzLCJ4bXNfYWN0X2ZjdCI6IjMgNSIsInhtc19jYyI6WyJjcDEiXSwieG1zX2Z0ZCI6ImNmNk1iT0F3NnBTVWo4N29hWHBzQmczWEc0eDl6YnROdUhScmVONXBqeXdCZFhOdWIzSjBhQzFrYzIxeiIsInhtc19pZHJlbCI6IjEgMjYiLCJ4bXNfcGZ0ZXhwIjoxNzc0NDM3OTM1LCJ4bXNfc3NtIjoiMSIsInhtc19zdCI6eyJzdWIiOiJWdWdOU29LWncwLVZ0eUtBMFRNMVZHSTlnY0U2UEJsLS1QSjZjUkJlZkswIn0sInhtc19zdWJfZmN0IjoiMyAxMiIsInhtc190Y2R0IjoxNDMxMTE2MzAzLCJ4bXNfdG50X2ZjdCI6IjMgMiJ9.IRqfzGSh1gQqqYYJI0esAfaYfHKX6y87Wta0a0bf4U9jHnE_ieVk14pmhiW7OA_j_FF3qljGcbYQEJdfKdtR-GK0y4t5E1BTh_DkDLA7qZ6eQUmSi-XdYb32qDE_lgiW5yviE1b7fp0dLLA3VSV3DG4bjt5xYnzzgg5Mu1AaOlIiCR_OdtzC6SE2WKmTbJYzlQUpMc1yrZHT05yQPHGTbL91-Yt9GcY97kLeVU9uiAyglsWSf2o8auzZJYtNp2JtvGHFaIzTkWmOFz1il8eGxzomtI3IXmhpFIhVhsGbrwv6l_61FreKyc8pK2a-cKBxlhXr-TM6NjPS42tg8lMvkg'
def planner (token, plan_id, tipo):
    ACCESS_TOKEN  = token

    # ID do Plano que você quer acessar (Você pode obter isso via URL do Planner ou listando grupos)

    BASE_URL = 'https://graph.microsoft.com/v1.0'

    headers = {
        'Authorization': ACCESS_TOKEN,
        'Content-Type': 'application/json'
    }

    response = requests.get(f'{BASE_URL}/planner/plans/{plan_id}/{tipo}', headers=headers)

    if response.status_code == 200:
        print(response.json())
    else:
        print("Token expirado ou inválido via navegador:", response.status_code)

    return response.json()

def planner (token, plan_id, tipo):
    ACCESS_TOKEN  = token

    # ID do Plano que você quer acessar (Você pode obter isso via URL do Planner ou listando grupos)

    BASE_URL = 'https://graph.microsoft.com/v1.0'

    headers = {
        'Authorization': ACCESS_TOKEN,
        'Content-Type': 'application/json'
    }

    response = requests.get(f'{BASE_URL}/planner/plans/{plan_id}/{tipo}', headers=headers)

    if response.status_code == 200:
        print(response.json())
    else:
        print("Token expirado ou inválido via navegador:", response.status_code)

    return response.json()

def corrigir_marcos(linha):
    if linha['Bucket'].lower() == 'Marcos'.lower():
        # Extrai o número que está depois de 'Marco ' na coluna Card
        match = re.search(r'(?i)Marco (\d+)', str(linha['Card']))
        if match:
            numero_marco = match.group(1)
            # Retorna o nome da Entrega correspondente (se não achar, mantém 'Marcos')
            return mapeamento_entregas.get(numero_marco, linha['Bucket'])
    
    # Se não for 'Marcos', mantém o valor original
    return linha['Bucket']

############################################################################################################################################################
#Telenordeste
############################################################################################################################################################

PLAN_ID = '5gba_uRxB0CW_jU02ozC3WUABJth'

response = planner(ACCESS_TOKEN, PLAN_ID, 'buckets')
buckets_entregas = pd.DataFrame(response['value'])

response = planner(ACCESS_TOKEN, PLAN_ID, 'tasks')
tasks_atividade = pd.DataFrame(response['value'])

df_merge = pd.merge(
    buckets_entregas,
    tasks_atividade,
    left_on='id',
    right_on='bucketId',
    how='left'
    )

columns_exec_fisica = ['id_y', 'name', 'title', 'referenceCount', 'checklistItemCount', 'activeChecklistItemCount', 'startDateTime', 'dueDateTime', 'completedDateTime', 'priority', 'completedBy', 'appliedCategories']

exec_fisica_telenordeste = df_merge[columns_exec_fisica].copy()

percentual_atividades_incompletas = np.abs(np.divide(exec_fisica_telenordeste['activeChecklistItemCount'].astype('float'),exec_fisica_telenordeste['checklistItemCount'].astype('float')))*100
percentual_atividades_completas = np.abs((np.divide(exec_fisica_telenordeste['activeChecklistItemCount'].astype('float'),exec_fisica_telenordeste['checklistItemCount'].astype('float'))-1))*100
exec_fisica_telenordeste.loc[:, 'percentual_atividades_completas'] = np.round(percentual_atividades_completas,1)
exec_fisica_telenordeste.loc[:, 'percentual_atividades_incompletas'] = np.round(percentual_atividades_incompletas,1)

task_id_list = exec_fisica_telenordeste.loc[exec_fisica_telenordeste['id_y'].notnull(), 'id_y']

headers = {
        'Authorization': ACCESS_TOKEN,
        'Content-Type': 'application/json'
    }

df = pd.DataFrame()    
for task in task_id_list:
    task_link = f'https://graph.microsoft.com/v1.0/planner/tasks/{task}?$expand=details'
    response = requests.get(task_link, headers=headers)
    task_id = pd.DataFrame(response.json()['details'])['id']
    checklist = pd.DataFrame(response.json()['details']['checklist']).T

    task_checklist = pd.merge(
        task_id,
        checklist,
        left_index=True,
        right_index=True,
        how='left'
    )
    
    df = pd.concat([df, task_checklist])

exec_fisica_telenordeste_final = pd.merge(
    exec_fisica_telenordeste,
    df,
    left_on='id_y',
    right_on='id',
    how='left'
)

exec_fisica_telenordeste_final ['completedDateTime'] = pd.to_datetime(exec_fisica_telenordeste_final ['completedDateTime'])
exec_fisica_telenordeste_final ['startDateTime'] = pd.to_datetime(exec_fisica_telenordeste_final ['startDateTime'])
exec_fisica_telenordeste_final ['dueDateTime'] = pd.to_datetime(exec_fisica_telenordeste_final ['dueDateTime'])
exec_fisica_telenordeste_final['Tempo de execução'] = exec_fisica_telenordeste_final['completedDateTime'] - exec_fisica_telenordeste_final['startDateTime'] #.dt.tz_localize(None)

columns_name_dict = {
    'id_y': 'planner_id',
    'name': 'Bucket',
    'title_x': 'Card',
    'title_y': 'Atividade do checklist',
    'startDateTime': 'Data de início da execução',
    'dueDateTime': 'Data de conclusão prevista',
    'completedDateTime': 'Data de conclusão'
}

exec_fisica_telenordeste_final.rename(columns=columns_name_dict, inplace=True)

exec_fisica_telenordeste_final[exec_fisica_telenordeste_final['isChecked'].isnull()]

exec_fisica_telenordeste_final.loc[(exec_fisica_telenordeste_final['isChecked'].isnull()) & (exec_fisica_telenordeste_final['Data de conclusão'].isnull()), 'isChecked'] = 'Não concluído'

exec_fisica_telenordeste_final.loc[(exec_fisica_telenordeste_final['isChecked'].isnull()) & (exec_fisica_telenordeste_final['Data de conclusão'].notnull()), 'isChecked'] = 'Concluído'

exec_fisica_telenordeste_final = exec_fisica_telenordeste_final[exec_fisica_telenordeste_final['planner_id'].notnull()]

exec_fisica_telenordeste_final.loc[:, 'Projeto'] = 'Telenordeste'

# Criar um dicionário (mapeamento) com os nomes reais das Entregas
# Isso vai procurar todos os 'Buckets' que contêm 'Entrega', extrair o número e salvar o nome completo.
mapeamento_entregas = {}
for bucket in exec_fisica_telenordeste_final['Bucket'].unique():
    if 'Entrega'.lower() in str(bucket).lower():
        # Procura o número após a palavra "Entrega "
        match = re.search(r'(?i)Entrega (\d+)', str(bucket))
        if match:
            numero_entrega = match.group(1)
            mapeamento_entregas[numero_entrega] = bucket

# Aplicar a função ao DataFrame
exec_fisica_telenordeste_final['Bucket'] = exec_fisica_telenordeste_final.apply(corrigir_marcos, axis=1)

exec_fisica_telenordeste_final[exec_fisica_telenordeste_final['Card'].str.contains('Marco', case=False)]

############################################################################################################################################################
#DNA HPV
############################################################################################################################################################

PLAN_ID = 'OK5KMdtgqUyLxQp6U6TEXWUADLme'

response = planner(ACCESS_TOKEN, PLAN_ID, 'buckets')
buckets_entregas = pd.DataFrame(response['value'])

response = planner(ACCESS_TOKEN, PLAN_ID, 'tasks')
tasks_atividade = pd.DataFrame(response['value'])

df_merge = pd.merge(
    buckets_entregas,
    tasks_atividade,
    left_on='id',
    right_on='bucketId',
    how='left'
    )

columns_exec_fisica = ['id_y', 'name', 'title', 'referenceCount', 'checklistItemCount', 'activeChecklistItemCount', 'startDateTime', 'dueDateTime', 'completedDateTime', 'priority', 'completedBy', 'appliedCategories']

exec_fisica_dna = df_merge[columns_exec_fisica].copy()

percentual_atividades_incompletas = np.abs(np.divide(exec_fisica_dna['activeChecklistItemCount'].astype('float'),exec_fisica_dna['checklistItemCount'].astype('float')))*100
percentual_atividades_completas = np.abs((np.divide(exec_fisica_dna['activeChecklistItemCount'].astype('float'),exec_fisica_dna['checklistItemCount'].astype('float'))-1))*100
exec_fisica_dna.loc[:, 'percentual_atividades_completas'] = np.round(percentual_atividades_completas,1)
exec_fisica_dna.loc[:, 'percentual_atividades_incompletas'] = np.round(percentual_atividades_incompletas,1)

task_id_list = exec_fisica_dna.loc[exec_fisica_dna['id_y'].notnull(), 'id_y']

headers = {
        'Authorization': ACCESS_TOKEN,
        'Content-Type': 'application/json'
    }

df = pd.DataFrame()    
for task in task_id_list:
    task_link = f'https://graph.microsoft.com/v1.0/planner/tasks/{task}?$expand=details'
    response = requests.get(task_link, headers=headers)
    task_id = pd.DataFrame(response.json()['details'])['id']
    checklist = pd.DataFrame(response.json()['details']['checklist']).T

    task_checklist = pd.merge(
        task_id,
        checklist,
        left_index=True,
        right_index=True,
        how='left'
    )
    
    df = pd.concat([df, task_checklist])

exec_fisica_dna_final = pd.merge(
    exec_fisica_dna,
    df,
    left_on='id_y',
    right_on='id',
    how='left'
)

exec_fisica_dna_final ['completedDateTime'] = pd.to_datetime(exec_fisica_dna_final ['completedDateTime'])
exec_fisica_dna_final ['startDateTime'] = pd.to_datetime(exec_fisica_dna_final ['startDateTime'])
exec_fisica_dna_final ['dueDateTime'] = pd.to_datetime(exec_fisica_dna_final ['dueDateTime'])
exec_fisica_dna_final['Tempo de execução'] = exec_fisica_dna_final['completedDateTime'] - exec_fisica_dna_final['startDateTime'].dt.tz_localize(None)

columns_name_dict = {
    'id_y': 'planner_id',
    'name': 'Bucket',
    'title_x': 'Card',
    'title_y': 'Atividade do checklist',
    'startDateTime': 'Data de início da execução',
    'dueDateTime': 'Data de conclusão prevista',
    'completedDateTime': 'Data de conclusão'
}

exec_fisica_dna_final.rename(columns=columns_name_dict, inplace=True)

exec_fisica_dna_final[exec_fisica_dna_final['isChecked'].isnull()]

exec_fisica_dna_final.loc[(exec_fisica_dna_final['isChecked'].isnull()) & (exec_fisica_dna_final['Data de conclusão'].isnull()), 'isChecked'] = 'Não concluído'

exec_fisica_dna_final.loc[(exec_fisica_dna_final['isChecked'].isnull()) & (exec_fisica_dna_final['Data de conclusão'].notnull()), 'isChecked'] = 'Concluído'

exec_fisica_dna_final = exec_fisica_dna_final[exec_fisica_dna_final['planner_id'].notnull()]

exec_fisica_dna_final.loc[:, 'Projeto'] = 'DNA HPV'

# 2. Criar um dicionário (mapeamento) com os nomes reais das Entregas
# Isso vai procurar todos os 'Buckets' que contêm 'Entrega', extrair o número e salvar o nome completo.
mapeamento_entregas = {}
for bucket in exec_fisica_dna_final['Bucket'].unique():
    if 'Entrega'.lower() in str(bucket).lower():
        # Procura o número após a palavra "Entrega "
        match = re.search(r'(?i)Entrega (\d+)', str(bucket))
        if match:
            numero_entrega = match.group(1)
            mapeamento_entregas[numero_entrega] = bucket

# Aplicar a função ao DataFrame
exec_fisica_dna_final['Bucket'] = exec_fisica_dna_final.apply(corrigir_marcos, axis=1)

exec_fisica_dna_final[exec_fisica_dna_final['Card'].str.contains('Marco', case=False)]

############################################################################################################################################################
#Boas práticas
############################################################################################################################################################

PLAN_ID = '7Sy7DE_GjUq9e8fYH880HmUAAuEw'

response = planner(ACCESS_TOKEN, PLAN_ID, 'buckets')
buckets_entregas = pd.DataFrame(response['value'])

response = planner(ACCESS_TOKEN, PLAN_ID, 'tasks')
tasks_atividade = pd.DataFrame(response['value'])

df_merge = pd.merge(
    buckets_entregas,
    tasks_atividade,
    left_on='id',
    right_on='bucketId',
    how='left'
    )

columns_exec_fisica = ['id_y', 'name', 'title', 'referenceCount', 'checklistItemCount', 'activeChecklistItemCount', 'startDateTime', 'dueDateTime', 'completedDateTime', 'priority', 'completedBy', 'appliedCategories']

exec_fisica_boas = df_merge[columns_exec_fisica].copy()

percentual_atividades_incompletas = np.abs(np.divide(exec_fisica_boas['activeChecklistItemCount'].astype('float'),exec_fisica_boas['checklistItemCount'].astype('float')))*100
percentual_atividades_completas = np.abs((np.divide(exec_fisica_boas['activeChecklistItemCount'].astype('float'),exec_fisica_boas['checklistItemCount'].astype('float'))-1))*100
exec_fisica_boas.loc[:, 'percentual_atividades_completas'] = np.round(percentual_atividades_completas,1)
exec_fisica_boas.loc[:, 'percentual_atividades_incompletas'] = np.round(percentual_atividades_incompletas,1)

task_id_list = exec_fisica_boas.loc[exec_fisica_boas['id_y'].notnull(), 'id_y']

headers = {
        'Authorization': ACCESS_TOKEN,
        'Content-Type': 'application/json'
    }

df = pd.DataFrame()    
for task in task_id_list:
    task_link = f'https://graph.microsoft.com/v1.0/planner/tasks/{task}?$expand=details'
    response = requests.get(task_link, headers=headers)
    task_id = pd.DataFrame(response.json()['details'])['id']
    checklist = pd.DataFrame(response.json()['details']['checklist']).T

    task_checklist = pd.merge(
        task_id,
        checklist,
        left_index=True,
        right_index=True,
        how='left'
    )
    
    df = pd.concat([df, task_checklist])

exec_fisica_boas_final = pd.merge(
    exec_fisica_boas,
    df,
    left_on='id_y',
    right_on='id',
    how='left'
)

exec_fisica_boas_final ['completedDateTime'] = pd.to_datetime(exec_fisica_boas_final ['completedDateTime'])
exec_fisica_boas_final ['startDateTime'] = pd.to_datetime(exec_fisica_boas_final ['startDateTime'])
exec_fisica_boas_final ['dueDateTime'] = pd.to_datetime(exec_fisica_boas_final ['dueDateTime'])
exec_fisica_boas_final['Tempo de execução'] = exec_fisica_boas_final['completedDateTime'] - exec_fisica_boas_final['startDateTime']

columns_name_dict = {
    'id_y': 'planner_id',
    'name': 'Bucket',
    'title_x': 'Card',
    'title_y': 'Atividade do checklist',
    'startDateTime': 'Data de início da execução',
    'dueDateTime': 'Data de conclusão prevista',
    'completedDateTime': 'Data de conclusão'
}

exec_fisica_boas_final.rename(columns=columns_name_dict, inplace=True)

exec_fisica_boas_final[exec_fisica_boas_final['isChecked'].isnull()]

exec_fisica_boas_final.loc[(exec_fisica_boas_final['isChecked'].isnull()) & (exec_fisica_boas_final['Data de conclusão'].isnull()), 'isChecked'] = 'Não concluído'

exec_fisica_boas_final.loc[(exec_fisica_boas_final['isChecked'].isnull()) & (exec_fisica_boas_final['Data de conclusão'].notnull()), 'isChecked'] = 'Concluído'

exec_fisica_boas_final = exec_fisica_boas_final[exec_fisica_boas_final['planner_id'].notnull()]

exec_fisica_boas_final.loc[:, 'Projeto'] = 'Boas práticas'

# Criar um dicionário (mapeamento) com os nomes reais das Entregas
# Isso vai procurar todos os 'Buckets' que contêm 'Entrega', extrair o número e salvar o nome completo.
mapeamento_entregas = {}
for bucket in exec_fisica_boas_final['Bucket'].unique():
    if 'Entrega'.lower() in str(bucket).lower():
        # Procura o número após a palavra "Entrega "
        match = re.search(r'(?i)Entrega (\d+)', str(bucket))
        if match:
            numero_entrega = match.group(1)
            mapeamento_entregas[numero_entrega] = bucket

# Aplicar a função ao DataFrame
exec_fisica_boas_final['Bucket'] = exec_fisica_boas_final.apply(corrigir_marcos, axis=1)

exec_fisica_boas_final[exec_fisica_boas_final['Card'].str.contains('Marco', case=False)]

############################################################################################################################################################
#Aprimora
############################################################################################################################################################

PLAN_ID = 'aL-u7J8hWUqXvQ3BRmkDXmUAGTlE'

response = planner(ACCESS_TOKEN, PLAN_ID, 'buckets')
buckets_entregas = pd.DataFrame(response['value'])

response = planner(ACCESS_TOKEN, PLAN_ID, 'tasks')
tasks_atividade = pd.DataFrame(response['value'])

df_merge = pd.merge(
    buckets_entregas,
    tasks_atividade,
    left_on='id',
    right_on='bucketId',
    how='left'
    )

columns_exec_fisica = ['id_y', 'name', 'title', 'referenceCount', 'checklistItemCount', 'activeChecklistItemCount', 'startDateTime', 'dueDateTime', 'completedDateTime', 'priority', 'completedBy', 'appliedCategories']

exec_fisica_aprimora = df_merge[columns_exec_fisica].copy()

percentual_atividades_incompletas = np.abs(np.divide(exec_fisica_aprimora['activeChecklistItemCount'].astype('float'),exec_fisica_aprimora['checklistItemCount'].astype('float')))*100
percentual_atividades_completas = np.abs((np.divide(exec_fisica_aprimora['activeChecklistItemCount'].astype('float'),exec_fisica_aprimora['checklistItemCount'].astype('float'))-1))*100
exec_fisica_aprimora.loc[:, 'percentual_atividades_completas'] = np.round(percentual_atividades_completas,1)
exec_fisica_aprimora.loc[:, 'percentual_atividades_incompletas'] = np.round(percentual_atividades_incompletas,1)

task_id_list = exec_fisica_aprimora['id_y'].unique()

headers = {
        'Authorization': ACCESS_TOKEN,
        'Content-Type': 'application/json'
    }

df = pd.DataFrame()    
for task in task_id_list:
    task_link = f'https://graph.microsoft.com/v1.0/planner/tasks/{task}?$expand=details'
    response = requests.get(task_link, headers=headers)
    task_id = pd.DataFrame(response.json()['details'])['id']
    checklist = pd.DataFrame(response.json()['details']['checklist']).T

    task_checklist = pd.merge(
        task_id,
        checklist,
        left_index=True,
        right_index=True,
        how='left'
    )
    
    df = pd.concat([df, task_checklist])

exec_fisica_aprimora_final = pd.merge(
    exec_fisica_aprimora,
    df,
    left_on='id_y',
    right_on='id',
    how='left'
)

exec_fisica_aprimora_final ['completedDateTime'] = pd.to_datetime(exec_fisica_aprimora_final ['completedDateTime'])
exec_fisica_aprimora_final ['startDateTime'] = pd.to_datetime(exec_fisica_aprimora_final ['startDateTime'])
exec_fisica_aprimora_final ['dueDateTime'] = pd.to_datetime(exec_fisica_aprimora_final ['dueDateTime'])
exec_fisica_aprimora_final['Tempo de execução'] = exec_fisica_aprimora_final ['completedDateTime'] - exec_fisica_aprimora_final['startDateTime']

exec_fisica_aprimora_final[exec_fisica_aprimora_final['title_x'].str.contains('MARCO 1', case=False)]

columns_name_dict = {
    'id_y': 'planner_id',
    'name': 'Bucket',
    'title_x': 'Card',
    'title_y': 'Atividade do checklist',
    'startDateTime': 'Data de início da execução',
    'dueDateTime': 'Data de conclusão prevista',
    'completedDateTime': 'Data de conclusão'
}

exec_fisica_aprimora_final.rename(columns=columns_name_dict, inplace=True)

exec_fisica_aprimora_final[exec_fisica_aprimora_final['isChecked'].isnull()]

exec_fisica_aprimora_final.loc[(exec_fisica_aprimora_final['isChecked'].isnull()) & (exec_fisica_aprimora_final['Data de conclusão'].isnull()), 'isChecked'] = 'Não concluído'

exec_fisica_aprimora_final.loc[(exec_fisica_aprimora_final['isChecked'].isnull()) & (exec_fisica_aprimora_final['Data de conclusão'].notnull()), 'isChecked'] = 'Concluído'

exec_fisica_aprimora_final.loc[:, 'Projeto'] = 'Aprimora SUS'

# Criar um dicionário (mapeamento) com os nomes reais das Entregas
# Isso vai procurar todos os 'Buckets' que contêm 'Entrega', extrair o número e salvar o nome completo.
mapeamento_entregas = {}
for bucket in exec_fisica_aprimora_final['Bucket'].unique():
    if 'Entrega'.lower() in str(bucket).lower():
        # Procura o número após a palavra "Entrega "
        match = re.search(r'(?i)Entrega (\d+)', str(bucket))
        if match:
            numero_entrega = match.group(1)
            mapeamento_entregas[numero_entrega] = bucket

# Aplicar a função ao DataFrame
exec_fisica_aprimora_final['Bucket'] = exec_fisica_aprimora_final.apply(corrigir_marcos, axis=1)

exec_fisica_aprimora_final[exec_fisica_aprimora_final['Card'].str.contains('Marco', case=False)]


############################################################################################################################################################
#Qualiguia APS
############################################################################################################################################################

PLAN_ID = 'i1mcchiJ4E28SizHvS_vUmUACIBy'

response = planner(ACCESS_TOKEN, PLAN_ID, 'buckets')
buckets_entregas = pd.DataFrame(response['value'])

response = planner(ACCESS_TOKEN, PLAN_ID, 'tasks')
tasks_atividade = pd.DataFrame(response['value'])


df_merge = pd.merge(
    buckets_entregas,
    tasks_atividade,
    left_on='id',
    right_on='bucketId',
    how='left'
    )

columns_exec_fisica = ['id_y', 'name', 'title', 'referenceCount', 'checklistItemCount', 'activeChecklistItemCount', 'startDateTime', 'dueDateTime', 'completedDateTime', 'priority', 'completedBy', 'appliedCategories']

exec_fisica_qg_aps = df_merge[columns_exec_fisica].copy()

percentual_atividades_incompletas = np.abs(np.divide(exec_fisica_qg_aps['activeChecklistItemCount'].astype('float'),exec_fisica_qg_aps['checklistItemCount'].astype('float')))*100
percentual_atividades_completas = np.abs((np.divide(exec_fisica_qg_aps['activeChecklistItemCount'].astype('float'),exec_fisica_qg_aps['checklistItemCount'].astype('float'))-1))*100
exec_fisica_qg_aps.loc[:, 'percentual_atividades_completas'] = np.round(percentual_atividades_completas,1)
exec_fisica_qg_aps.loc[:, 'percentual_atividades_incompletas'] = np.round(percentual_atividades_incompletas,1)


task_id_list = exec_fisica_qg_aps.loc[exec_fisica_qg_aps['id_y'].notnull(), 'id_y']

headers = {
        'Authorization': ACCESS_TOKEN,
        'Content-Type': 'application/json'
    }

df = pd.DataFrame()

for task in task_id_list:
    task_link = f'https://graph.microsoft.com/v1.0/planner/tasks/{task}?$expand=details'
    response = requests.get(task_link, headers=headers)
    task_id = pd.DataFrame(response.json()['details'])['id']
    checklist = pd.DataFrame(response.json()['details']['checklist']).T

    task_checklist = pd.merge(
        task_id,
        checklist,
        left_index=True,
        right_index=True,
        how='left'
    )
    
    df = pd.concat([df, task_checklist])


exec_fisica_qg_aps_final = pd.merge(
    exec_fisica_qg_aps,
    df,
    left_on='id_y',
    right_on='id',
    how='left'
)

exec_fisica_qg_aps_final ['completedDateTime'] = pd.to_datetime(exec_fisica_qg_aps_final ['completedDateTime'])
exec_fisica_qg_aps_final ['startDateTime'] = pd.to_datetime(exec_fisica_qg_aps_final ['startDateTime'])
exec_fisica_qg_aps_final ['dueDateTime'] = pd.to_datetime(exec_fisica_qg_aps_final ['dueDateTime'])
exec_fisica_qg_aps_final['Tempo de execução'] = exec_fisica_qg_aps_final['completedDateTime'] - exec_fisica_qg_aps_final['startDateTime']

columns_name_dict = {
    'id_y': 'planner_id',
    'name': 'Bucket',
    'title_x': 'Card',
    'title_y': 'Atividade do checklist',
    'startDateTime': 'Data de início da execução',
    'dueDateTime': 'Data de conclusão prevista',
    'completedDateTime': 'Data de conclusão'
}

exec_fisica_qg_aps_final.rename(columns=columns_name_dict, inplace=True)

exec_fisica_qg_aps_final[exec_fisica_qg_aps_final['isChecked'].isnull()]

exec_fisica_qg_aps_final.loc[(exec_fisica_qg_aps_final['isChecked'].isnull()) & (exec_fisica_qg_aps_final['Data de conclusão'].isnull()), 'isChecked'] = 'Não concluído'

exec_fisica_qg_aps_final.loc[(exec_fisica_qg_aps_final['isChecked'].isnull()) & (exec_fisica_qg_aps_final['Data de conclusão'].notnull()), 'isChecked'] = 'Concluído'

exec_fisica_qg_aps_final = exec_fisica_qg_aps_final[exec_fisica_qg_aps_final['planner_id'].notnull()]

exec_fisica_qg_aps_final.loc[:, 'Projeto'] = 'Qualiguia APS'

# Criar um dicionário (mapeamento) com os nomes reais das Entregas
# Isso vai procurar todos os 'Buckets' que contêm 'Entrega', extrair o número e salvar o nome completo.
mapeamento_entregas = {}
for bucket in exec_fisica_qg_aps_final['Bucket'].unique():
    if 'Entrega'.lower() in str(bucket).lower():
        # Procura o número após a palavra "Entrega "
        match = re.search(r'(?i)Entrega (\d+)', str(bucket))
        if match:
            numero_entrega = match.group(1)
            mapeamento_entregas[numero_entrega] = bucket

# Aplicar a função ao DataFrame
exec_fisica_qg_aps_final['Bucket'] = exec_fisica_qg_aps_final.apply(corrigir_marcos, axis=1)

exec_fisica_qg_aps_final[exec_fisica_qg_aps_final['Card'].str.contains('Marco', case=False)]

############################################################################################################################################################
#Qualiguia Hospitalar
############################################################################################################################################################

PLAN_ID = 'S_hW5HZYLEKpeqBoDI4d52UAF1Ou'

response = planner(ACCESS_TOKEN, PLAN_ID, 'buckets')
buckets_entregas = pd.DataFrame(response['value'])

response = planner(ACCESS_TOKEN, PLAN_ID, 'tasks')
tasks_atividade = pd.DataFrame(response['value'])

df_merge = pd.merge(
    buckets_entregas,
    tasks_atividade,
    left_on='id',
    right_on='bucketId',
    how='left'
    )

columns_exec_fisica = ['id_y', 'name', 'title', 'referenceCount', 'checklistItemCount', 'activeChecklistItemCount', 'startDateTime', 'dueDateTime', 'completedDateTime', 'priority', 'completedBy', 'appliedCategories']

exec_fisica_hospitalar = df_merge[columns_exec_fisica].copy()

percentual_atividades_incompletas = np.abs(np.divide(exec_fisica_hospitalar['activeChecklistItemCount'].astype('float'),exec_fisica_hospitalar['checklistItemCount'].astype('float')))*100
percentual_atividades_completas = np.abs((np.divide(exec_fisica_hospitalar['activeChecklistItemCount'].astype('float'),exec_fisica_hospitalar['checklistItemCount'].astype('float'))-1))*100
exec_fisica_hospitalar.loc[:, 'percentual_atividades_completas'] = np.round(percentual_atividades_completas,1)
exec_fisica_hospitalar.loc[:, 'percentual_atividades_incompletas'] = np.round(percentual_atividades_incompletas,1)

task_id_list = exec_fisica_hospitalar.loc[exec_fisica_hospitalar['id_y'].notnull(), 'id_y']

headers = {
        'Authorization': ACCESS_TOKEN,
        'Content-Type': 'application/json'
    }

df = pd.DataFrame()    
for task in task_id_list:
    task_link = f'https://graph.microsoft.com/v1.0/planner/tasks/{task}?$expand=details'
    response = requests.get(task_link, headers=headers)
    task_id = pd.DataFrame(response.json()['details'])['id']
    checklist = pd.DataFrame(response.json()['details']['checklist']).T

    task_checklist = pd.merge(
        task_id,
        checklist,
        left_index=True,
        right_index=True,
        how='left'
    )
    
    df = pd.concat([df, task_checklist])

exec_fisica_hospitalar_final = pd.merge(
    exec_fisica_hospitalar,
    df,
    left_on='id_y',
    right_on='id',
    how='left'
)

exec_fisica_hospitalar_final ['completedDateTime'] = pd.to_datetime(exec_fisica_hospitalar_final ['completedDateTime'])
exec_fisica_hospitalar_final ['startDateTime'] = pd.to_datetime(exec_fisica_hospitalar_final ['startDateTime'])
exec_fisica_hospitalar_final ['dueDateTime'] = pd.to_datetime(exec_fisica_hospitalar_final ['dueDateTime'])
exec_fisica_hospitalar_final['Tempo de execução'] = exec_fisica_hospitalar_final['completedDateTime'] - exec_fisica_hospitalar_final['startDateTime']

columns_name_dict = {
    'id_y': 'planner_id',
    'name': 'Bucket',
    'title_x': 'Card',
    'title_y': 'Atividade do checklist',
    'startDateTime': 'Data de início da execução',
    'dueDateTime': 'Data de conclusão prevista',
    'completedDateTime': 'Data de conclusão'
}

exec_fisica_hospitalar_final.rename(columns=columns_name_dict, inplace=True)

exec_fisica_hospitalar_final[exec_fisica_hospitalar_final['isChecked'].isnull()]

exec_fisica_hospitalar_final.loc[(exec_fisica_hospitalar_final['isChecked'].isnull()) & (exec_fisica_hospitalar_final['Data de conclusão'].isnull()), 'isChecked'] = 'Não concluído'

exec_fisica_hospitalar_final.loc[(exec_fisica_hospitalar_final['isChecked'].isnull()) & (exec_fisica_hospitalar_final['Data de conclusão'].notnull()), 'isChecked'] = 'Concluído'

exec_fisica_hospitalar_final = exec_fisica_hospitalar_final[exec_fisica_hospitalar_final['planner_id'].notnull()]

exec_fisica_hospitalar_final.loc[:, 'Projeto'] = 'Qualiguia Hospitalar'

# Criar um dicionário (mapeamento) com os nomes reais das Entregas
# Isso vai procurar todos os 'Buckets' que contêm 'Entrega', extrair o número e salvar o nome completo.
mapeamento_entregas = {}
for bucket in exec_fisica_hospitalar_final['Bucket'].unique():
    if 'Entrega'.lower() in str(bucket).lower():
        # Procura o número após a palavra "Entrega "
        match = re.search(r'(?i)Entrega (\d+)', str(bucket))
        if match:
            numero_entrega = match.group(1)
            mapeamento_entregas[numero_entrega] = bucket

# Aplicar a função ao DataFrame
exec_fisica_hospitalar_final['Bucket'] = exec_fisica_hospitalar_final.apply(corrigir_marcos, axis=1)

exec_fisica_hospitalar_final[exec_fisica_hospitalar_final['Card'].str.contains('Marco', case=False)]

############################################################################################################################################################
#Data frame final
############################################################################################################################################################

exec_fisica = pd.concat([exec_fisica_aprimora_final, exec_fisica_dna_final, exec_fisica_boas_final, exec_fisica_telenordeste_final, exec_fisica_qg_aps_final, exec_fisica_hospitalar_final])

#%%
exec_fisica