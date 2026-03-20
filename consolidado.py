#%%
import requests
import msal
import json
import pandas as pd
import numpy as np

# --- CONFIGURAÇÕES ---
# CLIENT_ID = 'SEU_CLIENT_ID_AQUI'
# CLIENT_SECRET = 'SEU_CLIENT_SECRET_AQUI'
# TENANT_ID = 'SEU_TENANT_ID_AQUI'
ACCESS_TOKEN = 'Bearer eyJ0eXAiOiJKV1QiLCJub25jZSI6Ii1mSGpJcWNxSjg0Q3NpZFVsZ1U2YjNMVEI4TURfWWNkMGdVeUFSVktCRmMiLCJhbGciOiJSUzI1NiIsIng1dCI6IlFaZ045SHFOa0dORU00R2VLY3pEMDJQY1Z2NCIsImtpZCI6IlFaZ045SHFOa0dORU00R2VLY3pEMDJQY1Z2NCJ9.eyJhdWQiOiJodHRwczovL2dyYXBoLm1pY3Jvc29mdC5jb20iLCJpc3MiOiJodHRwczovL3N0cy53aW5kb3dzLm5ldC9lZjY0ZDdkMC1mNmRkLTQyOGMtYmM5OC00OWVmMTgxMWU5YjUvIiwiaWF0IjoxNzc0MDI5NjA4LCJuYmYiOjE3NzQwMjk2MDgsImV4cCI6MTc3NDExNjMwOCwiYWNjdCI6MCwiYWNyIjoiMSIsImFjcnMiOlsicDEiXSwiYWlvIjoiQVhRQWkvOGJBQUFBV3VBVkJPSDFxWWpja0hCUHUrUzR6T0VXSENZakRBNitETVIySDZjMWhVdWlWZVZLTjkrRmxDV0grZTVTd09JTnN6bVd1ekJCajNPMTdrZm13aC9ZNHB5NzUrTnZNaTVPME55ZzZuMWhUWWVGSDkwdmJjOXdtdVhDdFYzSS8wMUFGalVoV213VkRWOUEwUTRRTHBjRG5RPT0iLCJhbXIiOlsicHdkIiwicnNhIiwibWZhIl0sImFwcF9kaXNwbGF5bmFtZSI6Ik1pY3Jvc29mdCBQbGFubmVyIENsaWVudCIsImFwcGlkIjoiNzVmMzE3OTctMzdjOS00OThlLThkYzktNTNjMTZhMzZhZmNhIiwiYXBwaWRhY3IiOiIwIiwiY2Fwb2xpZHNfbGF0ZWJpbmQiOlsiYmQyZWZkNDEtNWM2Zi00MGQ5LTk3ZjItYjY4NjNjNzMzNmIyIl0sImRldmljZWlkIjoiNTk2Y2UxZmEtOGZlYi00ZTI0LTlkZGItNjAzYzZkMzc0OGM2IiwiZmFtaWx5X25hbWUiOiJTaWx2YSIsImdpdmVuX25hbWUiOiJXZXNsbGV5IiwiaWR0eXAiOiJ1c2VyIiwiaXBhZGRyIjoiMjAwLjIxMy44MC4xNDYiLCJuYW1lIjoiV2VzbGxleSBMdWlzIFNpbHZhIiwib2lkIjoiNzJkOTVjNzgtOWIyNy00NWVkLWIzNjYtYzdmYmVkMTc0YmUzIiwib25wcmVtX3NpZCI6IlMtMS01LTIxLTc2MTk5NTM3NS0xNDkyMjg4ODcxLTIyMDI1NjkzMTktODk1ODkiLCJwbGF0ZiI6IjMiLCJwdWlkIjoiMTAwMzIwMDUyN0IzOENDNCIsInJoIjoiMS5BVFFBME5kazc5MzJqRUs4bUVudkdCSHB0UU1BQUFBQUFBQUF3QUFBQUFBQUFBQTBBSE0wQUEuIiwic2NwIjoiQ2FsZW5kYXJzLlJlYWRCYXNpYyBDaGFubmVsTWVtYmVyLlJlYWQuQWxsIENoYXQuUmVhZEJhc2ljIERpcmVjdG9yeS5SZWFkLkFsbCBlbWFpbCBGaWxlcy5SZWFkV3JpdGUuQWxsIEZpbGVTdG9yYWdlQ29udGFpbmVyLlNlbGVjdGVkIEdyb3VwLlJlYWRXcml0ZS5BbGwgR3JvdXBNZW1iZXIuUmVhZFdyaXRlLkFsbCBJbmZvcm1hdGlvblByb3RlY3Rpb25Qb2xpY3kuUmVhZCBvcGVuaWQgT3JnYW5pemF0aW9uLlJlYWQuQWxsIHByb2ZpbGUgU2Vuc2l0aXZpdHlMYWJlbC5SZWFkIFRhc2tzLlJlYWRXcml0ZSBVbmlmaWVkR3JvdXBNZW1iZXIuUmVhZC5Bc0d1ZXN0IFVzZXIuUmVhZC5BbGwgVXNlci5SZWFkQmFzaWMuQWxsIiwic2lkIjoiMDA5YWMxNDktMDYwMi00YzMyLWIwNzktNTYxNWQ3ZGEyN2E0Iiwic2lnbmluX3N0YXRlIjpbImR2Y19tbmdkIiwiZHZjX2NtcCIsImR2Y19kbWpkIiwia21zaSJdLCJzdWIiOiIwbm1qRXBuRXhwZGJkV19XX0pWRDlELWowZ1M4d0pXbE9uTFRpRWJrQjBRIiwidGVuYW50X3JlZ2lvbl9zY29wZSI6IlNBIiwidGlkIjoiZWY2NGQ3ZDAtZjZkZC00MjhjLWJjOTgtNDllZjE4MTFlOWI1IiwidW5pcXVlX25hbWUiOiJicDU2OTA5NEBicC5vcmcuYnIiLCJ1cG4iOiJicDU2OTA5NEBicC5vcmcuYnIiLCJ1dGkiOiJUamNYa3ZzNVprLWczN1UtRGdNZkFBIiwidmVyIjoiMS4wIiwid2lkcyI6WyJiNzlmYmY0ZC0zZWY5LTQ2ODktODE0My03NmIxOTRlODU1MDkiXSwieG1zX2FjZCI6MTY5MDg2NDI5MywieG1zX2FjdF9mY3QiOiIzIDUiLCJ4bXNfY2MiOlsiY3AxIl0sInhtc19mdGQiOiJIN1A3N0FKVDJXaE1IWW1YaDFkZjBGdjF1bHBDbTVCNkFUakZESHdfRFJnQmRYTnpiM1YwYUMxa2MyMXoiLCJ4bXNfaWRyZWwiOiIxIDYiLCJ4bXNfcGZ0ZXhwIjoxNzc0MjAyNzA4LCJ4bXNfc3NtIjoiMSIsInhtc19zdCI6eyJzdWIiOiJWdWdOU29LWncwLVZ0eUtBMFRNMVZHSTlnY0U2UEJsLS1QSjZjUkJlZkswIn0sInhtc19zdWJfZmN0IjoiMyAxOCIsInhtc190Y2R0IjoxNDMxMTE2MzAzLCJ4bXNfdG50X2ZjdCI6IjMgOCJ9.VlHsVZR6ZsuRYRQC_wWI7a9NAvswpDPBNz5l4Bza4TNtqz_lpjRnku4pxqQoavYWT_GJWGuug_jXLE3fk4S04Sw_4ff-2ZmYfH1-QIYR22mZp-ZLvOAjXSJjwd5HtMMHrt8x-fJcO-sFszcIp-8TccMrzkuImfhYAV9i1V4KONMS0oe8CuLBFenZxTs_5ztJ7ELNC2oQX1EBXQA617gbgMKb-UcE7y9Z3BhPFxH39ZActCCljAY5-DUWd9CTPp4boL2dPKFktEAfqGQl4xX9wlz_cCk92Azv5J2BzIxNw0CVatRWj29bA1NHGv1iqAr1GK1eFEhqFFAT7JXvq7kDzQ'
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

exec_fisica = pd.concat([exec_fisica_aprimora_final, exec_fisica_dna_final, exec_fisica_boas_final, exec_fisica_aprimora_final, exec_fisica_qg_aps_final, exec_fisica_hospitalar_final])