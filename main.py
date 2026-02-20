#%%
import requests
import msal
import json
import pandas as pd
import numpy as np
# --- CONFIGURAÇÕES ---
CLIENT_ID = 'SEU_CLIENT_ID_AQUI'
CLIENT_SECRET = 'SEU_CLIENT_SECRET_AQUI'
TENANT_ID = 'SEU_TENANT_ID_AQUI'

ACCESS_TOKEN = 'Bearer eyJ0eXAiOiJKV1QiLCJub25jZSI6IlZFTjZGVkQ3VXIycjFFdzZ0ZWlBejA4Qi1UUFpuMnZqbHBSWFZXMXFwUGsiLCJhbGciOiJSUzI1NiIsIng1dCI6InNNMV95QXhWOEdWNHlOLUI2ajJ4em1pazVBbyIsImtpZCI6InNNMV95QXhWOEdWNHlOLUI2ajJ4em1pazVBbyJ9.eyJhdWQiOiJodHRwczovL2dyYXBoLm1pY3Jvc29mdC5jb20iLCJpc3MiOiJodHRwczovL3N0cy53aW5kb3dzLm5ldC9lZjY0ZDdkMC1mNmRkLTQyOGMtYmM5OC00OWVmMTgxMWU5YjUvIiwiaWF0IjoxNzcxNjE0MDExLCJuYmYiOjE3NzE2MTQwMTEsImV4cCI6MTc3MTcwMDcxMSwiYWNjdCI6MCwiYWNyIjoiMSIsImFjcnMiOlsicDEiXSwiYWlvIjoiQVhRQWkvOGJBQUFBaGs0cWRxZzY5REtDZE9xbkdhK0p3dlJZRjFJK1kyTXlaNmRWZnBuUjErZk9RdWptdElZZFdlbHZjTE1rWHl4UVEwSHUrZGJyeSt5SWFJSDFNWVZKWkttK3NKNEtKNnByTUdCeFUwUDZZUHpybG1WeG1EVFZiNkVIYlZBa2dCbXA4a0EzK21JV3UxRWw4cms3Y0tUVStRPT0iLCJhbXIiOlsicHdkIiwicnNhIiwibWZhIl0sImFwcF9kaXNwbGF5bmFtZSI6Ik1pY3Jvc29mdCBQbGFubmVyIENsaWVudCIsImFwcGlkIjoiNzVmMzE3OTctMzdjOS00OThlLThkYzktNTNjMTZhMzZhZmNhIiwiYXBwaWRhY3IiOiIwIiwiY2Fwb2xpZHNfbGF0ZWJpbmQiOlsiYmQyZWZkNDEtNWM2Zi00MGQ5LTk3ZjItYjY4NjNjNzMzNmIyIl0sImRldmljZWlkIjoiNTk2Y2UxZmEtOGZlYi00ZTI0LTlkZGItNjAzYzZkMzc0OGM2IiwiZmFtaWx5X25hbWUiOiJTaWx2YSIsImdpdmVuX25hbWUiOiJXZXNsbGV5IiwiaWR0eXAiOiJ1c2VyIiwiaXBhZGRyIjoiMjAxLjMxLjE5My4xNTAiLCJuYW1lIjoiV2VzbGxleSBMdWlzIFNpbHZhIiwib2lkIjoiNzJkOTVjNzgtOWIyNy00NWVkLWIzNjYtYzdmYmVkMTc0YmUzIiwib25wcmVtX3NpZCI6IlMtMS01LTIxLTc2MTk5NTM3NS0xNDkyMjg4ODcxLTIyMDI1NjkzMTktODk1ODkiLCJwbGF0ZiI6IjMiLCJwdWlkIjoiMTAwMzIwMDUyN0IzOENDNCIsInJoIjoiMS5BVFFBME5kazc5MzJqRUs4bUVudkdCSHB0UU1BQUFBQUFBQUF3QUFBQUFBQUFBQTBBSE0wQUEuIiwic2NwIjoiQ2FsZW5kYXJzLlJlYWRCYXNpYyBDaGFubmVsTWVtYmVyLlJlYWQuQWxsIENoYXQuUmVhZEJhc2ljIERpcmVjdG9yeS5SZWFkLkFsbCBlbWFpbCBGaWxlcy5SZWFkV3JpdGUuQWxsIEZpbGVTdG9yYWdlQ29udGFpbmVyLlNlbGVjdGVkIEdyb3VwLlJlYWRXcml0ZS5BbGwgR3JvdXBNZW1iZXIuUmVhZFdyaXRlLkFsbCBJbmZvcm1hdGlvblByb3RlY3Rpb25Qb2xpY3kuUmVhZCBvcGVuaWQgT3JnYW5pemF0aW9uLlJlYWQuQWxsIHByb2ZpbGUgU2Vuc2l0aXZpdHlMYWJlbC5SZWFkIFRhc2tzLlJlYWRXcml0ZSBVbmlmaWVkR3JvdXBNZW1iZXIuUmVhZC5Bc0d1ZXN0IFVzZXIuUmVhZC5BbGwgVXNlci5SZWFkQmFzaWMuQWxsIiwic2lkIjoiMDA5YWMxNDktMDYwMi00YzMyLWIwNzktNTYxNWQ3ZGEyN2E0Iiwic2lnbmluX3N0YXRlIjpbImR2Y19tbmdkIiwiZHZjX2NtcCIsImR2Y19kbWpkIiwiaW5rbm93bm50d2siLCJrbXNpIl0sInN1YiI6IjBubWpFcG5FeHBkYmRXX1dfSlZEOUQtajBnUzh3SldsT25MVGlFYmtCMFEiLCJ0ZW5hbnRfcmVnaW9uX3Njb3BlIjoiU0EiLCJ0aWQiOiJlZjY0ZDdkMC1mNmRkLTQyOGMtYmM5OC00OWVmMTgxMWU5YjUiLCJ1bmlxdWVfbmFtZSI6ImJwNTY5MDk0QGJwLm9yZy5iciIsInVwbiI6ImJwNTY5MDk0QGJwLm9yZy5iciIsInV0aSI6IjVZbkUzaXlLSUVPWUkxaHJGakZOQUEiLCJ2ZXIiOiIxLjAiLCJ3aWRzIjpbImI3OWZiZjRkLTNlZjktNDY4OS04MTQzLTc2YjE5NGU4NTUwOSJdLCJ4bXNfYWNkIjoxNjkwODY0MjkzLCJ4bXNfYWN0X2ZjdCI6IjMgNSIsInhtc19jYyI6WyJjcDEiXSwieG1zX2Z0ZCI6ImJfRDdxdzZmb0M0RGp6U000TVoydzZnQ0xNY3p6TVc3clpBdFFodTIwc2tCZFhObFlYTjBMV1J6YlhNIiwieG1zX2lkcmVsIjoiMzAgMSIsInhtc19zc20iOiIxIiwieG1zX3N0Ijp7InN1YiI6IlZ1Z05Tb0tadzAtVnR5S0EwVE0xVkdJOWdjRTZQQmwtLVBKNmNSQmVmSzAifSwieG1zX3N1Yl9mY3QiOiIzIDQiLCJ4bXNfdGNkdCI6MTQzMTExNjMwMywieG1zX3RudF9mY3QiOiIzIDEyIn0.DIQ_1_uG4GpniQvu9dh9LgqD82h9UWpSNscNssIlkmdCoX2uIbXnRpQ1P0PgIAyEUkADRhJgL-YPrL4dSgVRdZ2PoZ1J2WopdsZNir6iS8DveNC6I_LhDggkXAJA_63KGk9tv42_xscC0kRsGs4WxUJNpeNJOLWpPI5vZhkMYE-WOX97_ISkhtCfPXnIK9pPZaO0saFVXCyw9IGAaq22j5x9HIVycvhBwUMGQr-dKika8p58masXH-SmyCVezczL9JMvh45EWirZy2f24J2N8qYgxQxlNWY-Q9OcKmXOZNvC4YFkZGNwunk2reE6g4LRa7Yl66erqVuwKU1BdsO6pA'
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

exec_fisica_aprimora = df_merge[columns_exec_fisica].copy()

percentual_atividades_incompletas = np.abs(np.divide(exec_fisica_aprimora['activeChecklistItemCount'].astype('float'),exec_fisica_aprimora['checklistItemCount'].astype('float')))*100
percentual_atividades_completas = np.abs((np.divide(exec_fisica_aprimora['activeChecklistItemCount'].astype('float'),exec_fisica_aprimora['checklistItemCount'].astype('float'))-1))*100
exec_fisica_aprimora.loc[:, 'percentual_atividades_completas'] = np.round(percentual_atividades_completas,1)
exec_fisica_aprimora.loc[:, 'percentual_atividades_incompletas'] = np.round(percentual_atividades_incompletas,1)

task_id_list = exec_fisica_aprimora.loc[exec_fisica_aprimora['id_y'].notnull(), 'id_y']
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


exec_fisica_aprimora_final ['completedDateTime'] = pd.to_datetime(exec_fisica_aprimora_final['completedDateTime'])
exec_fisica_aprimora_final ['startDateTime'] = pd.to_datetime(exec_fisica_aprimora_final['startDateTime'])
exec_fisica_aprimora_final ['dueDateTime'] = pd.to_datetime(exec_fisica_aprimora_final['dueDateTime'])
exec_fisica_aprimora_final['Tempo de execução'] = exec_fisica_aprimora_final['completedDateTime'] - exec_fisica_aprimora_final['startDateTime'].dt.tz_localize(None)

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

exec_fisica_aprimora_final = exec_fisica_aprimora_final[exec_fisica_aprimora_final['planner_id'].notnull()]
#%%
exec_fisica_aprimora_final
