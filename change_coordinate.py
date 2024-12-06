import requests


headers = {'authorization': 'yv4c1D0QJqyE1pXfy7nZFWngGayQa4NwABAYzwvhHIHuJBD3Nru2Dj9FCWeXJeyk'}

id_doc = '672242bc809bcbe30a73803e' # ID документа
id_r = '288ecfe96b6231197a32f2a6' # ID раздела

latitude_old = '47.244902'
longitude_old = '40.7180928'

latitude_new = '53.268304'
longitude_new = '34.273162'


url = f'https://autocheck.fsa.gov.ru/api/RegistryProcesses/{id_doc}/stepsArray/{id_r}'

response = requests.get(url, headers=headers)

data = response.text
data = data.replace(latitude_old, latitude_new)
data = data.replace(longitude_old, longitude_new)


files = {
    'data': (None, data)
}

response = requests.put(url, headers=headers, files=files)
print(response.status_code)