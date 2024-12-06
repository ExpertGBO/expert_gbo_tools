import base64
import json
import time
import requests
import openpyxl


token = "LIM_aTv_biysULDX9nNlU8PyMq7BVxuJTPnmqgVRfwLp7_Hak-0KgJWJgdGkYAdU"
headers = {"Authorization": token,}

def extract_file_id(json_string):
    try:
        # Парсим JSON-строку в Python-словарь
        data = json.loads(json_string)

        # Извлекаем значение по ключу 'src'
        src_value = data.get('src', '')

        # Извлекаем идентификатор из строки после последнего '/'
        file_id = src_value.split('/')[-1]

        return file_id
    except json.JSONDecodeError:
        print("Ошибка декодирования JSON")
        return None


def get_first_id(num_pte):
    url = "https://autocheck.fsa.gov.ru/api/ViewPreliminaryConclusions?filter=%7B%22skip%22:0,%22limit%22:50,%22order%22:%22creationDate+DESC%22,%22where%22:%7B%22stepsByKeyTechnicalFieldsResultFormNumber%22:%7B%22like%22:%22%D0%9F%D0%A2%D0%AD%D0%9E%D0%9249-" +  num_pte[8:] + "%22,%22options%22:%22i%22%7D,%22resultStatesStatusKey%22:%7B%22inq%22:[%22new%22,%22submitted%22,%22error%22,%22expertSigning%22,%22directorALSigning%22,%22done%22]%7D%7D%7D"
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        json_data = response.json()
        if json_data:
            first_id = json_data[0]['id']
            return first_id
        else:
            print("Список JSON данных пуст.")
    else:
        print("Ошибка при выполнении запроса:", response.status_code)
    return None


def download_pdf(num_pte, id_pte):
    url = "https://autocheck.fsa.gov.ru/api/RegistryProcesses/downloadGeneratedDoc?processId=" + id_pte + "&fileKey=pteConclusionsEP.pdf"
    response = requests.get(url, headers=headers)
    num_pte = num_pte.replace("/", "_")
    # Указать полный путь для сохранения файла
    output_path = "D:\\Заключения архив Октябрь\\Заключение предварительной технической экспертизы № " + num_pte + ".pdf"
    if response.status_code == 200:

        with open(output_path, 'wb') as file:
            file.write(base64.b64decode(response.content))
        print("PDF успешно сохранен.")
    else:
        print(f"Ошибка при выполнении запроса: {response.status_code}")


def download_archive(num_pte, id_pte):
    url = "https://autocheck.fsa.gov.ru/api/RegistryProcesses/" + id_pte + "/getArchiveDownloadSrc?archiveKey=fullArchive"
    num_pte = num_pte.replace("/", "_")
    archive_path = "D:\\Заключения архив Октябрь\\Архив\\Заключение ПТЭ № " + num_pte + ".zip"
    response = requests.get(url, headers=headers)

    if response.status_code == 200:
        while response.text[:10] == '{"message"':
            response = requests.get(url, headers=headers)
            time.sleep(1)


        id_archive =  extract_file_id (response.text)

    else:
        print(f"Ошибка при выполнении запроса: {response.status_code}")

    url = "https://autocheck.fsa.gov.ru/api/v1/file-storage/downloadFile/" + id_archive

    response = requests.get(url, headers=headers)

    if response.status_code == 200:
        # Сохранение архива
        with open(archive_path, "wb") as f:
            f.write(response.content)
        print("Архив успешно сохранен.")
    else:
        print(f"Ошибка при выполнении запроса: {response.status_code}")


def main_download():
    # Открываем файл Excel
    workbook = openpyxl.load_workbook("C:/Users/Pasha Sagura/Desktop/ПТЭ.xlsx")
    # Выбираем лист "Данные"
    sheet = workbook['Данные']
    # Создаем пустой список для хранения значений из столбца A
    column_a_values = []
    # Перебираем все значения в столбце A до тех пор, пока не достигнем пустой ячейки
    for cell in sheet['A']:
        if cell.value:
            column_a_values.append(cell.value)
    # Перебираем значения из столбца A в цикле
    for value in column_a_values:
        num_pte = value
        print(num_pte)
        id_pte = get_first_id(num_pte)
        download_pdf(num_pte, id_pte)
        download_archive(num_pte, id_pte)


main_download()













