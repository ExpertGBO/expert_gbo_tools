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


def get_first_id(num_pb):
    url = "https://autocheck.fsa.gov.ru/api/ViewSecurityProtocols?filter=%7B%22skip%22:0,%22limit%22:50,%22order%22:%22creationDate+DESC%22,%22where%22:%7B%22stepsByKeyTechnicalFieldsResultNumber%22:%7B%22like%22:%22%D0%9F%D0%91%D0%9E%D0%9249-" +  num_pb[7:] + "%22,%22options%22:%22i%22%7D,%22resultStatesStatusKey%22:%7B%22inq%22:[%22new%22,%22submitted%22,%22error%22,%22expertSigning%22,%22directorALSigning%22,%22done%22]%7D%7D%7D"
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


def download_pdf(num_pb, id_pb):
    print(id_pb)
    url = "https://autocheck.fsa.gov.ru/api/RegistryProcesses/downloadGeneratedDoc?processId=" + id_pb + "&fileKey=securityProtocolEP.pdf"
    response = requests.get(url, headers=headers)
    num_pb = num_pb.replace("/", "_")
    # Указать полный путь для сохранения файла
    output_path = "D:\\Протоколы архив Октябрь\\Протокол безопасности № " + num_pb + ".pdf"
    if response.status_code == 200:

        with open(output_path, 'wb') as file:
            file.write(base64.b64decode(response.content))
        print("PDF успешно сохранен.")
    else:
        print(f"Ошибка при выполнении запроса: {response.status_code}")


def download_archive(num_pb, id_pb):
    url = "https://autocheck.fsa.gov.ru/api/RegistryProcesses/" + id_pb + "/getArchiveDownloadSrc?archiveKey=fullArchive"
    num_pb = num_pb.replace("/", "_")
    archive_path = "D:\\Протоколы архив Октябрь\\Архив\\Протокол безопасности № " + num_pb + ".zip"
    response = requests.get(url, headers=headers)

    if response.status_code == 200:
        while response.text[:10] == '{"message"':
            response = requests.get(url, headers=headers)
            time.sleep(1)


        id_archive = extract_file_id (response.text)

        print(id_archive)

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
    # Создаем пустой список для хранения значений из столбца B
    column_a_values = []
    # Перебираем все значения в столбце A до тех пор, пока не достигнем пустой ячейки
    for cell in sheet['B']:
        if cell.value:
            column_a_values.append(cell.value)
    # Перебираем значения из столбца A в цикле
    for value in column_a_values:
        num_pb = value
        print(num_pb)
        id_pb = get_first_id(num_pb)
        download_pdf(num_pb, id_pb)
        download_archive(num_pb, id_pb)


main_download()


