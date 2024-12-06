import pandas as pd
import psycopg2
# Подключение к базе данных PostgreSQL
def get_customer_name(document_number):
    try:
        # Параметры подключения (измените на свои)
        conn = psycopg2.connect(
            dbname="expert_pool",
            user="admin",
            password="Eon7a27i1ZczZ59onvFQ",
            host="192.168.2.111",
            port="5432"
        )
        cursor = conn.cursor()

        # Выполнение SQL запроса для получения имени заказчика по номеру документа
        cursor.execute("SELECT ppto FROM conclusions_protocols_view  WHERE doc_number = %s", (document_number,))
        result = cursor.fetchone()

        # Если найден результат, вернуть имя, иначе вернуть None
        if result:
            return result[0]
        else:
            return None
    except Exception as e:
        print(f"Ошибка при выполнении запроса: {e}")
        return None
    finally:
        cursor.close()
        conn.close()


# Путь к файлу
file_path = 'Выгрузка carcoin заключения 2024-10-01-2024-10-31.xlsx'

# Чтение Excel файла
df = pd.read_excel(file_path)

# Указание столбцов для удаления
columns_to_remove = [
    'Дата изменения', 'Вид переоборудования', 'Исполнитель организации', 'ИНН Заказчика',
    'Брокер', 'ИНН Брокера', 'Собственник ТС', 'Марка ТС', 'Модель ТС', 'VIN',
    'Категория ТС (ТР ТС 0118/2011)', 'Тип ТС', 'Дата публикации', 'Место осмотра',
    'Склонирован из процесса', 'Дата и время создания процесса', 'Номер заявки',
    'Дата создания заявки', 'Адрес Заказчика', 'Статус процесса'
]

# Удаление указанных столбцов
df.drop(columns=columns_to_remove, inplace=True)

# Фильтрация строк, где в столбце "Решение по документу" значение не равно "Разрешено"
df_filtered = df[df['Решение по документу'] == 'Разрешено'].copy()

# Оставляем только дату в столбце "Дата и время создания" с параметром dayfirst=True
df_filtered.loc[:, 'Дата и время создания'] = pd.to_datetime(df_filtered['Дата и время создания'],
                                                             dayfirst=True).dt.date

# Сортировка по столбцу "Дата и время создания" по возрастанию
df_filtered = df_filtered.sort_values(by='Дата и время создания', ascending=True)

# Обработка пустых значений в столбце "Заказчик"
for index, row in df_filtered.iterrows():
    if pd.isna(row['Заказчик']):
        # Используем значение из столбца "Номер документа" для выполнения запроса в базу данных
        document_number = row['Номер документа']
        customer_name = get_customer_name(document_number)

        # Если имя найдено, заменяем пустое значение
        if customer_name:
            df_filtered.loc[index, 'Заказчик'] = customer_name

# Сохранение результата в новый Excel файл
output_file_path = 'Реестр заключений Сентябрь.xlsx'
df_filtered.to_excel(output_file_path, index=False)

print(f"Файл сохранен по пути: {output_file_path}")
