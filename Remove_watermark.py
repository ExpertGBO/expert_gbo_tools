import os
from pathlib import Path
import re
import pikepdf
from io import BytesIO
import pdf_redactor

# Папка для чтения файлов с водяным знаком
input_folder = Path(os.path.expanduser("D:\ОТТС со знаком"))

# Папка для сохранения файлов без водяного знака
output_folder = Path(os.path.expanduser("D:\ОТТС без знака"))

# Убедитесь, что выходная папка существует
output_folder.mkdir(parents=True, exist_ok=True)

def remove_objects_with_length_66(input_file_path, output_file_path):
    # Преобразуем путь в объект Path, если это еще не объект Path
    output_file_path = Path(output_file_path)

    # Получаем родительскую директорию, имя файла и расширение
    directory = output_file_path.parent
    stem = output_file_path.stem  # Имя файла без расширения
    suffix = output_file_path.suffix  # Расширение файла (например, .pdf)

    # Добавляем текст "(не обработан)" перед расширением
    new_filename = f"{stem}{suffix}"

    # Создаем новый полный путь
    new_output_file_path = directory / new_filename


    target_data = b'SERTAUTO.RU'  # Байтовое представление текста, который нужно искать
    target_data_2 =b'q\n0.001 w\n1'

    with pikepdf.open(input_file_path) as pdf:
        for page_num, page in enumerate(pdf.pages):

            if '/Contents' in page:
                contents = page['/Contents']

                if isinstance(contents, pikepdf.Array):
                    to_remove = []
                    # Перебираем каждый объект на странице
                    for i, content in enumerate(contents):
                        if isinstance(content, pikepdf.Stream):
                            stream_data = content.read_bytes()
                            # Проверяем, содержит ли поток строку "SERTAUTO.RU"
                            if target_data in stream_data:
                                to_remove.append(i)  # Добавляем объект для удаления
                            if target_data_2 in stream_data:
                                to_remove.append(i)  # Добавляем объект для удаления

                    # Удаляем объекты в обратном порядке, чтобы не нарушить индексацию
                    for index in reversed(to_remove):
                        del contents[index]

                    # Обновляем содержимое страницы
                    page['/Contents'] = contents
                else:
                    print(f"Один объект на странице {page_num + 1}")

        # Сохраняем результат
        pdf.save(new_output_file_path)

def remove_watermark_from_pdf(input_file_path, output_file_path, watermark_text='SERTAUTO.RU'):
    watermark_found = False
    try:
        options = pdf_redactor.RedactorOptions()
        watermark_pattern = re.compile(re.escape(watermark_text), re.IGNORECASE)

        def replace_with_empty(match):
            nonlocal watermark_found
            watermark_found = True
            return ''

        options.content_filters = [(watermark_pattern, replace_with_empty)]

        with open(input_file_path, 'rb') as pdf_file:
            pdf_content = pdf_file.read()

        options.input_stream = BytesIO(pdf_content)
        output_stream = BytesIO()
        options.output_stream = output_stream

        pdf_redactor.redactor(options)

        output_stream.seek(0)

        with open(output_file_path, 'wb') as f:
            f.write(output_stream.read())
            print(f'Файл {output_file_path} успешно сохранен. Способ 1')
    except Exception as e:
        print(e)

    if not watermark_found:
        remove_objects_with_length_66(input_file_path, output_file_path)
        print(f'Файл {output_file_path} успешно сохранен. Способ 2')

# Получаем список имен файлов без расширений из папки "ОТТС без знака"
existing_files = {file.stem for file in output_folder.glob("*.pdf")}


## Обрабатываем файлы из папки "ОТТС со знаком", которых нет в "ОТТС без знака"
for input_file in input_folder.glob("*.pdf"):

    if input_file.stem not in existing_files:
        # Задаем имя выходного файла
        output_file = output_folder / input_file.name
        # Удаляем водяной знак и сохраняем файл в выходную папку
        remove_watermark_from_pdf(input_file, output_file)
    else:
        print(f'Файл {input_file.name} уже существует в папке "ОТТС без знака". Пропускаем.')