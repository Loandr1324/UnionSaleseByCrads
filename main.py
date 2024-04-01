# Author Loik Andrey 7034@balancedv.ru
import config
import send_mail  # Универсальный модуль для отправки сообщений на почту
from loguru import logger
import smbclient
from datetime import date, timedelta  # Загружаем библиотеку для работы с текущим временем
import pandas as pd
import os  # Загружаем библиотеку для работы с файлами
from smbclient import shutil as smb_shutil  # Универсальный модуль для копирования файлов

logger.add(config.FILE_NAME_CONFIG,
           format="{time:DD/MM/YY HH:mm:ss} - {file} - {level} - {message}",
           level="INFO",
           rotation="1 month",
           compression="zip")

# Создаём подключение для работы с файлами на сервере
smbclient.ClientConfig(username=config.LOCAL_PATH['USER'], password=config.LOCAL_PATH['PSW'])
path = config.LOCAL_PATH['PATH']


def read_data() -> pd.DataFrame:
    """
    Получаем данные из фалов с отчётами
    """
    df_sales_card = pd.DataFrame()
    names = ['Номер карты', 'Владелец карты', 'Выручка']

    for item in smbclient.listdir(path):
        if item.endswith('.xlsx'):
            # Считываем файл с сервера
            df = pd.read_excel(
                smbclient.open_file(path + "/" + item, 'rb'), header=6, engine='openpyxl'
            )
            df.dropna(axis=1, how='all', inplace=True)  # Удаляем пустые колонки
            del_col = [df.columns[2], df.columns[3]]  # Выбираем колонки для удаления
            df.drop(del_col, axis=1, inplace=True)  # Удаляем колонки с количеством продаж и карт
            df.columns = names  # Переименовываем колонки
            df_sales_card = pd.concat([df_sales_card, df], ignore_index=True)  # Добавляем данные в общий DataFrame

    # Преобразуем колонку "Номер карты" в строку и дополняем недостающими нулями до 5 символов
    if len(df_sales_card) > 0:
        df_sales_card['Номер карты'] = df_sales_card['Номер карты'].apply(lambda col: str(col).zfill(5))
    return df_sales_card


def df_group(df: pd.DataFrame) -> pd.DataFrame:
    """
    Группируем данные в pd.DataFrame
    :param df: pd.DataFrame с исходными данными
    :return: DataFrame с группированными данными
    """
    df = df.groupby(['Номер карты', 'Владелец карты']).sum()
    df.reset_index(inplace=True)
    return df


def df_add_total(df: pd.DataFrame) -> list:
    """
    Суммируем значение выручки и кол-во карт и возвращаем строку для записи в отчёт
    """
    sales_amount = df['Выручка'].sum()  # Получаем сумму выручки
    count = len(df)  # Считаем количество карт
    return ['Компания MaCar:', None, sales_amount, count]


def format_custom():
    """Готовые форматы для эксель"""
    year_format = {
        'font_name': 'Arial',
        'font_size': '10',
        'align': 'left',
        'bold': True,
        'bg_color': '#F4ECC5',
        'border': True,
        'border_color': '#CCC085'
    }
    columns_name_format = {
        'font_name': 'Arial',
        'font_size': '8',
        'align': 'center',
        'border': True,
        'border_color': '#CCC085',
        'bg_color': '#F8F2D8'
    }
    month_format = {
        'font_name': 'Arial',
        'font_size': '8',
        'align': 'right',
        'bold': False,
        'border': True,
        'border_color': '#CCC085'
    }
    sum_format = {
        'num_format': '# ### ##0.00"р.";[red]-# ##0.00"р."',
        'font_name': 'Arial',
        'font_size': '8',
        'border': True,
        'border_color': '#CCC085'
    }
    quantity_format = {
        'num_format': '# ### ##0',
        'font_name': 'Arial',
        'font_size': '8',
        'border': True,
        'border_color': '#CCC085'
    }
    caption_format = {
        'font_name': 'Arial',
        'font_size': '14',
        'bold': True,
        'border': False
    }

    return year_format, caption_format, columns_name_format, month_format, sum_format, quantity_format


def write_to_excel(df: pd.DataFrame, exel_file: str) -> None:
    """Записываем данные в эксель в нужном формате"""
    sheet_name = 'Данные'  # Наименование вкладки для сводной таблицы
    caption = 'Продажи по картам лояльности СТО' # Заголовок таблицы

    with pd.ExcelWriter(exel_file, engine='xlsxwriter') as writer:  # Открываем файл для записи
        workbook = writer.book

        # Получаем форматы для эксель
        _, caption_format, columns_name_format, _, sum_format, _ = format_custom()
        # Записываем данные в эксель
        df.to_excel(writer, sheet_name=sheet_name, startrow=3, header=False, index=False)

        wks1 = writer.sheets[sheet_name]  # Открываем вкладку для форматирования

        # Изменяем формат колонок
        table_format = columns_name_format.copy()
        table_format['bg_color'] = '#FFFFFF'
        wks1.set_column('A:A', 11, workbook.add_format(table_format))
        table_format['align'] = 'left'
        wks1.set_column('B:B', 65, workbook.add_format(table_format))
        wks1.set_column('C:C', 15, workbook.add_format(sum_format))
        wks1.set_column('D:D', 11, None)

        # Записываем заголовок
        wks1.write('A1', caption, workbook.add_format(caption_format))
        # Записываем наименование колонок
        columns_name = df.columns.values.tolist()
        columns_name.append('Кол-во карт')
        for i, v in enumerate(columns_name):
            wks1.write(1, i, v, workbook.add_format(columns_name_format))
        # Записываем итоговые данные
        for i, v in enumerate(df_add_total(df)):
            w_format = sum_format
            w_format['bold'] = True
            w_format['bg_color']= '#F8F2D8'
            if i == 3:
                w_format['num_format'] = '# ### ##0'
            wks1.write(2, i, v, workbook.add_format(w_format))

        # Изменяем шрифт в нумерациях ячеек эксель, чтобы уменьшить размер ячеек
        workbook.formats[0].set_font_size(9)
    return


def date_xlsx():
    """
    Определяем необходимые даты, в том числе русское строковое наименование месяца для добавления в эксель.
    :return: month_name_str - наименование на русском прошлого месяца январь, ..., декабрь.
    :return: year - числовое обозначение года прошлого месяца, четыре цифры. Например, 2020.
    :return: month_name_int - числовое обозначение прошлого месяца двухзначное число 01...12.
    """
    ru_month_values = {
        '01': 'Январь',
        '02': 'Февраль',
        '03': 'Март',
        '04': 'Апрель',
        '05': 'Май',
        '06': 'Июнь',
        '07': 'Июль',
        '08': 'Август',
        '09': 'Сентябрь',
        '10': 'Октябрь',
        '11': 'Ноябрь',
        '12': 'Декабрь'
    }
    last_month = date.today() - timedelta(days=25)
    month_name_str = ru_month_values[last_month.strftime('%m')]
    month_name_int = last_month.strftime('%m')
    year = last_month.strftime('%Y')
    return month_name_str, year, month_name_int


def send_file_to_mail(files: list) -> None:
    """Отправляем созданный эксель файл на почту"""
    # Получаем месяц и год из функции по определению дат
    year, month = date_xlsx()[1:]
    # Текст сообщения в формате html
    email_content = f"""
        <html>
          <head></head>
          <body>
            <p>
                Объединенный отчет по картам лояльности СТО за {month}.{year}г. во вложении<br>
                Отчет сформирован автоматически без участия сотрудников.<br>
                Если обнаружены ошибки, то прошу сообщить администраторам 1С.<br>
            </p>
          </body>
        </html>
    """
    message = {
        'Subject': f'Продажи по картам лояльности СТО за {month}.{year}',  # Тема сообщения,
        'email_content': email_content,
        'To': config.TO_EMAILS['TO_CORRECT'],
        'File_name': files,
        'Temp_file': files

    }

    send_mail.send(message)
    return


def send_mail_error() -> None:
    """
    Отправляем сообщения об ошибке при выполнении программы
    :return: None
    """
    logger.info(f"Нет отчета продаж по картам СТО за предыдущий месяц.")
    message = {
        'Subject': f"Ошибка при формировании ежемесячного отчета по Картам СТО",
        'email_content': (f"Нет отчета продаж по картам СТО за предыдущий месяц.<br>"
                                    f"Разместите отчет в папке:<br>"
                                    f"{config.LOCAL_PATH['PATH']}"),
        'To': config.TO_EMAILS['TO_ERROR'],
        'File_name': '',
        'Temp_file': ''
    }

    # Оправка письма со сформированными параметрами
    send_mail.send(message)
    return


def remove_files():
    """
    Копируем отчеты с исходной папки и файлы с данными в папку за месяц и удаляем все файлы из исходной папки с отчётами
    :return: None
    """
    path1 = config.LOCAL_PATH['PATH']
    year, month = date_xlsx()[1:]
    path2 = path1 + f"/Отчёты за {month}.{year}"

    # Создаём резервную папку за месяц отчета
    smbclient.mkdir(path2)

    # Переносим файлы из Исходной директории в резервную
    for item in smbclient.listdir(path1):
        if item.endswith('.xlsx'):
            smbclient.copyfile(path1 + "/" + item, path2 + "/" + item)
            smbclient.remove(path1 + "/" + item)
    # Переносим файлы с данными из директории скрипта в резервную папку
    for item in os.listdir():
        if item.endswith('.xlsx'):
            smb_shutil.copyfile(item, path2 + "/" + item)
    return


def run():
    """
    Весь рабочий процесс программы.
    """
    logger.info(f"... Начало работы программы")
    logger.info("Считываем данные из отчетов")
    df_sales = read_data()

    # Если данных нет, то отправляем письмо об ошибке и прекращаем работу программы
    if len(df_sales) == 0:
        send_mail_error()
        return

    logger.info("Группируем полученные данные")
    df_sales = df_group(df_sales)

    logger.info("Записываем данные в эксель")
    write_to_excel(df_sales, config.FILE_NAME_OUTPUT)

    logger.info("Отправляем итоговый файл на почту")
    send_file_to_mail([config.FILE_NAME_OUTPUT])

    logger.info("Резервируем данные")
    remove_files()

    logger.info(f"... Завершение работы программы")
    return


if __name__ == '__main__':
    run()
