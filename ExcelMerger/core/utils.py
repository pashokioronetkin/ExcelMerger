from datetime import datetime

from . import constants

def letter_to_number(column_letter):
    """Конвертация буквенного обозначения колонки в числовое."""
    if not column_letter:
        return None

    num = 0
    for i, char in enumerate(reversed(column_letter.upper())):
        num += (ord(char) - ord('A') + 1) * (26 ** i)
    return num

def number_to_letter(n):
    """Конвертация числового обозначения колонки в буквенное."""
    string = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        string = chr(65 + remainder) + string
    return string

def form_date(dates):
    """Форматирование дат из формата YYYY-MM-DD в DD.MM.YYYY."""
    date_ = []
    for i in dates:
        try:
            date_obj = datetime.strptime(i, "%Y-%m-%d")
            formatted_date = date_obj.strftime("%d.%m.%Y")
            date_.append(formatted_date)
        except ValueError as e:
            print(f"Ошибка форматирования даты {i}: {e}")
    return date_

def form_date_add_id(date_str):
    """Форматирование одиночной даты из строки в объект datetime."""
    try:
        date_obj = datetime.strptime(str(date_str), "%Y-%m-%d")
        return date_obj.strftime("%d.%m.%Y")
    except ValueError as e:
        print(f"Ошибка форматирования даты {date_str}: {e}")
        return None