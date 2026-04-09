import logging 
import time 
from typing import Literal 

import typer 
from typer import colors 
from pywinauto import WindowSpecification, mouse
from pywinauto.controls.uiawrapper import UIAWrapper
from pywinauto.keyboard import send_keys

from enums import MyApp, DetailsTableColumns
from user_docs import LEXICON

logger = logging.getLogger(__name__)


def error_exit(
    msg: str, 
    exit_code: int = 1,
    original_exception: Exception | None = None
) -> None:
    """
    Выводит сообщение об ошибке красным цветом, логирует и завершает программу с указанным кодом.
    
    Args:
        msg: Сообщение об ошибке для вывода пользователю
        exit_code: Код выхода (по умолчанию 1)
        original_exception: Исходное исключение для сохранения traceback (опционально)
    """ 
    typer.secho(msg, fg=colors.RED)
    logger.error(msg)
    
    if original_exception is None:
        raise typer.Exit(exit_code)
    else:
        raise typer.Exit(exit_code) from original_exception


def print_log(msg: str, color: str = colors.GREEN) -> None:
    """
    Выводит сообщение в консоль и пишет в лог
    """
    typer.secho(msg, fg=color)
    logger.info(msg)


def generate_docs(cmd: MyApp):
    """Декоратор для установки документации из LEXICON"""
    def decorator(func):
        func.__doc__ = LEXICON[cmd]
        return func

    return decorator


def extract_details_row(table_row_elements: list[UIAWrapper], row_number: int) -> dict:
    """
    Извлекает данные из строки таблицы Детали и возвращает словарь с соответствием столбец: значение.
    
    Таблица в 1С представлена сплошным списком элементов, где каждая строка состоит из 12 ячеек.
    Функция находит начало строки по номеру строки (ячейка с форматом "{row_number} N"),
    вырезает следующие 12 элементов и сопоставляет их с названиями столбцов из DetailsTableColumns.
    
    Args:
        table_row_elements: Список всех Custom-элементов таблицы (полученный через .children())
        row_number: Номер строки для извлечения (отображается в первом столбце)
    
    Returns:
        dict: Словарь, где ключ — название столбца (из DetailsTableColumns), 
            значение — UIAWrapper элемент ячейки (для дальнейшего клика или получения текста)
    """
    idx = 0
    row_dict = {}

    for idx, item in enumerate(table_row_elements):
        if item.window_text() == f'{row_number} N':
            break 
    
    row_cells = table_row_elements[idx: idx + 12]
    for cell, column in zip(row_cells, DetailsTableColumns):
        row_dict[column] = cell 
    
    return row_dict


def perform_search_with_retry(
    window: WindowSpecification,
    search_type: Literal['job', 'detail'],
    where_text: str,
    what_text: str, 
    max_attempts: int = 5
) -> bool:
    """
    Открывает окно поиска, заполняет поля и повторяет попытку при неудаче.
    
    Режимы работы:
        - 'job':   открывает поиск через Ctrl+F, предварительно кликает по первой записи таблицы
        - 'detail': открывает поиск через кнопку "Найти..."
    
    Args:
        window: Окно 1С (geely_window)
        where_text: Текст для поля "Где искать:" (например, "Работа" или "Артикул")
        what_text: Текст для поля "Что искать:" (код работы или артикул детали)
        search_type: Тип поиска — 'job' для кодов работ, 'detail' для деталей
        max_attempts: Количество попыток (по умолчанию 5)
    
    Returns:
        bool: True если удалось открыть окно поиска и заполнить поля, False если нет
    """
    for _ in range(max_attempts):
        if search_type == 'job':
            # set_focus() + клик по первой записи — стабильная комбинация для перевода фокуса.
            # По отдельности оба метода ненадёжны.
            nomenclature_table = window['Отбор по модели и деталиTable'].wrapper_object()
            nomenclature_table.set_focus()

            if nomenclature_table.children():
                nomenclature_table.children()[0].click_input()
            else:
                rect = nomenclature_table.rectangle()
                mouse.click(coords=(rect.left + rect.width() // 2, rect.top + 35))

            time.sleep(0.5)
            send_keys('^f') 
        elif search_type == 'detail':
            window.child_window(title='Найти...', control_type='Button').click_input() 
        else:
            return False  
        
        time.sleep(0.5)  # время на прогрузку окна
        
        if fill_search_fields(window, where_text, what_text):
            for item in window.wrapper_object().descendants(control_type='Text'):
                if item.window_text() == 'По точному совпадению':
                    item.click_input()
                    break 
            return True 
    
    return False 


def fill_search_fields(window: WindowSpecification, where_text: str, what_text: str) -> bool:
    """
    Заполняет поля поиска "Где искать:" и "Что искать:" в окне поиска 1С.
    
    Открытое окно поиска должно быть активным. Функция находит текстовые поля
    по их заголовкам и вводит указанные значения, после чего переключается на следующее поле (TAB).
    
    Args:
        window: Окно 1С (обычно geely_window), в котором ищем descendants
        where_text: Текст для поля "Где искать:" (например, "Работа")
        what_text: Текст для поля "Что искать:" (например, код работы или детали)
    
    Returns:
        bool: True если оба поля найдены и значения введены, False если хотя бы одно не найдено
    """
    where_found = False 
    what_found = False 

    # Особенность 1С (UIA):
    # print_control_identifiers() нестабилен (падает на None),
    # а "Где искать" не определяется как ComboBox (идёт как Text),
    # поэтому поиск через descendants(Text) по label

    for item in window.wrapper_object().descendants(control_type='Text'):
        item_text = item.window_text()

        if item_text == '&Где искать:':
            where_found = True 
            item.click_input()
            item.type_keys(where_text + '{TAB}')
        elif item_text == '&Что искать:':
            what_found = True 
            item.click_input()
            item.type_keys(what_text + '{TAB}')
    
    return where_found and what_found


def delete_empty_rows(table: UIAWrapper, empty_marker: str) -> bool:
    """
    Удаляет строки, у которых ячейка содержит указанный маркер пустой строки.
    
    Удаление происходит снизу вверх, чтобы не сбивать индексы.
    
    Args:
        table: Таблица (UIAWrapper)
        empty_marker: Текст-маркер пустой строки (например, ' Наименование работы' или ' Наименование детали')
    
    Returns:
        bool: True, если была удалена хотя бы одна строка, иначе False.
    """
    delete_flag = False 
    table.set_focus()

    for item in reversed(table.children()):
        if item.window_text() == empty_marker:
            delete_flag = True 
            item.click_input()
            item.type_keys('{DEL}')
    
    return delete_flag