import logging 

import typer 
from typer import colors 
from pywinauto.controls.uiawrapper import UIAWrapper

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
    row_dict = {}

    for idx, item in enumerate(table_row_elements, 1):
        if item.window_text() == f'{row_number} N':
            break 
    
    row_cells = table_row_elements[idx: idx + 12]
    for cell, column in zip(row_cells, DetailsTableColumns):
        row_dict[column] = cell 
    
    return row_dict


