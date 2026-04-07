import os
import sys
import time 
import logging 
import warnings
import json 
from pathlib import Path

import typer 
from typer import colors
from pywinauto import Application, ElementNotFoundError, ElementAmbiguousError, WindowSpecification, mouse 
from pywinauto.keyboard import send_keys
from pydantic import ValidationError

from enums import DetailsTableColumns, OneCWebWMS, MyApp
from utils import error_exit, print_log, generate_docs, extract_details_row
from models import DetailItem


app = typer.Typer(
    name=MyApp.NAME,
    help=f"""Утилита для заполнения данных рекламации в клиентском приложении {OneCWebWMS.APP_NAME}\n
            Требования:
                • Windows OS
                • Запущенный экземпляр {OneCWebWMS.APP_NAME}
                • Открыта вкладка "{OneCWebWMS.CLAIM_CREATE_TAB_TITLE}"
            
            С исходным кодом утилиты можно ознакомиться по ссылке - https://github.com/bupbib/onec-claim-geely-util
    """,
    no_args_is_help=True,
    add_completion=False
)


@app.callback()
def main(ctx: typer.Context):
    f"""
    Callback-функция, выполняемая перед каждой командой.

    Вызывает ошибку, если:
        • Операционная система не Windows
        • Не запущено клиентское приложение {OneCWebWMS.APP_NAME} или запущено более одного окна
        • В приложении не открыта вкладка "{OneCWebWMS.CLAIM_CREATE_TAB_TITLE}"
    """ 
    if not sys.platform.startswith('win') or os.name != 'nt':
        error_exit(msg='Ошибка: Утилита работает только на Windows')
    
    try:
        geely_app = Application(backend='uia').connect(title=OneCWebWMS.MAIN_WINDOW_TITLE)
        # top_window() может вернуть не главное окно (всплывающие окна имеют более высокий Z-order),
        # поэтому используем поиск по title
        geely_window: WindowSpecification = geely_app.window(title=OneCWebWMS.MAIN_WINDOW_TITLE)

        claim_tab = geely_window.child_window(title_re=OneCWebWMS.CLAIM_CREATE_TAB_PATTERN, control_type='Pane')
        
        if not claim_tab.exists():
            error_exit(
                msg=f'Ошибка: Не открыта вкладка "{OneCWebWMS.CLAIM_CREATE_TAB_TITLE}". Откройте вкладку создания рекламации и повторите попытку'
            )

        ctx.obj = geely_window
        logger.info(f'Успешное подключение к {OneCWebWMS.APP_NAME}')
    except ElementNotFoundError as no_app_err:
        error_exit(
            msg=f'Ошибка: Клиентское приложение {OneCWebWMS.APP_NAME} не запущено. Запустите приложение и повторите попытку',
            original_exception=no_app_err
        )
    except ElementAmbiguousError as many_window_err:
        error_exit(
            msg=f'Ошибка: Запущено несколько окон {OneCWebWMS.APP_NAME}. Оставьте только одно активное окно и повторите попытку',
            original_exception=many_window_err
        )


@app.command(MyApp.ADD_JOBS)
@generate_docs(MyApp.ADD_JOBS)
def add_jobs(
    ctx: typer.Context,
    jobs: list[str] = typer.Argument(..., help='Коды работ через пробел')
):
    """
    Добавляет коды работ в рекламацию.
    """
    not_found_jobs = []
    total_jobs = len(jobs)
    geely_window: WindowSpecification = ctx.obj

    geely_window.set_focus()
    
    logger.info(f'Выполнение команды {MyApp.ADD_JOBS}, количество кодов: {total_jobs}')
    
    geely_window.child_window(title_re='Работы', control_type='TabItem').click_input()
    add_job_btn = geely_window.child_window(title='Добавить', control_type='Button').wrapper_object()
    
    # Флаг: нужно ли открывать вкладку заново
    need_open_tab = True 

    for idx, job in enumerate(jobs, 1):
        if need_open_tab:
            add_job_btn.click_input() 

        # set_focus() и клик по центру таблицы нестабильны. 
        # Клик по первой записи (или по месту для неё) гарантированно переводит фокус.
        # Повторяем до 5 раз на случай редких подвисаний 1С.
        # Если после 5 попыток окно поиска не появилось — добавляем job в not_found_jobs.
        search_dialog_found = False 
        for _ in range(5):
            nomenclature_table = geely_window['Отбор по модели и деталиTable'].wrapper_object()
            if nomenclature_table.children():
                nomenclature_table.children()[0].click_input()
            else:
                rect = nomenclature_table.rectangle()
                mouse.click(coords=(rect.left + rect.width() // 2, rect.top + 35))

            time.sleep(0.5)
            send_keys('%f')

            # Особенность 1С (UIA):
            # print_control_identifiers() нестабилен (падает на None),
            # а "Где искать" не определяется как ComboBox (идёт как Text),
            # поэтому поиск через descendants(Text) по label
            for item in geely_window.wrapper_object().descendants(control_type='Text'):
                item_text = item.window_text()

                if item_text in ('&Где искать:', '&Что искать:'):
                    search_dialog_found = True 

                    item.click_input()
                    item.type_keys(
                        ('Работа' if 'Где' in item_text else job) + '{TAB}'
                    )
            
            if search_dialog_found: break
        else: 
            # else выполнится, если не было break 
            not_found_jobs.append(job)
            logger.info(f'Не удалось открыть окно поиска для кода работы: {job}')
            send_keys('{ESC}')
        
        if not search_dialog_found: continue 

        # Если появляется надпись "Показать все" значит код работы не найден в номенклатуре
        show_all = geely_window.wrapper_object().descendants(control_type='Text', title='Показать все')
        if show_all:
            not_found_jobs.append(job)
            logger.debug(f'Код работы {job} не был найден в номенклатуре')

            # Если не нашли текущий код и есть еще коды - не закрываем вкладку
            if idx < total_jobs:
                need_open_tab = False
                send_keys('{ESC 3}', pause=0.2)  # ESC 3 = закрыть только окно поиска, остаться на вкладке
            else:
                send_keys('{ESC 4}', pause=0.2)  # ESC 4 = закрыть все (и поиск, и вкладку)
        else:
            need_open_tab = True 
            send_keys('^{ENTER}')
            logger.debug(f'Код работы {job} успешно найден в номенклатуре')

            for item in nomenclature_table.children():
                if f'{job} Работа' in item.window_text(): 
                    item.click_input(double=True)
    
    added_jobs_table = geely_window['Дата:Table'].wrapper_object()
    added_jobs_table.set_focus()

    # Удаляем пустые строки снизу вверх, чтобы не смещать индексы остальных элементов
    for job in reversed(added_jobs_table.children()):
        if job.window_text() == ' Наименование работы':
            job.click_input()
            job.type_keys('{DEL}')
    
    if not not_found_jobs:
        print_log(msg='Все коды работ успешно добавлены!')
    elif len(not_found_jobs) == total_jobs:
        print_log(msg='Ни один код не был найден. Проверьте корректность передаваемых кодов', color=colors.YELLOW)
    else:
        found_count = total_jobs - len(not_found_jobs)
        print_log(
            msg=f'Частичный успех. Добавлено {found_count} из {total_jobs}. Ненайденные коды: {";".join(not_found_jobs)}',
            color=colors.YELLOW
        )
             

@app.command(MyApp.ADD_DETAILS)
# @generate_docs(MyApp.ADD_DETAILS)
def add_details(
    ctx: typer.Context,
    file_path: Path = typer.Argument(
        ..., 
        help='Путь к JSON файлу с деталями', 
        exists=True,
        file_okay=True,
        dir_okay=False,
        readable=True
    )
):
    """
    Добавить детали в рекламацию
    """  
    with open(file_path, encoding='utf-8') as file:
        try:
            data = json.load(file)
        except json.JSONDecodeError as json_err:
            error_exit(
                msg='Ошибка: файл не является валидным JSON',
                original_exception=json_err
            )
    
    invalid_details = []
    not_found_details = []
    total_details = len(data)
    geely_window: WindowSpecification = ctx.obj 

    geely_window.set_focus()

    logger.info(f'Выполнение команды {MyApp.ADD_DETAILS}, количество кодов: {total_details}')

    geely_window.child_window(title_re='Детали', control_type='TabItem').click_input()
    add_detail_btn = geely_window.child_window(title='Добавить', control_type='Button').wrapper_object()
    added_details_table = geely_window['Дата:Table']

    # geely_window.print_control_identifiers()
    valid_count = 1
    for detail_dict in data:
        try:
            detail_item = DetailItem.model_validate(detail_dict)
        except ValidationError:
            logger.debug(f'Не удалось провалидировать деталь: {detail_dict}')
            invalid_details.append(detail_dict)
            continue 
    
        add_detail_btn.click_input()

        row = extract_details_row(
            table_row_elements=added_details_table.wrapper_object().children(), 
            row_number=valid_count
        )

        row[DetailsTableColumns.PART_NUMBER].type_keys('{F4}')
        # geely_window.print_control_identifiers()

        # Если появляется надпись "Показать все" значит номер детали не найден в номенклатуре
        # show_all = geely_window.wrapper_object().descendants(control_type='Text', title='Показать все')
        # if show_all:

        
        break 


        

 
if __name__ == '__main__':
    logging.basicConfig(
        level=logging.DEBUG,
        filename='app.log',
        encoding='utf-8',
        format='[{asctime}] #{levelname:8} {filename}:{lineno} - {name} - {message}',
        style='{'
    )

    logger = logging.getLogger(__name__)

    # Некоторые элементы 1С не поддерживают set_focus() через UIA, но это не критично
    warnings.filterwarnings('ignore', message='The window has not been focused due to COMError') 

    logger.info('Запуск приложения onec-claim-geely-util')
    logger.info(f'Аргументы командной строки: {" ".join(sys.argv)}')

    start = time.perf_counter()

    try:
        app()
    except SystemExit as err:
        code = err.code if err.code is not None else 0 
        end = time.perf_counter() - start 
        running_time_msg = f'Время работы: {end:.2f}'

        if len(sys.argv) == 1 and code == 2:
            logger.info('Вывод справки о приложении')
        elif code == 0:
            logger.info(f'Приложение завершилось успешно. {running_time_msg}')
        else:
            logger.error(f'Приложение завершилось с ошибкой, коды выхода: {err.code}. {running_time_msg}')

        raise err 