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

from enums import DetailsTableColumns, OneCWebWMS, MyApp, ReportAddDetails
from utils import (
    delete_empty_rows, error_exit, print_log, 
    generate_docs, extract_details_row, perform_search_with_retry
)
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
    Добавить коды работ в рекламацию
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
            geely_window.child_window(title='Отбор только по модели', control_type='CheckBox').click_input()

        search_dialog_found = perform_search_with_retry(
            window=geely_window,
            search_type='job',
            where_text='Работа',
            what_text=job
        )
        
        if not search_dialog_found: 
            not_found_jobs.append(job)
            logger.info(f'Не удалось открыть окно поиска для кода работы: {job}')
            send_keys('{ESC}')
            continue 

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
            send_keys('^{ENTER}')

            # UPD 13.04: Здесь тоже могут быть случаи, когда мы запустим поиск необходимого кода, 
            # который, в целом, присутствует в номенклатуре, но отсутствует у конкретной машины.
            # И тогда тут тоже, как и в add-details 1С генерирует порядка 80 пустых элементов.
            nomenclature_table = geely_window['Отбор по модели и деталиTable'].wrapper_object()
            code_job = nomenclature_table.children()[1]  # под индексом 1 будет ячейка с кодом работ, либо пустой элемент

            if code_job.window_text() == f'{job} Работа':
                code_job.click_input(double=True)
                logger.debug(f'Код работы {job} успешно найден в номенклатуре')
                need_open_tab = True 
            else:
                not_found_jobs.append(job)
                logger.debug(f'Код работы {job} не был найден в номенклатуре')

                if idx < total_jobs:
                    need_open_tab = False 
                else:
                    send_keys('{ESC}')             
    
    delete_empty_rows(
        table=geely_window['Дата:Table'].wrapper_object(),
        empty_marker=' Наименование работы'
    )
    
    if not not_found_jobs:
        print_log(msg='Все коды работ успешно добавлены!')
    elif len(not_found_jobs) == total_jobs:
        print_log(msg='Ни один код не был найден. Проверьте корректность передаваемых кодов', color=colors.YELLOW)
    else:
        found_count = total_jobs - len(not_found_jobs)
        print_log(
            msg = f"Частичный успех. Добавлено {found_count} из {total_jobs}. Ненайденные коды: {', '.join(f'\"{job}\"' for job in not_found_jobs)}",
            color=colors.YELLOW
        )
             

@app.command(MyApp.ADD_DETAILS)
@generate_docs(MyApp.ADD_DETAILS)
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
    Добавить детали в рекламацию из JSON файла
    """  
    with open(file_path, encoding='utf-8') as file:
        try:
            data = json.load(file)
            all_details = data.get('payload', {}).get('Детали')

            if all_details is None: raise typer.BadParameter('В передаваемом json отсутствует обязательный ключ "Детали"')
        except json.JSONDecodeError as json_err:
            error_exit(
                msg='Ошибка: файл не является валидным JSON',
                original_exception=json_err
            )
    
    invalid_details = []
    not_found_details = []
    total_details = len(all_details)
    geely_window: WindowSpecification = ctx.obj 

    geely_window.set_focus()

    logger.info(f'Выполнение команды {MyApp.ADD_DETAILS}, количество кодов: {total_details}')

    geely_window.child_window(title_re='Детали', control_type='TabItem').click_input()
    add_detail_btn = geely_window.child_window(title='Добавить', control_type='Button').wrapper_object()
    added_details_table = geely_window['Дата:Table']

    valid_count = 0
    need_open_tab = True 

    for idx, detail_dict in enumerate(all_details, 1):
        try:
            detail_item = DetailItem.model_validate(detail_dict)
        except ValidationError:
            logger.debug(f'Не удалось провалидировать деталь: {detail_dict}')
            invalid_details.append(detail_dict)
            continue 
        
        if need_open_tab: add_detail_btn.click_input()

        row = extract_details_row(
            table_row_elements=added_details_table.wrapper_object().children(), 
            row_number=(valid_count + 1)
        )

        if need_open_tab:
            row[DetailsTableColumns.PART_NUMBER].click_input(double=True)
            row[DetailsTableColumns.PART_NUMBER].type_keys('{F4}')
        
        search_dialog_found = perform_search_with_retry(
            window=geely_window,
            search_type='detail',
            where_text='Артикул',
            what_text=detail_item.part_number
        )
        
        if not search_dialog_found:
            not_found_details.append(detail_dict)
            logger.info(f'Не удалось открыть окно поиска для детали: {detail_dict}')
            send_keys('{ESC}')
            continue  
    
        send_keys('^{ENTER}')

        # Используем children() вместо child_window():
        # 1С при отсутствии детали генерирует ~80 пустых Custom-элементов, которые уже попадают в children().
        # child_window() в таком случае значительно медленнее (каждый вызов делает полный обход UIA-дерева).
        # children() получает элементы один раз, без повторных поисков, поэтому это самый быстрый вариант здесь.
        nomenclature_table = geely_window['№ производителя:Table'].wrapper_object()
        first_cell = nomenclature_table.children()[0]

        if first_cell.window_text() == f'{detail_item.part_number} Артикул':
            first_cell.click_input(double=True)
            need_open_tab = True 
            valid_count += 1
            logger.debug(f'Деталь {detail_dict} была успешно найдена в номенклатуре')

            row[DetailsTableColumns.QUANTITY].click_input(double=True)
            row[DetailsTableColumns.QUANTITY].type_keys('^a{DEL}')
            row[DetailsTableColumns.QUANTITY].type_keys(detail_item.quantity)

            row[DetailsTableColumns.UPD_NUMBER].click_input(double=True)
            row[DetailsTableColumns.UPD_NUMBER].type_keys(detail_item.upd)
            row[DetailsTableColumns.UPD_NUMBER].type_keys('{ENTER}')

            if geely_window.child_window(title='Накладные по данной позиции не найдены', control_type='Text').exists():
                send_keys('{ENTER}')
        else:
            if idx < total_details:
                need_open_tab = False 
            else:
                send_keys('{ESC}')
            
            logger.debug(f'Не удалось найти деталь {detail_dict} в номенклатуре')
            not_found_details.append(detail_dict)

    # Если предпоследняя деталь не была найдена, need_open_tab = False, а последняя деталь - невалидна,
    # то окно с поиском может остаться
    if geely_window.child_window(title='Найти...', control_type='Button').exists():
        send_keys('{ESC}')

    if not delete_empty_rows(
        table=added_details_table.wrapper_object(),
        empty_marker=' Наименование детали'
    ):
        # Если не была удалена ни одна строка, то фокус остается на ячейке "Количество"
        # последней добавленной строки, кликая по центру мы этот фокус убираем
        geely_window.click_input() 

    if not not_found_details and not invalid_details:
        print_log(msg='Все детали успешно добавлены!') 
    elif len(not_found_details) + len(invalid_details) == total_details:
        print_log(
            msg='Ни одна деталь не была найдена. Проверьте корректность передаваемых деталей', 
            color=colors.YELLOW
        )
    else:
        found_count = total_details - len(not_found_details) - len(invalid_details)
        report = {'not_found': not_found_details, 'invalid': invalid_details}

        report_path = Path(ReportAddDetails.DIR) / ReportAddDetails.FILENAME

        # Вообще папка создается роботом, но на всякий случай сюда тоже добавил, если папка есть - ничего не будет
        report_path.parent.mkdir(parents=True, exist_ok=True)

        with open(report_path, 'w', encoding='utf-8') as file:
            json.dump(report, file, ensure_ascii=False, indent=4)

        print_log(
            msg=f'Частичный успех. Добавлено {found_count} из {total_details}. '
                f'Не найдены: {len(not_found_details)}. '
                f'Ошибки: {len(invalid_details)}. '
                f'Подробности сохранены в {report_path}',
            color=colors.YELLOW
        )

 
if __name__ == '__main__':
    # Робот (Robin RPA) корректно показывает кириллицу только в кодировке cp866.
    # При тестировании в обычной консоли Windows эта кодировка даёт кракозябры.
    # Поэтому при тестировании закомментируйте строку ниже, а перед запуском в роботе — раскомментируйте.
    sys.stdout.reconfigure(encoding='cp866', errors='replace')  # pyright: ignore[reportAttributeAccessIssue]

    log_file = Path(__file__).parent / 'app.log'

    logging.basicConfig(
        level=logging.DEBUG,
        filename=str(log_file),
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