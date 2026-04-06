import os 
import sys
import time 
import logging 

import typer 
from typer import colors
from pywinauto import Application, ElementNotFoundError, ElementAmbiguousError, WindowSpecification

from enums import OneCWebWMS
from utils import error_exit


app = typer.Typer(
    name='onec-claim-geely-util',
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

        geely_window.set_focus()
        ctx.obj = geely_window
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


@app.command('works')
def works(
        ctx: typer.Context,
):
    """Тестовая команда"""
    window: WindowSpecification = ctx.obj
    typer.secho("✅ Утилита работает, подключение к Geely установлено", fg=colors.GREEN)


logging.basicConfig(
    level=logging.DEBUG,
    filename='app.log',
    encoding='utf-8',
    format='[{asctime}] #{levelname:8} {filename}:{lineno} - {name} - {message}',
    style='{'
)

logger = logging.getLogger(__name__)

if __name__ == '__main__':
    logging.info('Запуск приложения onec-claim-geely-util')
    logging.info(f'Аргументы командной строки: {" ".join(sys.argv)}')

    start = time.perf_counter()

    try:
        app()
    except SystemExit as err:
        code = err.code if err.code is not None else 0 
        end = time.perf_counter() - start 
        running_time_msg = f'Время работы: {end:.2f}'

        if len(sys.argv) == 1 and code == 2:
            logging.info('Вывод справки о приложении')
        elif code == 0:
            logging.info(f'Приложение завершилось успешно. {running_time_msg}')
        else:
            logging.error(f'Приложение завершилось с ошибкой, коды выхода: {err.code}. {running_time_msg}')

        raise err 