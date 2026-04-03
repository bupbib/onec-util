import os 
import sys

import typer 
from typer import colors
from pywinauto import Application, ElementNotFoundError, ElementAmbiguousError, WindowSpecification

from enums import OneCWebWMS


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
        typer.secho(
            'Ошибка: Утилита работает только на Windows',
            fg=colors.RED
        )
        raise typer.Exit(1)
    
    try:
        geely_app = Application(backend='uia').connect(title=OneCWebWMS.MAIN_WINDOW_TITLE)
        # top_window() может вернуть не главное окно (всплывающие окна имеют более высокий Z-order),
        # поэтому используем поиск по title
        geely_window: WindowSpecification = geely_app.window(title=OneCWebWMS.MAIN_WINDOW_TITLE)

        claim_tab = geely_window.child_window(title_re=OneCWebWMS.CLAIM_CREATE_TAB_PATTERN, control_type='Pane')
        
        if not claim_tab.exists():
            typer.secho(
                f'Ошибка: Не открыта вкладка "{OneCWebWMS.CLAIM_CREATE_TAB_TITLE}". Откройте вкладку создания рекламации и повторите попытку',
                fg=colors.RED
            )
            raise typer.Exit(1)

        geely_window.set_focus()
        ctx.obj = geely_window
    except ElementNotFoundError as no_app_err:
        typer.secho(
            f'Ошибка: Клиентское приложение {OneCWebWMS.APP_NAME} не запущено. Запустите приложение и повторите попытку',
            fg=colors.RED
        )
        raise typer.Exit(1) from no_app_err
    except ElementAmbiguousError as many_window_err:
        typer.secho(
            f'Ошибка: Запущено несколько окон {OneCWebWMS.APP_NAME}. Оставьте только одно активное окно и повторите попытку',
            fg=colors.RED
        )
        raise typer.Exit(1) from many_window_err


@app.command('works')
def works(
        ctx: typer.Context,
):
    """Тестовая команда"""
    window: WindowSpecification = ctx.obj
    typer.secho("✅ Утилита работает, подключение к Geely установлено", fg=colors.GREEN)


if __name__ == '__main__':
    app()