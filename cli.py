import os 
import sys 
import typer 
from typer import colors
from pywinauto import Application, ElementNotFoundError, ElementAmbiguousError


app = typer.Typer(
    name='onec-util',
    help="""Утилита для работы с табличными частями 1С.""",
    no_args_is_help=True,
    add_completion=False
)


@app.callback()
def main(ctx: typer.Context):
    """
    Callback-функция, выполняемая перед каждой командой.

    Вызывает ошибку, если:
        • Операционная система не Windows
        • Не запущено клиентское приложение 1С WEB VMS (Geely)
        • Запущено несколько активных окон приложения 1С WEB VMS (Geely)
    """ 
    if not sys.platform.startswith('win') or os.name != 'nt':
        typer.secho(
            'Ошибка: Утилита работает только на Windows',
            fg=colors.RED
        )
        raise typer.Exit(1)
    
    try:
        geely_app = Application(backend='uia').connect(title='WEB VMS - ООО «Бизнес Кар М» - Москва')
        geely_window = geely_app.top_window()
        geely_window.set_focus()
        ctx.obj = geely_window
    except ElementNotFoundError as no_app_err:
        typer.secho(
            'Ошибка: клиентское приложение 1С WEB VMS (Geely) не запущено. Запустите приложение и повторите попытку',
            fg=colors.RED
        )
        raise typer.Exit(1) from no_app_err
    except ElementAmbiguousError as many_window_err:
        typer.secho(
            'Ошибка: запущено несколько окон 1С WEB VMS (Geely). Оставьте только одно активное окно и повторите попытку',
            fg=colors.RED
        )
        raise typer.Exit(1) from many_window_err


@app.command('works')
def works(
        ctx: typer.Context,
):
    """Тестовая команда"""
    return 'hello world' 


if __name__ == '__main__':
    app()