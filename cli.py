import os 
import sys
import typer 
from typer import colors
from pywinauto import Application, ElementNotFoundError, ElementAmbiguousError, WindowSpecification


app = typer.Typer(
    name='onec-claim-geely-util',
    help="""Утилита для заполнения данных рекламации в клиентском приложении 1С WEB VMS (Geely)\n
            Требования:
                • Windows OS
                • Запущенный экземпляр 1С WEB VMS (Geely)
                • Открыта вкладка "Рекламация VMS (создание) *"
            
            С исходным кодом утилиты можно ознакомиться по ссылке - https://github.com/bupbib/onec-util
    """,
    no_args_is_help=True,
    add_completion=False
)


@app.callback()
def main(ctx: typer.Context):
    """
    Callback-функция, выполняемая перед каждой командой.

    Вызывает ошибку, если:
        • Операционная система не Windows
        • Не запущено клиентское приложение 1С WEB VMS (Geely) или запущено более одного окна
        • В приложении не открыта вкладка "Рекламация VMS (создание)"
    """ 
    if not sys.platform.startswith('win') or os.name != 'nt':
        typer.secho(
            'Ошибка: Утилита работает только на Windows',
            fg=colors.RED
        )
        raise typer.Exit(1)
    
    try:
        geely_app = Application(backend='uia').connect(title='WEB VMS - ООО «Бизнес Кар М» - Москва')
        geely_window: WindowSpecification = geely_app.top_window()

        claim_tab = geely_window.child_window(title_re=r'Рекламация VMS \(создание\)\s?\*?', control_type='Pane')
        
        if not claim_tab.exists():
            typer.secho(
                'Ошибка: Не открыта вкладка "Рекламация VMS (создание)". Откройте вкладку создания рекламации и повторите попытку',
                fg=colors.RED
            )
            raise typer.Exit(1)

        geely_window.set_focus()
        ctx.obj = geely_window
    except ElementNotFoundError as no_app_err:
        typer.secho(
            'Ошибка: Клиентское приложение 1С WEB VMS (Geely) не запущено. Запустите приложение и повторите попытку',
            fg=colors.RED
        )
        raise typer.Exit(1) from no_app_err
    except ElementAmbiguousError as many_window_err:
        typer.secho(
            'Ошибка: Запущено несколько окон 1С WEB VMS (Geely). Оставьте только одно активное окно и повторите попытку',
            fg=colors.RED
        )
        raise typer.Exit(1) from many_window_err


@app.command('works')
def works(
        ctx: typer.Context,
):
    """Тестовая команда"""
    claim: WindowSpecification = ctx.obj 
    # claim.print_control_identifiers()


if __name__ == '__main__':
    app()