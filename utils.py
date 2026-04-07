import logging 

import typer 
from typer import colors 


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