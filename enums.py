from enum import StrEnum 


class OneCWebWMS(StrEnum):
    # Приложение/окно
    APP_NAME = '1С WEB VMS (Geely)'
    MAIN_WINDOW_TITLE = 'WEB VMS - ООО «Бизнес Кар М» - Москва'

    # Вкладки
    CLAIM_CREATE_TAB_TITLE = 'Рекламация VMS (создание)'
    CLAIM_CREATE_TAB_PATTERN = r'Рекламация VMS \(создание\)\s?\*?'


class MyApp(StrEnum):
    NAME = 'onec-claim-geely-util'
    ADD_JOBS = 'add-jobs'
    ADD_DETAILS = 'add-details'

    def __repr__(self) -> str:
        return f"'{self.value}'"
