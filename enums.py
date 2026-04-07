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


class DetailsTableColumns(StrEnum):
    """Столбцы таблицы деталей в форме рекламации"""
    ROW_NUMBER = 'N'
    PART_NUMBER = 'Деталь/Каталожный номер'
    PART_NAME = 'Наименование детали'
    UNIT = 'Ед. изм.'
    UPD_NUMBER = '№ УПД, Реализации'
    UPD_DATE = 'Дата накладной'
    CONTRACTOR = 'Контрагент накладной'
    QUANTITY = 'Количество'
    PRICE = 'Цена'
    AMOUNT_WITHOUT_NDS = 'Сумма без НДС'
    AMOUNT_WITH_NDS = 'Сумма с НДС'
    MODIFIED_DATE = 'Дата изменения'

