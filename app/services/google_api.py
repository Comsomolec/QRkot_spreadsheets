from copy import deepcopy
from datetime import datetime

from aiogoogle import Aiogoogle

from app.core.config import settings
from app.models import CharityProject


FORMAT = "%Y/%m/%d %H:%M:%S"

COLUMN_COUNT = 3
LOCALE = 'ru_RU'
ROW_COUNT = 100
TITLE_SHEET = 'Список закрытых проектов'
TITLE_FILE = 'Отчёт от {}'
SHEET_ID = 0
SHEET_TYPE = 'GRID'

SHEET_BODY = dict(
    properties=dict(
        title=TITLE_FILE,
        locale=LOCALE,
    ),
    sheets=[dict(properties=dict(
        sheetType=SHEET_TYPE,
        sheetId=SHEET_ID,
        title=TITLE_SHEET,
        gridProperties=dict(
            rowCount=ROW_COUNT,
            columnCount=COLUMN_COUNT,
        )
    ))]
)

FIELDS = 'id'
PERMISSION_TYPE = 'user'
PERMISSION_ROLE = 'writer'

DIMENSIONS = 'ROWS'
INPUT_OPTIONS = 'USER_ENTERED'
TABLE_HEAD = [
    ['Отчёт от', ''],
    ['Топ проектов по скорости закрытия'],
    ['Название проекта', 'Время сбора', 'Описание']
]

COLUMN_ERROR = (
    'Передано неверное количество столбцов.'
    'Передано {input}. Должно быть {expect}'
)
ROW_ERROR = (
    'Передано неверное количество строк.'
    'Передано {input}. Должно быть {expect}'
)


async def spreadsheets_create(
    wrapper_services: Aiogoogle,
    spreadsheet_body: dict[str, dict] = deepcopy(SHEET_BODY)
) -> str:
    spreadsheet_body['properties']['title'] = datetime.now().strftime(FORMAT)
    service = await wrapper_services.discover('sheets', 'v4')
    response = await wrapper_services.as_service_account(
        service.spreadsheets.create(json=spreadsheet_body)
    )
    spreadsheet_id = response['spreadsheetId']
    return spreadsheet_id


async def set_user_permissions(
        spreadsheet_id: str,
        wrapper_services: Aiogoogle
) -> None:
    permissions_body = {'type': PERMISSION_TYPE,
                        'role': PERMISSION_ROLE,
                        'emailAddress': settings.email}
    service = await wrapper_services.discover('drive', 'v3')
    await wrapper_services.as_service_account(
        service.permissions.create(
            fileId=spreadsheet_id,
            json=permissions_body,
            fields=FIELDS
        )
    )


async def spreadsheets_update_value(
    spreadsheet_id: str,
    projects: list[CharityProject],
    wrapper_services: Aiogoogle,
    head: list[list[str]] = deepcopy(TABLE_HEAD)
) -> None:
    service = await wrapper_services.discover('sheets', 'v4')
    head[0][1] = datetime.now().strftime(FORMAT)
    table_values = [
        *head,
        *[list(map(str, (
            project.name,
            project.close_date - project.create_date,
            project.description
        ))) for project in projects]]
    update_body = {
        'majorDimension': DIMENSIONS,
        'values': table_values
    }
    new_row_count = len(table_values)
    if new_row_count > ROW_COUNT:
        raise ValueError(
            ROW_ERROR.format(input=new_row_count, expect=ROW_COUNT)
        )
    len_columns = max(len(row) for row in head)
    if len_columns > COLUMN_COUNT:
        raise ValueError(
            COLUMN_ERROR.format(input=len_columns, expect=COLUMN_COUNT)
        )
    await wrapper_services.as_service_account(
        service.spreadsheets.values.update(
            spreadsheetId=spreadsheet_id,
            range=f'R1C1:R{new_row_count}C{len_columns}',
            valueInputOption=INPUT_OPTIONS,
            json=update_body
        )
    )
