from copy import deepcopy
from datetime import datetime

from aiogoogle import Aiogoogle

from app.core.config import settings
from app.models import CharityProject


FORMAT = "%Y/%m/%d %H:%M:%S"

COLUMN_COUNT = 3
ROW_COUNT = 100

SHEET_BODY = dict(
    properties=dict(
        title='Отчёт от {}',
        locale='ru_RU',
    ),
    sheets=[dict(properties=dict(
        sheetType='GRID',
        sheetId=0,
        title='Список закрытых проектов',
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
    'Передано неверное количество столбцов. Передано {input}.'
    f'Ограничение по максимуму: {COLUMN_COUNT}'
)
ROW_ERROR = (
    'Передано неверное количество строк. Передано {input}.'
    f'Ограничение по максимуму: {ROW_COUNT}'
)


async def spreadsheets_create(
    wrapper_services: Aiogoogle,
    body: dict[str, dict] = SHEET_BODY
) -> str:
    spreadsheet_body = deepcopy(body)
    spreadsheet_body['properties']['title'] = datetime.now().strftime(FORMAT)
    service = await wrapper_services.discover('sheets', 'v4')
    response = await wrapper_services.as_service_account(
        service.spreadsheets.create(json=spreadsheet_body)
    )
    return response['spreadsheetId'], response['spreadsheetUrl']


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
    table_head: list[list[str]] = TABLE_HEAD
) -> None:
    head = deepcopy(table_head)
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
            ROW_ERROR.format(input=new_row_count)
        )
    len_columns = max(map(len, head))
    if len_columns > COLUMN_COUNT:
        raise ValueError(
            COLUMN_ERROR.format(input=len_columns)
        )
    await wrapper_services.as_service_account(
        service.spreadsheets.values.update(
            spreadsheetId=spreadsheet_id,
            range=f'R1C1:R{new_row_count}C{len_columns}',
            valueInputOption=INPUT_OPTIONS,
            json=update_body
        )
    )
