from datetime import datetime

from aiogoogle import Aiogoogle

from app.core.config import settings
from app.models import CharityProject


FORMAT = "%Y/%m/%d %H:%M:%S"

DATE_NOW = datetime.now().strftime(FORMAT)

COLUMN_COUNT = 3
LOCALE = 'ru_RU'
ROW_COUNT = 100
TITLE_SHEET = 'Список закрытых проектов'
SHEET_ID = 0
SHEET_TYPE = 'GRID'

FIELDS = 'id'
PERMISSION_TYPE = 'user'
PERMISSION_ROLE = 'writer'

COUNT_TABLE_HEAD_ROW = 4
DIMENSIONS = 'ROWS'
INPUT_OPTIONS = 'USER_ENTERED'
RANGE_FIELD = 'A1:C{len}'
TABLE_HEAD = [
    ['Отчёт от', f'{DATE_NOW}'],
    ['Топ проектов по скорости закрытия'],
    ['Название проекта', 'Время сбора']
]
TITLE_FILE = f'Отчёт от {DATE_NOW}'


async def spreadsheets_create(wrapper_services: Aiogoogle) -> str:
    service = await wrapper_services.discover('sheets', 'v4')
    spreadsheet_body = {
        'properties': {'title': TITLE_FILE,
                       'locale': LOCALE},
        'sheets': [
            {
                'properties': {
                    'sheetType': SHEET_TYPE,
                    'sheetId': SHEET_ID,
                    'title': TITLE_SHEET,
                    'gridProperties': {
                        'rowCount': ROW_COUNT,
                        'columnCount': COLUMN_COUNT
                    }
                }
            }
        ]
    }
    response = await wrapper_services.as_service_account(
        service.spreadsheets.create(json=spreadsheet_body)
    )
    spreadsheetid = response['spreadsheetId']
    return spreadsheetid


async def set_user_permissions(
        spreadsheetid: str,
        wrapper_services: Aiogoogle
) -> None:
    permissions_body = {'type': PERMISSION_TYPE,
                        'role': PERMISSION_ROLE,
                        'emailAddress': settings.email}
    service = await wrapper_services.discover('drive', 'v3')
    await wrapper_services.as_service_account(
        service.permissions.create(
            fileId=spreadsheetid,
            json=permissions_body,
            fields=FIELDS
        )
    )


async def spreadsheets_update_value(
        spreadsheetid: str,
        projects: list[CharityProject],
        wrapper_services: Aiogoogle
) -> None:
    service = await wrapper_services.discover('sheets', 'v4')
    table_values = TABLE_HEAD
    for project in projects:
        new_row = [
            project.name,
            str(project.close_date - project.create_date),
            project.description,
        ]
        table_values.append(new_row)
    update_body = {
        'majorDimension': DIMENSIONS,
        'values': table_values
    }
    new_row_count = len(projects)
    await wrapper_services.as_service_account(
        service.spreadsheets.values.update(
            spreadsheetId=spreadsheetid,
            range=RANGE_FIELD.format(len=new_row_count + COUNT_TABLE_HEAD_ROW),
            valueInputOption=INPUT_OPTIONS,
            json=update_body
        )
    )
