from aiogoogle import Aiogoogle
from fastapi import APIRouter, Depends
from sqlalchemy.ext.asyncio import AsyncSession

from app.core.db import get_async_session
from app.core.google_client import get_service
from app.core.user import current_superuser

from app.crud.charityproject import charityproject_crud
from app.services.google_api import (
    spreadsheets_create,
    set_user_permissions,
    spreadsheets_update_value
)

CONNECTION_ERROR = 'Возникла ошибка при обращении к серверу {error}'
URL = 'https://docs.google.com/spreadsheets/d/{spreadsheet_id}'

router = APIRouter()


@router.get(
    '/',
    dependencies=[Depends(current_superuser)]
)
async def get_sheet_invested_projects(
        session: AsyncSession = Depends(get_async_session),
        wrapper_services: Aiogoogle = Depends(get_service)
):
    projects = await charityproject_crud.get_projects_by_completion_rate(
        session,
    )
    spreadsheet_id = await spreadsheets_create(wrapper_services)
    await set_user_permissions(spreadsheet_id, wrapper_services)
    try:
        await spreadsheets_update_value(spreadsheet_id,
                                        projects,
                                        wrapper_services)
    except ConnectionError as error:
        return CONNECTION_ERROR.format(error=error)
    return URL.format(spreadsheet_id=spreadsheet_id)
