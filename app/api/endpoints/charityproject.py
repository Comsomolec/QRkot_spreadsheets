from sqlalchemy.ext.asyncio import AsyncSession
from fastapi import APIRouter, Depends

from app.core.db import get_async_session
from app.crud import charityproject_crud, donation_crud
from app.core.user import current_superuser
from app.schemas.charityproject import (
    CharityProjectCreate,
    CharityProjectDB,
    CharityProjectUpdate
)
from app.api.validators import (
    check_name_duplicate,
    check_project_exists,
    check_project_invested_amount,
    check_project_full_amount,
    check_project_status
)
from app.services.investment_process import investment_process


router = APIRouter()


@router.get(
    '/',
    response_model=list[CharityProjectDB],
    response_model_exclude_none=True
)
async def get_all_charityprojects(
    session: AsyncSession = Depends(get_async_session),
):
    return await charityproject_crud.get_multi(session)


@router.post(
    '/',
    response_model=CharityProjectDB,
    response_model_exclude_none=True,
    dependencies=[Depends(current_superuser)],
)
async def create_charityproject(
    charity_project: CharityProjectCreate,
    session: AsyncSession = Depends(get_async_session),
):
    await check_name_duplicate(charity_project.name, session)
    charity_project = await charityproject_crud.create(
        charity_project, session, commit=False
    )
    session.add_all(
        investment_process(
            target=charity_project,
            sources=await donation_crud.get_multi_open_objects(session)
        )
    )
    await session.commit()
    await session.refresh(charity_project)
    return charity_project


@router.delete(
    '/{project_id}',
    response_model=CharityProjectDB,
    dependencies=[Depends(current_superuser)],
)
async def delete_charityproject(
    project_id: int,
    session: AsyncSession = Depends(get_async_session),
):
    charity_project = await check_project_exists(project_id, session)
    await check_project_invested_amount(project_id, session)
    charity_project = await charityproject_crud.remove(
        charity_project, session
    )
    return charity_project


@router.patch(
    '/{project_id}',
    response_model=CharityProjectDB,
    dependencies=[Depends(current_superuser)],
)
async def partially_update_charity_project(
    project_id: int,
    obj_in: CharityProjectUpdate,
    session: AsyncSession = Depends(get_async_session),
):
    charity_project = await check_project_exists(project_id, session)
    await check_project_status(project_id, session)
    if obj_in.name is not None:
        await check_name_duplicate(obj_in.name, session)
    if obj_in.full_amount is not None:
        await check_project_full_amount(
            project_id, obj_in.full_amount, session
        )
    charity_project = await charityproject_crud.update(
        charity_project, obj_in, session
    )
    return charity_project
