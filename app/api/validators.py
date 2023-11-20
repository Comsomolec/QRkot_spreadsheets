from http import HTTPStatus

from fastapi import HTTPException
from sqlalchemy.ext.asyncio import AsyncSession

from app.crud.charityproject import charityproject_crud
from app.models import CharityProject


PROJECT_NAME_TAKEN_ERROR = 'Проект с таким именем уже существует!'
PROJECT_NOT_EXISTS_ERROR = 'Проект не найден!'
PROJECT_INVESTED_AMOUNT_IS_NOT_NULL_ERROR = (
    'В проект были внесены средства, не подлежит удалению!'
)
PROJECT_FULL_AMOUNT_IS_LOWER_ERROR = 'Нельзя уменьшать требуемую сумму проекта'
PROJECT_CLOSE_ERROR = 'Закрытый проект нельзя редактировать!'


async def check_name_duplicate(
        project_name: str,
        session: AsyncSession,
) -> None:
    charity_project_id = await charityproject_crud.get_project_id_by_name(
        project_name, session
    )
    if charity_project_id is not None:
        raise HTTPException(
            status_code=HTTPStatus.BAD_REQUEST,
            detail=PROJECT_NAME_TAKEN_ERROR,
        )


async def check_project_exists(
    charity_project_id: int,
    session: AsyncSession,
) -> CharityProject:
    charity_project = await charityproject_crud.get(
        charity_project_id, session
    )
    if not charity_project:
        raise HTTPException(
            status_code=HTTPStatus.NOT_FOUND,
            detail=PROJECT_NOT_EXISTS_ERROR
        )
    return charity_project


async def check_project_invested_amount(
    charity_project_id: int,
    session: AsyncSession,
) -> None:
    charity_project = await charityproject_crud.get(
        charity_project_id, session
    )
    if charity_project.invested_amount:
        raise HTTPException(
            status_code=HTTPStatus.BAD_REQUEST,
            detail=PROJECT_INVESTED_AMOUNT_IS_NOT_NULL_ERROR
        )


async def check_project_full_amount(
    charity_project_id: int,
    new_full_amount: int,
    session: AsyncSession,
) -> None:
    charity_project = await charityproject_crud.get(
        charity_project_id, session
    )
    if new_full_amount < charity_project.invested_amount:
        raise HTTPException(
            status_code=HTTPStatus.BAD_REQUEST,
            detail=PROJECT_FULL_AMOUNT_IS_LOWER_ERROR
        )


async def check_project_status(
    charity_project_id: int,
    session: AsyncSession,
) -> None:
    charity_project = await charityproject_crud.get(
        charity_project_id, session
    )
    if charity_project.fully_invested:
        raise HTTPException(
            status_code=HTTPStatus.BAD_REQUEST,
            detail=PROJECT_CLOSE_ERROR
        )
