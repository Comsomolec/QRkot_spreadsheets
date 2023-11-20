from sqlalchemy.ext.asyncio import AsyncSession
from fastapi import APIRouter, Depends

from app.core.db import get_async_session
from app.crud import donation_crud, charityproject_crud
from app.core.user import current_superuser, current_user
from app.models import User
from app.schemas.donation import (
    DonationAllDB,
    DonationCreate,
    DonationUserDB
)
from app.services.investment_process import investment_process

router = APIRouter()


@router.get(
    '/',
    response_model=list[DonationAllDB],
    response_model_exclude_none=True,
    dependencies=[Depends(current_superuser)],
)
async def get_all_donations(
    session: AsyncSession = Depends(get_async_session),
):
    return await donation_crud.get_multi(session)


@router.get(
    '/my',
    response_model=list[DonationUserDB],
    response_model_exclude_none=True,
)
async def get_user_donations(
    session: AsyncSession = Depends(get_async_session),
    user: User = Depends(current_user),
):
    return await donation_crud.get_by_user(user, session)


@router.post(
    '/',
    response_model=DonationUserDB,
    response_model_exclude_none=True,
)
async def create_dontaion(
    donation: DonationCreate,
    session: AsyncSession = Depends(get_async_session),
    user: User = Depends(current_user),
):
    donation = await donation_crud.create(
        donation, session, user, commit=False
    )
    session.add_all(
        investment_process(
            target=donation,
            sources=await charityproject_crud.get_multi_open_objects(session)
        )
    )
    await session.commit()
    await session.refresh(donation)
    return donation
