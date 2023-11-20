from datetime import datetime
from typing import Optional

from pydantic import BaseModel, Extra, PositiveInt


class DonationBase(BaseModel):
    comment: Optional[str]
    full_amount: PositiveInt

    class Config:
        extra = Extra.forbid


class DonationCreate(DonationBase):
    pass


class DonationUserDB(DonationBase):
    id: int
    create_date: datetime

    class Config:
        orm_mode = True


class DonationAllDB(DonationUserDB):
    user_id: int
    invested_amount: int = 0
    fully_invested: bool = False
    close_date: Optional[datetime]

    class Config:
        orm_mode = True
