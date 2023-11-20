from datetime import datetime

from sqlalchemy import Boolean, Column, DateTime, Integer, CheckConstraint

from app.core.db import Base


PATTERN = (
    'Сумма: {amount}. Внесено {invested}. Отметка о закрытии: {close}.'
    'Дата создания {create_date}. Дата закрытия {close_date}.'
)


class BaseInvestModel(Base):
    __abstract__ = True
    __table_args__ = (
        CheckConstraint('full_amount > 0'),
        CheckConstraint('invested_amount <= full_amount'),
    )

    full_amount = Column(Integer)
    invested_amount = Column(Integer, default=0)
    fully_invested = Column(Boolean, default=False)
    create_date = Column(DateTime, default=datetime.now)
    close_date = Column(DateTime)

    def __repr__(self) -> str:
        return PATTERN.format(
            amount=self.full_amount,
            invested=self.invested_amount,
            close=self.fully_invested,
            create_date=self.create_date,
            close_date=self.close_date,
        )
