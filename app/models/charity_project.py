from sqlalchemy import Column, String, Text

from app.models.base import BaseInvestModel


MAX_LENGHT_NAME = 100
PATTERN = 'Благотворительный проект: {name}. {data}'


class CharityProject(BaseInvestModel):
    __tablename__ = 'charityproject'

    name = Column(String(MAX_LENGHT_NAME), unique=True, nullable=False)
    description = Column(Text, nullable=False)

    def __repr__(self):
        return PATTERN.format(
            name=self.name,
            data=super().__repr__()
        )
