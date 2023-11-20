from sqlalchemy import Column, ForeignKey, Integer, Text

from app.models.base import BaseInvestModel


PATTERN = (
    'Пожертвование внес: {user}. {data}'
)


class Donation(BaseInvestModel):
    __tablename__ = 'donation'

    user_id = Column(Integer, ForeignKey('user.id'))
    comment = Column(Text)

    def __repr__(self):
        return PATTERN.format(
            user=self.user_id,
            data=super().__repr__()
        )
