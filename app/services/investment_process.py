from datetime import datetime as dt

from app.models.base import BaseInvestModel


def investment_process(
    target: BaseInvestModel, sources: list[BaseInvestModel],
) -> list[BaseInvestModel]:
    result = []
    for object in sources:
        if target.fully_invested:
            break
        invest_amount = min(
            object.full_amount - object.invested_amount,
            target.full_amount - target.invested_amount
        )
        for element in [target, object]:
            element.invested_amount += invest_amount
            if element.full_amount == element.invested_amount:
                element.fully_invested = True
                element.close_date = dt.now()
        result.append(object)
    return result
