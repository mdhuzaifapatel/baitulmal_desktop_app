from sqlalchemy import func
from models import Receipt, MadCategory, Register

def mad_wise(db):

    data = (
        db.query(
            MadCategory.name,
            func.sum(Receipt.amount)
        )
        .join(Receipt, Receipt.mad_id == MadCategory.id)
        .group_by(MadCategory.name)
        .all()
    )

    return data


def register_wise(db):

    data = (
        db.query(
            Register.name,
            func.sum(Receipt.amount)
        )
        .join(Receipt, Receipt.register_id == Register.id)
        .group_by(Register.name)
        .all()
    )

    return data