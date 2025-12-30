from fastapi import APIRouter, Depends, HTTPException, status
from sqlalchemy.orm import Session
from typing import List
from .. import models, schemas, database
from ..auth import auth

router = APIRouter(
    prefix="/currencies",
    tags=["Currencies"],
    dependencies=[Depends(auth.get_current_active_user)]
)

@router.get("/", response_model=List[schemas.Currency])
def read_currencies(
    skip: int = 0,
    limit: int = 100,
    db: Session = Depends(database.get_db),
    current_user: models.User = Depends(auth.get_current_active_user)
):
    currencies = db.query(models.Currency).offset(skip).limit(limit).all()
    return currencies

@router.post("/", response_model=schemas.Currency)
def create_currency(
    currency: schemas.CurrencyCreate,
    db: Session = Depends(database.get_db),
    current_user: models.User = Depends(auth.get_current_active_user)
):
    # Check if currency code already exists
    db_currency = db.query(models.Currency).filter(
        models.Currency.code == currency.code
    ).first()
    if db_currency:
        raise HTTPException(status_code=400, detail="Currency code already registered")

    db_currency = models.Currency(
        code=currency.code,
        name=currency.name,
        symbol=currency.symbol,
        is_active=currency.is_active
    )
    db.add(db_currency)
    db.commit()
    db.refresh(db_currency)
    return db_currency

@router.get("/{currency_id}", response_model=schemas.Currency)
def read_currency(
    currency_id: str,
    db: Session = Depends(database.get_db),
    current_user: models.User = Depends(auth.get_current_active_user)
):
    currency = db.query(models.Currency).filter(models.Currency.id == currency_id).first()
    if currency is None:
        raise HTTPException(status_code=404, detail="Currency not found")
    return currency

@router.put("/{currency_id}", response_model=schemas.Currency)
def update_currency(
    currency_id: str,
    currency_update: schemas.CurrencyUpdate,
    db: Session = Depends(database.get_db),
    current_user: models.User = Depends(auth.get_current_active_user)
):
    db_currency = db.query(models.Currency).filter(models.Currency.id == currency_id).first()
    if db_currency is None:
        raise HTTPException(status_code=404, detail="Currency not found")

    # Update fields
    for field, value in currency_update.dict(exclude_unset=True).items():
        setattr(db_currency, field, value)

    db.commit()
    db.refresh(db_currency)
    return db_currency

@router.delete("/{currency_id}", status_code=status.HTTP_204_NO_CONTENT)
def delete_currency(
    currency_id: str,
    db: Session = Depends(database.get_db),
    current_user: models.User = Depends(auth.get_current_active_user)
):
    db_currency = db.query(models.Currency).filter(models.Currency.id == currency_id).first()
    if db_currency is None:
        raise HTTPException(status_code=404, detail="Currency not found")

    db.delete(db_currency)
    db.commit()
    return