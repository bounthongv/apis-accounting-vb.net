from fastapi import APIRouter, Depends, HTTPException, status
from sqlalchemy.orm import Session
from typing import List
from .. import models, schemas, database
from ..auth import auth

router = APIRouter(
    prefix="/account_types",
    tags=["Account Types"],
    dependencies=[Depends(auth.get_current_active_user)]
)

@router.get("/", response_model=List[schemas.AccountType])
def read_account_types(
    skip: int = 0,
    limit: int = 100,
    db: Session = Depends(database.get_db),
    current_user: models.User = Depends(auth.get_current_active_user)
):
    account_types = db.query(models.AccountType).offset(skip).limit(limit).all()
    return account_types

@router.post("/", response_model=schemas.AccountType)
def create_account_type(
    account_type: schemas.AccountTypeCreate,
    db: Session = Depends(database.get_db),
    current_user: models.User = Depends(auth.get_current_active_user)
):
    # Check if account type name or code already exists
    db_account_type = db.query(models.AccountType).filter(
        (models.AccountType.name == account_type.name) |
        (models.AccountType.code == account_type.code)
    ).first()
    if db_account_type:
        raise HTTPException(status_code=400, detail="Account type name or code already registered")

    db_account_type = models.AccountType(
        name=account_type.name,
        code=account_type.code,
        description=account_type.description,
        normal_balance=account_type.normal_balance
    )
    db.add(db_account_type)
    db.commit()
    db.refresh(db_account_type)
    return db_account_type

@router.get("/{account_type_id}", response_model=schemas.AccountType)
def read_account_type(
    account_type_id: str,
    db: Session = Depends(database.get_db),
    current_user: models.User = Depends(auth.get_current_active_user)
):
    account_type = db.query(models.AccountType).filter(models.AccountType.id == account_type_id).first()
    if account_type is None:
        raise HTTPException(status_code=404, detail="Account type not found")
    return account_type

@router.put("/{account_type_id}", response_model=schemas.AccountType)
def update_account_type(
    account_type_id: str,
    account_type_update: schemas.AccountTypeCreate,
    db: Session = Depends(database.get_db),
    current_user: models.User = Depends(auth.get_current_active_user)
):
    db_account_type = db.query(models.AccountType).filter(models.AccountType.id == account_type_id).first()
    if db_account_type is None:
        raise HTTPException(status_code=404, detail="Account type not found")

    # Update fields
    for field, value in account_type_update.dict(exclude_unset=True).items():
        setattr(db_account_type, field, value)

    db.commit()
    db.refresh(db_account_type)
    return db_account_type

@router.delete("/{account_type_id}", status_code=status.HTTP_204_NO_CONTENT)
def delete_account_type(
    account_type_id: str,
    db: Session = Depends(database.get_db),
    current_user: models.User = Depends(auth.get_current_active_user)
):
    db_account_type = db.query(models.AccountType).filter(models.AccountType.id == account_type_id).first()
    if db_account_type is None:
        raise HTTPException(status_code=404, detail="Account type not found")

    db.delete(db_account_type)
    db.commit()
    return