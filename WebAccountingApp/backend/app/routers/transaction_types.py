from fastapi import APIRouter, Depends, HTTPException, status
from sqlalchemy.orm import Session
from typing import List
from .. import models, schemas, database
from ..auth import auth

router = APIRouter(
    prefix="/transaction_types",
    tags=["Transaction Types"],
    dependencies=[Depends(auth.get_current_active_user)]
)

@router.get("/", response_model=List[schemas.TransactionType])
def read_transaction_types(
    skip: int = 0,
    limit: int = 100,
    db: Session = Depends(database.get_db),
    current_user: models.User = Depends(auth.get_current_active_user)
):
    transaction_types = db.query(models.TransactionType).offset(skip).limit(limit).all()
    return transaction_types

@router.post("/", response_model=schemas.TransactionType)
def create_transaction_type(
    transaction_type: schemas.TransactionTypeCreate,
    db: Session = Depends(database.get_db),
    current_user: models.User = Depends(auth.get_current_active_user)
):
    # Check if transaction type name or code already exists
    db_transaction_type = db.query(models.TransactionType).filter(
        (models.TransactionType.name == transaction_type.name) |
        (models.TransactionType.code == transaction_type.code)
    ).first()
    if db_transaction_type:
        raise HTTPException(status_code=400, detail="Transaction type name or code already registered")

    db_transaction_type = models.TransactionType(
        name=transaction_type.name,
        code=transaction_type.code,
        description=transaction_type.description
    )
    db.add(db_transaction_type)
    db.commit()
    db.refresh(db_transaction_type)
    return db_transaction_type

@router.get("/{transaction_type_id}", response_model=schemas.TransactionType)
def read_transaction_type(
    transaction_type_id: str,
    db: Session = Depends(database.get_db),
    current_user: models.User = Depends(auth.get_current_active_user)
):
    transaction_type = db.query(models.TransactionType).filter(models.TransactionType.id == transaction_type_id).first()
    if transaction_type is None:
        raise HTTPException(status_code=404, detail="Transaction type not found")
    return transaction_type

@router.put("/{transaction_type_id}", response_model=schemas.TransactionType)
def update_transaction_type(
    transaction_type_id: str,
    transaction_type_update: schemas.TransactionTypeCreate,
    db: Session = Depends(database.get_db),
    current_user: models.User = Depends(auth.get_current_active_user)
):
    db_transaction_type = db.query(models.TransactionType).filter(models.TransactionType.id == transaction_type_id).first()
    if db_transaction_type is None:
        raise HTTPException(status_code=404, detail="Transaction type not found")

    # Update fields
    for field, value in transaction_type_update.dict(exclude_unset=True).items():
        setattr(db_transaction_type, field, value)

    db.commit()
    db.refresh(db_transaction_type)
    return db_transaction_type

@router.delete("/{transaction_type_id}", status_code=status.HTTP_204_NO_CONTENT)
def delete_transaction_type(
    transaction_type_id: str,
    db: Session = Depends(database.get_db),
    current_user: models.User = Depends(auth.get_current_active_user)
):
    db_transaction_type = db.query(models.TransactionType).filter(models.TransactionType.id == transaction_type_id).first()
    if db_transaction_type is None:
        raise HTTPException(status_code=404, detail="Transaction type not found")

    db.delete(db_transaction_type)
    db.commit()
    return