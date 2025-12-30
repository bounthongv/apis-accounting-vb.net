from fastapi import APIRouter, Depends, HTTPException, status
from sqlalchemy.orm import Session
from typing import List
from .. import models, schemas, database
from ..auth import auth

router = APIRouter(
    prefix="/chart_of_accounts",
    tags=["Chart of Accounts"],
    dependencies=[Depends(auth.get_current_active_user)]
)

@router.get("/", response_model=List[schemas.ChartOfAccounts])
def read_chart_of_accounts(
    skip: int = 0,
    limit: int = 100,
    db: Session = Depends(database.get_db),
    current_user: models.User = Depends(auth.get_current_active_user)
):
    accounts = db.query(models.ChartOfAccounts).offset(skip).limit(limit).all()
    return accounts

@router.post("/", response_model=schemas.ChartOfAccounts)
def create_chart_of_account(
    account: schemas.ChartOfAccountsCreate,
    db: Session = Depends(database.get_db),
    current_user: models.User = Depends(auth.get_current_active_user)
):
    # Check if account code already exists
    db_account = db.query(models.ChartOfAccounts).filter(
        models.ChartOfAccounts.account_code == account.account_code
    ).first()
    if db_account:
        raise HTTPException(status_code=400, detail="Account code already registered")

    db_account = models.ChartOfAccounts(
        account_code=account.account_code,
        account_name=account.account_name,
        account_type_id=account.account_type_id,
        account_group_id=account.account_group_id,
        parent_account_id=account.parent_account_id,
        level=account.level,
        is_active=account.is_active,
        is_system_account=account.is_system_account,
        opening_balance=account.opening_balance,
        current_balance=account.current_balance
    )
    db.add(db_account)
    db.commit()
    db.refresh(db_account)
    return db_account

@router.get("/{account_id}", response_model=schemas.ChartOfAccounts)
def read_chart_of_account(
    account_id: str,
    db: Session = Depends(database.get_db),
    current_user: models.User = Depends(auth.get_current_active_user)
):
    account = db.query(models.ChartOfAccounts).filter(models.ChartOfAccounts.id == account_id).first()
    if account is None:
        raise HTTPException(status_code=404, detail="Account not found")
    return account

@router.put("/{account_id}", response_model=schemas.ChartOfAccounts)
def update_chart_of_account(
    account_id: str,
    account_update: schemas.ChartOfAccountsUpdate,
    db: Session = Depends(database.get_db),
    current_user: models.User = Depends(auth.get_current_active_user)
):
    db_account = db.query(models.ChartOfAccounts).filter(models.ChartOfAccounts.id == account_id).first()
    if db_account is None:
        raise HTTPException(status_code=404, detail="Account not found")

    # Update fields
    for field, value in account_update.dict(exclude_unset=True).items():
        setattr(db_account, field, value)

    db.commit()
    db.refresh(db_account)
    return db_account

@router.delete("/{account_id}", status_code=status.HTTP_204_NO_CONTENT)
def delete_chart_of_account(
    account_id: str,
    db: Session = Depends(database.get_db),
    current_user: models.User = Depends(auth.get_current_active_user)
):
    db_account = db.query(models.ChartOfAccounts).filter(models.ChartOfAccounts.id == account_id).first()
    if db_account is None:
        raise HTTPException(status_code=404, detail="Account not found")

    db.delete(db_account)
    db.commit()
    return