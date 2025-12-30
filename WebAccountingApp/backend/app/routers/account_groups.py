from fastapi import APIRouter, Depends, HTTPException, status
from sqlalchemy.orm import Session
from typing import List
from .. import models, schemas, database
from ..auth import auth

router = APIRouter(
    prefix="/account_groups",
    tags=["Account Groups"],
    dependencies=[Depends(auth.get_current_active_user)]
)

@router.get("/", response_model=List[schemas.AccountGroup])
def read_account_groups(
    skip: int = 0,
    limit: int = 100,
    db: Session = Depends(database.get_db),
    current_user: models.User = Depends(auth.get_current_active_user)
):
    account_groups = db.query(models.AccountGroup).offset(skip).limit(limit).all()
    return account_groups

@router.post("/", response_model=schemas.AccountGroup)
def create_account_group(
    account_group: schemas.AccountGroupCreate,
    db: Session = Depends(database.get_db),
    current_user: models.User = Depends(auth.get_current_active_user)
):
    # Check if account group code already exists
    db_account_group = db.query(models.AccountGroup).filter(
        models.AccountGroup.code == account_group.code
    ).first()
    if db_account_group:
        raise HTTPException(status_code=400, detail="Account group code already registered")

    db_account_group = models.AccountGroup(
        name=account_group.name,
        code=account_group.code,
        parent_id=account_group.parent_id,
        level=account_group.level,
        is_active=account_group.is_active
    )
    db.add(db_account_group)
    db.commit()
    db.refresh(db_account_group)
    return db_account_group

@router.get("/{account_group_id}", response_model=schemas.AccountGroup)
def read_account_group(
    account_group_id: str,
    db: Session = Depends(database.get_db),
    current_user: models.User = Depends(auth.get_current_active_user)
):
    account_group = db.query(models.AccountGroup).filter(models.AccountGroup.id == account_group_id).first()
    if account_group is None:
        raise HTTPException(status_code=404, detail="Account group not found")
    return account_group

@router.put("/{account_group_id}", response_model=schemas.AccountGroup)
def update_account_group(
    account_group_id: str,
    account_group_update: schemas.AccountGroupUpdate,
    db: Session = Depends(database.get_db),
    current_user: models.User = Depends(auth.get_current_active_user)
):
    db_account_group = db.query(models.AccountGroup).filter(models.AccountGroup.id == account_group_id).first()
    if db_account_group is None:
        raise HTTPException(status_code=404, detail="Account group not found")

    # Update fields
    for field, value in account_group_update.dict(exclude_unset=True).items():
        setattr(db_account_group, field, value)

    db.commit()
    db.refresh(db_account_group)
    return db_account_group

@router.delete("/{account_group_id}", status_code=status.HTTP_204_NO_CONTENT)
def delete_account_group(
    account_group_id: str,
    db: Session = Depends(database.get_db),
    current_user: models.User = Depends(auth.get_current_active_user)
):
    db_account_group = db.query(models.AccountGroup).filter(models.AccountGroup.id == account_group_id).first()
    if db_account_group is None:
        raise HTTPException(status_code=404, detail="Account group not found")

    db.delete(db_account_group)
    db.commit()
    return