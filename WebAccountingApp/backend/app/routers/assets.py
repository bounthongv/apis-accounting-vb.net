from fastapi import APIRouter, Depends, HTTPException, status
from sqlalchemy.orm import Session
from typing import List
from .. import models, schemas, database
from ..auth import auth

router = APIRouter(
    prefix="/assets",
    tags=["Assets"],
    dependencies=[Depends(auth.get_current_active_user)]
)

@router.get("/", response_model=List[schemas.Asset])
def read_assets(
    skip: int = 0,
    limit: int = 100,
    db: Session = Depends(database.get_db),
    current_user: models.User = Depends(auth.get_current_active_user)
):
    assets = db.query(models.Asset).offset(skip).limit(limit).all()
    return assets

@router.post("/", response_model=schemas.Asset)
def create_asset(
    asset: schemas.AssetCreate,
    db: Session = Depends(database.get_db),
    current_user: models.User = Depends(auth.get_current_active_user)
):
    # Check if asset code already exists
    db_asset = db.query(models.Asset).filter(
        models.Asset.asset_code == asset.asset_code
    ).first()
    if db_asset:
        raise HTTPException(status_code=400, detail="Asset code already registered")

    db_asset = models.Asset(
        asset_code=asset.asset_code,
        asset_name=asset.asset_name,
        category_id=asset.category_id,
        description=asset.description,
        purchase_date=asset.purchase_date,
        purchase_price=asset.purchase_price,
        current_value=asset.current_value,
        depreciation_start_date=asset.depreciation_start_date,
        useful_life=asset.useful_life,
        salvage_value=asset.salvage_value,
        status=asset.status,
        location=asset.location,
        is_active=asset.is_active,
        accumulated_depreciation=0,  # Initialize to 0
        net_book_value=asset.current_value  # Initially equal to current value
    )
    db.add(db_asset)
    db.commit()
    db.refresh(db_asset)
    return db_asset

@router.get("/{asset_id}", response_model=schemas.Asset)
def read_asset(
    asset_id: str,
    db: Session = Depends(database.get_db),
    current_user: models.User = Depends(auth.get_current_active_user)
):
    asset = db.query(models.Asset).filter(models.Asset.id == asset_id).first()
    if asset is None:
        raise HTTPException(status_code=404, detail="Asset not found")
    return asset

@router.put("/{asset_id}", response_model=schemas.Asset)
def update_asset(
    asset_id: str,
    asset_update: schemas.AssetUpdate,
    db: Session = Depends(database.get_db),
    current_user: models.User = Depends(auth.get_current_active_user)
):
    db_asset = db.query(models.Asset).filter(models.Asset.id == asset_id).first()
    if db_asset is None:
        raise HTTPException(status_code=404, detail="Asset not found")

    # Update fields
    for field, value in asset_update.dict(exclude_unset=True).items():
        setattr(db_asset, field, value)

    # Recalculate net book value
    db_asset.net_book_value = db_asset.current_value - db_asset.accumulated_depreciation

    db.commit()
    db.refresh(db_asset)
    return db_asset

@router.delete("/{asset_id}", status_code=status.HTTP_204_NO_CONTENT)
def delete_asset(
    asset_id: str,
    db: Session = Depends(database.get_db),
    current_user: models.User = Depends(auth.get_current_active_user)
):
    db_asset = db.query(models.Asset).filter(models.Asset.id == asset_id).first()
    if db_asset is None:
        raise HTTPException(status_code=404, detail="Asset not found")

    db.delete(db_asset)
    db.commit()
    return