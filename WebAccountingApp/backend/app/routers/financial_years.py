from fastapi import APIRouter, Depends, HTTPException, status
from sqlalchemy.orm import Session
from typing import List
from .. import models, schemas, database
from ..auth import auth

router = APIRouter(
    prefix="/financial_years",
    tags=["Financial Years"],
    dependencies=[Depends(auth.get_current_active_user)]
)

@router.get("/", response_model=List[schemas.FinancialYear])
def read_financial_years(
    skip: int = 0,
    limit: int = 100,
    db: Session = Depends(database.get_db),
    current_user: models.User = Depends(auth.get_current_active_user)
):
    financial_years = db.query(models.FinancialYear).offset(skip).limit(limit).all()
    return financial_years

@router.post("/", response_model=schemas.FinancialYear)
def create_financial_year(
    financial_year: schemas.FinancialYearCreate,
    db: Session = Depends(database.get_db),
    current_user: models.User = Depends(auth.get_current_active_user)
):
    # Check if financial year name already exists
    db_financial_year = db.query(models.FinancialYear).filter(
        models.FinancialYear.year_name == financial_year.year_name
    ).first()
    if db_financial_year:
        raise HTTPException(status_code=400, detail="Financial year name already registered")

    # Check if the date range overlaps with existing financial years
    overlapping_year = db.query(models.FinancialYear).filter(
        (models.FinancialYear.start_date <= financial_year.end_date) &
        (models.FinancialYear.end_date >= financial_year.start_date)
    ).first()
    
    if overlapping_year:
        raise HTTPException(status_code=400, detail="Financial year date range overlaps with an existing year")

    db_financial_year = models.FinancialYear(
        year_name=financial_year.year_name,
        start_date=financial_year.start_date,
        end_date=financial_year.end_date,
        is_closed=financial_year.is_closed,
        is_active=financial_year.is_active
    )
    db.add(db_financial_year)
    db.commit()
    db.refresh(db_financial_year)
    return db_financial_year

@router.get("/{financial_year_id}", response_model=schemas.FinancialYear)
def read_financial_year(
    financial_year_id: str,
    db: Session = Depends(database.get_db),
    current_user: models.User = Depends(auth.get_current_active_user)
):
    financial_year = db.query(models.FinancialYear).filter(models.FinancialYear.id == financial_year_id).first()
    if financial_year is None:
        raise HTTPException(status_code=404, detail="Financial year not found")
    return financial_year

@router.put("/{financial_year_id}", response_model=schemas.FinancialYear)
def update_financial_year(
    financial_year_id: str,
    financial_year_update: schemas.FinancialYearUpdate,
    db: Session = Depends(database.get_db),
    current_user: models.User = Depends(auth.get_current_active_user)
):
    db_financial_year = db.query(models.FinancialYear).filter(models.FinancialYear.id == financial_year_id).first()
    if db_financial_year is None:
        raise HTTPException(status_code=404, detail="Financial year not found")

    # Check if the date range overlaps with other financial years (excluding current one)
    overlapping_year = db.query(models.FinancialYear).filter(
        (models.FinancialYear.id != financial_year_id) &
        (models.FinancialYear.start_date <= financial_year_update.end_date) &
        (models.FinancialYear.end_date >= financial_year_update.start_date)
    ).first()
    
    if overlapping_year:
        raise HTTPException(status_code=400, detail="Financial year date range overlaps with an existing year")

    # Update fields
    for field, value in financial_year_update.dict(exclude_unset=True).items():
        setattr(db_financial_year, field, value)

    db.commit()
    db.refresh(db_financial_year)
    return db_financial_year

@router.delete("/{financial_year_id}", status_code=status.HTTP_204_NO_CONTENT)
def delete_financial_year(
    financial_year_id: str,
    db: Session = Depends(database.get_db),
    current_user: models.User = Depends(auth.get_current_active_user)
):
    db_financial_year = db.query(models.FinancialYear).filter(models.FinancialYear.id == financial_year_id).first()
    if db_financial_year is None:
        raise HTTPException(status_code=404, detail="Financial year not found")

    # Check if financial year is in use (has related transactions)
    # This is a simplified check - in a real app, you'd check for related transactions
    db.commit()
    db.refresh(db_financial_year)
    
    db.delete(db_financial_year)
    db.commit()
    return