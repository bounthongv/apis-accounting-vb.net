from fastapi import APIRouter, Depends, HTTPException, status
from sqlalchemy.orm import Session
from typing import List
from .. import models, schemas, database
from ..auth import auth

router = APIRouter(
    prefix="/parties",
    tags=["Parties"],
    dependencies=[Depends(auth.get_current_active_user)]
)

@router.get("/", response_model=List[schemas.Party])
def read_parties(
    skip: int = 0,
    limit: int = 100,
    db: Session = Depends(database.get_db),
    current_user: models.User = Depends(auth.get_current_active_user)
):
    parties = db.query(models.Party).offset(skip).limit(limit).all()
    return parties

@router.post("/", response_model=schemas.Party)
def create_party(
    party: schemas.PartyCreate,
    db: Session = Depends(database.get_db),
    current_user: models.User = Depends(auth.get_current_active_user)
):
    # Check if party code already exists
    db_party = db.query(models.Party).filter(
        models.Party.party_code == party.party_code
    ).first()
    if db_party:
        raise HTTPException(status_code=400, detail="Party code already registered")

    db_party = models.Party(
        party_code=party.party_code,
        party_name=party.party_name,
        party_type_id=party.party_type_id,
        address=party.address,
        phone=party.phone,
        email=party.email,
        tax_id=party.tax_id,
        is_active=party.is_active,
        credit_limit=party.credit_limit
    )
    db.add(db_party)
    db.commit()
    db.refresh(db_party)
    
    # Create party detail
    db_party_detail = models.PartyDetail(
        party_id=db_party.id,
        is_customer=party.is_customer if hasattr(party, 'is_customer') else False,
        is_supplier=party.is_supplier if hasattr(party, 'is_supplier') else False,
        payment_terms=party.payment_terms if hasattr(party, 'payment_terms') else 0
    )
    db.add(db_party_detail)
    db.commit()
    
    return db_party

@router.get("/{party_id}", response_model=schemas.Party)
def read_party(
    party_id: str,
    db: Session = Depends(database.get_db),
    current_user: models.User = Depends(auth.get_current_active_user)
):
    party = db.query(models.Party).filter(models.Party.id == party_id).first()
    if party is None:
        raise HTTPException(status_code=404, detail="Party not found")
    return party

@router.put("/{party_id}", response_model=schemas.Party)
def update_party(
    party_id: str,
    party_update: schemas.PartyUpdate,
    db: Session = Depends(database.get_db),
    current_user: models.User = Depends(auth.get_current_active_user)
):
    db_party = db.query(models.Party).filter(models.Party.id == party_id).first()
    if db_party is None:
        raise HTTPException(status_code=404, detail="Party not found")

    # Update fields
    for field, value in party_update.dict(exclude_unset=True).items():
        setattr(db_party, field, value)

    db.commit()
    db.refresh(db_party)
    return db_party

@router.delete("/{party_id}", status_code=status.HTTP_204_NO_CONTENT)
def delete_party(
    party_id: str,
    db: Session = Depends(database.get_db),
    current_user: models.User = Depends(auth.get_current_active_user)
):
    db_party = db.query(models.Party).filter(models.Party.id == party_id).first()
    if db_party is None:
        raise HTTPException(status_code=404, detail="Party not found")

    # Delete related party details
    db_party_detail = db.query(models.PartyDetail).filter(models.PartyDetail.party_id == party_id).first()
    if db_party_detail:
        db.delete(db_party_detail)
    
    db.delete(db_party)
    db.commit()
    return