from fastapi import APIRouter, Depends, HTTPException, status
from sqlalchemy.orm import Session
from typing import List
from .. import models, schemas, database
from ..auth import auth

router = APIRouter(
    prefix="/journal_entries",
    tags=["Journal Entries"],
    dependencies=[Depends(auth.get_current_active_user)]
)

@router.get("/", response_model=List[schemas.JournalEntry])
def read_journal_entries(
    skip: int = 0,
    limit: int = 100,
    db: Session = Depends(database.get_db),
    current_user: models.User = Depends(auth.get_current_active_user)
):
    entries = db.query(models.JournalEntry).offset(skip).limit(limit).all()
    return entries

@router.post("/", response_model=schemas.JournalEntry)
def create_journal_entry(
    entry: schemas.JournalEntryCreate,
    db: Session = Depends(database.get_db),
    current_user: models.User = Depends(auth.get_current_active_user)
):
    # Generate entry number if not provided
    import uuid
    entry_number = f"JE-{str(uuid.uuid4())[:8].upper()}"
    
    db_entry = models.JournalEntry(
        entry_number=entry_number,
        entry_date=entry.entry_date,
        reference_number=entry.reference_number,
        description=entry.description,
        transaction_type_id=entry.transaction_type_id,
        created_by=entry.created_by,
        is_posted=entry.is_posted
    )
    db.add(db_entry)
    db.commit()
    db.refresh(db_entry)
    return db_entry

@router.get("/{entry_id}", response_model=schemas.JournalEntry)
def read_journal_entry(
    entry_id: str,
    db: Session = Depends(database.get_db),
    current_user: models.User = Depends(auth.get_current_active_user)
):
    entry = db.query(models.JournalEntry).filter(models.JournalEntry.id == entry_id).first()
    if entry is None:
        raise HTTPException(status_code=404, detail="Journal entry not found")
    return entry

@router.put("/{entry_id}", response_model=schemas.JournalEntry)
def update_journal_entry(
    entry_id: str,
    entry_update: schemas.JournalEntryUpdate,
    db: Session = Depends(database.get_db),
    current_user: models.User = Depends(auth.get_current_active_user)
):
    db_entry = db.query(models.JournalEntry).filter(models.JournalEntry.id == entry_id).first()
    if db_entry is None:
        raise HTTPException(status_code=404, detail="Journal entry not found")

    # Update fields
    for field, value in entry_update.dict(exclude_unset=True).items():
        setattr(db_entry, field, value)

    db.commit()
    db.refresh(db_entry)
    return db_entry

@router.delete("/{entry_id}", status_code=status.HTTP_204_NO_CONTENT)
def delete_journal_entry(
    entry_id: str,
    db: Session = Depends(database.get_db),
    current_user: models.User = Depends(auth.get_current_active_user)
):
    db_entry = db.query(models.JournalEntry).filter(models.JournalEntry.id == entry_id).first()
    if db_entry is None:
        raise HTTPException(status_code=404, detail="Journal entry not found")

    db.delete(db_entry)
    db.commit()
    return