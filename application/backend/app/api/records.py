"""Record API routes"""
from typing import Optional
from fastapi import APIRouter, Depends, HTTPException, Query
from sqlalchemy.orm import Session
from datetime import datetime

from app.db.database import get_db
from app.models.record import PropertyRecord
from app.schemas.record import (
    PropertyRecordCreate,
    PropertyRecordUpdate,
    PropertyRecordResponse,
    RecordListResponse
)
from app.core.config import settings

router = APIRouter()


@router.get("/", response_model=RecordListResponse)
def get_records(
    skip: int = Query(0, ge=0),
    limit: int = Query(100, ge=1, le=1000),
    client: Optional[str] = None,
    clerk: Optional[str] = None,
    address: Optional[str] = None,
    status: Optional[str] = None,
    db: Session = Depends(get_db)
):
    """Get all records with optional filtering"""
    query = db.query(PropertyRecord)
    
    if client:
        query = query.filter(PropertyRecord.client.ilike(f"%{client}%"))
    if clerk:
        query = query.filter(PropertyRecord.clerk.ilike(f"%{clerk}%"))
    if address:
        query = query.filter(PropertyRecord.property_address.ilike(f"%{address}%"))
    if status:
        query = query.filter(PropertyRecord.status.ilike(f"%{status}%"))
    
    total = query.count()
    records = query.order_by(PropertyRecord.created_at.desc()).offset(skip).limit(limit).all()
    
    # Format date for display
    formatted_records = []
    for record in records:
        record_data = PropertyRecordResponse.from_orm(record)
        try:
            record_data.date = datetime.strptime(record.date, "%Y-%m-%d").strftime("%d-%m-%Y")
        except:
            pass
        formatted_records.append(record_data)
    
    return {
        "items": formatted_records,
        "total": total
    }


@router.post("/", response_model=PropertyRecordResponse, status_code=201)
def create_record(
    record: PropertyRecordCreate,
    db: Session = Depends(get_db)
):
    """Create a new record"""
    # Validate clerk
    if record.clerk not in settings.CLERK_OPTIONS:
        raise HTTPException(
            status_code=400,
            detail=f"Invalid clerk. Must be one of: {', '.join(settings.CLERK_OPTIONS)}"
        )
    
    # Validate status
    if record.status not in settings.STATUS_OPTIONS:
        record.status = "Inspected"
    
    db_record = PropertyRecord(
        **record.model_dump(),
        date=datetime.now().strftime("%Y-%m-%d")
    )
    db.add(db_record)
    db.commit()
    db.refresh(db_record)
    
    # Format date for response
    response = PropertyRecordResponse.from_orm(db_record)
    try:
        response.date = datetime.strptime(db_record.date, "%Y-%m-%d").strftime("%d-%m-%Y")
    except:
        pass
    
    return response


@router.get("/{record_id}", response_model=PropertyRecordResponse)
def get_record(
    record_id: int,
    db: Session = Depends(get_db)
):
    """Get a single record by ID"""
    record = db.query(PropertyRecord).filter(PropertyRecord.id == record_id).first()
    if not record:
        raise HTTPException(status_code=404, detail="Record not found")
    
    response = PropertyRecordResponse.from_orm(record)
    try:
        response.date = datetime.strptime(record.date, "%Y-%m-%d").strftime("%d-%m-%Y")
    except:
        pass
    
    return response


@router.put("/{record_id}", response_model=PropertyRecordResponse)
def update_record(
    record_id: int,
    record_update: PropertyRecordUpdate,
    db: Session = Depends(get_db)
):
    """Update a record"""
    db_record = db.query(PropertyRecord).filter(PropertyRecord.id == record_id).first()
    if not db_record:
        raise HTTPException(status_code=404, detail="Record not found")
    
    # Don't allow status change if already completed
    if db_record.status == "Completed" and record_update.status:
        record_update.status = None
    
    update_data = record_update.model_dump(exclude_unset=True)
    for field, value in update_data.items():
        setattr(db_record, field, value)
    
    db.commit()
    db.refresh(db_record)
    
    response = PropertyRecordResponse.from_orm(db_record)
    try:
        response.date = datetime.strptime(db_record.date, "%Y-%m-%d").strftime("%d-%m-%Y")
    except:
        pass
    
    return response


@router.delete("/{record_id}", status_code=204)
def delete_record(
    record_id: int,
    db: Session = Depends(get_db)
):
    """Delete a record"""
    db_record = db.query(PropertyRecord).filter(PropertyRecord.id == record_id).first()
    if not db_record:
        raise HTTPException(status_code=404, detail="Record not found")
    
    db.delete(db_record)
    db.commit()
    return None


@router.get("/options/clerks")
def get_clerk_options():
    """Get available clerk options"""
    return {"clerks": settings.CLERK_OPTIONS}


@router.get("/options/statuses")
def get_status_options():
    """Get available status options"""
    return {"statuses": settings.STATUS_OPTIONS}
