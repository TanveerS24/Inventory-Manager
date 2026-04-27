"""Pydantic schemas"""
from pydantic import BaseModel, Field
from typing import Optional
from datetime import datetime


class PropertyRecordBase(BaseModel):
    clerk: str = Field(..., min_length=1, max_length=100)
    property_address: str = Field(..., min_length=1, max_length=500)
    client: str = Field(..., min_length=1, max_length=200)
    inv_type: str = Field(..., min_length=1, max_length=100)
    status: str = Field(default="Inspected")


class PropertyRecordCreate(PropertyRecordBase):
    pass


class PropertyRecordUpdate(BaseModel):
    clerk: Optional[str] = Field(None, min_length=1, max_length=100)
    property_address: Optional[str] = Field(None, min_length=1, max_length=500)
    client: Optional[str] = Field(None, min_length=1, max_length=200)
    inv_type: Optional[str] = Field(None, min_length=1, max_length=100)
    status: Optional[str] = None


class PropertyRecordResponse(PropertyRecordBase):
    id: int
    created_at: datetime
    updated_at: Optional[datetime] = None
    date: str
    final_doc_path: Optional[str] = None

    class Config:
        from_attributes = True


class RecordListResponse(BaseModel):
    items: list[PropertyRecordResponse]
    total: int
