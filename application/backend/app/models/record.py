"""Database models"""
from sqlalchemy import Column, Integer, String, DateTime, func
from app.db.database import Base


class PropertyRecord(Base):
    __tablename__ = "property_records"

    id = Column(Integer, primary_key=True, index=True)
    created_at = Column(DateTime(timezone=True), server_default=func.now())
    updated_at = Column(DateTime(timezone=True), onupdate=func.now())
    
    date = Column(String, nullable=False)
    clerk = Column(String, nullable=False)
    property_address = Column(String, nullable=False)
    client = Column(String, nullable=False)
    inv_type = Column(String, nullable=False)
    status = Column(String, nullable=False, default="Inspected")
    
    final_doc_path = Column(String, nullable=True)
