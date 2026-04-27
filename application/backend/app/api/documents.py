"""Document generation API routes"""
import os
import shutil
from fastapi import APIRouter, Depends, HTTPException, UploadFile, File, Form
from sqlalchemy.orm import Session

from app.db.database import get_db
from app.models.record import PropertyRecord
from app.services.document_service import DocumentService

router = APIRouter()


@router.post("/generate-report/{record_id}")
async def generate_report(
    record_id: int,
    middle_doc: UploadFile = File(..., description="Transcription DOCX file"),
    photos_folder: str = Form(..., description="Path to photos folder"),
    db: Session = Depends(get_db)
):
    """
    Generate complete report:
    1. Force landscape on transcription
    2. Generate template
    3. Generate photo index
    4. Merge all documents
    5. Save to photos folder
    6. Update status to Completed
    """
    # Validate record
    record = db.query(PropertyRecord).filter(PropertyRecord.id == record_id).first()
    if not record:
        raise HTTPException(status_code=404, detail="Record not found")
    
    # Validate photos folder
    if not os.path.exists(photos_folder):
        raise HTTPException(status_code=400, detail="Photos folder not found")
    
    # Check for images
    image_files = [
        f for f in os.listdir(photos_folder)
        if f.lower().endswith(('.png', '.jpg', '.jpeg'))
    ]
    if not image_files:
        raise HTTPException(status_code=400, detail="No images found in folder")
    
    # Validate middle doc type
    allowed_types = [
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        "application/msword"
    ]
    if middle_doc.content_type not in allowed_types:
        raise HTTPException(
            status_code=400,
            detail="Invalid document type. Must be .docx or .doc"
        )
    
    try:
        # Save uploaded middle doc temporarily
        temp_dir = "/tmp" if os.name != 'nt' else os.environ.get('TEMP', 'C:\\temp')
        middle_path = os.path.join(temp_dir, f"middle_{record_id}_{middle_doc.filename}")
        
        with open(middle_path, "wb") as f:
            shutil.copyfileobj(middle_doc.file, f)
        
        # Generate complete report
        service = DocumentService()
        final_path = service.generate_complete_report(
            record=record,
            middle_doc_path=middle_path,
            photos_folder=photos_folder,
            output_folder=photos_folder
        )
        
        # Update record status
        record.status = "Completed"
        record.final_doc_path = final_path
        db.commit()
        
        # Clean up temp file
        if os.path.exists(middle_path):
            os.remove(middle_path)
        
        return {
            "success": True,
            "message": "Report generated successfully",
            "document_path": final_path,
            "filename": os.path.basename(final_path)
        }
        
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Document generation failed: {str(e)}")


@router.post("/upload-photos")
async def upload_photos(
    files: list[UploadFile] = File(...),
    destination: str = Form(...)
):
    """Upload photos to a specific folder"""
    if not os.path.exists(destination):
        os.makedirs(destination, exist_ok=True)
    
    uploaded = []
    for file in files:
        if file.content_type not in ["image/jpeg", "image/png", "image/jpg"]:
            continue
        
        file_path = os.path.join(destination, file.filename)
        with open(file_path, "wb") as f:
            shutil.copyfileobj(file.file, f)
        uploaded.append(file.filename)
    
    return {
        "success": True,
        "uploaded": uploaded,
        "count": len(uploaded)
    }
