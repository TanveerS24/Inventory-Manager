"""InventoryHouse Pro - FastAPI Backend"""
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles

from app.core.config import settings
from app.db.database import init_db
from app.api import records, documents

# Initialize database
init_db()

# Create app
app = FastAPI(
    title=settings.PROJECT_NAME,
    version=settings.VERSION,
    description="Modern inventory management system with document generation"
)

# CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=settings.CORS_ORIGINS,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Include routers
app.include_router(records.router, prefix=f"{settings.API_V1_STR}/records", tags=["records"])
app.include_router(documents.router, prefix=f"{settings.API_V1_STR}/documents", tags=["documents"])


@app.get("/health")
def health_check():
    return {
        "status": "healthy",
        "app": settings.PROJECT_NAME,
        "version": settings.VERSION
    }


@app.get("/")
def root():
    return {
        "message": f"Welcome to {settings.PROJECT_NAME}",
        "version": settings.VERSION,
        "docs": "/api/v1/docs"
    }


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="127.0.0.1", port=8000)
