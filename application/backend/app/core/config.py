"""Application configuration"""
from pydantic_settings import BaseSettings
from typing import List


class Settings(BaseSettings):
    PROJECT_NAME: str = "InventoryHouse Pro"
    VERSION: str = "2.0.0"
    API_V1_STR: str = "/api/v1"
    
    # Database
    DATABASE_URL: str = "sqlite:///./inventory.db"
    
    # CORS
    CORS_ORIGINS: List[str] = [
        "http://localhost:3000",
        "http://localhost:5173",
        "http://127.0.0.1:3000",
        "http://127.0.0.1:5173",
        "app://."
    ]
    
    # Storage
    BASE_PATH: str = "."
    ASSETS_PATH: str = "./assets"
    
    # Document Settings
    CLERK_OPTIONS: List[str] = ["Tom Tyrrel", "Kevin Crack", "Bill West"]
    STATUS_OPTIONS: List[str] = ["Inspected", "Audio Recorded", "Completed"]
    
    # Photo Grid
    PHOTOS_PER_PAGE: int = 8
    PHOTOS_PER_ROW: int = 4
    PHOTO_WIDTH_CM: float = 5.85
    PHOTO_HEIGHT_CM: float = 6.11
    
    # Company Info
    COMPANY_NAME: str = "Inventory House"
    COMPANY_PHONE: str = "08700 336969"
    COMPANY_EMAIL: str = "info@inventoryhouse.co.uk"
    COMPANY_WEBSITE: str = "www.inventoryhouse.co.uk"
    COMPANY_ADDRESS: str = "Head Office: 3 County Gate London SE9 3UB"
    COMPANY_REGISTRATION: str = "Inventory House Limited. Registered in England & Wales Company No. 5250554"
    
    class Config:
        env_file = ".env"


settings = Settings()
