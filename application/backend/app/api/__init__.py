# API module
from .records import router as records_router
from .documents import router as documents_router

__all__ = ['records_router', 'documents_router']
