"""
Excel QA Validator - FastAPI Main Application
Main entry point for the Excel comparison and validation API
"""

from fastapi import FastAPI, HTTPException, Depends
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from fastapi.responses import JSONResponse
import os
from dotenv import load_dotenv
import logging
from contextlib import asynccontextmanager

# Import routers
from app.routers import upload_router, validation_router

# Load environment variables
load_dotenv()

# Configure logging
logging.basicConfig(
    level=getattr(logging, os.getenv("LOG_LEVEL", "INFO")),
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    handlers=[
        logging.FileHandler(os.getenv("LOG_FILE", "./logs/app.log")),
        logging.StreamHandler()
    ]
)

logger = logging.getLogger(__name__)

# Application lifecycle management
@asynccontextmanager
async def lifespan(app: FastAPI):
    """Application startup and shutdown events"""
    # Startup
    logger.info("🚀 Starting Excel QA Validator API...")
    
    # Create necessary directories
    reports_path = os.getenv("REPORTS_BASE_PATH", "./reports")
    logs_path = os.path.dirname(os.getenv("LOG_FILE", "./logs/app.log"))
    
    os.makedirs(reports_path, exist_ok=True)
    os.makedirs(logs_path, exist_ok=True)
    
    logger.info(f"📁 Reports directory: {reports_path}")
    logger.info(f"📄 Logs directory: {logs_path}")
    
    yield
    
    # Shutdown
    logger.info("🛑 Shutting down Excel QA Validator API...")

# Create FastAPI application
app = FastAPI(
    title=os.getenv("API_TITLE", "Excel QA Validator API"),
    description=os.getenv("API_DESCRIPTION", "Precision Excel file comparison and validation system"),
    version=os.getenv("API_VERSION", "1.0.0"),
    docs_url="/docs" if os.getenv("ENABLE_API_DOCS", "True").lower() == "true" else None,
    redoc_url="/redoc" if os.getenv("ENABLE_API_DOCS", "True").lower() == "true" else None,
    lifespan=lifespan
)

# CORS Configuration
cors_origins = os.getenv("CORS_ORIGINS", "http://localhost:3000,http://localhost:5173").split(",")

app.add_middleware(
    CORSMiddleware,
    allow_origins=cors_origins,
    allow_credentials=True,
    allow_methods=["GET", "POST", "PUT", "DELETE", "OPTIONS"],
    allow_headers=["*"],
)

# Custom exception handler
@app.exception_handler(HTTPException)
async def http_exception_handler(request, exc):
    """Custom HTTP exception handler with detailed error info"""
    logger.error(f"HTTP Exception: {exc.status_code} - {exc.detail}")
    return JSONResponse(
        status_code=exc.status_code,
        content={
            "error": True,
            "message": exc.detail,
            "status_code": exc.status_code,
            "path": str(request.url)
        }
    )

@app.exception_handler(Exception)
async def general_exception_handler(request, exc):
    """General exception handler for unexpected errors"""
    logger.error(f"Unexpected error: {str(exc)}", exc_info=True)
    return JSONResponse(
        status_code=500,
        content={
            "error": True,
            "message": "Internal server error occurred",
            "status_code": 500,
            "path": str(request.url)
        }
    )

# Health check endpoint
@app.get("/", tags=["Health"])
async def root():
    """Root endpoint - API health check"""
    return {
        "message": "Excel QA Validator API is running",
        "status": "healthy",
        "version": os.getenv("API_VERSION", "1.0.0"),
        "docs_url": "/docs" if os.getenv("ENABLE_API_DOCS", "True").lower() == "true" else None
    }

@app.get("/health", tags=["Health"])
async def health_check():
    """Detailed health check endpoint"""
    try:
        # Check if reports directory is accessible
        reports_path = os.getenv("REPORTS_BASE_PATH", "./reports")
        reports_accessible = os.path.exists(reports_path) and os.access(reports_path, os.W_OK)
        
        # Get available reports count
        report_count = 0
        if reports_accessible:
            try:
                report_count = len([d for d in os.listdir(reports_path) 
                                 if os.path.isdir(os.path.join(reports_path, d))])
            except:
                report_count = 0
        
        return {
            "status": "healthy",
            "timestamp": "",
            "version": os.getenv("API_VERSION", "1.0.0"),
            "environment": {
                "debug": os.getenv("DEBUG", "False").lower() == "true",
                "reports_path": reports_path,
                "reports_accessible": reports_accessible,
                "available_reports": report_count,
                "max_file_size_mb": os.getenv("MAX_FILE_SIZE_MB", "50"),
                "allowed_extensions": os.getenv("ALLOWED_EXTENSIONS", "xlsx,xls").split(",")
            }
        }
    except Exception as e:
        logger.error(f"Health check failed: {str(e)}")
        return JSONResponse(
            status_code=503,
            content={
                "status": "unhealthy",
                "error": str(e)
            }
        )

# Include routers
app.include_router(
    upload_router.router,
    prefix="/api/v1/upload",
    tags=["File Upload"],
    responses={404: {"description": "Not found"}}
)

app.include_router(
    validation_router.router,
    prefix="/api/v1/validation",
    tags=["Report Validation"],
    responses={404: {"description": "Not found"}}
)

# Static files (for serving uploaded reports if needed)
reports_path = os.getenv("REPORTS_BASE_PATH", "./reports")
if os.path.exists(reports_path):
    app.mount("/reports", StaticFiles(directory=reports_path), name="reports")

# Development middleware for profiling (if enabled)
if os.getenv("ENABLE_PROFILING", "False").lower() == "true":
    import time
    
    @app.middleware("http")
    async def add_process_time_header(request, call_next):
        start_time = time.time()
        response = await call_next(request)
        process_time = time.time() - start_time
        response.headers["X-Process-Time"] = str(process_time)
        logger.debug(f"Request {request.url} processed in {process_time:.4f}s")
        return response

if __name__ == "__main__":
    import uvicorn
    
    # Development server configuration
    host = os.getenv("HOST", "127.0.0.1")
    port = int(os.getenv("PORT", "8000"))
    debug = os.getenv("DEBUG", "True").lower() == "true"
    reload = os.getenv("RELOAD", "True").lower() == "true"
    
    logger.info(f"🌟 Starting server at http://{host}:{port}")
    
    uvicorn.run(
        "app.main:app",
        host=host,
        port=port,
        reload=reload,
        log_level=os.getenv("LOG_LEVEL", "info").lower(),
        access_log=debug
    )