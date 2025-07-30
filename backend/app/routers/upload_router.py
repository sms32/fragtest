"""
Excel QA Validator - Upload Router
API endpoints for file upload and report management
"""

from fastapi import APIRouter, UploadFile, File, Form, HTTPException, Depends
from fastapi.responses import JSONResponse
from typing import List, Optional
import os
import logging
from datetime import datetime

from app.models.comparison_models import (
    UploadResponse, ReportsListResponse, FileInfo, ReportSummary, 
    ErrorResponse, ValidationError, get_file_info
)
from app.utils.file_utils import FileManager, validate_file_extension, get_safe_filename
from app.services.excel_parser import ExcelParser

logger = logging.getLogger(__name__)

# Create router
router = APIRouter()

# Dependencies
def get_file_manager():
    """Dependency to get FileManager instance"""
    base_path = os.getenv("REPORTS_BASE_PATH", "./reports")
    return FileManager(base_path)

def get_excel_parser():
    """Dependency to get ExcelParser instance"""
    return ExcelParser()

# Configuration
MAX_FILE_SIZE_MB = int(os.getenv("MAX_FILE_SIZE_MB", "50"))
MAX_FILE_SIZE_BYTES = MAX_FILE_SIZE_MB * 1024 * 1024

@router.post("/", response_model=UploadResponse)
async def upload_files(
    report_name: str = Form(..., description="Name of the report"),
    source_file: UploadFile = File(..., description="Source Excel file"),
    dest_file: UploadFile = File(..., description="Destination Excel file"),
    file_manager: FileManager = Depends(get_file_manager),
    excel_parser: ExcelParser = Depends(get_excel_parser)
):
    """
    Upload source and destination Excel files for comparison
    """
    logger.info(f"Upload request for report: {report_name}")
    
    try:
        # Validate report name
        if not report_name or len(report_name.strip()) == 0:
            raise HTTPException(
                status_code=400, 
                detail="Report name cannot be empty"
            )
        
        safe_report_name = get_safe_filename(report_name.strip())
        
        # Validate files
        validation_errors = []
        
        # Check file extensions
        if not validate_file_extension(source_file.filename):
            validation_errors.append(ValidationError(
                error_type="INVALID_FILE_TYPE",
                message=f"Source file '{source_file.filename}' must be .xlsx or .xls",
                field="source_file"
            ))
        
        if not validate_file_extension(dest_file.filename):
            validation_errors.append(ValidationError(
                error_type="INVALID_FILE_TYPE",
                message=f"Destination file '{dest_file.filename}' must be .xlsx or .xls",
                field="dest_file"
            ))
        
        # Check file sizes
        source_content = await source_file.read()
        dest_content = await dest_file.read()
        
        if len(source_content) > MAX_FILE_SIZE_BYTES:
            validation_errors.append(ValidationError(
                error_type="FILE_TOO_LARGE",
                message=f"Source file size ({len(source_content)} bytes) exceeds maximum allowed size ({MAX_FILE_SIZE_MB}MB)",
                field="source_file"
            ))
        
        if len(dest_content) > MAX_FILE_SIZE_BYTES:
            validation_errors.append(ValidationError(
                error_type="FILE_TOO_LARGE",
                message=f"Destination file size ({len(dest_content)} bytes) exceeds maximum allowed size ({MAX_FILE_SIZE_MB}MB)",
                field="dest_file"
            ))
        
        # Check if files are empty
        if len(source_content) == 0:
            validation_errors.append(ValidationError(
                error_type="EMPTY_FILE",
                message="Source file is empty",
                field="source_file"
            ))
        
        if len(dest_content) == 0:
            validation_errors.append(ValidationError(
                error_type="EMPTY_FILE",
                message="Destination file is empty",
                field="dest_file"
            ))
        
        # If there are validation errors, return them
        if validation_errors:
            return JSONResponse(
                status_code=400,
                content=ErrorResponse(
                    error="Validation failed",
                    details=validation_errors,
                    timestamp=datetime.now()
                ).dict()
            )
        
        # Save files
        logger.info(f"Saving files for report: {safe_report_name}")
        
        source_path = await file_manager.save_uploaded_file(
            source_content, safe_report_name, "source"
        )
        dest_path = await file_manager.save_uploaded_file(
            dest_content, safe_report_name, "dest"
        )
        
        # Validate Excel file structure
        source_validation_errors = excel_parser.validate_file_structure(source_path)
        dest_validation_errors = excel_parser.validate_file_structure(dest_path)
        
        all_validation_errors = source_validation_errors + dest_validation_errors
        
        # Create file info objects
        source_file_info = get_file_info(source_path, "source.xlsx")
        dest_file_info = get_file_info(dest_path, "dest.xlsx")
        
        # Prepare response
        upload_response = UploadResponse(
            success=len(all_validation_errors) == 0,
            message="Files uploaded successfully" if len(all_validation_errors) == 0 else f"Files uploaded with {len(all_validation_errors)} validation warnings",
            report_name=safe_report_name,
            source_file=source_file_info,
            dest_file=dest_file_info,
            upload_path=str(file_manager.base_path / safe_report_name)
        )
        
        if all_validation_errors:
            # Return with warnings but still success (files are saved)
            return JSONResponse(
                status_code=200,  # Still successful upload
                content={
                    **upload_response.dict(),
                    "warnings": [error.dict() for error in all_validation_errors]
                }
            )
        
        logger.info(f"Successfully uploaded files for report: {safe_report_name}")
        return upload_response
        
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Upload failed for report {report_name}: {str(e)}")
        raise HTTPException(
            status_code=500,
            detail=f"Upload failed: {str(e)}"
        )

@router.get("/reports", response_model=ReportsListResponse)
async def list_reports(
    file_manager: FileManager = Depends(get_file_manager)
):
    """
    Get list of all available reports
    """
    logger.info("Fetching available reports list")
    
    try:
        reports = file_manager.list_available_reports()
        
        return ReportsListResponse(
            success=True,
            reports=reports,
            total_count=len(reports)
        )
        
    except Exception as e:
        logger.error(f"Failed to list reports: {str(e)}")
        raise HTTPException(
            status_code=500,
            detail=f"Failed to retrieve reports: {str(e)}"
        )

@router.get("/reports/{report_name}")
async def get_report_info(
    report_name: str,
    file_manager: FileManager = Depends(get_file_manager)
):
    """
    Get detailed information about a specific report
    """
    logger.info(f"Fetching info for report: {report_name}")
    
    try:
        safe_report_name = get_safe_filename(report_name)
        source_path, dest_path = file_manager.get_report_files(safe_report_name)
        
        if not source_path and not dest_path:
            raise HTTPException(
                status_code=404,
                detail=f"Report '{report_name}' not found"
            )
        
        # Get file information
        source_info = get_file_info(source_path, "source.xlsx") if source_path else None
        dest_info = get_file_info(dest_path, "dest.xlsx") if dest_path else None
        
        # Get creation time
        created_at = datetime.now()
        if source_info and source_info.exists:
            created_at = min(created_at, source_info.upload_timestamp)
        if dest_info and dest_info.exists:
            created_at = min(created_at, dest_info.upload_timestamp)
        
        report_summary = ReportSummary(
            report_name=safe_report_name,
            source_file=source_info,
            dest_file=dest_info,
            created_at=created_at,
            has_both_files=(source_info and source_info.exists) and (dest_info and dest_info.exists)
        )
        
        return {
            "success": True,
            "report": report_summary.dict(),
            "can_validate": report_summary.has_both_files
        }
        
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Failed to get report info for {report_name}: {str(e)}")
        raise HTTPException(
            status_code=500,
            detail=f"Failed to get report information: {str(e)}"
        )

@router.delete("/reports/{report_name}")
async def delete_report(
    report_name: str,
    file_manager: FileManager = Depends(get_file_manager)
):
    """
    Delete a report and all its files
    """
    logger.info(f"Delete request for report: {report_name}")
    
    try:
        safe_report_name = get_safe_filename(report_name)
        
        # Check if report exists
        source_path, dest_path = file_manager.get_report_files(safe_report_name)
        if not source_path and not dest_path:
            raise HTTPException(
                status_code=404,
                detail=f"Report '{report_name}' not found"
            )
        
        # Delete the report
        success = file_manager.delete_report(safe_report_name)
        
        if not success:
            raise HTTPException(
                status_code=500,
                detail=f"Failed to delete report '{report_name}'"
            )
        
        return {
            "success": True,
            "message": f"Report '{report_name}' deleted successfully",
            "report_name": safe_report_name
        }
        
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Failed to delete report {report_name}: {str(e)}")
        raise HTTPException(
            status_code=500,
            detail=f"Failed to delete report: {str(e)}"
        )

@router.get("/reports/{report_name}/preview")
async def get_file_preview(
    report_name: str,
    file_type: str,  # 'source' or 'dest'
    max_rows: int = 10,
    file_manager: FileManager = Depends(get_file_manager),
    excel_parser: ExcelParser = Depends(get_excel_parser)
):
    """
    Get preview of Excel file structure and sample data
    """
    logger.info(f"Preview request for report {report_name}, file type: {file_type}")
    
    try:
        if file_type not in ['source', 'dest']:
            raise HTTPException(
                status_code=400,
                detail="file_type must be 'source' or 'dest'"
            )
        
        safe_report_name = get_safe_filename(report_name)
        source_path, dest_path = file_manager.get_report_files(safe_report_name)
        
        file_path = source_path if file_type == 'source' else dest_path
        
        if not file_path:
            raise HTTPException(
                status_code=404,
                detail=f"{file_type.capitalize()} file not found for report '{report_name}'"
            )
        
        # Get file preview
        preview = excel_parser.get_file_preview(file_path, max_rows)
        
        return {
            "success": True,
            "report_name": safe_report_name,
            "file_type": file_type,
            "preview": preview
        }
        
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Failed to get preview for {report_name} ({file_type}): {str(e)}")
        raise HTTPException(
            status_code=500,
            detail=f"Failed to get file preview: {str(e)}"
        )

@router.post("/reports/{report_name}/validate-structure")
async def validate_file_structure(
    report_name: str,
    file_manager: FileManager = Depends(get_file_manager),
    excel_parser: ExcelParser = Depends(get_excel_parser)
):
    """
    Validate the structure of both source and destination files
    """
    logger.info(f"Structure validation request for report: {report_name}")
    
    try:
        safe_report_name = get_safe_filename(report_name)
        source_path, dest_path = file_manager.get_report_files(safe_report_name)
        
        if not source_path or not dest_path:
            raise HTTPException(
                status_code=404,
                detail=f"Both source and destination files required for report '{report_name}'"
            )
        
        # Validate both files
        source_errors = excel_parser.validate_file_structure(source_path)
        dest_errors = excel_parser.validate_file_structure(dest_path)
        
        # Combine results
        validation_result = {
            "success": len(source_errors) == 0 and len(dest_errors) == 0,
            "report_name": safe_report_name,
            "source_validation": {
                "valid": len(source_errors) == 0,
                "errors": [error.dict() for error in source_errors],
                "error_count": len(source_errors)
            },
            "dest_validation": {
                "valid": len(dest_errors) == 0,
                "errors": [error.dict() for error in dest_errors],
                "error_count": len(dest_errors)
            },
            "total_errors": len(source_errors) + len(dest_errors),
            "ready_for_comparison": len(source_errors) == 0 and len(dest_errors) == 0
        }
        
        return validation_result
        
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Structure validation failed for {report_name}: {str(e)}")
        raise HTTPException(
            status_code=500,
            detail=f"Structure validation failed: {str(e)}"
        )

@router.get("/storage/info")
async def get_storage_info(
    file_manager: FileManager = Depends(get_file_manager)
):
    """
    Get storage information and statistics
    """
    try:
        from app.utils.file_utils import get_directory_size
        
        # Get storage statistics
        reports = file_manager.list_available_reports()
        
        total_reports = len(reports)
        complete_reports = sum(1 for r in reports if r.has_both_files)
        incomplete_reports = total_reports - complete_reports
        
        # Calculate total storage used
        storage_used_mb = get_directory_size(str(file_manager.base_path))
        
        # Get individual file sizes
        total_source_size = 0
        total_dest_size = 0
        
        for report in reports:
            if report.source_file and report.source_file.exists:
                total_source_size += report.source_file.size_mb
            if report.dest_file and report.dest_file.exists:
                total_dest_size += report.dest_file.size_mb
        
        return {
            "success": True,
            "storage_info": {
                "base_path": str(file_manager.base_path),
                "total_storage_mb": storage_used_mb,
                "total_reports": total_reports,
                "complete_reports": complete_reports,
                "incomplete_reports": incomplete_reports,
                "total_source_files_mb": round(total_source_size, 2),
                "total_dest_files_mb": round(total_dest_size, 2),
                "max_file_size_mb": MAX_FILE_SIZE_MB
            }
        }
        
    except Exception as e:
        logger.error(f"Failed to get storage info: {str(e)}")
        raise HTTPException(
            status_code=500,
            detail=f"Failed to get storage information: {str(e)}"
        )

@router.post("/storage/cleanup")
async def cleanup_old_reports(
    days_old: int = 30,
    file_manager: FileManager = Depends(get_file_manager)
):
    """
    Clean up reports older than specified days
    """
    logger.info(f"Cleanup request for reports older than {days_old} days")
    
    try:
        if days_old < 1:
            raise HTTPException(
                status_code=400,
                detail="days_old must be at least 1"
            )
        
        deleted_count = file_manager.cleanup_old_reports(days_old)
        
        return {
            "success": True,
            "message": f"Cleanup completed",
            "deleted_reports": deleted_count,
            "days_threshold": days_old
        }
        
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Cleanup failed: {str(e)}")
        raise HTTPException(
            status_code=500,
            detail=f"Cleanup failed: {str(e)}"
        )