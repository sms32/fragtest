"""
Excel QA Validator - Validation Router
API endpoints for report validation and comparison
"""

from fastapi import APIRouter, HTTPException, Depends, BackgroundTasks
from fastapi.responses import FileResponse, JSONResponse
from typing import Optional
import os
import logging
from datetime import datetime
import time
import asyncio
from pathlib import Path

from app.models.comparison_models import (
    ValidationRequest, ValidationResponse, ValidationConfig,
    ErrorResponse, get_file_info
)
from app.services.excel_parser import ExcelParser
from app.services.report_comparator import ReportComparator
from app.utils.file_utils import FileManager, ExcelExporter, get_safe_filename

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

def get_excel_exporter():
    """Dependency to get ExcelExporter instance"""
    return ExcelExporter()

# Configuration
MAX_PROCESSING_TIME = int(os.getenv("VALIDATION_TIMEOUT_SECONDS", "300"))

@router.post("/{report_name}", response_model=ValidationResponse)
async def validate_report(
    report_name: str,
    validation_request: Optional[ValidationRequest] = None,
    file_manager: FileManager = Depends(get_file_manager),
    excel_parser: ExcelParser = Depends(get_excel_parser)
):
    """
    Perform comprehensive validation of a report by comparing source and destination files
    """
    logger.info(f"Validation request for report: {report_name}")
    start_time = time.time()
    
    try:
        # Use default validation request if none provided
        if validation_request is None:
            validation_request = ValidationRequest(report_name=report_name)
        
        safe_report_name = get_safe_filename(report_name)
        
        # Get file paths
        source_path, dest_path = file_manager.get_report_files(safe_report_name)
        
        if not source_path or not dest_path:
            missing_files = []
            if not source_path:
                missing_files.append("source")
            if not dest_path:
                missing_files.append("destination")
            
            raise HTTPException(
                status_code=404,
                detail=f"Missing files for report '{report_name}': {', '.join(missing_files)}"
            )
        
        # Validate files exist and are accessible
        if not os.path.exists(source_path):
            raise HTTPException(
                status_code=404,
                detail=f"Source file not found: {source_path}"
            )
        
        if not os.path.exists(dest_path):
            raise HTTPException(
                status_code=404,
                detail=f"Destination file not found: {dest_path}"
            )
        
        logger.info(f"Starting validation process for {safe_report_name}")
        logger.info(f"Source: {source_path}")
        logger.info(f"Dest: {dest_path}")
        
        # Parse source file
        logger.info("Parsing source file...")
        source_data = excel_parser.parse_excel_file(source_path)
        
        if source_data.parsing_errors:
            logger.warning(f"Source parsing errors: {source_data.parsing_errors}")
        
        # Parse destination file
        logger.info("Parsing destination file...")
        dest_data = excel_parser.parse_excel_file(dest_path)
        
        if dest_data.parsing_errors:
            logger.warning(f"Destination parsing errors: {dest_data.parsing_errors}")
        
        # Check if parsing was successful
        if not source_data.sections and not dest_data.sections:
            raise HTTPException(
                status_code=422,
                detail="Both files failed to parse. Please check file formats and content."
            )
        
        if not source_data.sections:
            raise HTTPException(
                status_code=422,
                detail=f"Source file parsing failed: {', '.join(source_data.parsing_errors)}"
            )
        
        if not dest_data.sections:
            raise HTTPException(
                status_code=422,
                detail=f"Destination file parsing failed: {', '.join(dest_data.parsing_errors)}"
            )
        
        # Create validation configuration
        config = ValidationConfig(
            precision=validation_request.precision,
            enable_calculation_validation=validation_request.include_calculation_validation,
            enable_structure_validation=validation_request.include_structure_validation
        )
        
        # Perform comparison
        logger.info("Starting report comparison...")
        comparator = ReportComparator(config)
        
        # Check for timeout
        elapsed_time = time.time() - start_time
        if elapsed_time > MAX_PROCESSING_TIME:
            raise HTTPException(
                status_code=408,
                detail=f"Validation timeout after {elapsed_time:.2f} seconds"
            )
        
        comparison_results, calculation_validations = comparator.compare_reports(source_data, dest_data)
        
        # Generate summaries
        logger.info("Generating validation summaries...")
        validation_summary = comparator.generate_summary(comparison_results, calculation_validations)
        section_summaries = comparator.generate_section_summaries(comparison_results)
        
        # Get file information
        source_file_info = get_file_info(source_path, "source.xlsx")
        dest_file_info = get_file_info(dest_path, "dest.xlsx")
        
        # Calculate processing time
        processing_time = time.time() - start_time
        
        logger.info(f"Validation completed in {processing_time:.2f} seconds")
        logger.info(f"Results: {len(comparison_results)} comparisons, {validation_summary.total_mismatches} mismatches")
        
        # Create response
        validation_response = ValidationResponse(
            success=True,
            message=f"Validation completed successfully with {validation_summary.total_mismatches} discrepancies found",
            report_name=safe_report_name,
            validation_timestamp=datetime.now(),
            processing_time_seconds=round(processing_time, 3),
            summary=validation_summary,
            section_summaries=section_summaries,
            comparison_results=comparison_results,
            calculation_validations=calculation_validations,
            source_file_info=source_file_info,
            dest_file_info=dest_file_info
        )
        
        return validation_response
        
    except HTTPException:
        raise
    except Exception as e:
        processing_time = time.time() - start_time
        logger.error(f"Validation failed for {report_name} after {processing_time:.2f}s: {str(e)}")
        
        return JSONResponse(
            status_code=500,
            content=ErrorResponse(
                error=f"Validation failed: {str(e)}",
                timestamp=datetime.now(),
                request_id=f"{report_name}_{int(time.time())}"
            ).dict()
        )

@router.get("/{report_name}/status")
async def get_validation_status(
    report_name: str,
    file_manager: FileManager = Depends(get_file_manager)
):
    """
    Get the current validation status and readiness of a report
    """
    logger.info(f"Status check for report: {report_name}")
    
    try:
        safe_report_name = get_safe_filename(report_name)
        source_path, dest_path = file_manager.get_report_files(safe_report_name)
        
        # Check file existence
        source_exists = source_path and os.path.exists(source_path)
        dest_exists = dest_path and os.path.exists(dest_path)
        
        # Get file info
        source_info = get_file_info(source_path, "source.xlsx") if source_exists else None
        dest_info = get_file_info(dest_path, "dest.xlsx") if dest_exists else None
        
        # Determine readiness
        ready_for_validation = source_exists and dest_exists
        
        status_info = {
            "report_name": safe_report_name,
            "ready_for_validation": ready_for_validation,
            "source_file": {
                "exists": source_exists,
                "info": source_info.dict() if source_info else None
            },
            "dest_file": {
                "exists": dest_exists,
                "info": dest_info.dict() if dest_info else None
            },
            "requirements": {
                "both_files_present": source_exists and dest_exists,
                "files_accessible": True,  # If we got here, they're accessible
                "valid_formats": True  # Assuming Excel format validation passed during upload
            }
        }
        
        if not ready_for_validation:
            missing = []
            if not source_exists:
                missing.append("source file")
            if not dest_exists:
                missing.append("destination file")
            
            status_info["message"] = f"Cannot validate: missing {' and '.join(missing)}"
        else:
            status_info["message"] = "Ready for validation"
        
        return {
            "success": True,
            **status_info
        }
        
    except Exception as e:
        logger.error(f"Failed to get status for {report_name}: {str(e)}")
        raise HTTPException(
            status_code=500,
            detail=f"Failed to get validation status: {str(e)}"
        )

@router.post("/{report_name}/export")
async def export_validation_results(
    report_name: str,
    format: str = "excel",  # excel, json, csv
    validation_request: Optional[ValidationRequest] = None,
    background_tasks: BackgroundTasks = BackgroundTasks(),
    file_manager: FileManager = Depends(get_file_manager),
    excel_parser: ExcelParser = Depends(get_excel_parser),
    excel_exporter: ExcelExporter = Depends(get_excel_exporter)
):
    """
    Export validation results to specified format
    """
    logger.info(f"Export request for report {report_name} in {format} format")
    
    try:
        if format not in ["excel", "json", "csv"]:
            raise HTTPException(
                status_code=400,
                detail="Format must be 'excel', 'json', or 'csv'"
            )
        
        # First, perform validation to get results
        safe_report_name = get_safe_filename(report_name)
        source_path, dest_path = file_manager.get_report_files(safe_report_name)
        
        if not source_path or not dest_path:
            raise HTTPException(
                status_code=404,
                detail=f"Missing files for report '{report_name}'"
            )
        
        # Use default validation request if none provided
        if validation_request is None:
            validation_request = ValidationRequest(report_name=report_name)
        
        # Parse files
        source_data = excel_parser.parse_excel_file(source_path)
        dest_data = excel_parser.parse_excel_file(dest_path)
        
        # Create validation configuration
        config = ValidationConfig(
            precision=validation_request.precision,
            enable_calculation_validation=validation_request.include_calculation_validation,
            enable_structure_validation=validation_request.include_structure_validation
        )
        
        # Perform comparison
        comparator = ReportComparator(config)
        comparison_results, calculation_validations = comparator.compare_reports(source_data, dest_data)
        
        # Generate summaries
        validation_summary = comparator.generate_summary(comparison_results, calculation_validations)
        section_summaries = comparator.generate_section_summaries(comparison_results)
        
        # Create export directory
        export_dir = file_manager.base_path / safe_report_name / "exports"
        export_dir.mkdir(exist_ok=True)
        
        # Generate timestamp for filename
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        if format == "excel":
            # Export to Excel
            export_filename = f"{safe_report_name}_validation_{timestamp}.xlsx"
            export_path = export_dir / export_filename
            
            excel_exporter.export_validation_results(
                safe_report_name,
                validation_summary,
                section_summaries,
                comparison_results,
                calculation_validations,
                str(export_path)
            )
            
            # Schedule cleanup of export file after 1 hour
            background_tasks.add_task(cleanup_export_file, str(export_path), 3600)
            
            return FileResponse(
                path=str(export_path),
                filename=export_filename,
                media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        
        elif format == "json":
            # Export to JSON
            import json
            
            export_filename = f"{safe_report_name}_validation_{timestamp}.json"
            export_path = export_dir / export_filename
            
            export_data = {
                "report_name": safe_report_name,
                "validation_timestamp": datetime.now().isoformat(),
                "summary": validation_summary.dict(),
                "section_summaries": [s.dict() for s in section_summaries],
                "comparison_results": [r.dict() for r in comparison_results],
                "calculation_validations": [c.dict() for c in calculation_validations]
            }
            
            with open(export_path, 'w') as f:
                json.dump(export_data, f, indent=2, default=str)
            
            # Schedule cleanup
            background_tasks.add_task(cleanup_export_file, str(export_path), 3600)
            
            return FileResponse(
                path=str(export_path),
                filename=export_filename,
                media_type="application/json"
            )
        
        elif format == "csv":
            # Export to CSV
            import pandas as pd
            
            export_filename = f"{safe_report_name}_validation_{timestamp}.csv"
            export_path = export_dir / export_filename
            
            # Convert comparison results to DataFrame
            csv_data = []
            for result in comparison_results:
                csv_data.append({
                    'Key': result.key,
                    'Section': result.section,
                    'Field': result.field,
                    'Source_Value': result.source_value,
                    'Dest_Value': result.dest_value,
                    'Status': result.status.value,
                    'Severity': result.severity.value,
                    'Difference': result.difference,
                    'Notes': result.notes
                })
            
            df = pd.DataFrame(csv_data)
            df.to_csv(export_path, index=False)
            
            # Schedule cleanup
            background_tasks.add_task(cleanup_export_file, str(export_path), 3600)
            
            return FileResponse(
                path=str(export_path),
                filename=export_filename,
                media_type="text/csv"
            )
        
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Export failed for {report_name}: {str(e)}")
        raise HTTPException(
            status_code=500,
            detail=f"Export failed: {str(e)}"
        )

@router.get("/{report_name}/summary")
async def get_validation_summary(
    report_name: str,
    file_manager: FileManager = Depends(get_file_manager),
    excel_parser: ExcelParser = Depends(get_excel_parser)
):
    """
    Get a quick summary of validation results without full comparison
    """
    logger.info(f"Summary request for report: {report_name}")
    
    try:
        safe_report_name = get_safe_filename(report_name)
        source_path, dest_path = file_manager.get_report_files(safe_report_name)
        
        if not source_path or not dest_path:
            raise HTTPException(
                status_code=404,
                detail=f"Missing files for report '{report_name}'"
            )
        
        # Quick parse to get basic info
        source_data = excel_parser.parse_excel_file(source_path)
        dest_data = excel_parser.parse_excel_file(dest_path)
        
        # Basic comparison without detailed field matching
        source_sections = set(source_data.sections.keys())
        dest_sections = set(dest_data.sections.keys())
        
        quick_summary = {
            "report_name": safe_report_name,
            "source_info": {
                "sections": list(source_sections),
                "total_records": source_data.total_records,
                "parsing_errors": len(source_data.parsing_errors)
            },
            "dest_info": {
                "sections": list(dest_sections),
                "total_records": dest_data.total_records,
                "parsing_errors": len(dest_data.parsing_errors)
            },
            "structural_analysis": {
                "common_sections": list(source_sections & dest_sections),
                "missing_in_dest": list(source_sections - dest_sections),
                "extra_in_dest": list(dest_sections - source_sections),
                "structural_match": source_sections == dest_sections
            },
            "ready_for_detailed_validation": (
                len(source_data.parsing_errors) == 0 and 
                len(dest_data.parsing_errors) == 0 and
                bool(source_sections & dest_sections)
            )
        }
        
        return {
            "success": True,
            **quick_summary
        }
        
    except Exception as e:
        logger.error(f"Failed to get summary for {report_name}: {str(e)}")
        raise HTTPException(
            status_code=500,
            detail=f"Failed to get validation summary: {str(e)}"
        )

# Background task functions
async def cleanup_export_file(file_path: str, delay_seconds: int):
    """Background task to clean up export files after specified delay"""
    await asyncio.sleep(delay_seconds)
    try:
        if os.path.exists(file_path):
            os.remove(file_path)
            logger.info(f"Cleaned up export file: {file_path}")
    except Exception as e:
        logger.error(f"Failed to cleanup export file {file_path}: {e}")

# Health check for validation service
@router.get("/health")
async def validation_health_check():
    """Health check for validation service"""
    return {
        "service": "validation",
        "status": "healthy",
        "timestamp": datetime.now().isoformat(),
        "max_processing_time": MAX_PROCESSING_TIME,
        "supported_formats": ["excel", "json", "csv"]
    }