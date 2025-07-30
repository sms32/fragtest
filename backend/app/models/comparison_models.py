"""
Excel QA Validator - Data Models
Pydantic models for validation requests, responses, and comparison results
"""

from pydantic import BaseModel, Field, validator
from typing import List, Dict, Any, Optional, Union
from enum import Enum
from datetime import datetime
import os

class ValidationStatus(str, Enum):
    """Validation status enumeration"""
    MATCH = "MATCH"
    MISMATCH = "MISMATCH"
    MISSING_IN_SOURCE = "MISSING_IN_SOURCE"
    MISSING_IN_DEST = "MISSING_IN_DEST"
    CALCULATION_ERROR = "CALCULATION_ERROR"
    STRUCTURAL_ERROR = "STRUCTURAL_ERROR"

class FileType(str, Enum):
    """Supported file types"""
    XLSX = "xlsx"
    XLS = "xls"

class SeverityLevel(str, Enum):
    """Issue severity levels"""
    CRITICAL = "CRITICAL"      # Major data discrepancies, calculation errors
    HIGH = "HIGH"              # Missing records, significant mismatches
    MEDIUM = "MEDIUM"          # Minor formatting differences
    LOW = "LOW"                # Cosmetic differences

# Request Models
class UploadRequest(BaseModel):
    """File upload request model"""
    report_name: str = Field(..., min_length=1, max_length=100, description="Name of the report")
    
    @validator('report_name')
    def validate_report_name(cls, v):
        # Remove invalid characters for filesystem
        invalid_chars = ['<', '>', ':', '"', '/', '\\', '|', '?', '*']
        for char in invalid_chars:
            if char in v:
                raise ValueError(f"Report name cannot contain '{char}'")
        return v.strip()

class ValidationRequest(BaseModel):
    """Validation request model"""
    report_name: str = Field(..., description="Name of the report to validate")
    include_calculation_validation: bool = Field(True, description="Enable calculation validation")
    include_structure_validation: bool = Field(True, description="Enable structure validation")
    precision: float = Field(0.01, description="Numerical comparison precision")

# Response Models
class FileInfo(BaseModel):
    """File information model"""
    filename: str
    size_bytes: int
    size_mb: float
    file_type: FileType
    upload_timestamp: datetime
    exists: bool

class UploadResponse(BaseModel):
    """File upload response model"""
    success: bool
    message: str
    report_name: str
    source_file: Optional[FileInfo] = None
    dest_file: Optional[FileInfo] = None
    upload_path: str

class ReportSummary(BaseModel):
    """Report summary information"""
    report_name: str
    source_file: Optional[FileInfo] = None
    dest_file: Optional[FileInfo] = None
    created_at: datetime
    has_both_files: bool

class ReportsListResponse(BaseModel):
    """Response for listing available reports"""
    success: bool
    reports: List[ReportSummary]
    total_count: int

# Comparison Result Models
class ComparisonResult(BaseModel):
    """Individual comparison result"""
    key: str = Field(..., description="Composite key (Region_Supervisor_Area)")
    section: str = Field(..., description="Section name (BQ, NA, COMBINED)")
    field: str = Field(..., description="Field/column name")
    source_value: Optional[Union[str, int, float]] = Field(None, description="Value from source file")
    dest_value: Optional[Union[str, int, float]] = Field(None, description="Value from destination file")
    status: ValidationStatus = Field(..., description="Comparison status")
    severity: SeverityLevel = Field(..., description="Issue severity level")
    difference: Optional[Union[str, float]] = Field(None, description="Calculated difference (for numbers)")
    notes: Optional[str] = Field(None, description="Additional notes about the comparison")

class SectionSummary(BaseModel):
    """Summary for a specific section"""
    section_name: str
    total_records: int
    total_fields: int
    matches: int
    mismatches: int
    missing_in_source: int
    missing_in_dest: int
    match_percentage: float

class CalculationValidation(BaseModel):
    """Calculation validation result"""
    field: str
    expected_value: Union[int, float]
    actual_value: Union[int, float]
    difference: Union[int, float]
    percentage_error: float
    status: ValidationStatus
    formula_used: Optional[str] = None

class ValidationSummary(BaseModel):
    """Overall validation summary"""
    total_records_compared: int
    total_fields_compared: int
    total_matches: int
    total_mismatches: int
    total_missing_in_source: int
    total_missing_in_dest: int
    overall_match_percentage: float
    critical_issues: int
    high_issues: int
    medium_issues: int
    low_issues: int
    has_calculation_errors: bool
    has_structural_errors: bool

class ValidationResponse(BaseModel):
    """Complete validation response"""
    success: bool
    message: str
    report_name: str
    validation_timestamp: datetime
    processing_time_seconds: float
    
    # Summary information
    summary: ValidationSummary
    section_summaries: List[SectionSummary]
    
    # Detailed results
    comparison_results: List[ComparisonResult]
    calculation_validations: List[CalculationValidation]
    
    # File information
    source_file_info: FileInfo
    dest_file_info: FileInfo
    
    # Export information
    export_available: bool = True
    export_formats: List[str] = ["excel", "json", "csv"]

# Error Models
class ValidationError(BaseModel):
    """Validation error model"""
    error_type: str
    message: str
    field: Optional[str] = None
    value: Optional[str] = None
    suggestion: Optional[str] = None

class ErrorResponse(BaseModel):
    """Error response model"""
    success: bool = False
    error: str
    details: Optional[List[ValidationError]] = None
    timestamp: datetime
    request_id: Optional[str] = None

# Internal Processing Models
class ParsedExcelData(BaseModel):
    """Parsed Excel data structure"""
    sections: Dict[str, List[Dict[str, Any]]] = Field(..., description="Parsed sections with data")
    headers: Dict[str, Dict[int, str]] = Field(..., description="Headers for each section")
    metadata: Dict[str, Any] = Field(default_factory=dict, description="File metadata")
    total_records: int = Field(0, description="Total number of data records")
    parsing_errors: List[str] = Field(default_factory=list, description="Any parsing errors encountered")

class ComparisonContext(BaseModel):
    """Context for comparison operations"""
    source_data: ParsedExcelData
    dest_data: ParsedExcelData
    comparison_settings: Dict[str, Any]
    start_time: datetime
    report_name: str

# Configuration Models
class ValidationConfig(BaseModel):
    """Validation configuration"""
    precision: float = Field(0.01, description="Numerical comparison precision")
    enable_calculation_validation: bool = True
    enable_structure_validation: bool = True
    max_processing_time_seconds: int = 300
    severity_thresholds: Dict[str, float] = Field(
        default_factory=lambda: {
            "critical_percentage_error": 10.0,
            "high_value_threshold": 1000.0,
            "medium_percentage_error": 1.0
        }
    )

# Utility functions for model validation
def get_file_info(file_path: str, filename: str) -> FileInfo:
    """Create FileInfo object from file path"""
    if not os.path.exists(file_path):
        return FileInfo(
            filename=filename,
            size_bytes=0,
            size_mb=0.0,
            file_type=FileType.XLSX,
            upload_timestamp=datetime.now(),
            exists=False
        )
    
    stat = os.stat(file_path)
    file_extension = filename.lower().split('.')[-1]
    
    return FileInfo(
        filename=filename,
        size_bytes=stat.st_size,
        size_mb=round(stat.st_size / (1024 * 1024), 2),
        file_type=FileType.XLSX if file_extension == 'xlsx' else FileType.XLS,
        upload_timestamp=datetime.fromtimestamp(stat.st_mtime),
        exists=True
    )

def determine_severity(comparison_result: ComparisonResult, config: ValidationConfig) -> SeverityLevel:
    """Determine severity level based on comparison result"""
    if comparison_result.status == ValidationStatus.CALCULATION_ERROR:
        return SeverityLevel.CRITICAL
    
    if comparison_result.status in [ValidationStatus.MISSING_IN_SOURCE, ValidationStatus.MISSING_IN_DEST]:
        return SeverityLevel.HIGH
    
    if comparison_result.status == ValidationStatus.MISMATCH:
        # Check if it's a numerical difference
        if isinstance(comparison_result.difference, (int, float)):
            if abs(comparison_result.difference) > config.severity_thresholds["high_value_threshold"]:
                return SeverityLevel.CRITICAL
            elif abs(comparison_result.difference) > config.severity_thresholds["medium_percentage_error"]:
                return SeverityLevel.HIGH
            else:
                return SeverityLevel.MEDIUM
        else:
            return SeverityLevel.MEDIUM
    
    return SeverityLevel.LOW