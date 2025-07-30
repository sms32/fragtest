"""
Excel QA Validator - Report Comparator Service
Precision comparison engine for Excel reports with comprehensive validation
"""

from typing import Dict, List, Any, Tuple, Optional
import pandas as pd
import logging
from datetime import datetime
import math

from app.models.comparison_models import (
    ComparisonResult, ValidationStatus, SeverityLevel, 
    ValidationSummary, SectionSummary, CalculationValidation,
    ParsedExcelData, ValidationConfig, determine_severity
)

logger = logging.getLogger(__name__)

class NumericalComparator:
    """Handles numerical value comparisons with precision"""
    
    def __init__(self, precision: float = 0.01):
        self.precision = precision
    
    def compare_numbers(self, source_val: Any, dest_val: Any) -> Tuple[bool, Optional[float]]:
        """Compare two numerical values with precision tolerance"""
        # Convert to numbers if possible
        source_num = self._to_number(source_val)
        dest_num = self._to_number(dest_val)
        
        # If either is not a number, do string comparison
        if source_num is None or dest_num is None:
            return str(source_val) == str(dest_val), None
        
        # Calculate difference
        difference = abs(source_num - dest_num)
        
        # Check if within precision tolerance
        is_match = difference <= self.precision
        
        return is_match, source_num - dest_num
    
    def _to_number(self, value: Any) -> Optional[float]:
        """Convert value to number if possible"""
        if value is None:
            return None
        
        if isinstance(value, (int, float)):
            return float(value)
        
        if isinstance(value, str):
            # Handle percentages
            if value.strip().endswith('%'):
                try:
                    return float(value.strip()[:-1])
                except ValueError:
                    return None
            
            # Handle parentheses (negative)
            if value.strip().startswith('(') and value.strip().endswith(')'):
                try:
                    return -float(value.strip()[1:-1].replace(',', ''))
                except ValueError:
                    return None
            
            # Handle regular numbers with commas
            try:
                return float(value.replace(',', ''))
            except ValueError:
                return None
        
        return None

class StructuralValidator:
    """Validates structural consistency between source and destination"""
    
    def validate_sections(self, source_data: ParsedExcelData, dest_data: ParsedExcelData) -> List[ComparisonResult]:
        """Validate that sections are consistent between source and destination"""
        results = []
        
        source_sections = set(source_data.sections.keys())
        dest_sections = set(dest_data.sections.keys())
        
        # Check for missing sections
        missing_in_dest = source_sections - dest_sections
        missing_in_source = dest_sections - source_sections
        
        for section in missing_in_dest:
            results.append(ComparisonResult(
                key=f"SECTION_{section}",
                section="STRUCTURAL",
                field="ENTIRE_SECTION",
                source_value="EXISTS",
                dest_value="MISSING",
                status=ValidationStatus.MISSING_IN_DEST,
                severity=SeverityLevel.CRITICAL,
                notes=f"Section '{section}' exists in source but missing in destination"
            ))
        
        for section in missing_in_source:
            results.append(ComparisonResult(
                key=f"SECTION_{section}",
                section="STRUCTURAL",
                field="ENTIRE_SECTION",
                source_value="MISSING",
                dest_value="EXISTS",
                status=ValidationStatus.MISSING_IN_SOURCE,
                severity=SeverityLevel.HIGH,
                notes=f"Section '{section}' exists in destination but missing in source"
            ))
        
        return results
    
    def validate_record_counts(self, source_data: ParsedExcelData, dest_data: ParsedExcelData) -> List[ComparisonResult]:
        """Validate record counts between sections"""
        results = []
        
        common_sections = set(source_data.sections.keys()) & set(dest_data.sections.keys())
        
        for section in common_sections:
            source_count = len(source_data.sections[section])
            dest_count = len(dest_data.sections[section])
            
            if source_count != dest_count:
                results.append(ComparisonResult(
                    key=f"COUNT_{section}",
                    section=section,
                    field="RECORD_COUNT",
                    source_value=source_count,
                    dest_value=dest_count,
                    status=ValidationStatus.MISMATCH,
                    severity=SeverityLevel.HIGH,
                    difference=source_count - dest_count,
                    notes=f"Record count mismatch in section '{section}'"
                ))
        
        return results

class CalculationValidator:
    """Validates calculations and totals"""
    
    def __init__(self, numerical_comparator: NumericalComparator):
        self.numerical_comparator = numerical_comparator
    
    def validate_totals(self, source_data: ParsedExcelData, dest_data: ParsedExcelData) -> List[CalculationValidation]:
        """Validate that totals match the sum of individual records"""
        validations = []
        
        # Define fields that should be summable
        summable_fields = [
            'WK Slab', 'Day Sale', 'Day Slab', 'Day Stale', 
            'WTD Slab', 'WTD Sale', 'WTD Stale', 'Wk Sale LY'
        ]
        
        # Check if destination has a COMBINED or TOTAL section
        dest_sections = dest_data.sections
        total_section = None
        
        for section_name in ['COMBINED', 'CENTRAL', 'TOTAL']:
            if section_name in dest_sections:
                total_section = section_name
                break
        
        if not total_section:
            return validations
        
        # Get individual data sections (BQ, NA)
        individual_sections = [s for s in dest_sections.keys() if s not in ['COMBINED', 'CENTRAL', 'TOTAL']]
        
        if not individual_sections:
            return validations
        
        # Calculate expected totals and compare with actual totals
        total_records = dest_sections[total_section]
        
        if not total_records:
            return validations
        
        # Assume the first record in total section is the summary
        total_record = total_records[0]
        
        for field in summable_fields:
            expected_total = 0
            field_found = False
            
            # Sum from individual sections
            for section in individual_sections:
                for record in dest_sections[section]:
                    if field in record and record[field] is not None:
                        field_value = self.numerical_comparator._to_number(record[field])
                        if field_value is not None:
                            expected_total += field_value
                            field_found = True
            
            if not field_found:
                continue
            
            # Compare with actual total
            if field in total_record and total_record[field] is not None:
                actual_total = self.numerical_comparator._to_number(total_record[field])
                
                if actual_total is not None:
                    difference = expected_total - actual_total
                    percentage_error = abs(difference / expected_total * 100) if expected_total != 0 else 0
                    
                    status = ValidationStatus.MATCH
                    if abs(difference) > self.numerical_comparator.precision:
                        status = ValidationStatus.CALCULATION_ERROR
                    
                    validations.append(CalculationValidation(
                        field=field,
                        expected_value=expected_total,
                        actual_value=actual_total,
                        difference=difference,
                        percentage_error=percentage_error,
                        status=status,
                        formula_used=f"Sum of {', '.join(individual_sections)} sections"
                    ))
        
        return validations

class ReportComparator:
    """Main report comparison engine"""
    
    def __init__(self, config: Optional[ValidationConfig] = None):
        self.config = config or ValidationConfig()
        self.numerical_comparator = NumericalComparator(self.config.precision)
        self.structural_validator = StructuralValidator()
        self.calculation_validator = CalculationValidator(self.numerical_comparator)
    
    def compare_reports(self, source_data: ParsedExcelData, dest_data: ParsedExcelData) -> Tuple[List[ComparisonResult], List[CalculationValidation]]:
        """Main comparison method that returns all results"""
        logger.info("Starting comprehensive report comparison")
        
        comparison_results = []
        calculation_validations = []
        
        # 1. Structural validation
        if self.config.enable_structure_validation:
            logger.info("Performing structural validation")
            structural_results = self._perform_structural_validation(source_data, dest_data)
            comparison_results.extend(structural_results)
        
        # 2. Data comparison
        logger.info("Performing data comparison")
        data_results = self._perform_data_comparison(source_data, dest_data)
        comparison_results.extend(data_results)
        
        # 3. Calculation validation
        if self.config.enable_calculation_validation:
            logger.info("Performing calculation validation")
            calc_validations = self.calculation_validator.validate_totals(source_data, dest_data)
            calculation_validations.extend(calc_validations)
            
            # Convert calculation errors to comparison results
            for calc_val in calc_validations:
                if calc_val.status == ValidationStatus.CALCULATION_ERROR:
                    comparison_results.append(ComparisonResult(
                        key="TOTAL_CALCULATION",
                        section="CALCULATION",
                        field=calc_val.field,
                        source_value=calc_val.expected_value,
                        dest_value=calc_val.actual_value,
                        status=ValidationStatus.CALCULATION_ERROR,
                        severity=SeverityLevel.CRITICAL,
                        difference=calc_val.difference,
                        notes=f"Calculation error: {calc_val.formula_used}"
                    ))
        
        logger.info(f"Comparison completed: {len(comparison_results)} results, {len(calculation_validations)} calculation validations")
        
        return comparison_results, calculation_validations
    
    def _perform_structural_validation(self, source_data: ParsedExcelData, dest_data: ParsedExcelData) -> List[ComparisonResult]:
        """Perform structural validation"""
        results = []
        
        # Validate sections
        section_results = self.structural_validator.validate_sections(source_data, dest_data)
        results.extend(section_results)
        
        # Validate record counts
        count_results = self.structural_validator.validate_record_counts(source_data, dest_data)
        results.extend(count_results)
        
        return results
    
    def _perform_data_comparison(self, source_data: ParsedExcelData, dest_data: ParsedExcelData) -> List[ComparisonResult]:
        """Perform detailed data comparison"""
        results = []
        
        # Get common sections
        common_sections = set(source_data.sections.keys()) & set(dest_data.sections.keys())
        
        for section_name in common_sections:
            logger.debug(f"Comparing section: {section_name}")
            section_results = self._compare_section(
                source_data.sections[section_name],
                dest_data.sections[section_name],
                section_name
            )
            results.extend(section_results)
        
        return results
    
    def _compare_section(self, source_section: List[Dict[str, Any]], dest_section: List[Dict[str, Any]], section_name: str) -> List[ComparisonResult]:
        """Compare data within a specific section"""
        results = []
        
        # Create lookup dictionaries using composite keys
        source_lookup = {record.get('composite_key', f"row_{i}"): record 
                        for i, record in enumerate(source_section)}
        dest_lookup = {record.get('composite_key', f"row_{i}"): record 
                      for i, record in enumerate(dest_section)}
        
        # Get all unique keys
        all_keys = set(source_lookup.keys()) | set(dest_lookup.keys())
        
        for key in all_keys:
            source_record = source_lookup.get(key)
            dest_record = dest_lookup.get(key)
            
            if source_record is None:
                # Record exists in dest but not in source
                results.append(ComparisonResult(
                    key=key,
                    section=section_name,
                    field="ENTIRE_RECORD",
                    source_value="MISSING",
                    dest_value="EXISTS",
                    status=ValidationStatus.MISSING_IN_SOURCE,
                    severity=SeverityLevel.HIGH,
                    notes=f"Record exists in destination but missing in source"
                ))
            elif dest_record is None:
                # Record exists in source but not in dest
                results.append(ComparisonResult(
                    key=key,
                    section=section_name,
                    field="ENTIRE_RECORD",
                    source_value="EXISTS",
                    dest_value="MISSING",
                    status=ValidationStatus.MISSING_IN_DEST,
                    severity=SeverityLevel.HIGH,
                    notes=f"Record exists in source but missing in destination"
                ))
            else:
                # Both records exist, compare field by field
                field_results = self._compare_records(source_record, dest_record, key, section_name)
                results.extend(field_results)
        
        return results
    
    def _compare_records(self, source_record: Dict[str, Any], dest_record: Dict[str, Any], key: str, section: str) -> List[ComparisonResult]:
        """Compare individual fields between two records"""
        results = []
        
        # Get all fields from both records
        all_fields = set(source_record.keys()) | set(dest_record.keys())
        
        # Skip composite_key field in comparison
        all_fields.discard('composite_key')
        
        for field in all_fields:
            source_value = source_record.get(field)
            dest_value = dest_record.get(field)
            
            # Perform comparison
            is_match, difference = self.numerical_comparator.compare_numbers(source_value, dest_value)
            
            if not is_match:
                # Determine severity
                comparison_result = ComparisonResult(
                    key=key,
                    section=section,
                    field=field,
                    source_value=source_value,
                    dest_value=dest_value,
                    status=ValidationStatus.MISMATCH,
                    severity=SeverityLevel.MEDIUM,  # Will be updated by determine_severity
                    difference=difference
                )
                
                # Update severity based on context
                comparison_result.severity = determine_severity(comparison_result, self.config)
                
                results.append(comparison_result)
        
        return results
    
    def generate_summary(self, comparison_results: List[ComparisonResult], calculation_validations: List[CalculationValidation]) -> ValidationSummary:
        """Generate validation summary from results"""
        
        # Count by status
        matches = sum(1 for r in comparison_results if r.status == ValidationStatus.MATCH)
        mismatches = sum(1 for r in comparison_results if r.status == ValidationStatus.MISMATCH)
        missing_in_source = sum(1 for r in comparison_results if r.status == ValidationStatus.MISSING_IN_SOURCE)
        missing_in_dest = sum(1 for r in comparison_results if r.status == ValidationStatus.MISSING_IN_DEST)
        
        # Count by severity
        critical_issues = sum(1 for r in comparison_results if r.severity == SeverityLevel.CRITICAL)
        high_issues = sum(1 for r in comparison_results if r.severity == SeverityLevel.HIGH)
        medium_issues = sum(1 for r in comparison_results if r.severity == SeverityLevel.MEDIUM)
        low_issues = sum(1 for r in comparison_results if r.severity == SeverityLevel.LOW)
        
        # Calculate percentages
        total_comparisons = len(comparison_results)
        match_percentage = (matches / total_comparisons * 100) if total_comparisons > 0 else 100
        
        # Check for errors
        has_calculation_errors = any(cv.status == ValidationStatus.CALCULATION_ERROR for cv in calculation_validations)
        has_structural_errors = any(r.status in [ValidationStatus.STRUCTURAL_ERROR, ValidationStatus.MISSING_IN_SOURCE, ValidationStatus.MISSING_IN_DEST] 
                                  for r in comparison_results)
        
        return ValidationSummary(
            total_records_compared=total_comparisons,
            total_fields_compared=total_comparisons,  # Approximation
            total_matches=matches,
            total_mismatches=mismatches,
            total_missing_in_source=missing_in_source,
            total_missing_in_dest=missing_in_dest,
            overall_match_percentage=round(match_percentage, 2),
            critical_issues=critical_issues,
            high_issues=high_issues,
            medium_issues=medium_issues,
            low_issues=low_issues,
            has_calculation_errors=has_calculation_errors,
            has_structural_errors=has_structural_errors
        )
    
    def generate_section_summaries(self, comparison_results: List[ComparisonResult]) -> List[SectionSummary]:
        """Generate section-wise summaries"""
        sections = {}
        
        for result in comparison_results:
            section = result.section
            if section not in sections:
                sections[section] = {
                    'total': 0,
                    'matches': 0,
                    'mismatches': 0,
                    'missing_in_source': 0,
                    'missing_in_dest': 0
                }
            
            sections[section]['total'] += 1
            
            if result.status == ValidationStatus.MATCH:
                sections[section]['matches'] += 1
            elif result.status == ValidationStatus.MISMATCH:
                sections[section]['mismatches'] += 1
            elif result.status == ValidationStatus.MISSING_IN_SOURCE:
                sections[section]['missing_in_source'] += 1
            elif result.status == ValidationStatus.MISSING_IN_DEST:
                sections[section]['missing_in_dest'] += 1
        
        summaries = []
        for section_name, counts in sections.items():
            match_percentage = (counts['matches'] / counts['total'] * 100) if counts['total'] > 0 else 100
            
            summaries.append(SectionSummary(
                section_name=section_name,
                total_records=counts['total'],
                total_fields=counts['total'],  # Approximation
                matches=counts['matches'],
                mismatches=counts['mismatches'],
                missing_in_source=counts['missing_in_source'],
                missing_in_dest=counts['missing_in_dest'],
                match_percentage=round(match_percentage, 2)
            ))
        
        return summaries