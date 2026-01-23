"""
Report Merger - Validation Service
Centralized validation for all merger inputs.

Built by Reid Havens of Analytic Endeavors
Copyright (c) 2024 Analytic Endeavors LLC. All rights reserved.
"""

import json
import uuid
from pathlib import Path
from typing import Dict, Any

from core.constants import AppConstants
from tools.report_merger.merger_exceptions import ValidationError

# Ownership fingerprint component
_AE_FP = "VmFsaWRhdGlvblNlcnZpY2U6QUUtMjAyNA=="


class ValidationService:
    """Centralized validation service for all input validation needs."""

    @staticmethod
    def validate_input_paths(path_a: str, path_b: str) -> None:
        """Validate input PBIP file paths with comprehensive error reporting."""
        errors = []

        # Validate Report A
        try:
            ValidationService._validate_single_pbip_path(path_a, "Report A")
        except ValidationError as e:
            errors.append(str(e))

        # Validate Report B
        try:
            ValidationService._validate_single_pbip_path(path_b, "Report B")
        except ValidationError as e:
            errors.append(str(e))

        # Check for same file
        if path_a and path_b and Path(path_a).resolve() == Path(path_b).resolve():
            errors.append("Report A and Report B cannot be the same file")

        if errors:
            raise ValidationError("Input validation failed:\n• " + "\n• ".join(errors))

    @staticmethod
    def _validate_single_pbip_path(file_path: str, report_name: str) -> None:
        """Validate a single PBIP file path."""
        if not file_path:
            raise ValidationError(f"{report_name} path is required")

        path_obj = Path(file_path)

        # Check file extension
        if not file_path.lower().endswith('.pbip'):
            raise ValidationError(f"{report_name} must be a .pbip file (got: {path_obj.suffix})")

        # Check file exists
        if not path_obj.exists():
            raise ValidationError(f"{report_name} file not found: {file_path}")

        # Check file is actually a file (not directory)
        if not path_obj.is_file():
            raise ValidationError(f"{report_name} path must point to a file, not a directory")

        # Check read permissions
        try:
            with path_obj.open('r') as f:
                pass  # Just test if we can open it
        except PermissionError:
            raise ValidationError(f"{report_name} file cannot be read (permission denied): {file_path}")
        except Exception as e:
            raise ValidationError(f"{report_name} file cannot be accessed: {str(e)}")

        # Check corresponding .Report directory exists
        report_dir = path_obj.parent / f"{path_obj.stem}.Report"
        if not report_dir.exists():
            raise ValidationError(f"{report_name} missing corresponding .Report directory: {report_dir}")

        if not report_dir.is_dir():
            raise ValidationError(f"{report_name} .Report path exists but is not a directory: {report_dir}")

    @staticmethod
    def validate_output_path(output_path: str) -> None:
        """Validate output path for write access and structure."""
        if not output_path:
            raise ValidationError("Output path is required")

        path_obj = Path(output_path)

        # Check file extension
        if not output_path.lower().endswith('.pbip'):
            raise ValidationError(f"Output file must be a .pbip file (got: {path_obj.suffix})")

        # Check parent directory exists and is writable
        parent_dir = path_obj.parent
        if not parent_dir.exists():
            raise ValidationError(f"Output directory does not exist: {parent_dir}")

        # Check write permissions on parent directory
        try:
            test_file = parent_dir / f"write_test_{uuid.uuid4().hex[:8]}.tmp"
            test_file.touch()
            test_file.unlink()
        except PermissionError:
            raise ValidationError(f"Cannot write to output directory (permission denied): {parent_dir}")
        except Exception as e:
            raise ValidationError(f"Cannot write to output directory: {str(e)}")

    @staticmethod
    def validate_thin_report_structure(report_dir: Path, report_name: str) -> None:
        """Comprehensive validation of thin report structure."""
        errors = []

        # Check .Report directory structure
        if not report_dir.exists():
            errors.append(f"{report_name} .Report directory not found: {report_dir}")
            raise ValidationError("\n• ".join(errors))

        # Check for semantic model (should not exist for thin reports)
        semantic_model_dir = report_dir.parent / f"{report_name}.SemanticModel"
        if semantic_model_dir.exists():
            errors.append(f"{report_name} appears to have a semantic model (not a thin report): {semantic_model_dir}")

        # Check required directories and files
        required_paths = [
            (report_dir / "definition", "definition directory"),
            (report_dir / "definition" / "report.json", "report.json file"),
            (report_dir / ".pbi", ".pbi directory"),
            (report_dir / ".platform", ".platform file")
        ]

        for path, description in required_paths:
            if not path.exists():
                errors.append(f"{report_name} missing {description}: {path}")

        # Validate key JSON files
        if (report_dir / "definition" / "report.json").exists():
            try:
                ValidationService.validate_json_structure(report_dir / "definition" / "report.json")
            except ValidationError as e:
                errors.append(f"{report_name} has invalid report.json: {str(e)}")

        if (report_dir / ".platform").exists():
            try:
                ValidationService.validate_json_structure(report_dir / ".platform", "platform")
            except ValidationError as e:
                errors.append(f"{report_name} has invalid .platform file: {str(e)}")

        if errors:
            raise ValidationError("Report structure validation failed:\n• " + "\n• ".join(errors))

    @staticmethod
    def validate_json_structure(file_path: Path, expected_schema_key: str = None) -> Dict[str, Any]:
        """Validate JSON file structure with optional schema checking."""
        if not file_path.exists():
            raise ValidationError(f"JSON file not found: {file_path}")

        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
        except json.JSONDecodeError as e:
            raise ValidationError(f"Invalid JSON in {file_path.name}: {str(e)}")
        except UnicodeDecodeError as e:
            raise ValidationError(f"Cannot read {file_path.name} (encoding issue): {str(e)}")
        except Exception as e:
            raise ValidationError(f"Cannot read {file_path.name}: {str(e)}")

        # Basic structure validation
        if not isinstance(data, dict):
            raise ValidationError(f"{file_path.name} must contain a JSON object, not {type(data).__name__}")

        # Optional schema validation
        if expected_schema_key and expected_schema_key in AppConstants.SCHEMA_URLS:
            expected_schema = AppConstants.SCHEMA_URLS[expected_schema_key]
            actual_schema = data.get('$schema', '')

            if actual_schema and expected_schema not in actual_schema:
                # Log warning but don't fail - schemas can be flexible
                pass

        return data


__all__ = ['ValidationService']
