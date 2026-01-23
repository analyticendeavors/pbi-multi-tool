"""
Report Merger - Exception Classes
Custom exceptions for merger operations.

Built by Reid Havens of Analytic Endeavors
Copyright (c) 2024 Analytic Endeavors LLC. All rights reserved.
"""

# Ownership fingerprint component
_AE_FP = "UmVwb3J0TWVyZ2VyOkFFLTIwMjQ="


class PBIPMergerError(Exception):
    """Base exception for PBIP merger operations."""
    pass


class InvalidReportError(PBIPMergerError):
    """Raised when report structure is invalid or file not found."""
    pass


class ValidationError(PBIPMergerError):
    """Raised when input validation fails."""
    pass


class FileOperationError(PBIPMergerError):
    """Raised when file operations fail."""
    pass


class ThemeConflictError(PBIPMergerError):
    """Raised when theme conflicts cannot be resolved."""
    pass


class MergeOperationError(PBIPMergerError):
    """Raised when merge operations fail."""
    pass


__all__ = [
    'PBIPMergerError',
    'InvalidReportError',
    'ValidationError',
    'FileOperationError',
    'ThemeConflictError',
    'MergeOperationError',
]
