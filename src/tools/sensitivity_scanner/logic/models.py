"""
Data models for sensitivity scanner findings.

This module defines the data structures used to represent:
- Individual pattern matches
- Risk assessments
- Scan results and reports
"""

from dataclasses import dataclass, field
from enum import Enum
from typing import List, Optional, Dict, Any
from datetime import datetime


class RiskLevel(Enum):
    """Risk levels for sensitivity findings."""
    HIGH = "HIGH"
    MEDIUM = "MEDIUM"
    LOW = "LOW"


class FindingCategory(Enum):
    """Categories for types of sensitive data."""
    PII = "pii"                          # Personally Identifiable Information
    FINANCIAL = "financial"              # Financial data
    CREDENTIALS = "credentials"          # Passwords, API keys, tokens
    GOVERNMENT_ID = "government_id"      # SSN, passport, etc.
    CONTACT = "contact"                  # Email, phone
    INFRASTRUCTURE = "infrastructure"    # Server names, IPs, paths
    SECURITY = "security"                # Security expressions, RLS
    CONNECTION_STRING = "connection_string"  # Database connections
    DAX = "dax"                          # DAX expressions
    COLUMN_NAME = "column_name"          # Column naming patterns
    ORGANIZATIONAL = "organizational"    # Dept, division, etc.
    BUSINESS = "business"                # Business entity references
    FILE_PATH = "file_path"              # File system paths
    TRAVEL = "travel"                    # Travel documents (passport, visa)
    MEDICAL = "medical"                  # Medical/health data
    HIPAA = "hipaa"                      # HIPAA-regulated data
    API = "api"                          # API tokens and credentials
    ENCRYPTION = "encryption"            # Encryption keys
    AZURE = "azure"                      # Azure cloud credentials
    CLOUD = "cloud"                      # Cloud service credentials
    AWS = "aws"                          # AWS cloud credentials
    NETWORK = "network"                  # Network identifiers (MAC, etc.)
    HR = "hr"                            # HR-related data
    BIOMETRIC = "biometric"              # Biometric data
    DEMOGRAPHIC = "demographic"          # Demographic information
    EEO = "eeo"                          # EEO/protected class data


@dataclass
class PatternMatch:
    """
    Represents a single pattern match in the scanned content.
    
    Attributes:
        pattern_id: Unique identifier for the pattern that matched
        pattern_name: Human-readable name of the pattern
        matched_text: The actual text that was matched
        start_pos: Starting position in the content
        end_pos: Ending position in the content
        line_number: Line number where match was found (if applicable)
        context_before: Text before the match (for context)
        context_after: Text after the match (for context)
    """
    pattern_id: str
    pattern_name: str
    matched_text: str
    start_pos: int
    end_pos: int
    line_number: Optional[int] = None
    context_before: str = ""
    context_after: str = ""


@dataclass
class Finding:
    """
    Represents a sensitivity finding with risk assessment.
    
    Attributes:
        risk_level: HIGH, MEDIUM, or LOW risk
        pattern_match: The pattern match that triggered this finding
        file_path: Path to the file where finding was discovered
        file_type: Type of file (e.g., "TMDL", "JSON", "Report")
        location_description: Human-readable description of where it was found
        categories: List of applicable categories (PII, financial, etc.)
        confidence_score: 0.0-1.0 score indicating confidence in the finding
        context_risk_adjustment: Risk adjustment based on surrounding context
        description: Description of why this is sensitive
        recommendation: Suggested action to take
    """
    risk_level: RiskLevel
    pattern_match: PatternMatch
    file_path: str
    file_type: str
    location_description: str
    categories: List[FindingCategory]
    confidence_score: float
    description: str
    recommendation: str = ""
    context_risk_adjustment: float = 0.0  # +/- adjustment based on context
    metadata: Dict[str, Any] = field(default_factory=dict)


@dataclass
class ScanResult:
    """
    Complete results from scanning a file or collection of files.
    
    Attributes:
        scan_id: Unique identifier for this scan
        scan_time: When the scan was performed
        source_path: Path to the scanned PBIP or file
        findings: List of all findings, sorted by risk level
        high_risk_count: Number of HIGH risk findings
        medium_risk_count: Number of MEDIUM risk findings
        low_risk_count: Number of LOW risk findings
        total_files_scanned: Number of files processed
        scan_duration_seconds: How long the scan took
        patterns_used: Which pattern set version was used
    """
    scan_id: str
    scan_time: datetime
    source_path: str
    findings: List[Finding]
    high_risk_count: int = 0
    medium_risk_count: int = 0
    low_risk_count: int = 0
    total_files_scanned: int = 0
    scan_duration_seconds: float = 0.0
    patterns_used: str = "1.0.0"
    metadata: Dict[str, Any] = field(default_factory=dict)
    
    def __post_init__(self):
        """Calculate risk counts from findings."""
        self.high_risk_count = sum(1 for f in self.findings if f.risk_level == RiskLevel.HIGH)
        self.medium_risk_count = sum(1 for f in self.findings if f.risk_level == RiskLevel.MEDIUM)
        self.low_risk_count = sum(1 for f in self.findings if f.risk_level == RiskLevel.LOW)
    
    @property
    def total_findings(self) -> int:
        """Total number of findings across all risk levels."""
        return len(self.findings)
    
    @property
    def has_high_risk(self) -> bool:
        """Whether any HIGH risk findings were detected."""
        return self.high_risk_count > 0
    
    @property
    def has_medium_risk(self) -> bool:
        """Whether any MEDIUM risk findings were detected."""
        return self.medium_risk_count > 0
    
    def get_findings_by_risk(self, risk_level: RiskLevel) -> List[Finding]:
        """Get all findings for a specific risk level."""
        return [f for f in self.findings if f.risk_level == risk_level]
    
    def get_findings_by_category(self, category: FindingCategory) -> List[Finding]:
        """Get all findings for a specific category."""
        return [f for f in self.findings if category in f.categories]


@dataclass
class PatternDefinition:
    """
    Definition of a sensitivity pattern loaded from configuration.
    
    Attributes:
        id: Unique pattern identifier
        name: Human-readable pattern name
        pattern: Regex pattern string
        description: What this pattern detects
        risk_level: Base risk level for this pattern
        categories: List of applicable categories
        examples: Example strings that match this pattern
    """
    id: str
    name: str
    pattern: str
    description: str
    risk_level: RiskLevel
    categories: List[FindingCategory]
    examples: List[str] = field(default_factory=list)
