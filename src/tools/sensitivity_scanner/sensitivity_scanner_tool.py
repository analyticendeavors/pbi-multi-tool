"""
Sensitivity Scanner Tool - Main tool class implementation.

This module provides the SensitivityScannerTool class which:
- Registers the tool with the AE Multi-Tool framework
- Creates the UI tab
- Provides tool metadata and help content
- Checks dependencies
"""

from typing import Dict, Any
from core.tool_manager import BaseTool
from tools.sensitivity_scanner.ui.sensitivity_scanner_tab import SensitivityScannerTab

# Ownership fingerprint
_AE_FP = "U2Vuc2l0aXZpdHlTY2FubmVyOkFFLTIwMjQ="


class SensitivityScannerTool(BaseTool):
    """
    Sensitivity Scanner - Detect sensitive data in Power BI semantic models.
    
    Scans TMDL files for:
    - PII (SSN, email, phone, etc.)
    - Financial data (credit cards, salaries)
    - Credentials (passwords, API keys, connection strings)
    - Infrastructure details (server names, IPs, file paths)
    - Security expressions (RLS)
    """
    
    def __init__(self):
        """Initialize the Sensitivity Scanner tool."""
        super().__init__(
            tool_id="sensitivity_scanner",
            name="Sensitivity Scanner",
            description="Scan Power BI semantic models for sensitive data and security risks",
            version="1.2.0"
        )
    
    def create_ui_tab(self, parent, main_app) -> 'BaseToolTab':
        """
        Create the UI tab for this tool.
        
        Args:
            parent: Parent widget (notebook)
            main_app: Main application instance
        
        Returns:
            SensitivityScannerTab instance
        """
        return SensitivityScannerTab(parent, main_app)
    
    def get_tab_title(self) -> str:
        """
        Get the display title for the tab.
        
        Returns:
            Tab title with emoji
        """
        return "🔍 Sensitivity Scanner"
    
    def get_help_content(self) -> Dict[str, Any]:
        """
        Get help content for this tool.
        
        Returns:
            Dictionary with help sections
        """
        return {
            "title": "Sensitivity Scanner - Help",
            "sections": [
                {
                    "title": "IMPORTANT DISCLAIMERS & REQUIREMENTS",
                    "items": [
                        "• This is an INFORMATIONAL TOOL ONLY for awareness and education",
                        "• NOT certified for security audits or compliance validation",
                        "• NOT a replacement for professional security assessments",
                        "• NOT a substitute for certified compliance tools (GDPR, HIPAA, PCI-DSS)",
                        "• Results should be reviewed by qualified security professionals",
                        "• Always use certified tools and professionals for compliance requirements",
                        "• This tool performs STATIC ANALYSIS only - scans model structure, not data values",
                        "• Always backup your reports before making any changes"
                    ]
                },
                {
                    "title": "What This Tool Does",
                    "items": [
                        "Scans Power BI semantic model TMDL files for sensitive data",
                        "Detects PII (SSN, email, phone), financial data, and credentials",
                        "Identifies infrastructure details and security expressions",
                        "Provides risk-level assessments (HIGH/MEDIUM/LOW)",
                        "Offers actionable Power BI-specific recommendations"
                    ]
                },
                {
                    "title": "Quick Start",
                    "items": [
                        "Step 1: Select a PBIP file (requires companion .SemanticModel folder)",
                        "Step 2: Click 'SCAN FOR SENSITIVE DATA' to start scanning",
                        "Step 3: Review findings organized by risk level",
                        "Step 4: Export detailed report with recommendations"
                    ]
                },
                {
                    "title": "What Gets Scanned",
                    "items": [
                        "HIGH RISK: SSN, credit cards, passwords, API keys, connection strings",
                        "MEDIUM RISK: Salaries, employee IDs, server names, IPs, file paths",
                        "LOW RISK: Names, addresses, departments, company references",
                        "Scans all TMDL files: tables, measures, columns, relationships, RLS"
                    ]
                },
                {
                    "title": "Risk Levels",
                    "items": [
                        "HIGH: Critical security/privacy risks - remove or protect immediately",
                        "MEDIUM: Potential exposure - review and protect with RLS/security",
                        "LOW: Minor concerns - verify appropriateness for sharing context"
                    ]
                }
            ]
        }
    
    def can_run(self) -> bool:
        """
        Check if the tool can run (dependencies available).
        
        Returns:
            True if all dependencies are available
        """
        try:
            # Check core dependencies
            from core.pbi_file_reader import PBIPReader
            from tools.sensitivity_scanner.logic.pattern_detector import PatternDetector
            from tools.sensitivity_scanner.logic.risk_scorer import RiskScorer
            from tools.sensitivity_scanner.logic.tmdl_scanner import TmdlScanner
            from tools.sensitivity_scanner.logic.models import (
                PatternMatch, Finding, ScanResult, RiskLevel
            )
            
            # Check if patterns file exists
            from pathlib import Path
            patterns_file = Path(__file__).parent.parent.parent / "data" / "sensitivity_patterns.json"
            if not patterns_file.exists():
                self.logger.error(f"Patterns file not found: {patterns_file}")
                return False
            
            return True
            
        except ImportError as e:
            self.logger.error(f"Missing dependency: {e}")
            return False
        except Exception as e:
            self.logger.error(f"Error checking dependencies: {e}")
            return False
