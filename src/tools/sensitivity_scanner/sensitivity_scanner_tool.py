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
            version="1.0.0"
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
                    "title": "⚠️ IMPORTANT DISCLAIMERS & REQUIREMENTS",
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
                    "title": "🎯 What This Tool Does",
                    "items": [
                        "Scans Power BI semantic model TMDL files for sensitive data",
                        "Detects PII (SSN, email, phone), financial data, and credentials",
                        "Identifies infrastructure details and security expressions",
                        "Provides risk-level assessments (HIGH/MEDIUM/LOW)",
                        "Offers actionable Power BI-specific recommendations"
                    ]
                },
                {
                    "title": "🚀 Quick Start",
                    "items": [
                        "Step 1: Select a PBIP file containing your semantic model",
                        "Step 2: Click 'SCAN FOR SENSITIVE DATA' to start scanning",
                        "Step 3: Review findings organized by risk level",
                        "Step 4: Export detailed report with recommendations",
                        "Step 5: Apply recommended fixes before sharing your model"
                    ]
                },
                {
                    "title": "🔍 What Gets Scanned",
                    "items": [
                        "HIGH RISK: SSN, credit cards, passwords, API keys, connection strings",
                        "MEDIUM RISK: Salaries, employee IDs, server names, IPs, file paths, RLS",
                        "LOW RISK: Names, addresses, departments, company references",
                        "All TMDL files: tables, measures, columns, relationships, RLS roles"
                    ]
                },
                {
                    "title": "📊 Risk Levels Explained",
                    "items": [
                        "HIGH (🔴): Critical security/privacy risks - remove immediately",
                        "MEDIUM (🟡): Potential exposure - review and protect with RLS/security",
                        "LOW (🟢): Minor concerns - verify appropriateness for sharing context",
                        "Confidence Score: 0.0-1.0 indicating detection accuracy"
                    ]
                },
                {
                    "title": "💡 Best Practices",
                    "items": [
                        "✅ Always scan models before sharing externally",
                        "✅ Use Power Query parameters for credentials (cleared in PBIT)",
                        "✅ Implement Row-Level Security (RLS) for sensitive data",
                        "✅ Hide sensitive columns (IsHidden = true in TMDL)",
                        "✅ Use aggregated measures instead of exposing raw data",
                        "✅ Test RLS with 'View As' feature before publishing",
                        "✅ Replace hardcoded values with parameters",
                        "✅ Document security decisions separately"
                    ]
                },
                {
                    "title": "📁 File Requirements",
                    "items": [
                        "Input: Power BI PBIP file (enhanced PBIP format)",
                        "Requires: Companion .Report folder with TMDL files",
                        "Scans: definition.tmdl, model.tmd, table TMLs, role TMLs",
                        "Output: Detailed findings report with recommendations"
                    ]
                },
                {
                    "title": "🎯 Use Cases",
                    "items": [
                        "Pre-publication security awareness check",
                        "Compliance awareness and education (NOT certification)",
                        "Template sanitization before distribution",
                        "Client data model review for common patterns",
                        "Internal security awareness initiatives",
                        "Training and education purposes"
                    ]
                },
                {
                    "title": "⚠️ Important Notes",
                    "items": [
                        "This tool performs STATIC ANALYSIS only (doesn't query data)",
                        "Scans model structure, not actual data values",
                        "Pattern-based detection may have false positives/negatives",
                        "Review all findings - automated detection isn't perfect",
                        "NOT a replacement for proper security practices",
                        "NOT officially supported by Microsoft",
                        "NOT certified for compliance validation",
                        "Always backup files before making changes",
                        "Always consult qualified security professionals for compliance"
                    ]
                },
                {
                    "title": "🔒 Privacy & Security",
                    "items": [
                        "All scanning happens locally on your machine",
                        "No data is sent to external servers",
                        "No file modifications are made",
                        "Results are saved only if you export them",
                        "Your data remains completely private"
                    ]
                },
                {
                    "title": "📋 Example Findings",
                    "items": [
                        "HIGH: 'SSN' column detected in Employees table → Implement RLS",
                        "HIGH: API key in connection string → Use parameters",
                        "MEDIUM: 'salary' column in Payroll table → Hide or aggregate",
                        "MEDIUM: Server name 'sqldb-prod-01' → Use parameters",
                        "LOW: 'customer_name' column → Verify sharing context"
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
            from core.pbip_reader import PBIPReader
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
