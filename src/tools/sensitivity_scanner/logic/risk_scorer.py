"""
Risk scoring and assessment for sensitivity findings.

This module provides the RiskScorer class which:
- Converts pattern matches into findings with risk assessments
- Calculates confidence scores based on context
- Adjusts risk levels based on surrounding content
- Provides recommendations for each finding
"""

import logging
from typing import List, Optional
from .models import (
    PatternMatch, PatternDefinition, Finding, 
    RiskLevel, FindingCategory
)
from .pattern_detector import PatternDetector


logger = logging.getLogger(__name__)


class RiskScorer:
    """
    Assigns risk levels and confidence scores to pattern matches.
    
    Takes pattern matches and converts them into full Finding objects
    with context-aware risk assessment and actionable recommendations.
    """
    
    def __init__(self, pattern_detector: PatternDetector):
        """
        Initialize the risk scorer.
        
        Args:
            pattern_detector: PatternDetector instance with loaded patterns
        """
        self.pattern_detector = pattern_detector
    
    def create_finding(
        self,
        pattern_match: PatternMatch,
        file_path: str,
        file_type: str,
        location_description: str,
        full_context: str = ""
    ) -> Finding:
        """
        Create a Finding from a PatternMatch with risk assessment.
        
        Args:
            pattern_match: The pattern match to convert
            file_path: Path to the file where it was found
            file_type: Type of file (TMDL, JSON, etc.)
            location_description: Human-readable location description
            full_context: Full surrounding context for risk adjustment
        
        Returns:
            Finding object with risk level and confidence score
        """
        # Get the pattern definition
        pattern_def = self.pattern_detector.get_pattern_by_id(
            pattern_match.pattern_id
        )
        
        if not pattern_def:
            logger.warning(
                f"Pattern definition not found for {pattern_match.pattern_id}"
            )
            # Create a minimal finding with LOW risk as fallback
            return self._create_fallback_finding(
                pattern_match, 
                file_path, 
                file_type, 
                location_description
            )
        
        # Base risk level from pattern definition
        base_risk_level = pattern_def.risk_level
        
        # Calculate confidence score
        confidence_score = self._calculate_confidence(
            pattern_match, 
            pattern_def, 
            full_context
        )
        
        # Check for context-based risk adjustment
        context_adjustment = self._assess_context_risk(
            pattern_match, 
            full_context
        )
        
        # Adjust risk level if needed
        final_risk_level = self._adjust_risk_level(
            base_risk_level, 
            context_adjustment,
            confidence_score
        )
        
        # Generate recommendation
        recommendation = self._generate_recommendation(
            final_risk_level, 
            pattern_def,
            location_description
        )
        
        # Create the finding
        finding = Finding(
            risk_level=final_risk_level,
            pattern_match=pattern_match,
            file_path=file_path,
            file_type=file_type,
            location_description=location_description,
            categories=pattern_def.categories,
            confidence_score=confidence_score,
            description=pattern_def.description,
            recommendation=recommendation,
            context_risk_adjustment=context_adjustment
        )
        
        return finding
    
    def _calculate_confidence(
        self,
        pattern_match: PatternMatch,
        pattern_def: PatternDefinition,
        full_context: str
    ) -> float:
        """
        Calculate confidence score (0.0-1.0) for a match.
        
        Args:
            pattern_match: The pattern match
            pattern_def: Pattern definition
            full_context: Surrounding context
        
        Returns:
            Confidence score between 0.0 and 1.0
        """
        confidence = 0.5  # Base confidence
        
        # High confidence for specific patterns
        high_confidence_patterns = [
            'ssn_us', 'credit_card', 'email_address', 
            'phone_us', 'api_key', 'connection_string_creds'
        ]
        if pattern_def.id in high_confidence_patterns:
            confidence += 0.3
        
        # Medium confidence for column name patterns
        if 'column_name' in [cat.value for cat in pattern_def.categories]:
            confidence += 0.2
        
        # Adjust based on match length for certain patterns
        match_length = len(pattern_match.matched_text)
        if pattern_def.id in ['api_key', 'password']:
            if match_length > 30:
                confidence += 0.1
            elif match_length < 10:
                confidence -= 0.1
        
        # Check context for validation
        context_keywords = self.pattern_detector.check_context_keywords(
            full_context
        )
        
        # Increase confidence if high-risk keywords present
        if context_keywords['high_risk'] > 0:
            confidence += min(0.15, context_keywords['high_risk'] * 0.05)
        
        # Decrease confidence if low-risk keywords present
        if context_keywords['low_risk'] > 0:
            confidence -= min(0.3, context_keywords['low_risk'] * 0.1)
        
        # Clamp between 0.0 and 1.0
        return max(0.0, min(1.0, confidence))
    
    def _assess_context_risk(
        self,
        pattern_match: PatternMatch,
        full_context: str
    ) -> float:
        """
        Assess risk adjustment based on surrounding context.
        
        Returns a value between -0.5 and +0.5 for risk adjustment.
        
        Args:
            pattern_match: The pattern match
            full_context: Surrounding context text
        
        Returns:
            Risk adjustment value
        """
        adjustment = 0.0
        
        combined_context = (
            f"{pattern_match.context_before} "
            f"{pattern_match.matched_text} "
            f"{pattern_match.context_after} "
            f"{full_context}"
        ).lower()
        
        context_counts = self.pattern_detector.check_context_keywords(
            combined_context
        )
        
        if context_counts['high_risk'] > 0:
            adjustment += min(0.3, context_counts['high_risk'] * 0.1)
        
        if context_counts['low_risk'] > 0:
            adjustment -= min(0.5, context_counts['low_risk'] * 0.15)
        
        return adjustment
    
    def _adjust_risk_level(
        self,
        base_risk: RiskLevel,
        context_adjustment: float,
        confidence: float
    ) -> RiskLevel:
        """Adjust risk level based on context and confidence."""
        risk_map = {
            RiskLevel.LOW: 1,
            RiskLevel.MEDIUM: 2,
            RiskLevel.HIGH: 3
        }
        
        reverse_map = {
            1: RiskLevel.LOW,
            2: RiskLevel.MEDIUM,
            3: RiskLevel.HIGH
        }
        
        numeric_risk = risk_map[base_risk]
        
        if context_adjustment < -0.3 and confidence < 0.5:
            numeric_risk = max(1, numeric_risk - 1)
        elif context_adjustment > 0.2 and confidence > 0.7:
            numeric_risk = min(3, numeric_risk + 1)
        
        return reverse_map[numeric_risk]
    
    def _generate_recommendation(
        self,
        risk_level: RiskLevel,
        pattern_def: PatternDefinition,
        location: str
    ) -> str:
        """
        Generate Power BI-specific actionable recommendations.
        
        Returns detailed guidance with:
        - Immediate action required
        - Power BI best practices
        - Implementation examples where applicable
        """
        # Check for pattern-specific recommendations first (most detailed)
        pattern_recs = self._get_pattern_specific_recommendations()
        
        if pattern_def.id in pattern_recs:
            return pattern_recs[pattern_def.id]
        
        # Fall back to category-based recommendations
        category_recs = self._get_category_recommendations(risk_level)
        
        for category in pattern_def.categories:
            if category.value in category_recs:
                return category_recs[category.value]
        
        # Final fallback to generic recommendation
        return self._get_generic_recommendation(risk_level)
    
    def _get_pattern_specific_recommendations(self) -> dict:
        """
        Get pattern-specific Power BI recommendations.
        
        Returns:
            Dictionary mapping pattern IDs to detailed recommendations
        """
        return {
            # === HIGH RISK PATTERNS ===
            
            # Credentials & Authentication
            'api_key': (
                "CRITICAL: Remove API key immediately. "
                "Power BI Best Practice: Store API keys in Power Query parameters "
                "(Home > Manage Parameters > New Parameter). PBIT files automatically "
                "clear parameter values on save, protecting credentials. For deployed "
                "reports, use gateway credentials or Azure Key Vault integration."
            ),
            'password': (
                "CRITICAL: Remove password immediately. "
                "Power BI Best Practice: Never hardcode passwords. Use Windows "
                "credentials (File > Options > Data source settings) or gateway "
                "connections with service accounts. For development, use parameters "
                "that are cleared when saving as PBIT."
            ),
            'connection_string_creds': (
                "CRITICAL: Remove connection string with embedded credentials. "
                "Power BI Best Practice: Configure data sources through "
                "'Data source settings' (File > Options) instead of hardcoding. "
                "Use organizational data sources for shared datasets, or implement "
                "parameter-based connections with credentials managed via gateway."
            ),
            
            # Personal Identifiable Information
            'ssn_us': (
                "CRITICAL: Remove Social Security Number. "
                "Power BI Best Practice: If SSN data is required, implement "
                "Row-Level Security (RLS) to restrict access by user. Consider "
                "masking (e.g., showing only last 4 digits) or removing from the "
                "semantic model entirely. Use object-level security to hide sensitive "
                "columns (set IsHidden = true in TMDL)."
            ),
            'ssn_column': (
                "CRITICAL: Column name suggests SSN data. "
                "Power BI Best Practice: If this column contains actual SSNs, "
                "implement Row-Level Security (RLS) and consider column-level security. "
                "Rename column to generic name (e.g., 'ID') if sharing model structure. "
                "Set IsHidden = true in TMDL to prevent discovery."
            ),
            'credit_card': (
                "CRITICAL: Remove credit card number. "
                "Power BI Best Practice: Credit card data should not be stored in "
                "Power BI models. If absolutely necessary, implement strict RLS, "
                "mask all but last 4 digits in source query, and hide column from "
                "report view. Consider using tokenized references instead."
            ),
            'credit_card_column': (
                "CRITICAL: Column name suggests credit card data. "
                "Power BI Best Practice: Verify if column contains actual card numbers. "
                "If yes, remove or implement strict RLS. Rename to generic name if "
                "sharing model structure. Use aggregated measures instead of exposing "
                "raw column."
            ),
            'email_address': (
                "HIGH RISK: Email address detected. "
                "Power BI Best Practice: If exposing user emails, implement RLS to "
                "ensure users only see their own data. Consider hashing emails for "
                "join keys (e.g., HASHBYTES in SQL) or using user IDs instead. "
                "For reporting, create measures that show counts rather than lists."
            ),
            'phone_us': (
                "HIGH RISK: Phone number detected. "
                "Power BI Best Practice: Implement RLS if users should only see their "
                "own contact info. Consider masking format (e.g., (555) XXX-1234). "
                "Hide column from report view and use measures for aggregated counts."
            ),
            
            # === MEDIUM RISK PATTERNS ===
            
            # Organizational Data
            'salary_column': (
                "MEDIUM RISK: Salary/compensation data detected. "
                "Power BI Best Practice: Implement Row-Level Security (RLS) based on "
                "manager hierarchy or HR roles. Use salary bands/ranges instead of "
                "exact amounts. Create aggregated measures (AVG, MEDIAN) rather than "
                "exposing individual salaries. Hide column from report view."
            ),
            'dob_column': (
                "MEDIUM RISK: Date of birth column detected. "
                "Power BI Best Practice: Replace with calculated age or age range. "
                "Implement RLS if DOB must be visible. Consider removing if not "
                "essential for analysis. Use measures for age-based calculations "
                "instead of exposing raw DOB."
            ),
            'employee_id': (
                "MEDIUM RISK: Employee ID detected. "
                "Power BI Best Practice: Verify if IDs need masking. Implement RLS "
                "so employees see only their own data. Consider using hashed IDs or "
                "generic identifiers when sharing model externally. Hide column if "
                "only used for relationships."
            ),
            'username_column': (
                "MEDIUM RISK: Username/login detected. "
                "Power BI Best Practice: If using for RLS, ensure USERNAME() or "
                "USERPRINCIPALNAME() functions match exact format. Test with "
                "'View As' role feature (Modeling > View As). Consider replacing "
                "with user IDs or email addresses for more reliable RLS."
            ),
            
            # Infrastructure & Connection Details
            'server_name': (
                "MEDIUM RISK: Server/database name detected. "
                "Power BI Best Practice: Use parameters for environment-based "
                "deployment (Dev/Test/Prod). Replace hardcoded server names with "
                "Power Query parameters that can be updated via gateway settings. "
                "Document server names separately rather than in shared models."
            ),
            'ip_address': (
                "MEDIUM RISK: IP address detected. "
                "Power BI Best Practice: Replace with DNS names or use parameters for "
                "environment flexibility. Avoid exposing internal IP ranges when "
                "sharing models externally. Use organizational data sources instead "
                "of direct IP connections."
            ),
            'file_path_windows': (
                "MEDIUM RISK: Windows file path detected. "
                "Power BI Best Practice: Use relative paths or UNC paths instead of "
                "absolute paths. For data sources, configure via gateway or use "
                "SharePoint/OneDrive URLs. Replace hardcoded paths with parameters "
                "for flexibility across environments."
            ),
            'unc_path': (
                "MEDIUM RISK: UNC network path detected. "
                "Power BI Best Practice: Document network paths externally rather than "
                "in model. If path must be shared, ensure it doesn't reveal sensitive "
                "organizational structure. Use parameters for dynamic path resolution."
            ),
            
            # Security Expressions
            'rls_expression': (
                "MEDIUM RISK: Row-Level Security expression detected. "
                "Power BI Best Practice: Review RLS logic for correctness. Test with "
                "'View As' role feature (Modeling > View As) before publishing. "
                "Ensure USERNAME() or USERPRINCIPALNAME() matches expected format. "
                "Document RLS logic externally and verify it doesn't leak sensitive "
                "business rules when sharing model structure."
            ),
            
            # === LOW RISK PATTERNS ===
            
            # Generic Personal Info
            'name_column': (
                "LOW RISK: Name column detected. "
                "Power BI Best Practice: If exposing employee/customer names, verify "
                "RLS is configured appropriately. Consider using initials or hashed "
                "identifiers if full names aren't necessary. For external sharing, "
                "replace with generic identifiers."
            ),
            'address_column': (
                "LOW RISK: Address field detected. "
                "Power BI Best Practice: Consider aggregating to city/state/region "
                "level if street addresses aren't required for analysis. Implement "
                "RLS if addresses should be restricted by user. Verify data sharing "
                "agreements allow address exposure."
            ),
            
            # Organizational Structure
            'department': (
                "LOW RISK: Department/organizational field detected. "
                "Power BI Best Practice: Verify if org structure should be shared "
                "externally. Consider using generic categories (e.g., 'Operations', "
                "'Sales') instead of specific department names if sharing with "
                "external parties."
            ),
            'company_specific': (
                "LOW RISK: Company/client reference detected. "
                "Power BI Best Practice: If sharing model externally, replace client "
                "names with generic identifiers (Client A, Client B). Verify "
                "confidentiality agreements allow client name exposure. Consider "
                "using industry categories instead of specific names."
            )
        }
    
    def _get_category_recommendations(self, risk_level: RiskLevel) -> dict:
        """
        Get category-based Power BI recommendations.
        
        Args:
            risk_level: The risk level for this finding
        
        Returns:
            Dictionary mapping categories to recommendations
        """
        if risk_level == RiskLevel.HIGH:
            return {
                'pii': (
                    "Remove personally identifiable information (PII) or implement "
                    "Row-Level Security (RLS) to restrict access by user."
                ),
                'credentials': (
                    "Remove credentials immediately. Use Power Query parameters, "
                    "gateway credentials, or organizational data sources instead."
                ),
                'financial': (
                    "Remove or protect financial data with RLS. Consider aggregated "
                    "measures instead of exposing raw values."
                ),
                'government_id': (
                    "Remove government-issued IDs (SSN, passport, etc.). If required, "
                    "implement strict RLS and column-level security."
                ),
                'connection_string': (
                    "Remove connection strings. Configure data sources via "
                    "'Data source settings' (File > Options) or gateway."
                ),
                'security': (
                    "Review security expressions before sharing. Test RLS with "
                    "'View As' feature to ensure proper access control."
                ),
                'contact': (
                    "Review email/phone exposure. Implement RLS if users should only "
                    "see their own contact information."
                )
            }
        elif risk_level == RiskLevel.MEDIUM:
            return {
                'infrastructure': (
                    "Review server/IP/path references. Use parameters for environment "
                    "flexibility. Avoid exposing internal infrastructure details."
                ),
                'organizational': (
                    "Review organizational data exposure. Consider if department/team "
                    "structure should be visible when sharing externally."
                ),
                'column_name': (
                    "Column name may reveal sensitive logic. Consider renaming to "
                    "generic name if sharing model structure."
                ),
                'file_path': (
                    "Replace hardcoded paths with parameters or relative references. "
                    "Avoid exposing organizational file structure."
                ),
                'dax': (
                    "Review DAX expression for sensitive logic or hardcoded values. "
                    "Consider if calculation reveals confidential business rules."
                )
            }
        else:  # LOW risk
            return {
                'business': (
                    "Review if business references are appropriate for sharing context. "
                    "Consider using generic identifiers for external audiences."
                )
            }
    
    def _get_generic_recommendation(self, risk_level: RiskLevel) -> str:
        """
        Get generic recommendation based on risk level.
        
        Args:
            risk_level: The risk level for this finding
        
        Returns:
            Generic recommendation string
        """
        if risk_level == RiskLevel.HIGH:
            return (
                "CRITICAL: Review and remove this item before sharing. Consider "
                "Power BI security features (RLS, object-level security) if data "
                "must be included."
            )
        elif risk_level == RiskLevel.MEDIUM:
            return (
                "Review this item to determine if it should be removed or protected. "
                "Consider Power BI parameters and organizational data sources for "
                "better security."
            )
        else:  # LOW
            return (
                "Verify this item is appropriate for your sharing context. Consider "
                "using generic identifiers if sharing externally."
            )
    
    def _create_fallback_finding(
        self,
        pattern_match: PatternMatch,
        file_path: str,
        file_type: str,
        location_description: str
    ) -> Finding:
        """Create a minimal fallback finding."""
        return Finding(
            risk_level=RiskLevel.LOW,
            pattern_match=pattern_match,
            file_path=file_path,
            file_type=file_type,
            location_description=location_description,
            categories=[FindingCategory.BUSINESS],
            confidence_score=0.3,
            description="Unknown pattern match",
            recommendation="Review this item.",
            context_risk_adjustment=0.0
        )
    
    def score_matches(
        self,
        matches: List[PatternMatch],
        file_path: str,
        file_type: str,
        location_prefix: str = "",
        full_context: str = ""
    ) -> List[Finding]:
        """Convert multiple pattern matches into scored findings with deduplication."""
        findings = []
        
        for match in matches:
            location_desc = f"{location_prefix} → Line {match.line_number}"
            if not location_prefix:
                location_desc = f"Line {match.line_number}"
            
            finding = self.create_finding(
                pattern_match=match,
                file_path=file_path,
                file_type=file_type,
                location_description=location_desc,
                full_context=full_context
            )
            
            findings.append(finding)
        
        # Deduplicate findings from the same line
        # Keep highest risk finding per line, or combine if same pattern type
        deduplicated = self._deduplicate_findings(findings)
        
        # Sort by risk level and line number
        deduplicated.sort(
            key=lambda f: (
                -{'HIGH': 3, 'MEDIUM': 2, 'LOW': 1}[f.risk_level.value],
                f.pattern_match.line_number or 0
            )
        )
        
        return deduplicated
    
    def _deduplicate_findings(self, findings: List[Finding]) -> List[Finding]:
        """
        Deduplicate findings from the same location.
        
        Strategy:
        - Group by file_path + line_number
        - For connection strings: Keep one finding, mention multiple credentials in description
        - For other patterns: Keep highest risk finding per line
        
        Args:
            findings: List of findings to deduplicate
        
        Returns:
            Deduplicated list of findings
        """
        from collections import defaultdict
        
        # Group findings by location (file + line)
        grouped = defaultdict(list)
        for finding in findings:
            key = f"{finding.file_path}:{finding.pattern_match.line_number}"
            grouped[key].append(finding)
        
        deduplicated = []
        
        for location_key, location_findings in grouped.items():
            if len(location_findings) == 1:
                # Single finding at this location - keep it
                deduplicated.append(location_findings[0])
            else:
                # Multiple findings at same location - deduplicate
                
                # Check if these are connection string credential patterns
                cred_patterns = {'connection_string_creds', 'password', 'api_key'}
                pattern_ids = {f.pattern_match.pattern_id for f in location_findings}
                
                if pattern_ids.intersection(cred_patterns):
                    # Connection string with multiple credentials - combine
                    combined = self._combine_credential_findings(location_findings)
                    deduplicated.append(combined)
                else:
                    # Different patterns - keep highest risk
                    highest_risk = max(
                        location_findings,
                        key=lambda f: {'HIGH': 3, 'MEDIUM': 2, 'LOW': 1}[f.risk_level.value]
                    )
                    deduplicated.append(highest_risk)
        
        return deduplicated
    
    def _combine_credential_findings(self, findings: List[Finding]) -> Finding:
        """
        Combine multiple credential findings from same line into one.
        
        Args:
            findings: List of credential-related findings
        
        Returns:
            Combined finding
        """
        # Take the first finding as base
        combined = findings[0]
        
        # Count credential types found
        cred_types = []
        for finding in findings:
            pattern_name = finding.pattern_match.pattern_name
            if pattern_name not in cred_types:
                cred_types.append(pattern_name)
        
        # Update description to mention multiple credentials
        if len(cred_types) > 1:
            combined.description = (
                f"Connection string with multiple embedded credentials: "
                f"{', '.join(cred_types)}"
            )
        
        return combined
