"""
Core pattern detection engine for sensitivity scanning.

This module provides the PatternDetector class which:
- Loads pattern definitions from JSON configuration
- Compiles regex patterns for efficient matching
- Scans text content for sensitive patterns
- Extracts context around matches
- Filters whitelist patterns (false positives)
"""

import re
import json
import logging
from pathlib import Path
from typing import List, Dict, Optional, Tuple
from .models import (
    PatternDefinition, PatternMatch, RiskLevel, 
    FindingCategory
)


logger = logging.getLogger(__name__)


class PatternDetector:
    """
    Core engine for detecting sensitive patterns in text content.
    
    This class loads pattern definitions and performs regex-based
    detection with context extraction and whitelist filtering.
    """
    
    def __init__(self, patterns_file: Optional[Path] = None):
        """
        Initialize the pattern detector.
        
        Args:
            patterns_file: Path to sensitivity_patterns.json
                          If None, uses default location in src/data/
        """
        self.patterns_file = patterns_file
        self.patterns: Dict[RiskLevel, List[PatternDefinition]] = {
            RiskLevel.HIGH: [],
            RiskLevel.MEDIUM: [],
            RiskLevel.LOW: []
        }
        self.compiled_patterns: Dict[str, re.Pattern] = {}
        self.whitelist_patterns: List[re.Pattern] = []
        self.context_keywords: Dict[str, List[str]] = {
            'high_risk': [],
            'low_risk': []
        }
        
        # Load patterns
        self._load_patterns()
    
    def _get_default_patterns_file(self) -> Path:
        """Get the default patterns file location."""
        # Assuming this file is in src/tools/sensitivity_scanner/logic/
        # Navigate up to src/data/sensitivity_patterns.json
        current_file = Path(__file__).resolve()
        src_dir = current_file.parent.parent.parent.parent
        
        # Check for custom patterns first
        custom_file = src_dir / "data" / "sensitivity_patterns_custom.json"
        default_file = src_dir / "data" / "sensitivity_patterns.json"
        
        if custom_file.exists():
            logger.info(f"Using custom patterns file: {custom_file}")
            return custom_file
        else:
            logger.info(f"Using default patterns file: {default_file}")
            return default_file
    
    def _load_patterns(self) -> None:
        """Load pattern definitions from JSON configuration."""
        if self.patterns_file is None:
            self.patterns_file = self._get_default_patterns_file()
        
        if not self.patterns_file.exists():
            logger.error(f"Patterns file not found: {self.patterns_file}")
            raise FileNotFoundError(f"Cannot find patterns file: {self.patterns_file}")
        
        try:
            with open(self.patterns_file, 'r', encoding='utf-8') as f:
                config = json.load(f)
            
            logger.info(f"Loading patterns version: {config.get('version', 'unknown')}")
            
            # Load patterns for each risk level
            for risk_str, patterns_list in config['patterns'].items():
                risk_level = self._parse_risk_level(risk_str)
                
                for pattern_dict in patterns_list:
                    pattern_def = self._create_pattern_definition(
                        pattern_dict, 
                        risk_level
                    )
                    self.patterns[risk_level].append(pattern_def)
                    
                    # Compile regex pattern
                    try:
                        compiled = re.compile(pattern_dict['pattern'])
                        self.compiled_patterns[pattern_def.id] = compiled
                        logger.debug(f"Compiled pattern: {pattern_def.id}")
                    except re.error as e:
                        logger.error(f"Failed to compile pattern {pattern_def.id}: {e}")
            
            # Load whitelist patterns
            for whitelist_dict in config['whitelist']['patterns']:
                try:
                    compiled = re.compile(whitelist_dict['pattern'])
                    self.whitelist_patterns.append(compiled)
                except re.error as e:
                    logger.error(f"Failed to compile whitelist pattern: {e}")
            
            # Load context keywords
            self.context_keywords['high_risk'] = config['context_keywords']['high_risk_context']
            self.context_keywords['low_risk'] = config['context_keywords']['low_risk_context']
            
            logger.info(f"Loaded {len(self.compiled_patterns)} patterns")
            logger.info(f"Loaded {len(self.whitelist_patterns)} whitelist patterns")
            
        except Exception as e:
            logger.error(f"Error loading patterns: {e}")
            raise
    
    def _parse_risk_level(self, risk_str: str) -> RiskLevel:
        """Convert risk level string to RiskLevel enum."""
        mapping = {
            'high_risk': RiskLevel.HIGH,
            'medium_risk': RiskLevel.MEDIUM,
            'low_risk': RiskLevel.LOW
        }
        return mapping.get(risk_str, RiskLevel.LOW)
    
    def _create_pattern_definition(
        self, 
        pattern_dict: Dict, 
        risk_level: RiskLevel
    ) -> PatternDefinition:
        """Create a PatternDefinition from JSON dictionary."""
        categories = [
            FindingCategory(cat) 
            for cat in pattern_dict.get('categories', [])
        ]
        
        return PatternDefinition(
            id=pattern_dict['id'],
            name=pattern_dict['name'],
            pattern=pattern_dict['pattern'],
            description=pattern_dict['description'],
            risk_level=risk_level,
            categories=categories,
            examples=pattern_dict.get('examples', [])
        )
    
    def scan_text(
        self, 
        text: str, 
        risk_levels: Optional[List[RiskLevel]] = None
    ) -> List[PatternMatch]:
        """
        Scan text content for sensitive patterns.
        
        Args:
            text: The text content to scan
            risk_levels: Which risk levels to check (default: all)
        
        Returns:
            List of PatternMatch objects for all matches found
        """
        if risk_levels is None:
            risk_levels = [RiskLevel.HIGH, RiskLevel.MEDIUM, RiskLevel.LOW]
        
        matches = []
        
        for risk_level in risk_levels:
            for pattern_def in self.patterns[risk_level]:
                pattern_matches = self._find_matches(
                    text, 
                    pattern_def
                )
                matches.extend(pattern_matches)
        
        # Filter out whitelist matches
        filtered_matches = self._filter_whitelist(text, matches)
        
        logger.debug(f"Found {len(filtered_matches)} matches after whitelist filtering")
        return filtered_matches
    
    def _find_matches(
        self, 
        text: str, 
        pattern_def: PatternDefinition
    ) -> List[PatternMatch]:
        """
        Find all matches for a specific pattern in text.
        
        Args:
            text: Text to search
            pattern_def: Pattern definition to match against
        
        Returns:
            List of PatternMatch objects
        """
        matches = []
        compiled_pattern = self.compiled_patterns.get(pattern_def.id)
        
        if not compiled_pattern:
            logger.warning(f"No compiled pattern for {pattern_def.id}")
            return matches
        
        for match in compiled_pattern.finditer(text):
            # Extract context around the match
            context_before, context_after = self._extract_context(
                text, 
                match.start(), 
                match.end()
            )
            
            # Calculate line number
            line_number = self._calculate_line_number(text, match.start())
            
            pattern_match = PatternMatch(
                pattern_id=pattern_def.id,
                pattern_name=pattern_def.name,
                matched_text=match.group(0),
                start_pos=match.start(),
                end_pos=match.end(),
                line_number=line_number,
                context_before=context_before,
                context_after=context_after
            )
            
            matches.append(pattern_match)
        
        return matches
    
    def _extract_context(
        self, 
        text: str, 
        start: int, 
        end: int, 
        context_chars: int = 100
    ) -> Tuple[str, str]:
        """
        Extract context text before and after a match.
        
        Args:
            text: Full text content
            start: Start position of match
            end: End position of match
            context_chars: Number of characters to extract on each side
        
        Returns:
            Tuple of (context_before, context_after)
        """
        context_start = max(0, start - context_chars)
        context_end = min(len(text), end + context_chars)
        
        context_before = text[context_start:start].strip()
        context_after = text[end:context_end].strip()
        
        return context_before, context_after
    
    def _calculate_line_number(self, text: str, position: int) -> int:
        """
        Calculate the line number for a position in text.
        
        Args:
            text: Full text content
            position: Character position
        
        Returns:
            Line number (1-indexed)
        """
        return text[:position].count('\n') + 1
    
    def _filter_whitelist(
        self, 
        text: str, 
        matches: List[PatternMatch]
    ) -> List[PatternMatch]:
        """
        Filter out matches that are in the whitelist.
        
        Args:
            text: Original text
            matches: List of pattern matches
        
        Returns:
            Filtered list with whitelist items removed
        """
        if not self.whitelist_patterns:
            return matches
        
        filtered = []
        
        for match in matches:
            # Check if the matched text matches any whitelist pattern
            is_whitelisted = False
            
            for whitelist_pattern in self.whitelist_patterns:
                if whitelist_pattern.search(match.matched_text):
                    is_whitelisted = True
                    logger.debug(
                        f"Whitelisted match: {match.matched_text} "
                        f"(pattern: {match.pattern_id})"
                    )
                    break
            
            if not is_whitelisted:
                filtered.append(match)
        
        return filtered
    
    def check_context_keywords(self, text: str) -> Dict[str, int]:
        """
        Check for risk-adjusting keywords in text context.
        
        Args:
            text: Text to analyze
        
        Returns:
            Dict with counts of high_risk and low_risk keyword matches
        """
        text_lower = text.lower()
        
        high_risk_count = sum(
            1 for keyword in self.context_keywords['high_risk']
            if keyword.lower() in text_lower
        )
        
        low_risk_count = sum(
            1 for keyword in self.context_keywords['low_risk']
            if keyword.lower() in text_lower
        )
        
        return {
            'high_risk': high_risk_count,
            'low_risk': low_risk_count
        }
    
    def get_pattern_by_id(self, pattern_id: str) -> Optional[PatternDefinition]:
        """
        Retrieve a pattern definition by its ID.
        
        Args:
            pattern_id: The pattern identifier
        
        Returns:
            PatternDefinition if found, None otherwise
        """
        for risk_level in [RiskLevel.HIGH, RiskLevel.MEDIUM, RiskLevel.LOW]:
            for pattern_def in self.patterns[risk_level]:
                if pattern_def.id == pattern_id:
                    return pattern_def
        return None
    
    def get_all_patterns(self) -> List[PatternDefinition]:
        """Get all loaded pattern definitions."""
        all_patterns = []
        for risk_level in [RiskLevel.HIGH, RiskLevel.MEDIUM, RiskLevel.LOW]:
            all_patterns.extend(self.patterns[risk_level])
        return all_patterns
