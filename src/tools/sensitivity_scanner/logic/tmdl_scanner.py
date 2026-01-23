"""
TMDL Scanner - Scans Power BI semantic model TMDL files for sensitive content.

This scanner:
- Reads TMDL files from PBIP folders
- Scans tables, measures, columns, relationships, RLS expressions
- Uses PatternDetector to find sensitive patterns
- Uses RiskScorer to assess findings
- Generates comprehensive sensitivity reports
"""

import logging
from pathlib import Path
from typing import List, Optional
from datetime import datetime
import uuid

from core.pbi_file_reader import PBIPReader
from .pattern_detector import PatternDetector
from .risk_scorer import RiskScorer
from .models import ScanResult, Finding


logger = logging.getLogger(__name__)


class TmdlScanner:
    """
    Scans TMDL files in Power BI semantic models for sensitive content.
    
    Uses the shared PBIPReader for file access and PatternDetector/RiskScorer
    for sensitivity analysis.
    """
    
    def __init__(self):
        """Initialize the TMDL scanner with dependencies."""
        self.logger = logging.getLogger(__name__)
        self.pbip_reader = PBIPReader()
        self.pattern_detector = PatternDetector()
        self.risk_scorer = RiskScorer(self.pattern_detector)
    
    def scan_pbip(self, pbip_folder: str) -> ScanResult:
        """
        Scan a complete PBIP folder for sensitive content.
        
        Args:
            pbip_folder: Path to the PBIP folder
        
        Returns:
            ScanResult with all findings aggregated
        """
        start_time = datetime.now()
        scan_id = str(uuid.uuid4())
        
        self.logger.info(f"Starting PBIP scan: {pbip_folder}")
        
        # Validate PBIP folder
        validation = self.pbip_reader.validate_pbip_folder(pbip_folder)
        if not validation['valid']:
            self.logger.error(f"Invalid PBIP folder: {validation['error']}")
            return ScanResult(
                scan_id=scan_id,
                scan_time=start_time,
                source_path=pbip_folder,
                findings=[],
                metadata={'error': validation['error']}
            )
        
        # Find all TMDL files
        tmdl_files = self.pbip_reader.find_tmdl_files(pbip_folder)
        if not tmdl_files:
            self.logger.warning("No TMDL files found in PBIP folder")
            return ScanResult(
                scan_id=scan_id,
                scan_time=start_time,
                source_path=pbip_folder,
                findings=[],
                metadata={'warning': 'No TMDL files found'}
            )
        
        self.logger.info(f"Found {len(tmdl_files)} TMDL files to scan")
        
        # Scan each TMDL file
        all_findings = []
        files_scanned = 0
        
        for name, tmdl_path in tmdl_files.items():
            self.logger.debug(f"Scanning {name}: {tmdl_path}")
            
            # Read file content
            content = self.pbip_reader.read_tmdl_file(tmdl_path)
            if not content:
                self.logger.warning(f"Could not read {name}, skipping")
                continue
            
            # Determine file type for context
            file_type = self._determine_file_type(name)
            location_prefix = self._get_location_prefix(name, tmdl_path)
            
            # Scan content for patterns
            matches = self.pattern_detector.scan_text(content)
            
            if matches:
                self.logger.info(f"Found {len(matches)} matches in {name}")
                
                # Convert matches to findings
                findings = self.risk_scorer.score_matches(
                    matches=matches,
                    file_path=str(tmdl_path),
                    file_type=file_type,
                    location_prefix=location_prefix,
                    full_context=content
                )
                
                all_findings.extend(findings)
            
            files_scanned += 1
        
        # Calculate duration
        end_time = datetime.now()
        duration = (end_time - start_time).total_seconds()
        
        self.logger.info(
            f"Scan complete: {len(all_findings)} findings in "
            f"{files_scanned} files ({duration:.2f}s)"
        )
        
        # Create scan result
        scan_result = ScanResult(
            scan_id=scan_id,
            scan_time=start_time,
            source_path=pbip_folder,
            findings=all_findings,
            total_files_scanned=files_scanned,
            scan_duration_seconds=duration,
            metadata={
                'pbip_info': self.pbip_reader.get_pbip_info(pbip_folder),
                'tmdl_files_found': len(tmdl_files),
                'pattern_version': self.pattern_detector.patterns_file.name 
                    if self.pattern_detector.patterns_file else 'unknown'
            }
        )
        
        return scan_result
    
    def scan_specific_files(
        self, 
        pbip_folder: str, 
        file_names: List[str]
    ) -> ScanResult:
        """
        Scan only specific TMDL files in a PBIP folder.
        
        Args:
            pbip_folder: Path to the PBIP folder
            file_names: List of file names to scan (e.g., ['Sales', 'model'])
        
        Returns:
            ScanResult with findings from specified files only
        """
        start_time = datetime.now()
        scan_id = str(uuid.uuid4())
        
        self.logger.info(f"Scanning specific files in {pbip_folder}: {file_names}")
        
        # Find all TMDL files
        tmdl_files = self.pbip_reader.find_tmdl_files(pbip_folder)
        
        # Filter to requested files
        files_to_scan = {
            name: path for name, path in tmdl_files.items() 
            if name in file_names
        }
        
        if not files_to_scan:
            self.logger.warning(f"None of the requested files found: {file_names}")
            return ScanResult(
                scan_id=scan_id,
                scan_time=start_time,
                source_path=pbip_folder,
                findings=[],
                metadata={'warning': f'Requested files not found: {file_names}'}
            )
        
        # Scan selected files
        all_findings = []
        for name, tmdl_path in files_to_scan.items():
            content = self.pbip_reader.read_tmdl_file(tmdl_path)
            if not content:
                continue
            
            file_type = self._determine_file_type(name)
            location_prefix = self._get_location_prefix(name, tmdl_path)
            
            matches = self.pattern_detector.scan_text(content)
            if matches:
                findings = self.risk_scorer.score_matches(
                    matches=matches,
                    file_path=str(tmdl_path),
                    file_type=file_type,
                    location_prefix=location_prefix,
                    full_context=content
                )
                all_findings.extend(findings)
        
        duration = (datetime.now() - start_time).total_seconds()
        
        return ScanResult(
            scan_id=scan_id,
            scan_time=start_time,
            source_path=pbip_folder,
            findings=all_findings,
            total_files_scanned=len(files_to_scan),
            scan_duration_seconds=duration,
            metadata={
                'scan_mode': 'specific_files',
                'requested_files': file_names,
                'files_found': list(files_to_scan.keys())
            }
        )
    
    def scan_by_category(
        self, 
        pbip_folder: str, 
        category: str
    ) -> ScanResult:
        """
        Scan only files of a specific category.
        
        Args:
            pbip_folder: Path to the PBIP folder
            category: Category to scan ('tables', 'roles', 'expressions', 'model', 'relationships')
        
        Returns:
            ScanResult with findings from specified category only
        """
        start_time = datetime.now()
        scan_id = str(uuid.uuid4())
        
        self.logger.info(f"Scanning category '{category}' in {pbip_folder}")
        
        # Find all TMDL files
        tmdl_files = self.pbip_reader.find_tmdl_files(pbip_folder)
        
        # Filter by category
        if category == 'tables':
            files_to_scan = {
                name: path for name, path in tmdl_files.items()
                if not name.startswith(('model', 'database', 'relationships', 'role_', 'expression_'))
            }
        elif category == 'roles':
            files_to_scan = {
                name: path for name, path in tmdl_files.items()
                if name.startswith('role_')
            }
        elif category == 'expressions':
            files_to_scan = {
                name: path for name, path in tmdl_files.items()
                if name.startswith('expression_')
            }
        elif category == 'model':
            files_to_scan = {
                name: path for name, path in tmdl_files.items()
                if name in ('model', 'database')
            }
        elif category == 'relationships':
            files_to_scan = {
                name: path for name, path in tmdl_files.items()
                if name == 'relationships'
            }
        else:
            self.logger.error(f"Unknown category: {category}")
            return ScanResult(
                scan_id=scan_id,
                scan_time=start_time,
                source_path=pbip_folder,
                findings=[],
                metadata={'error': f'Unknown category: {category}'}
            )
        
        if not files_to_scan:
            self.logger.warning(f"No files found for category: {category}")
            return ScanResult(
                scan_id=scan_id,
                scan_time=start_time,
                source_path=pbip_folder,
                findings=[],
                metadata={'warning': f'No files found for category: {category}'}
            )
        
        # Scan category files
        all_findings = []
        for name, tmdl_path in files_to_scan.items():
            content = self.pbip_reader.read_tmdl_file(tmdl_path)
            if not content:
                continue
            
            file_type = self._determine_file_type(name)
            location_prefix = self._get_location_prefix(name, tmdl_path)
            
            matches = self.pattern_detector.scan_text(content)
            if matches:
                findings = self.risk_scorer.score_matches(
                    matches=matches,
                    file_path=str(tmdl_path),
                    file_type=file_type,
                    location_prefix=location_prefix,
                    full_context=content
                )
                all_findings.extend(findings)
        
        duration = (datetime.now() - start_time).total_seconds()
        
        return ScanResult(
            scan_id=scan_id,
            scan_time=start_time,
            source_path=pbip_folder,
            findings=all_findings,
            total_files_scanned=len(files_to_scan),
            scan_duration_seconds=duration,
            metadata={
                'scan_mode': 'category',
                'category': category,
                'files_scanned': list(files_to_scan.keys())
            }
        )
    
    def _determine_file_type(self, file_name: str) -> str:
        """
        Determine the type of TMDL file for better context in findings.
        
        Args:
            file_name: Normalized file name
        
        Returns:
            File type string
        """
        if file_name.startswith('role_'):
            return "RLS Role"
        elif file_name.startswith('expression_'):
            return "Power Query Expression"
        elif file_name == 'model':
            return "Model Definition"
        elif file_name == 'database':
            return "Database Definition"
        elif file_name == 'relationships':
            return "Relationships"
        else:
            return "Table TMDL"
    
    def _get_location_prefix(self, file_name: str, file_path: Path) -> str:
        """
        Generate a human-readable location prefix for findings.
        
        Args:
            file_name: Normalized file name
            file_path: Path to the file
        
        Returns:
            Location description string
        """
        if file_name.startswith('role_'):
            role_name = file_name.replace('role_', '')
            return f"RLS Role '{role_name}'"
        elif file_name.startswith('expression_'):
            expr_name = file_name.replace('expression_', '')
            return f"Expression '{expr_name}'"
        elif file_name == 'model':
            return "Model Definition"
        elif file_name == 'database':
            return "Database Definition"
        elif file_name == 'relationships':
            return "Relationships"
        else:
            return f"Table '{file_name}'"
