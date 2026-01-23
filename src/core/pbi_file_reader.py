"""
Shared Power BI File Reading Utilities

This module provides common utilities for reading and parsing Power BI files,
including both PBIP (Power BI Project) files and PBIX files saved in PBIR format.
Can be used by any tool in the suite that needs to interact with Power BI file structures.

Extracted from pbip_layout_optimizer for reuse across the tool suite.
"""

import json
import logging
import shutil
import tempfile
import zipfile
from pathlib import Path
from typing import Dict, Optional, List, Any


logger = logging.getLogger(__name__)


class PBIPReader:
    """
    Shared utility class for reading Power BI Project (PBIP) and PBIR format files.

    Provides methods to:
    - Validate PBIP folder structure
    - Validate report files (PBIP and PBIX with PBIR format)
    - Find semantic model components
    - Read TMDL files
    - Parse JSON configuration files
    - Locate report definition files
    """

    def __init__(self):
        """Initialize the reader."""
        self.logger = logging.getLogger(__name__)

    def validate_pbip_folder(self, pbip_folder: str) -> Dict[str, Any]:
        """
        Validate that a folder contains a valid PBIP structure.

        Args:
            pbip_folder: Path to the PBIP folder

        Returns:
            Dictionary with validation results:
            {
                'valid': bool,
                'error': str (if invalid),
                'semantic_model_path': str (if valid),
                'definition_path': str (if valid),
                'diagram_layout_path': str (if valid, optional)
            }
        """
        try:
            folder_path = Path(pbip_folder)

            if not folder_path.exists():
                return {'valid': False, 'error': 'Folder does not exist'}

            if not folder_path.is_dir():
                return {'valid': False, 'error': 'Path is not a directory'}

            # Look for .SemanticModel folder
            semantic_folders = list(folder_path.glob("*.SemanticModel"))

            if not semantic_folders:
                return {
                    'valid': False,
                    'error': 'No .SemanticModel folder found. This requires PBIP format files.'
                }

            semantic_model_path = semantic_folders[0]

            # Check for definition folder
            definition_path = semantic_model_path / "definition"
            if not definition_path.exists():
                return {
                    'valid': False,
                    'error': 'No definition folder found in SemanticModel'
                }

            result = {
                'valid': True,
                'semantic_model_path': str(semantic_model_path),
                'definition_path': str(definition_path)
            }

            # Check for diagramLayout.json (optional - not all tools need it)
            diagram_layout_path = semantic_model_path / "diagramLayout.json"
            if not diagram_layout_path.exists():
                diagram_layout_path = definition_path / "diagramLayout.json"

            if diagram_layout_path.exists():
                result['diagram_layout_path'] = str(diagram_layout_path)

            return result

        except Exception as e:
            return {'valid': False, 'error': f'Error validating folder: {str(e)}'}

    def find_tmdl_files(self, pbip_folder: str) -> Dict[str, Path]:
        """
        Find all TMDL files in a PBIP folder.

        Args:
            pbip_folder: Path to the PBIP folder

        Returns:
            Dictionary mapping normalized names to file paths
            Example: {'Sales': Path('tables/Sales.tmdl'), 'model': Path('model.tmdl')}
        """
        try:
            validation = self.validate_pbip_folder(pbip_folder)
            if not validation['valid']:
                self.logger.error(f"Invalid PBIP folder: {validation['error']}")
                return {}

            definition_path = Path(validation['definition_path'])
            tmdl_files = {}

            # Find table TMDL files
            tables_path = definition_path / "tables"
            if tables_path.exists():
                for table_file in tables_path.glob("*.tmdl"):
                    table_name = table_file.stem
                    normalized_name = self._normalize_name(table_name)
                    tmdl_files[normalized_name] = table_file

            # Find relationships TMDL file
            relationships_file = definition_path / "relationships.tmdl"
            if relationships_file.exists():
                tmdl_files['relationships'] = relationships_file

            # Find model TMDL file
            model_file = definition_path / "model.tmdl"
            if model_file.exists():
                tmdl_files['model'] = model_file

            # Find database TMDL file
            database_file = definition_path / "database.tmdl"
            if database_file.exists():
                tmdl_files['database'] = database_file

            # Find roles folder
            roles_path = definition_path / "roles"
            if roles_path.exists():
                for role_file in roles_path.glob("*.tmdl"):
                    role_name = role_file.stem
                    normalized_name = f"role_{self._normalize_name(role_name)}"
                    tmdl_files[normalized_name] = role_file

            # Find expressions folder (Power Query)
            expressions_path = definition_path / "expressions"
            if expressions_path.exists():
                for expr_file in expressions_path.glob("*.tmdl"):
                    expr_name = expr_file.stem
                    normalized_name = f"expression_{self._normalize_name(expr_name)}"
                    tmdl_files[normalized_name] = expr_file

            self.logger.info(f"Found {len(tmdl_files)} TMDL files in {pbip_folder}")
            return tmdl_files

        except Exception as e:
            self.logger.error(f"Error finding TMDL files: {str(e)}")
            return {}

    def read_tmdl_file(self, tmdl_path: Path) -> Optional[str]:
        """
        Read a TMDL file and return its content.

        Args:
            tmdl_path: Path to the TMDL file

        Returns:
            File content as string, or None if read fails
        """
        try:
            if not isinstance(tmdl_path, Path):
                tmdl_path = Path(tmdl_path)

            if not tmdl_path.exists():
                self.logger.error(f"TMDL file does not exist: {tmdl_path}")
                return None

            with open(tmdl_path, 'r', encoding='utf-8') as f:
                content = f.read()

            self.logger.debug(f"Read TMDL file: {tmdl_path.name} ({len(content)} chars)")
            return content

        except Exception as e:
            self.logger.error(f"Error reading TMDL file {tmdl_path}: {str(e)}")
            return None

    def find_report_definition(self, pbip_folder: str) -> Optional[Path]:
        """
        Find the report definition file in a PBIP folder.

        Args:
            pbip_folder: Path to the PBIP folder

        Returns:
            Path to report.json or definition.json, or None if not found
        """
        try:
            folder_path = Path(pbip_folder)

            # Look for .Report folder
            report_folders = list(folder_path.glob("*.Report"))

            if not report_folders:
                self.logger.warning(f"No .Report folder found in {pbip_folder}")
                return None

            report_folder = report_folders[0]

            # Check for report.json
            report_json = report_folder / "report.json"
            if report_json.exists():
                return report_json

            # Check for definition.json (alternative location)
            definition_json = report_folder / "definition.json"
            if definition_json.exists():
                return definition_json

            # Check for definition folder
            definition_folder = report_folder / "definition"
            if definition_folder.exists():
                report_json = definition_folder / "report.json"
                if report_json.exists():
                    return report_json

            self.logger.warning(f"No report definition file found in {report_folder}")
            return None

        except Exception as e:
            self.logger.error(f"Error finding report definition: {str(e)}")
            return None

    def read_json_file(self, json_path: Path) -> Optional[Dict[str, Any]]:
        """
        Read and parse a JSON file.

        Args:
            json_path: Path to the JSON file

        Returns:
            Parsed JSON as dictionary, or None if read fails
        """
        try:
            if not isinstance(json_path, Path):
                json_path = Path(json_path)

            if not json_path.exists():
                self.logger.error(f"JSON file does not exist: {json_path}")
                return None

            with open(json_path, 'r', encoding='utf-8') as f:
                data = json.load(f)

            self.logger.debug(f"Read JSON file: {json_path.name}")
            return data

        except json.JSONDecodeError as e:
            self.logger.error(f"Invalid JSON in {json_path}: {str(e)}")
            return None
        except Exception as e:
            self.logger.error(f"Error reading JSON file {json_path}: {str(e)}")
            return None

    def get_table_names(self, pbip_folder: str) -> List[str]:
        """
        Get a list of all table names in the semantic model.

        Args:
            pbip_folder: Path to the PBIP folder

        Returns:
            List of table names (sorted)
        """
        tmdl_files = self.find_tmdl_files(pbip_folder)

        # Extract table names (exclude special files like model, relationships, roles, expressions)
        table_names = [
            name for name in tmdl_files.keys()
            if not name.startswith(('model', 'database', 'relationships', 'role_', 'expression_'))
        ]

        return sorted(table_names)

    def get_semantic_model_path(self, pbip_folder: str) -> Optional[Path]:
        """
        Get the path to the semantic model folder.

        Args:
            pbip_folder: Path to the PBIP folder

        Returns:
            Path to .SemanticModel folder, or None if not found
        """
        validation = self.validate_pbip_folder(pbip_folder)
        if validation['valid']:
            return Path(validation['semantic_model_path'])
        return None

    def get_definition_path(self, pbip_folder: str) -> Optional[Path]:
        """
        Get the path to the definition folder.

        Args:
            pbip_folder: Path to the PBIP folder

        Returns:
            Path to definition folder, or None if not found
        """
        validation = self.validate_pbip_folder(pbip_folder)
        if validation['valid']:
            return Path(validation['definition_path'])
        return None

    def _normalize_name(self, name: str) -> str:
        """
        Normalize a name for consistent lookup.

        Args:
            name: Name to normalize

        Returns:
            Normalized name
        """
        # Remove quotes and extra whitespace
        normalized = name.strip().strip("'\"")
        return normalized

    def _extract_embedded_pbir(self, pbix_path: Path) -> Optional[Path]:
        """
        Check if PBIX has embedded PBIR structure and extract to temp directory.

        PBIX files saved in PBIR format can have the Report/definition/ structure
        embedded inside the ZIP archive instead of as an external .Report folder.
        Also extracts DAXQueries/ and TMDLScripts/ folders if present.

        Args:
            pbix_path: Path to the PBIX file

        Returns:
            Path to extracted Report folder, or None if not PBIR format
        """
        try:
            with zipfile.ZipFile(pbix_path, 'r') as zf:
                # Check for PBIR structure (Report/definition/ folder)
                pbir_files = [f for f in zf.namelist() if f.startswith('Report/definition/')]
                if not pbir_files:
                    return None  # Not PBIR format

                # Create temp directory for extraction
                temp_dir = Path(tempfile.mkdtemp(prefix='pbir_'))

                # Extract Report folder and additional content folders (DAX queries, TMDL scripts)
                extract_prefixes = ('Report/', 'DAXQueries/', 'TMDLScripts/')
                for file_info in zf.infolist():
                    if file_info.filename.startswith(extract_prefixes):
                        zf.extract(file_info, temp_dir)

                self.logger.info(f"Extracted embedded PBIR to: {temp_dir}")
                # Return path that matches external .Report folder structure
                return temp_dir / 'Report'

        except zipfile.BadZipFile:
            self.logger.error(f"Invalid ZIP file: {pbix_path}")
            return None
        except Exception as e:
            self.logger.error(f"Error extracting embedded PBIR: {e}")
            return None

    def cleanup_extracted_pbir(self, validation_result: Dict[str, Any]) -> None:
        """
        Clean up extracted temp directory after processing.

        Should be called after analysis is complete when working with
        PBIX files that had embedded PBIR extracted to temp directory.

        Args:
            validation_result: The result from validate_report_file()
        """
        if validation_result.get('is_extracted') and validation_result.get('temp_dir'):
            try:
                temp_dir = validation_result['temp_dir']
                if isinstance(temp_dir, Path):
                    shutil.rmtree(temp_dir)
                else:
                    shutil.rmtree(Path(temp_dir))
                self.logger.debug(f"Cleaned up temp directory: {temp_dir}")
            except Exception as e:
                self.logger.warning(f"Failed to cleanup temp dir: {e}")

    def validate_report_file(self, file_path: str) -> Dict[str, Any]:
        """
        Validate a Power BI report file (PBIP or PBIX with PBIR format).

        This method validates that a report file exists and has the required
        Report folder structure. Supports both:
        - External .Report folder (like PBIP projects create)
        - Embedded PBIR inside PBIX files (extracted to temp directory)

        Args:
            file_path: Path to the .pbip or .pbix file

        Returns:
            Dictionary with validation results:
            {
                'valid': bool,
                'error': str (if invalid),
                'error_type': str ('not_found', 'invalid_extension', 'not_pbir_format'),
                'report_dir': Path (if valid) - path to the .Report folder,
                'file_type': str ('pbip' or 'pbix'),
                'is_extracted': bool - True if report_dir is from temp extraction,
                'temp_dir': Path - temp directory root (for cleanup, only if is_extracted)
            }
        """
        try:
            report_file = Path(file_path)

            if not report_file.exists():
                return {
                    'valid': False,
                    'error': f'Report file not found: {file_path}',
                    'error_type': 'not_found'
                }

            ext = report_file.suffix.lower()
            if ext not in ['.pbip', '.pbix']:
                return {
                    'valid': False,
                    'error': f'Invalid file type: {ext}. Expected .pbip or .pbix',
                    'error_type': 'invalid_extension'
                }

            # Check for external .Report folder first
            report_dir = report_file.parent / f"{report_file.stem}.Report"

            if report_dir.exists():
                return {
                    'valid': True,
                    'report_dir': report_dir,
                    'file_type': ext[1:],  # 'pbip' or 'pbix'
                    'is_extracted': False
                }

            # For PBIX, try to extract embedded PBIR
            if ext == '.pbix':
                extracted_dir = self._extract_embedded_pbir(report_file)
                if extracted_dir:
                    return {
                        'valid': True,
                        'report_dir': extracted_dir,
                        'file_type': 'pbix',
                        'is_extracted': True,
                        'temp_dir': extracted_dir.parent  # For cleanup
                    }
                else:
                    return {
                        'valid': False,
                        'error': (
                            "PBIR_FORMAT_ERROR:"
                            "This PBIX file is not saved in PBIR format.\n\n"
                            "This tool requires the report definition files "
                            "which are only available in PBIR format.\n\n"
                            "PBIR became the default format in the January 2026 Power BI Desktop release. "
                            "If this file was saved before that update, simply re-save it in Power BI Desktop "
                            "to convert it to PBIR format."
                        ),
                        'error_type': 'not_pbir_format'
                    }

            # PBIP without .Report folder
            return {
                'valid': False,
                'error': f'Report directory not found: {report_dir}',
                'error_type': 'not_found'
            }

        except Exception as e:
            return {
                'valid': False,
                'error': f'Error validating report file: {str(e)}',
                'error_type': 'exception'
            }

    def get_pbip_info(self, pbip_folder: str) -> Dict[str, Any]:
        """
        Get comprehensive information about a PBIP folder.

        Args:
            pbip_folder: Path to the PBIP folder

        Returns:
            Dictionary with PBIP structure information
        """
        validation = self.validate_pbip_folder(pbip_folder)
        if not validation['valid']:
            return {
                'valid': False,
                'error': validation['error']
            }

        tmdl_files = self.find_tmdl_files(pbip_folder)
        table_names = self.get_table_names(pbip_folder)
        report_def = self.find_report_definition(pbip_folder)

        info = {
            'valid': True,
            'pbip_folder': str(Path(pbip_folder)),
            'semantic_model_path': validation['semantic_model_path'],
            'definition_path': validation['definition_path'],
            'has_report': report_def is not None,
            'report_path': str(report_def) if report_def else None,
            'tmdl_file_count': len(tmdl_files),
            'table_count': len(table_names),
            'table_names': table_names,
            'has_model': 'model' in tmdl_files,
            'has_relationships': 'relationships' in tmdl_files,
            'has_roles': any(k.startswith('role_') for k in tmdl_files.keys()),
            'has_expressions': any(k.startswith('expression_') for k in tmdl_files.keys())
        }

        if 'diagram_layout_path' in validation:
            info['diagram_layout_path'] = validation['diagram_layout_path']
            info['has_diagram_layout'] = True
        else:
            info['has_diagram_layout'] = False

        return info
