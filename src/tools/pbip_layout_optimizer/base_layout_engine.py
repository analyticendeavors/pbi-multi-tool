# tools/pbip_layout_optimizer/base_layout_engine.py
"""
Base Layout Engine for PBIP model diagrams
Provides common functionality for all layout engines
"""

import json
import logging
import urllib.parse
from typing import Dict, List, Tuple, Any, Optional, Set
from pathlib import Path
from collections import defaultdict

logger = logging.getLogger("pbip-tools-mcp")


class BaseLayoutEngine:
    """Base class for all layout engines with common PBIP functionality"""
    
    def __init__(self, pbip_folder: str):
        self.pbip_folder = Path(pbip_folder)
        self.semantic_model_path = self._find_semantic_model_path()
        
    def _find_semantic_model_path(self) -> Optional[Path]:
        """Find the actual SemanticModel folder path"""
        for item in self.pbip_folder.iterdir():
            if item.is_dir() and item.name.endswith('.SemanticModel'):
                return item
        return None
        
    def _normalize_table_name(self, table_name: str) -> str:
        """Normalize table names to handle special characters and encoding"""
        try:
            decoded = urllib.parse.unquote(table_name)
        except:
            decoded = table_name
            
        normalized = decoded.strip().strip("'").strip('"')
        return normalized
        
    def _find_tmdl_files(self) -> Dict[str, Path]:
        """Find all TMDL files in the semantic model"""
        if not self.semantic_model_path:
            return {}
            
        definition_path = self.semantic_model_path / "definition"
        if not definition_path.exists():
            return {}
        
        tmdl_files = {}
        
        # Find model.tmdl
        model_file = definition_path / "model.tmdl"
        if model_file.exists():
            tmdl_files['model'] = model_file
        
        # Find tables directory
        tables_dir = definition_path / "tables"
        if tables_dir.exists():
            for tmdl_file in tables_dir.glob("*.tmdl"):
                table_name = tmdl_file.stem
                normalized_name = self._normalize_table_name(table_name)
                tmdl_files[normalized_name] = tmdl_file
        
        return tmdl_files
        
    def _get_table_names_from_tmdl(self) -> List[str]:
        """Get all table names from TMDL files with proper normalization"""
        tmdl_files = self._find_tmdl_files()
        tables = []
        
        for name, file_path in tmdl_files.items():
            if name != 'model':  # Skip the main model file
                tables.append(name)
        
        logger.info(f"Found {len(tables)} tables: {sorted(tables)}")
        return sorted(tables)
        
    def _parse_diagram_layout(self) -> Dict[str, Any]:
        """Parse the diagramLayout.json file"""
        if not self.semantic_model_path:
            return {}
            
        diagram_file = self.semantic_model_path / "diagramLayout.json"
        
        if not diagram_file.exists():
            return {}
        
        try:
            with open(diagram_file, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            logger.error(f"Error parsing diagram layout: {e}")
            return {}
    
    def _save_diagram_layout(self, layout_data: Dict[str, Any]) -> bool:
        """Save the diagramLayout.json file"""
        if not self.semantic_model_path:
            return False
            
        diagram_file = self.semantic_model_path / "diagramLayout.json"
        
        try:
            with open(diagram_file, 'w', encoding='utf-8') as f:
                json.dump(layout_data, f, indent=2)
            return True
        except Exception as e:
            logger.error(f"Error saving diagram layout: {e}")
            return False
