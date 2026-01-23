"""
Enhanced PBIP Layout Core - Integrates Advanced Auto-Arrange Functionality
Built by Reid Havens of Analytic Endeavors

This enhanced version uses your migrated MCP components while maintaining 
compatibility with the existing tool architecture.
"""

import json
import math
import logging
from typing import Dict, List, Any, Tuple, Optional
from pathlib import Path
from .base_layout_engine import BaseLayoutEngine


class EnhancedPBIPLayoutCore(BaseLayoutEngine):
    """
    Enhanced PBIP Layout Core that integrates advanced auto-arrange functionality
    using the migrated MCP components.
    """
    
    def __init__(self, pbip_folder: str = None):
        if pbip_folder:
            super().__init__(pbip_folder)
        else:
            # For compatibility with old interface
            self.pbip_folder = None
            self.semantic_model_path = None
            
        self.logger = logging.getLogger("enhanced_pbip_layout_optimizer")
        
        # Advanced integration components (will be initialized after files are moved)
        self.table_categorizer = None
        self.relationship_analyzer = None
        self.middle_out_engine = None
        self.mcp_available = False
        
        # Try to initialize advanced components
        self._initialize_advanced_components()
    
    def _initialize_advanced_components(self):
        """Initialize advanced components if available"""
        try:
            # Import advanced components from the migrated structure
            self.logger.info("🔍 Attempting to import advanced components...")
            
            from .analyzers.table_categorizer import TableCategorizer
            self.logger.info("✅ TableCategorizer imported successfully")
            
            from .analyzers.relationship_analyzer import RelationshipAnalyzer
            self.logger.info("✅ RelationshipAnalyzer imported successfully")
            
            from .engines.middle_out_layout_engine import MiddleOutLayoutEngine
            self.logger.info("✅ MiddleOutLayoutEngine imported successfully")
            
            # All components will be initialized per-operation with the specific PBIP folder
            # This avoids initialization issues when no folder context is available
            self.relationship_analyzer = None  # Will be created on-demand
            self.table_categorizer = None      # Will be created on-demand
            self.middle_out_engine = None      # Will be created on-demand
            
            self.mcp_available = True
            self.logger.info("🎉 Advanced components successfully initialized - Middle-out design AVAILABLE!")
            
        except ImportError as e:
            self.logger.warning(f"📦 Import error - Advanced components not available: {e}")
            self.logger.warning(f"📁 Module path issue - Using basic layout functionality")
            self.mcp_available = False
        except AttributeError as e:
            self.logger.error(f"🔧 Attribute error in component initialization: {e}")
            self.logger.error(f"📁 Component structure issue - Using basic layout functionality")
            self.mcp_available = False
        except Exception as e:
            self.logger.error(f"❌ Unexpected error initializing advanced components: {e}")
            self.logger.error(f"📁 Using basic layout functionality")
            import traceback
            self.logger.error(f"📋 Full traceback: {traceback.format_exc()}")
            self.mcp_available = False
    
    # =============================================================================
    # BRIDGE METHODS FOR TOOL COMPATIBILITY
    # =============================================================================
    
    def validate_pbip_folder(self, pbip_folder: str) -> Dict[str, Any]:
        """Bridge method for tool compatibility"""
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
                    'error': 'No .SemanticModel folder found. This tool requires PBIP format files.'
                }
            
            semantic_model_path = semantic_folders[0]
            
            # Check for definition folder
            definition_path = semantic_model_path / "definition"
            if not definition_path.exists():
                return {'valid': False, 'error': 'No definition folder found in SemanticModel'}
            
            # Check for diagramLayout.json
            diagram_layout_path = semantic_model_path / "diagramLayout.json"
            if not diagram_layout_path.exists():
                diagram_layout_path = definition_path / "diagramLayout.json"
                if not diagram_layout_path.exists():
                    return {
                        'valid': False,
                        'error': 'No diagramLayout.json found. This file is required for layout optimization.'
                    }
            
            return {
                'valid': True,
                'semantic_model_path': str(semantic_model_path),
                'definition_path': str(definition_path),
                'diagram_layout_path': str(diagram_layout_path)
            }
            
        except Exception as e:
            return {'valid': False, 'error': f'Error validating folder: {str(e)}'}
    
    def find_tmdl_files(self, pbip_folder: str) -> Dict[str, Path]:
        """Bridge method - find TMDL files with tool-compatible interface"""
        try:
            validation = self.validate_pbip_folder(pbip_folder)
            if not validation['valid']:
                return {}
            
            definition_path = Path(validation['definition_path'])
            tmdl_files = {}
            
            # Find tables
            tables_path = definition_path / "tables"
            if tables_path.exists():
                for table_file in tables_path.glob("*.tmdl"):
                    table_name = table_file.stem
                    normalized_name = self._normalize_table_name(table_name)
                    tmdl_files[normalized_name] = table_file
            
            # Find model.tmdl
            model_file = definition_path / "model.tmdl"
            if model_file.exists():
                tmdl_files['model'] = model_file
            
            return tmdl_files
            
        except Exception as e:
            self.logger.error(f"Error finding TMDL files: {str(e)}")
            return {}
    
    def get_table_names_from_tmdl(self, tmdl_files: Dict[str, Path]) -> List[str]:
        """Bridge method - extract table names"""
        table_names = []
        for name, path in tmdl_files.items():
            if name != 'model':
                table_names.append(name)
        return sorted(table_names)
    
    def parse_diagram_layout(self, pbip_folder: str) -> Optional[Dict[str, Any]]:
        """Bridge method - parse diagram layout"""
        try:
            validation = self.validate_pbip_folder(pbip_folder)
            if not validation['valid']:
                return None
            
            diagram_layout_path = Path(validation['diagram_layout_path'])
            
            with open(diagram_layout_path, 'r', encoding='utf-8') as f:
                layout_data = json.load(f)
            
            return layout_data
            
        except Exception as e:
            self.logger.error(f"Error parsing diagram layout: {str(e)}")
            return None
    
    def save_diagram_layout(self, pbip_folder: str, layout_data: Dict[str, Any]) -> bool:
        """Bridge method - save diagram layout"""
        try:
            validation = self.validate_pbip_folder(pbip_folder)
            if not validation['valid']:
                return False

            diagram_layout_path = Path(validation['diagram_layout_path'])

            with open(diagram_layout_path, 'w', encoding='utf-8') as f:
                json.dump(layout_data, f, indent=2)

            self.logger.info(f"Saved diagram layout to {diagram_layout_path}")
            return True

        except Exception as e:
            self.logger.error(f"Error saving diagram layout: {str(e)}")
            return False

    def get_diagram_list(self, pbip_folder: str) -> List[Dict[str, Any]]:
        """
        Get list of all diagrams in the diagramLayout.json file with metadata.

        Returns:
            List of dicts: [{'ordinal': 0, 'name': 'All tables', 'table_count': 42}, ...]
        """
        try:
            layout_data = self.parse_diagram_layout(pbip_folder)
            if not layout_data:
                return []

            diagrams = layout_data.get('diagrams', [])
            diagram_list = []

            for i, diagram in enumerate(diagrams):
                if not isinstance(diagram, dict):
                    continue

                # Get table count - prefer 'tables' array (most reliable),
                # fall back to counting nodes
                tables = diagram.get('tables', [])
                if tables:
                    table_count = len(tables)
                else:
                    # Get nodes - check BOTH locations:
                    # 1. Direct on diagram (PBIP native format): diagram.nodes
                    # 2. Nested under layout (our optimizer output): diagram.layout.nodes
                    nodes = diagram.get('nodes', [])
                    if not nodes:
                        layout = diagram.get('layout', {})
                        nodes = layout.get('nodes', []) if isinstance(layout, dict) else []

                    # Count nodes that have position data (left/top or location)
                    table_count = 0
                    for n in nodes:
                        if isinstance(n, dict):
                            # Check both formats: direct left/top or nested location dict
                            has_position = ('left' in n and 'top' in n) or (
                                'location' in n and isinstance(n.get('location'), dict)
                            )
                            if has_position:
                                table_count += 1

                diagram_info = {
                    'ordinal': diagram.get('ordinal', i),
                    'name': diagram.get('name', f'Diagram {i}'),
                    'table_count': table_count,
                    'index': i  # Actual array index for selection
                }
                diagram_list.append(diagram_info)

            return diagram_list

        except Exception as e:
            self.logger.error(f"Error getting diagram list: {str(e)}")
            return []

    # =============================================================================
    # ADVANCED COMPONENT INTEGRATION
    # =============================================================================
    
    def _create_advanced_components(self, pbip_folder: str):
        """Create advanced components for the specific PBIP folder"""
        try:
            if not self.mcp_available:
                return False
                
            # Update semantic model context first
            self._update_semantic_model_context(pbip_folder)
            
            from .analyzers.table_categorizer import TableCategorizer
            from .analyzers.relationship_analyzer import RelationshipAnalyzer
            
            # Initialize relationship analyzer
            self.logger.info(f"🔧 Creating RelationshipAnalyzer for {pbip_folder}...")
            self.relationship_analyzer = RelationshipAnalyzer(self)
            
            # Initialize table categorizer with dependencies
            self.logger.info(f"🔧 Creating TableCategorizer for {pbip_folder}...")
            self.table_categorizer = TableCategorizer(self, self.relationship_analyzer)
            
            self.logger.info("✅ Advanced components created successfully for this operation")
            return True
            
        except Exception as e:
            self.logger.error(f"Error creating advanced components: {e}")
            return False
    
    def _create_middle_out_engine(self, pbip_folder: str):
        """Create a middle-out engine instance for the specific PBIP folder"""
        try:
            from .engines.middle_out_layout_engine import MiddleOutLayoutEngine
            
            # Create the actual middle-out engine with our base engine as parameter
            return MiddleOutLayoutEngine(pbip_folder, self)
            
        except Exception as e:
            self.logger.error(f"Error creating middle-out engine: {e}")
            import traceback
            self.logger.error(f"Full traceback: {traceback.format_exc()}")
            
            # Return a simplified adapter as fallback
            class SimpleMiddleOutAdapter:
                def __init__(self, pbip_folder: str, base_engine):
                    self.pbip_folder = Path(pbip_folder)
                    self.base_engine = base_engine
                    
                def apply_middle_out_layout(self, canvas_width=1400, canvas_height=900, save_changes=True):
                    # Fallback to basic grid layout with enhanced result formatting
                    result = self.base_engine.optimize_layout(
                        str(self.pbip_folder), canvas_width, canvas_height, save_changes, False
                    )
                    
                    if result.get('success'):
                        result['layout_method'] = 'basic_grid_fallback'
                        result['advanced_features'] = {
                            'middle_out_positioning': False,
                            'table_categorization': True,
                            'relationship_analysis': True
                        }
                    
                    return result
            
            return SimpleMiddleOutAdapter(pbip_folder, self)
    
    def get_mcp_status(self) -> Dict[str, Any]:
        """Get status of advanced component availability"""
        return {
            'mcp_available': self.mcp_available,
            'components': {
                'table_categorizer': self.table_categorizer is not None,
                'relationship_analyzer': self.relationship_analyzer is not None,
                'middle_out_engine': self.middle_out_engine is not None
            },
            'message': "Advanced components ready" if self.mcp_available else "Using basic layout functionality"
        }
    
    # =============================================================================
    # LAYOUT ANALYSIS AND OPTIMIZATION
    # =============================================================================
    
    def analyze_table_categorization(self, pbip_folder: str) -> Dict[str, Any]:
        """
        Analyze and categorize tables using advanced components.
        Shows the categorization that will be used for layout optimization.
        """
        try:
            # Validate PBIP folder first
            validation = self.validate_pbip_folder(pbip_folder)
            if not validation['valid']:
                return {
                    'success': False,
                    'error': validation['error'],
                    'operation': 'analyze_table_categorization'
                }
            
            # Check if advanced components are available
            if not self.mcp_available:
                return {
                    'success': False,
                    'error': 'Advanced components not available. Using basic layout functionality.',
                    'operation': 'analyze_table_categorization',
                    'mcp_status': self.get_mcp_status()
                }
            
            # Create advanced components for this operation
            if not self._create_advanced_components(pbip_folder):
                return {
                    'success': False,
                    'error': 'Failed to create advanced components for this operation.',
                    'operation': 'analyze_table_categorization',
                    'mcp_status': self.get_mcp_status()
                }
            
            # Get table names and relationships
            tmdl_files = self.find_tmdl_files(pbip_folder)
            table_names = self.get_table_names_from_tmdl(tmdl_files)
            
            if not table_names:
                return {
                    'success': False,
                    'error': 'No tables found in TMDL files',
                    'operation': 'analyze_table_categorization'
                }
            
            # Analyze relationships using advanced components
            connections = self.relationship_analyzer.build_relationship_graph()
            
            # Categorize tables using advanced components
            categories = self.table_categorizer.categorize_tables(table_names, connections)
            
            # Calculate categorization statistics
            total_tables = len(table_names)
            categorized_tables = sum(len(category_tables) for category_tables in categories.values() 
                                   if isinstance(category_tables, list))
            
            # Create categorization summary
            categorization_summary = {
                'fact_tables': {
                    'count': len(categories.get('fact_tables', [])),
                    'tables': categories.get('fact_tables', [])
                },
                'dimension_tables': {
                    'l1_count': len(categories.get('l1_dimensions', [])),
                    'l2_count': len(categories.get('l2_dimensions', [])),
                    'l3_count': len(categories.get('l3_dimensions', [])),
                    'l4_plus_count': len(categories.get('l4_plus_dimensions', [])),
                    'l1_tables': categories.get('l1_dimensions', []),
                    'l2_tables': categories.get('l2_dimensions', []),
                    'l3_tables': categories.get('l3_dimensions', []),
                    'l4_plus_tables': categories.get('l4_plus_dimensions', [])
                },
                'special_tables': {
                    'calendar_count': len(categories.get('calendar_tables', [])),
                    'metrics_count': len(categories.get('metrics_tables', [])),
                    'parameter_count': len(categories.get('parameter_tables', [])),
                    'calculation_groups_count': len(categories.get('calculation_groups', [])),
                    'calendar_tables': categories.get('calendar_tables', []),
                    'metrics_tables': categories.get('metrics_tables', []),
                    'parameter_tables': categories.get('parameter_tables', []),
                    'calculation_groups': categories.get('calculation_groups', [])
                },
                'disconnected_tables': {
                    'count': len(categories.get('disconnected_tables', [])),
                    'tables': categories.get('disconnected_tables', [])
                },
                'excluded_tables': {
                    'auto_date_count': len(categories.get('auto_date_tables', [])),
                    'auto_date_tables': categories.get('auto_date_tables', [])
                }
            }
            
            # Analyze extensions
            extensions = categories.get('dimension_extensions', {})
            extension_summary = []
            for ext_table, ext_info in extensions.items():
                if isinstance(ext_info, dict):
                    base_table = ext_info.get('base_table', 'Unknown')
                    ext_type = ext_info.get('type', 'extension')
                else:
                    base_table = ext_info[0] if ext_info else 'Unknown'
                    ext_type = ext_info[1] if len(ext_info) > 1 else 'extension'
                
                extension_summary.append({
                    'extension_table': ext_table,
                    'base_table': base_table,
                    'type': ext_type
                })
            
            return {
                'success': True,
                'operation': 'analyze_table_categorization',
                'pbip_folder': str(Path(pbip_folder)),
                'semantic_model_path': validation['semantic_model_path'],
                'mcp_status': self.get_mcp_status(),
                'model_info': {
                    'total_tables': total_tables,
                    'categorized_tables': categorized_tables,
                    'tmdl_files_found': len(tmdl_files)
                },
                'categorization': categorization_summary,
                'extensions': extension_summary,
                'relationship_connections': dict(connections),  # Convert sets to lists for JSON serialization
                'layout_ready': True
            }
            
        except Exception as e:
            self.logger.error(f"Error analyzing table categorization: {str(e)}")
            return {
                'success': False,
                'error': str(e),
                'operation': 'analyze_table_categorization',
                'mcp_status': self.get_mcp_status()
            }
    
    def optimize_layout_with_advanced(self, pbip_folder: str, canvas_width: int = 1400, canvas_height: int = 900,
                                    save_changes: bool = True, use_middle_out: bool = True,
                                    diagram_indices: List[int] = None) -> Dict[str, Any]:
        """
        Optimize layout using advanced middle-out engine with enhanced categorization.

        Args:
            pbip_folder: Path to PBIP folder
            canvas_width: Target canvas width
            canvas_height: Target canvas height
            save_changes: Whether to save changes to file
            use_middle_out: Whether to use middle-out layout algorithm
            diagram_indices: List of diagram indices to optimize (default: [0])
        """
        try:
            # Default to first diagram for backwards compatibility
            if diagram_indices is None:
                diagram_indices = [0]

            # First, analyze table categorization
            categorization_result = self.analyze_table_categorization(pbip_folder)
            if not categorization_result['success']:
                return categorization_result

            # Check if advanced components are available
            if not self.mcp_available:
                # Fallback to basic grid layout
                self.logger.warning("Advanced components not available, falling back to basic grid layout")
                return self._basic_grid_layout(pbip_folder, canvas_width, canvas_height, save_changes, diagram_indices)

            # Create middle-out engine for this operation
            middle_out_engine = self._create_middle_out_engine(pbip_folder)
            if not middle_out_engine:
                self.logger.warning("Could not create middle-out engine, falling back to basic layout")
                return self._basic_grid_layout(pbip_folder, canvas_width, canvas_height, save_changes, diagram_indices)

            # Get current layout to extract existing diagram nodes
            layout_data = self.parse_diagram_layout(pbip_folder)

            # Optimize each selected diagram
            diagrams_optimized = []
            total_tables_arranged = 0
            all_changes_saved = True

            for diagram_idx in diagram_indices:
                # Get existing table names from this diagram (filter to only these tables)
                filter_tables = None
                if layout_data and 'diagrams' in layout_data and diagram_idx < len(layout_data['diagrams']):
                    diagram = layout_data['diagrams'][diagram_idx]
                    if isinstance(diagram, dict):
                        # Prefer 'tables' array if available, otherwise get from nodes
                        tables_array = diagram.get('tables', [])
                        if tables_array:
                            filter_tables = list(tables_array)
                            self.logger.info(f"Diagram {diagram_idx}: using tables array with {len(filter_tables)} tables")
                        else:
                            # Get nodes - check BOTH locations:
                            # 1. Direct on diagram (PBIP native format): diagram.nodes
                            # 2. Nested under layout (our optimizer output): diagram.layout.nodes
                            existing_nodes = diagram.get('nodes', [])
                            if not existing_nodes:
                                layout = diagram.get('layout', {})
                                existing_nodes = layout.get('nodes', []) if isinstance(layout, dict) else []
                            # Safely extract nodeIndex from nodes (could be string or dict)
                            filter_tables = []
                            for n in existing_nodes:
                                if isinstance(n, dict):
                                    node_index = n.get('nodeIndex')
                                    if node_index:
                                        # nodeIndex could be string or dict with 'name' key
                                        if isinstance(node_index, str):
                                            filter_tables.append(node_index)
                                        elif isinstance(node_index, dict):
                                            name = node_index.get('name')
                                            if name:
                                                filter_tables.append(name)
                            self.logger.info(f"Diagram {diagram_idx}: filtering to {len(filter_tables)} existing tables")

                # Use middle-out engine for enhanced layout (filtered to diagram's tables)
                layout_result = middle_out_engine.apply_middle_out_layout(
                    canvas_width,
                    canvas_height,
                    save_changes,
                    diagram_index=diagram_idx,
                    filter_tables=filter_tables
                )
                if layout_result.get('success'):
                    diagrams_optimized.append(diagram_idx)
                    total_tables_arranged += layout_result.get('tables_arranged', 0)
                    if not layout_result.get('changes_saved', False):
                        all_changes_saved = False

            # Build combined result with aggregated metrics
            final_result = {
                'success': len(diagrams_optimized) > 0,
                'operation': 'optimize_layout_with_advanced',
                'diagrams_optimized': diagrams_optimized,
                'diagrams_requested': diagram_indices,
                'tables_arranged': total_tables_arranged,
                'changes_saved': all_changes_saved and save_changes,
                'canvas_size': {'width': canvas_width, 'height': canvas_height},
                'categorization_preview': categorization_result['categorization'],
                'extensions': categorization_result['extensions'],
                'layout_method': 'middle_out_advanced',
                'advanced_features': {
                    'table_categorization': True,
                    'relationship_analysis': True,
                    'middle_out_positioning': True,
                    'extension_handling': True,
                    'auto_date_filtering': True
                }
            }

            return final_result

        except Exception as e:
            self.logger.error(f"Error optimizing layout with advanced components: {str(e)}")
            return {
                'success': False,
                'error': str(e),
                'operation': 'optimize_layout_with_advanced',
                'mcp_status': self.get_mcp_status()
            }
    
    def optimize_layout(self, pbip_folder: str, canvas_width: int = 1400, canvas_height: int = 900,
                       save_changes: bool = True, use_middle_out: bool = True,
                       diagram_indices: List[int] = None) -> Dict[str, Any]:
        """
        Enhanced optimize_layout that uses advanced components when available.
        Falls back to basic grid layout if advanced components are not available.

        Args:
            pbip_folder: Path to PBIP folder
            canvas_width: Target canvas width
            canvas_height: Target canvas height
            save_changes: Whether to save changes to file
            use_middle_out: Whether to use middle-out layout algorithm
            diagram_indices: List of diagram indices to optimize (default: [0] for backwards compatibility)
        """
        # Default to first diagram for backwards compatibility
        if diagram_indices is None:
            diagram_indices = [0]

        if self.mcp_available and use_middle_out:
            return self.optimize_layout_with_advanced(pbip_folder, canvas_width, canvas_height, save_changes, use_middle_out, diagram_indices)
        else:
            # Use basic grid layout as fallback
            return self._basic_grid_layout(pbip_folder, canvas_width, canvas_height, save_changes, diagram_indices)
    
    def _basic_grid_layout(self, pbip_folder: str, canvas_width: int, canvas_height: int,
                          save_changes: bool, diagram_indices: List[int] = None) -> Dict[str, Any]:
        """Basic grid layout fallback supporting multiple diagrams"""
        try:
            # Default to first diagram for backwards compatibility
            if diagram_indices is None:
                diagram_indices = [0]

            validation = self.validate_pbip_folder(pbip_folder)
            if not validation['valid']:
                return {
                    'success': False,
                    'error': validation['error'],
                    'operation': 'optimize_layout'
                }

            # Get current layout
            layout_data = self.parse_diagram_layout(pbip_folder)
            if not layout_data:
                return {
                    'success': False,
                    'error': 'Could not parse diagram layout',
                    'operation': 'optimize_layout'
                }

            # Get ALL table names from TMDL (for reference)
            tmdl_files = self.find_tmdl_files(pbip_folder)
            all_table_names = self.get_table_names_from_tmdl(tmdl_files)

            if not all_table_names:
                return {
                    'success': False,
                    'error': 'No tables found',
                    'operation': 'optimize_layout'
                }

            # Update layout data for each selected diagram (filter to only existing tables)
            diagrams_optimized = []
            total_tables_arranged = 0
            if 'diagrams' in layout_data and layout_data['diagrams']:
                for diagram_idx in diagram_indices:
                    if diagram_idx < len(layout_data['diagrams']):
                        diagram = layout_data['diagrams'][diagram_idx]
                        if not isinstance(diagram, dict):
                            continue

                        # Get existing table names from this diagram only (safely)
                        # Prefer 'tables' array if available (most reliable)
                        tables_array = diagram.get('tables', [])
                        if tables_array:
                            diagram_table_names = list(tables_array)
                        else:
                            # Get nodes - check BOTH locations:
                            # 1. Direct on diagram (PBIP native format): diagram.nodes
                            # 2. Nested under layout (our optimizer output): diagram.layout.nodes
                            existing_nodes = diagram.get('nodes', [])
                            if not existing_nodes:
                                layout = diagram.get('layout', {})
                                existing_nodes = layout.get('nodes', []) if isinstance(layout, dict) else []
                            diagram_table_names = []
                            for n in existing_nodes:
                                if isinstance(n, dict):
                                    node_index = n.get('nodeIndex')
                                    if node_index:
                                        # nodeIndex could be string or dict with 'name' key
                                        if isinstance(node_index, str):
                                            diagram_table_names.append(node_index)
                                        elif isinstance(node_index, dict):
                                            name = node_index.get('name')
                                            if name:
                                                diagram_table_names.append(name)

                        if not diagram_table_names:
                            # Fallback to all tables if diagram has no existing nodes
                            diagram_table_names = all_table_names

                        # Generate positions only for this diagram's tables
                        new_positions = self._generate_grid_positions(diagram_table_names, canvas_width, canvas_height)
                        # Ensure layout object exists and write nodes there
                        if 'layout' not in layout_data['diagrams'][diagram_idx]:
                            layout_data['diagrams'][diagram_idx]['layout'] = {}
                        layout_data['diagrams'][diagram_idx]['layout']['nodes'] = new_positions
                        diagrams_optimized.append(diagram_idx)
                        total_tables_arranged += len(new_positions)

            # Save if requested
            saved = False
            if save_changes and diagrams_optimized:
                saved = self.save_diagram_layout(pbip_folder, layout_data)

            return {
                'success': len(diagrams_optimized) > 0,
                'operation': 'optimize_layout_basic_grid',
                'pbip_folder': pbip_folder,
                'layout_method': 'basic_grid',
                'tables_arranged': total_tables_arranged,
                'diagrams_optimized': diagrams_optimized,
                'diagrams_requested': diagram_indices,
                'changes_saved': saved,
                'canvas_size': {'width': canvas_width, 'height': canvas_height},
                'mcp_status': self.get_mcp_status()
            }

        except Exception as e:
            self.logger.error(f"Error optimizing layout: {str(e)}")
            return {
                'success': False,
                'error': str(e),
                'operation': 'optimize_layout'
            }
    
    def _generate_grid_positions(self, table_names: List[str], canvas_width: int, canvas_height: int) -> List[Dict[str, Any]]:
        """Generate basic grid positions for tables"""
        positions = []
        
        table_width = 200
        table_height = 104
        spacing = 50
        margin = 50
        
        # Calculate grid dimensions
        tables_per_row = max(1, (canvas_width - 2 * margin) // (table_width + spacing))
        rows_needed = math.ceil(len(table_names) / tables_per_row)
        
        for i, table_name in enumerate(table_names):
            row = i // tables_per_row
            col = i % tables_per_row
            
            x = margin + col * (table_width + spacing)
            y = margin + row * (table_height + spacing)
            
            positions.append({
                'nodeIndex': table_name,
                'location': {'x': x, 'y': y},
                'size': {'width': table_width, 'height': table_height},
                'zIndex': i
            })
        
        return positions
    
    def analyze_layout_quality(self, pbip_folder: str) -> Dict[str, Any]:
        """Analyze the current layout quality"""
        try:
            validation = self.validate_pbip_folder(pbip_folder)
            if not validation['valid']:
                return {
                    'success': False,
                    'error': validation['error'],
                    'operation': 'analyze_layout_quality'
                }
            
            # Parse current layout
            layout_data = self.parse_diagram_layout(pbip_folder)
            if not layout_data:
                return {
                    'success': False,
                    'error': 'Could not parse diagram layout',
                    'operation': 'analyze_layout_quality'
                }
            
            # Get table information
            tmdl_files = self.find_tmdl_files(pbip_folder)
            table_names = self.get_table_names_from_tmdl(tmdl_files)
            
            # Analyze current positions
            analysis = self.analyze_current_positions(layout_data)
            
            # Calculate quality score
            quality_score = self._calculate_quality_score(analysis)
            rating = self._get_quality_rating(quality_score)
            
            # Generate recommendations
            recommendations = self._generate_recommendations(analysis)
            
            return {
                'success': True,
                'operation': 'analyze_layout_quality',
                'pbip_folder': pbip_folder,
                'semantic_model_path': validation['semantic_model_path'],
                'quality_score': quality_score,
                'rating': rating,
                'layout_analysis': analysis,
                'recommendations': recommendations,
                'table_names': table_names[:10],  # First 10 for preview
                'total_tables': len(table_names)
            }
            
        except Exception as e:
            self.logger.error(f"Error analyzing layout quality: {str(e)}")
            return {
                'success': False,
                'error': str(e),
                'operation': 'analyze_layout_quality'
            }
    
    def analyze_current_positions(self, layout_data: Dict[str, Any]) -> Dict[str, Any]:
        """Analyze current table positions"""
        analysis = {
            'total_tables': 0,
            'positioned_tables': 0,
            'overlapping_tables': 0,
            'tables_outside_canvas': 0,
            'average_spacing': 0,
            'layout_efficiency': 0
        }

        try:
            if 'diagrams' not in layout_data or not layout_data['diagrams']:
                return analysis

            diagram = layout_data['diagrams'][0]
            # Get nodes - check BOTH locations:
            # 1. Direct on diagram (PBIP native format): diagram.nodes
            # 2. Nested under layout (our optimizer output): diagram.layout.nodes
            nodes = diagram.get('nodes', [])
            if not nodes:
                layout = diagram.get('layout', {})
                nodes = layout.get('nodes', []) if isinstance(layout, dict) else []

            analysis['total_tables'] = len(nodes)

            # Count positioned tables safely - check for both formats
            positioned_count = 0
            for n in nodes:
                if isinstance(n, dict):
                    # Check both formats: direct left/top or nested location dict
                    has_position = ('left' in n and 'top' in n) or (
                        'location' in n and isinstance(n.get('location'), dict)
                    )
                    if has_position:
                        positioned_count += 1
            analysis['positioned_tables'] = positioned_count

            # Calculate overlaps and spacing
            positions = []
            for node in nodes:
                # Skip if node is not a dict
                if not isinstance(node, dict):
                    continue

                # Get position - support both formats (left/top or location.x/y)
                if 'left' in node and 'top' in node:
                    x = node.get('left', 0)
                    y = node.get('top', 0)
                elif 'location' in node and isinstance(node.get('location'), dict):
                    location = node['location']
                    x = location.get('x', 0)
                    y = location.get('y', 0)
                else:
                    continue

                positions.append((x, y))
            
            if len(positions) > 1:
                # Calculate average spacing
                total_distance = 0
                count = 0
                
                for i in range(len(positions)):
                    for j in range(i + 1, len(positions)):
                        x1, y1 = positions[i]
                        x2, y2 = positions[j]
                        distance = math.sqrt((x2 - x1) ** 2 + (y2 - y1) ** 2)
                        total_distance += distance
                        count += 1
                
                if count > 0:
                    analysis['average_spacing'] = round(total_distance / count, 1)
            
            # Calculate layout efficiency (simplified)
            if analysis['total_tables'] > 0:
                efficiency = (analysis['positioned_tables'] / analysis['total_tables']) * 100
                if analysis['overlapping_tables'] > 0:
                    efficiency -= (analysis['overlapping_tables'] / analysis['total_tables']) * 20
                analysis['layout_efficiency'] = max(0, round(efficiency, 1))
            
        except Exception as e:
            self.logger.error(f"Error analyzing positions: {str(e)}")
        
        return analysis
    
    def _calculate_quality_score(self, analysis: Dict[str, Any]) -> float:
        """Calculate quality score with relationship awareness and size-adjusted spacing"""
        score = 0
        total_tables = analysis.get('total_tables', 0)
        positioned_tables = analysis.get('positioned_tables', 0)
        overlapping = analysis.get('overlapping_tables', 0)

        # SMALL DIAGRAM HANDLING (1-3 tables)
        # These diagrams can't be fairly judged on spacing/relationships
        if total_tables <= 3:
            if total_tables == 0:
                return 0
            # Base: 70 points for having tables positioned
            positioned_ratio = positioned_tables / total_tables
            score = 70 * positioned_ratio
            # Bonus: +20 if no overlaps
            if overlapping == 0:
                score += 20
            # Bonus: +10 for being fully positioned
            if positioned_tables == total_tables:
                score += 10
            return max(0, min(100, round(score, 1)))

        # STANDARD SCORING for diagrams with 4+ tables
        # Base positioning (25 points max)
        if total_tables > 0:
            positioned_ratio = positioned_tables / total_tables
            score += positioned_ratio * 25
            # Bonus for 100% positioned (+5)
            if positioned_ratio == 1.0:
                score += 5

        # Overlap scoring (up to 25 points for no overlaps, penalty for overlaps)
        if total_tables > 0:
            if overlapping == 0:
                score += 25  # Perfect - no overlaps (optimizer goal)
            else:
                overlap_ratio = overlapping / max(1, total_tables)
                score -= overlap_ratio * 25  # Penalty scales with overlap severity

        # Spacing quality (0-25 points) - SIZE ADJUSTED
        # Larger models naturally have wider average spacing
        spacing = analysis.get('average_spacing', 0)
        if spacing > 0 and total_tables > 0:
            # Calculate ideal spacing range based on table count
            # Small model (10): ideal 200-600, Large model (60+): ideal 400-2500
            # Optimizer creates wider spacing for large models to avoid overlaps
            ideal_min = 200 + (total_tables * 3)   # 10→230, 30→290, 60→380
            ideal_max = 600 + (total_tables * 35)  # 10→950, 30→1650, 60→2700

            if spacing < ideal_min * 0.5:
                score += 5    # Very cramped
            elif spacing < ideal_min:
                score += 15   # Cramped
            elif spacing <= ideal_max:
                score += 25   # Ideal range for this model size
            elif spacing <= ideal_max * 1.3:
                score += 20   # Slightly spread (less penalty)
            else:
                score += 12   # Very spread (less harsh)

        # Layout efficiency (up to 15 points)
        score += analysis.get('layout_efficiency', 0) * 0.15

        # Relationship-based score (up to 30 points) - key differentiator
        relationship_score = analysis.get('relationship_score', 0)
        score += relationship_score * 0.3

        return max(0, min(100, round(score, 1)))
    
    def _get_quality_rating(self, score: float) -> str:
        """Get quality rating from score"""
        if score >= 90:
            return "EXCELLENT"
        elif score >= 80:
            return "GOOD"
        elif score >= 60:
            return "OK"
        elif score >= 30:
            return "BAD"
        else:
            return "VERY BAD"
    
    def _generate_recommendations(self, analysis: Dict[str, Any]) -> List[str]:
        """Generate layout improvement recommendations"""
        recommendations = []

        if analysis['overlapping_tables'] > 0:
            recommendations.append("Resolve overlapping tables by adjusting positions")

        if analysis['average_spacing'] < 200:
            recommendations.append("Increase spacing between tables for better readability")
        elif analysis['average_spacing'] > 2000:
            recommendations.append("Reduce spacing to make the diagram more compact")

        if analysis['positioned_tables'] < analysis['total_tables']:
            missing = analysis['total_tables'] - analysis['positioned_tables']
            recommendations.append(f"Position {missing} unpositioned tables")

        if analysis['layout_efficiency'] < 50:
            recommendations.append("Consider using Haven's middle-out design for better organization")

        if not recommendations:
            recommendations.append("Layout looks good! Consider minor adjustments for optimal positioning")

        return recommendations

    def analyze_diagram_quality(self, pbip_folder: str) -> List[Dict[str, Any]]:
        """
        Analyze each diagram individually and return per-diagram quality scores.

        Returns:
            List of dicts with diagram quality information:
            [{'index': 0, 'name': 'All tables', 'table_count': 42, 'quality_score': 65,
              'rating': 'GOOD', 'overlapping': 0, 'avg_spacing': 450}, ...]
        """
        try:
            validation = self.validate_pbip_folder(pbip_folder)
            if not validation['valid']:
                return []

            layout_data = self.parse_diagram_layout(pbip_folder)
            if not layout_data:
                return []

            diagrams = layout_data.get('diagrams', [])
            diagram_scores = []

            for i, diagram in enumerate(diagrams):
                # Get nodes - check BOTH locations:
                # 1. Direct on diagram (PBIP native format): diagram.nodes
                # 2. Nested under layout (our optimizer output): diagram.layout.nodes
                nodes = diagram.get('nodes', [])
                if not nodes:
                    layout = diagram.get('layout', {})
                    nodes = layout.get('nodes', []) if isinstance(layout, dict) else []
                name = diagram.get('name', f'Diagram {i}')

                # Use 'tables' array if available (most reliable count)
                tables = diagram.get('tables', [])
                table_count = len(tables) if tables else len(nodes)

                # Extract table names for categorization filtering
                table_names = []
                if tables:
                    # Tables array contains table names directly
                    table_names = list(tables)
                else:
                    # Extract from nodes via nodeIndex
                    for n in nodes:
                        if isinstance(n, dict):
                            node_index = n.get('nodeIndex')
                            if node_index:
                                if isinstance(node_index, str):
                                    table_names.append(node_index)
                                elif isinstance(node_index, dict):
                                    tbl_name = node_index.get('name')
                                    if tbl_name:
                                        table_names.append(tbl_name)

                # Analyze positions for this specific diagram (pass pbip_folder for relationship analysis)
                analysis = self._analyze_diagram_positions(nodes, pbip_folder)

                # Calculate quality score using existing method
                score = self._calculate_quality_score(analysis)
                rating = self._get_quality_rating(score)

                diagram_scores.append({
                    'index': i,
                    'name': name,
                    'table_count': table_count,  # Use tables array count if available
                    'table_names': table_names,  # List of table names for filtering
                    'quality_score': score,
                    'rating': rating,
                    'overlapping': analysis.get('overlapping_tables', 0),
                    'avg_spacing': round(analysis.get('average_spacing', 0), 0)
                })

            return diagram_scores

        except Exception as e:
            self.logger.error(f"Error analyzing diagram quality: {str(e)}")
            return []

    def _analyze_diagram_positions(self, nodes: List[Dict[str, Any]], pbip_folder: str = None) -> Dict[str, Any]:
        """Analyze positions for a single diagram's nodes"""
        analysis = {
            'total_tables': len(nodes),
            'positioned_tables': 0,
            'overlapping_tables': 0,
            'tables_outside_canvas': 0,
            'average_spacing': 0,
            'layout_efficiency': 0,
            'relationship_score': 0,
            'clustering_score': 0,
            'orphan_penalty': 0
        }

        try:
            # Count positioned tables safely - check for both formats
            positioned_count = 0
            for n in nodes:
                if isinstance(n, dict):
                    # Check both formats: direct left/top or nested location dict
                    has_position = ('left' in n and 'top' in n) or (
                        'location' in n and isinstance(n.get('location'), dict)
                    )
                    if has_position:
                        positioned_count += 1
            analysis['positioned_tables'] = positioned_count

            # Calculate overlaps and spacing
            positions = []
            sizes = []
            for node in nodes:
                # Skip if node is not a dict
                if not isinstance(node, dict):
                    continue

                # Get position - support both formats (left/top or location.x/y)
                if 'left' in node and 'top' in node:
                    x = node.get('left', 0)
                    y = node.get('top', 0)
                elif 'location' in node and isinstance(node.get('location'), dict):
                    location = node['location']
                    x = location.get('x', 0)
                    y = location.get('y', 0)
                else:
                    continue

                # Get size - support both formats (direct width/height or nested size dict)
                if 'width' in node and 'height' in node:
                    width = node.get('width', 200)
                    height = node.get('height', 104)
                elif 'size' in node and isinstance(node.get('size'), dict):
                    size = node['size']
                    width = size.get('width', 200)
                    height = size.get('height', 104)
                else:
                    width = 200
                    height = 104

                positions.append((x, y))
                sizes.append((width, height))

            # Check for overlaps
            overlap_count = 0
            for i in range(len(positions)):
                for j in range(i + 1, len(positions)):
                    x1, y1 = positions[i]
                    w1, h1 = sizes[i]
                    x2, y2 = positions[j]
                    w2, h2 = sizes[j]

                    # Check if rectangles overlap
                    if (x1 < x2 + w2 and x1 + w1 > x2 and
                        y1 < y2 + h2 and y1 + h1 > y2):
                        overlap_count += 1

            analysis['overlapping_tables'] = overlap_count

            if len(positions) > 1:
                # Calculate average spacing between tables
                total_distance = 0
                count = 0

                for i in range(len(positions)):
                    for j in range(i + 1, len(positions)):
                        x1, y1 = positions[i]
                        x2, y2 = positions[j]
                        distance = math.sqrt((x2 - x1) ** 2 + (y2 - y1) ** 2)
                        total_distance += distance
                        count += 1

                if count > 0:
                    analysis['average_spacing'] = total_distance / count

            # Calculate layout efficiency
            if analysis['total_tables'] > 0:
                efficiency = (analysis['positioned_tables'] / analysis['total_tables']) * 100
                if analysis['overlapping_tables'] > 0:
                    efficiency -= (analysis['overlapping_tables'] / analysis['total_tables']) * 20
                analysis['layout_efficiency'] = max(0, round(efficiency, 1))

            # Add relationship-based analysis if pbip_folder provided
            if pbip_folder:
                rel_quality = self._analyze_relationship_quality(nodes, pbip_folder)
                analysis['relationship_score'] = rel_quality.get('relationship_score', 0)
                analysis['clustering_score'] = rel_quality.get('clustering_score', 0)
                analysis['orphan_penalty'] = rel_quality.get('orphan_penalty', 0)

        except Exception as e:
            self.logger.error(f"Error analyzing diagram positions: {str(e)}")

        return analysis

    def _analyze_relationship_quality(self, nodes: List[Dict], pbip_folder: str) -> Dict[str, Any]:
        """Analyze layout quality based on relationship structure"""
        result = {
            'clustering_score': 0,      # How well related tables are grouped (0-100)
            'orphan_penalty': 0,        # Penalty for disconnected tables (0-100)
            'stratification_score': 0,  # Vertical layering quality (0-100)
            'relationship_score': 0     # Combined score (0-100)
        }

        try:
            # Ensure relationship analyzer is available
            if not self.relationship_analyzer:
                return result

            # Update context to ensure paths are set
            self._update_semantic_model_context(pbip_folder)

            # Parse relationships using the relationship analyzer
            relationships = self.relationship_analyzer.parse_relationships_from_tmdl()
            if not relationships:
                return result

            # Build connection map
            connection_counts = {}
            for rel in relationships:
                # Skip if rel is not a dict (safety check)
                if not isinstance(rel, dict):
                    continue
                from_table = rel.get('fromTable', '')
                to_table = rel.get('toTable', '')
                connection_counts[from_table] = connection_counts.get(from_table, 0) + 1
                connection_counts[to_table] = connection_counts.get(to_table, 0) + 1

            # Extract node names and positions
            node_positions = {}
            for node in nodes:
                # Skip if node is not a dict (safety check)
                if not isinstance(node, dict):
                    continue
                # Safely get nodeIndex - might be string or dict
                node_index = node.get('nodeIndex', {})
                if isinstance(node_index, dict):
                    name = node_index.get('name', '')
                elif isinstance(node_index, str):
                    name = node_index
                else:
                    name = ''

                # Get position - support both formats (left/top or location.x/y)
                if 'left' in node and 'top' in node:
                    node_positions[name] = (node.get('left', 0), node.get('top', 0))
                elif 'location' in node and isinstance(node.get('location'), dict):
                    location = node['location']
                    node_positions[name] = (location.get('x', 0), location.get('y', 0))

            # 1. Orphan penalty - tables with 0 connections
            orphan_count = sum(1 for name in node_positions if connection_counts.get(name, 0) == 0)
            if len(node_positions) > 0:
                orphan_ratio = orphan_count / len(node_positions)
                result['orphan_penalty'] = orphan_ratio * 100

            # 2. Clustering score - related tables should be near each other
            if relationships and len(node_positions) > 1:
                related_distances = []

                for rel in relationships:
                    # Skip if rel is not a dict (safety check)
                    if not isinstance(rel, dict):
                        continue
                    from_table = rel.get('fromTable', '')
                    to_table = rel.get('toTable', '')

                    if from_table in node_positions and to_table in node_positions:
                        x1, y1 = node_positions[from_table]
                        x2, y2 = node_positions[to_table]
                        dist = math.sqrt((x2 - x1) ** 2 + (y2 - y1) ** 2)
                        related_distances.append(dist)

                if related_distances:
                    avg_related = sum(related_distances) / len(related_distances)
                    # Ideal: related tables within 400px of each other
                    if avg_related < 200:
                        result['clustering_score'] = 100
                    elif avg_related < 400:
                        result['clustering_score'] = 80
                    elif avg_related < 600:
                        result['clustering_score'] = 60
                    elif avg_related < 1000:
                        result['clustering_score'] = 40
                    else:
                        result['clustering_score'] = 20

            # 3. Stratification - are high-connection tables (facts) centrally located?
            if node_positions and connection_counts:
                positions_list = list(node_positions.values())
                if positions_list:
                    # Find centroid
                    avg_x = sum(p[0] for p in positions_list) / len(positions_list)
                    avg_y = sum(p[1] for p in positions_list) / len(positions_list)

                    # Check if high-connection tables are near center
                    high_conn_tables = [name for name, count in connection_counts.items() if count >= 3]
                    if high_conn_tables:
                        center_distances = []
                        for name in high_conn_tables:
                            if name in node_positions:
                                x, y = node_positions[name]
                                dist = math.sqrt((x - avg_x) ** 2 + (y - avg_y) ** 2)
                                center_distances.append(dist)

                        if center_distances:
                            avg_center_dist = sum(center_distances) / len(center_distances)
                            # Ideal: facts within 300px of center
                            if avg_center_dist < 200:
                                result['stratification_score'] = 100
                            elif avg_center_dist < 400:
                                result['stratification_score'] = 70
                            elif avg_center_dist < 600:
                                result['stratification_score'] = 50
                            else:
                                result['stratification_score'] = 30

            # Combined relationship score (weighted)
            result['relationship_score'] = (
                result['clustering_score'] * 0.5 +
                result['stratification_score'] * 0.3 +
                (100 - result['orphan_penalty']) * 0.2
            )

        except Exception as e:
            import traceback
            self.logger.error(f"Error analyzing relationship quality: {str(e)}")
            self.logger.error(f"Traceback: {traceback.format_exc()}")

        return result

    # =============================================================================
    # BRIDGE METHODS FOR MCP COMPATIBILITY
    # =============================================================================
    
    def _update_semantic_model_context(self, pbip_folder: str):
        """Update semantic model context for advanced components"""
        # Update the inherited properties to work with the new structure
        self.pbip_folder = Path(pbip_folder) if pbip_folder else None
        if self.pbip_folder:
            self.semantic_model_path = self._find_semantic_model_path()
    
    def find_semantic_model_path(self, pbip_folder: str) -> Optional[Path]:
        """Bridge method for advanced components - find semantic model path"""
        validation = self.validate_pbip_folder(pbip_folder)
        if validation['valid']:
            return Path(validation['semantic_model_path'])
        return None
    
    def analyze_relationships(self, pbip_folder: str) -> Dict[str, Any]:
        """Bridge method for relationship analyzer - analyze table relationships"""
        # Update context first
        self._update_semantic_model_context(pbip_folder)

        if self.relationship_analyzer:
            return self.relationship_analyzer.build_relationship_graph()
        else:
            # Fallback: return empty relationships
            return {}

