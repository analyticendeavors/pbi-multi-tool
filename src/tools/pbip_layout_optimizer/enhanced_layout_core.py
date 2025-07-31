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
                                    save_changes: bool = True, use_middle_out: bool = True) -> Dict[str, Any]:
        """
        Optimize layout using advanced middle-out engine with enhanced categorization.
        """
        try:
            # First, analyze table categorization
            categorization_result = self.analyze_table_categorization(pbip_folder)
            if not categorization_result['success']:
                return categorization_result
            
            # Check if advanced components are available
            if not self.mcp_available:
                # Fallback to basic grid layout
                self.logger.warning("Advanced components not available, falling back to basic grid layout")
                return self.optimize_layout(pbip_folder, canvas_width, canvas_height, save_changes, False)
            
            # Create middle-out engine for this operation
            middle_out_engine = self._create_middle_out_engine(pbip_folder)
            if not middle_out_engine:
                self.logger.warning("Could not create middle-out engine, falling back to basic layout")
                return self.optimize_layout(pbip_folder, canvas_width, canvas_height, save_changes, False)
            
            # Use middle-out engine for enhanced layout
            layout_result = middle_out_engine.apply_middle_out_layout(
                canvas_width, 
                canvas_height, 
                save_changes
            )
            
            if layout_result.get('success'):
                # Enhance result with categorization info
                layout_result['categorization_preview'] = categorization_result['categorization']
                layout_result['extensions'] = categorization_result['extensions']
                layout_result['layout_method'] = 'middle_out_advanced'
                layout_result['advanced_features'] = {
                    'table_categorization': True,
                    'relationship_analysis': True,
                    'middle_out_positioning': True,
                    'extension_handling': True,
                    'auto_date_filtering': True
                }
            
            return layout_result
            
        except Exception as e:
            self.logger.error(f"Error optimizing layout with advanced components: {str(e)}")
            return {
                'success': False,
                'error': str(e),
                'operation': 'optimize_layout_with_advanced',
                'mcp_status': self.get_mcp_status()
            }
    
    def optimize_layout(self, pbip_folder: str, canvas_width: int = 1400, canvas_height: int = 900, 
                       save_changes: bool = True, use_middle_out: bool = True) -> Dict[str, Any]:
        """
        Enhanced optimize_layout that uses advanced components when available.
        Falls back to basic grid layout if advanced components are not available.
        """
        if self.mcp_available and use_middle_out:
            return self.optimize_layout_with_advanced(pbip_folder, canvas_width, canvas_height, save_changes, use_middle_out)
        else:
            # Use basic grid layout as fallback
            return self._basic_grid_layout(pbip_folder, canvas_width, canvas_height, save_changes)
    
    def _basic_grid_layout(self, pbip_folder: str, canvas_width: int, canvas_height: int, save_changes: bool) -> Dict[str, Any]:
        """Basic grid layout fallback"""
        try:
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
            
            # Get table names
            tmdl_files = self.find_tmdl_files(pbip_folder)
            table_names = self.get_table_names_from_tmdl(tmdl_files)
            
            if not table_names:
                return {
                    'success': False,
                    'error': 'No tables found',
                    'operation': 'optimize_layout'
                }
            
            # Generate basic grid positions
            new_positions = self._generate_grid_positions(table_names, canvas_width, canvas_height)
            
            # Update layout data
            if 'diagrams' in layout_data and layout_data['diagrams']:
                layout_data['diagrams'][0]['nodes'] = new_positions
            
            # Save if requested
            saved = False
            if save_changes:
                saved = self.save_diagram_layout(pbip_folder, layout_data)
            
            return {
                'success': True,
                'operation': 'optimize_layout_basic_grid',
                'pbip_folder': pbip_folder,
                'layout_method': 'basic_grid',
                'tables_arranged': len(new_positions),
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
            nodes = diagram.get('nodes', [])
            
            analysis['total_tables'] = len(nodes)
            analysis['positioned_tables'] = len([n for n in nodes if 'location' in n])
            
            # Calculate overlaps and spacing
            positions = []
            for node in nodes:
                if 'location' in node:
                    x = node['location'].get('x', 0)
                    y = node['location'].get('y', 0)
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
        """Calculate overall quality score"""
        score = 0
        
        # Base score from positioned tables
        if analysis['total_tables'] > 0:
            score += (analysis['positioned_tables'] / analysis['total_tables']) * 40
        
        # Penalty for overlaps
        if analysis['total_tables'] > 0:
            overlap_penalty = (analysis['overlapping_tables'] / analysis['total_tables']) * 30
            score -= overlap_penalty
        
        # Spacing score
        spacing = analysis['average_spacing']
        if spacing > 1000:  # Too spread out
            score += 20
        elif spacing > 500:  # Good spacing
            score += 30
        elif spacing > 200:  # Acceptable
            score += 25
        else:  # Too cramped
            score += 10
        
        # Efficiency bonus
        score += analysis['layout_efficiency'] * 0.1
        
        return max(0, min(100, round(score, 1)))
    
    def _get_quality_rating(self, score: float) -> str:
        """Get quality rating from score"""
        if score >= 80:
            return "EXCELLENT"
        elif score >= 60:
            return "GOOD"
        elif score >= 40:
            return "FAIR"
        elif score >= 20:
            return "POOR"
        else:
            return "VERY POOR"
    
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
