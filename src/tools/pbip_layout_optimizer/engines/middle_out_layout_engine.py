# tools/pbip_layout_optimizer/engines/middle_out_layout_engine.py
"""
Middle-Out Layout Engine for PBIP model diagrams
Universal dynamic layout with modular components and consistent spacing
"""

import json
import math
import logging
from typing import Dict, List, Any, Optional
from pathlib import Path

logger = logging.getLogger("pbip-tools-mcp")

# Import our modular components with correct paths for our tool structure
POSITIONING_AVAILABLE = False  # Initialize at module level

try:
    # Use relative imports within our tool structure
    from ..analyzers.relationship_analyzer import RelationshipAnalyzer
    from ..analyzers.table_categorizer import TableCategorizer
    
    # Import positioning components
    try:
        from ..positioning.position_calculator import PositionCalculator
        from ..positioning.dimension_optimizer import DimensionOptimizer
        from ..positioning.universal_chain_alignment import UniversalChainAlignment
        from ..positioning.chain_aware_position_generator import ChainAwarePositionGenerator
        from ..positioning.family_aware_alignment import enhance_alignment_with_family_grouping
        logger.info("✅ All modular components imported successfully (WITH CHAIN-AWARE POSITIONING AND FAMILY GROUPING)")
        POSITIONING_AVAILABLE = True
    except ImportError as pos_e:
        logger.error(f"❌ Positioning components import failed: {pos_e}")
        POSITIONING_AVAILABLE = False
        
except ImportError as e:
    logger.error(f"❌ Core components import error: {e}")
    raise ImportError(f"Could not import core components: {e}")


class MiddleOutLayoutEngine:
    """Universal Middle-Out Layout Engine with modular architecture"""
    
    def __init__(self, pbip_folder: str, base_engine):
        self.pbip_folder = Path(pbip_folder)
        self.base_engine = base_engine  # Our enhanced layout core
        
        # Layout configuration - Universal consistent spacing
        self.spacing_config = {
            'table_width': 200,
            'base_collapsed_height': 104,
            'height_per_relationship': 24,
            'expanded_height': 180,
            'calendar_table_height': 140,
            'within_stack_spacing': 15,  # TIGHTER: Revert to closer spacing within stacks
            'table_grid_height': 210,
            'stack_gap': 150,  # CONSISTENT: 150px between ALL stacks
            'calendar_spacing': 80,
            'left_margin': 50
        }
        
        # Initialize modular components if positioning is available
        if POSITIONING_AVAILABLE:
            try:
                self.position_calculator = PositionCalculator(self.spacing_config)
                self.dimension_optimizer = DimensionOptimizer(base_engine.table_categorizer)
                self.universal_chain_alignment = UniversalChainAlignment(self.spacing_config)
                self.chain_aware_position_generator = ChainAwarePositionGenerator(self.spacing_config)
                
                # 🔗 CRITICAL: Pass alignment engine reference for reserved position access
                self.chain_aware_position_generator._alignment_engine = self.universal_chain_alignment
                
                logger.info("✅ Advanced positioning components initialized")
            except Exception as e:
                logger.warning(f"⚠️ Error initializing positioning components: {e}")
                raise Exception(f"Failed to initialize positioning components: {e}")
        else:
            raise Exception("Positioning components not available - cannot create full middle-out engine")
        
        # User preferences
        self.collapse_connected_tables = True
        self.uniform_table_heights = True  # Universal: consistent heights
        
    def generate_middle_out_layout(self, canvas_width: int = 1400, canvas_height: int = 900) -> List[Dict[str, Any]]:
        """Generate layout positions with universal middle-out philosophy"""
        
        logger.info("Initializing universal chain alignment engine")
        
        table_names = self.base_engine._get_table_names_from_tmdl()
        if not table_names:
            return []
            
        # Build relationship graph using base engine's components
        connections = self.base_engine.relationship_analyzer.build_relationship_graph()
        
        # Categorize tables using dynamic analysis
        logger.info("Analyzing table relationships and categorizing tables...")
        categorized = self.base_engine.table_categorizer.categorize_tables(table_names, connections)
        logger.info(f"Table categorization complete: {len(categorized.get('fact_tables', []))} facts, {len(categorized.get('l1_dimensions', []))} L1 dimensions")
        
        # Extract categories AFTER 1:1 extension processing
        fact_tables = categorized['fact_tables']
        l1_dimensions = categorized['l1_dimensions']
        l2_dimensions = categorized['l2_dimensions']
        l3_dimensions = categorized['l3_dimensions']
        l4_plus_dimensions = categorized['l4_plus_dimensions']
        calendar_tables = categorized['calendar_tables']
        metrics_tables = categorized['metrics_tables']
        parameter_tables = categorized['parameter_tables']
        calculation_groups = categorized['calculation_groups']
        disconnected_tables = categorized['disconnected_tables']
        
        logger.info(f"Layout distribution: L1({len(l1_dimensions)}) L2({len(l2_dimensions)}) L3({len(l3_dimensions)}) L4+({len(l4_plus_dimensions)}) Facts({len(fact_tables)})")
        
        # Optimize dimension placement with ENHANCED calendar awareness
        left_l1_dimensions, right_l1_dimensions = self.dimension_optimizer.optimize_dimension_placement(
            l1_dimensions, fact_tables, connections, calendar_tables)
            
        # ENHANCED: Identify calendar-connected specials FIRST (before L2 detection)
        calendar_connected_specials = self.base_engine.table_categorizer.identify_calendar_connected_specials(
            table_names, calendar_tables, connections, set())
            
        # CRITICAL FIX: Remove calendar specials from ALL categories (including facts AND calendar_tables)
        fact_tables = [t for t in fact_tables if t not in calendar_connected_specials]
        calendar_tables = [t for t in calendar_tables if t not in calendar_connected_specials]  # FIX THE DUPLICATE!
        left_l1_dimensions = [t for t in left_l1_dimensions if t not in calendar_connected_specials]
        right_l1_dimensions = [t for t in right_l1_dimensions if t not in calendar_connected_specials]
        
        # ENHANCED: Identify additional L2 tables based on single connections to L1s
        remaining_tables = [t for t in table_names if t not in categorized['fact_tables'] and 
                           t not in categorized['calendar_tables'] and 
                           t not in categorized['metrics_tables'] and 
                           t not in categorized['parameter_tables'] and 
                           t not in categorized['calculation_groups'] and 
                           t not in categorized['disconnected_tables'] and 
                           t not in calendar_connected_specials]
        
        # Add existing L2s and identify additional ones
        all_l1_dimensions = left_l1_dimensions + right_l1_dimensions
        additional_l2s = self.dimension_optimizer.identify_additional_l2_tables(
            remaining_tables, all_l1_dimensions, connections, set(l2_dimensions))
        
        # Add additional L2s to the L2 list
        l2_dimensions.extend(additional_l2s)
        
        # Remove additional L2s from L1 lists
        left_l1_dimensions = [t for t in left_l1_dimensions if t not in additional_l2s]
        right_l1_dimensions = [t for t in right_l1_dimensions if t not in additional_l2s]
        
        # RELATIONSHIP-FIRST: Place L2 dimensions based on actual TMDL connections
        left_l2_dimensions, right_l2_dimensions = self.dimension_optimizer.place_l2_dimensions_near_l1(
            l2_dimensions, left_l1_dimensions, right_l1_dimensions, connections)
            
        # RELATIONSHIP-FIRST: Place L3 dimensions based on their L2 connections
        left_l3_dimensions, right_l3_dimensions = self.dimension_optimizer.place_l3_dimensions_near_l2(
            l3_dimensions, left_l2_dimensions, right_l2_dimensions, connections)
            
        # RELATIONSHIP-FIRST: Place L4+ dimensions based on their L3 connections (or L2 fallback)
        left_l4_dimensions, right_l4_dimensions = self.dimension_optimizer.place_l4_dimensions_near_l3(
            l4_plus_dimensions, left_l3_dimensions, right_l3_dimensions, left_l2_dimensions, right_l2_dimensions, connections)
            
        # Extension repositioning phase
        logger.info("Processing dimension extensions...")
        confirmed_extensions = categorized.get('dimension_extensions', {})
        if confirmed_extensions:
            try:
                self._reposition_extensions(
                    confirmed_extensions,
                    left_l1_dimensions, right_l1_dimensions,
                    left_l2_dimensions, right_l2_dimensions,
                    left_l3_dimensions, right_l3_dimensions,
                    left_l4_dimensions, right_l4_dimensions
                )
                logger.info(f"Repositioned {len(confirmed_extensions)} extension tables")
            except Exception as e:
                logger.error(f"Error in extension repositioning: {e}")
        
        # ENHANCED: Apply opposite-side placement for 1:1 relationships
        categorized_for_placement = {
            'l1_dimension_left': left_l1_dimensions,
            'l1_dimension_right': right_l1_dimensions,
            'l2_dimension_left': left_l2_dimensions,
            'l2_dimension_right': right_l2_dimensions,
            'l3_dimension_left': left_l3_dimensions,
            'l3_dimension_right': right_l3_dimensions,
            'l4_plus_dimension_left': left_l4_dimensions,
            'l4_plus_dimension_right': right_l4_dimensions
        }
        
        categorized_for_placement = self.dimension_optimizer.apply_opposite_side_placement(categorized_for_placement, connections)
        
        # Update dimension lists after opposite-side placement
        left_l1_dimensions = categorized_for_placement.get('l1_dimension_left', [])
        right_l1_dimensions = categorized_for_placement.get('l1_dimension_right', [])
        left_l2_dimensions = categorized_for_placement.get('l2_dimension_left', [])
        right_l2_dimensions = categorized_for_placement.get('l2_dimension_right', [])
        left_l3_dimensions = categorized_for_placement.get('l3_dimension_left', [])
        right_l3_dimensions = categorized_for_placement.get('l3_dimension_right', [])
        left_l4_dimensions = categorized_for_placement.get('l4_plus_dimension_left', [])
        right_l4_dimensions = categorized_for_placement.get('l4_plus_dimension_right', [])
        
        # Apply universal chain alignment optimization
        logger.info("Optimizing table alignment and positioning...")
        all_dimension_stacks = {
            'left_l4_dimensions': left_l4_dimensions,
            'left_l3_dimensions': left_l3_dimensions,
            'left_l2_dimensions': left_l2_dimensions,
            'left_l1_dimensions': left_l1_dimensions,
            'fact_tables': fact_tables,
            'right_l1_dimensions': right_l1_dimensions,
            'right_l2_dimensions': right_l2_dimensions,
            'right_l3_dimensions': right_l3_dimensions,
            'right_l4_dimensions': right_l4_dimensions
        }
        
        # Apply universal chain alignment optimization
        aligned_stacks = self.universal_chain_alignment.optimize_universal_stack_alignment(
            all_dimension_stacks, connections)
        
        # Apply family-aware grouping for related tables
        extensions = categorized.get('dimension_extensions', {})
        if extensions:
            try:
                formatted_extensions = self._format_extensions_for_family_grouping(extensions)
                aligned_stacks = enhance_alignment_with_family_grouping(aligned_stacks, formatted_extensions, connections)
                logger.info(f"Applied family grouping for {len(extensions)} extension families")
            except Exception as e:
                logger.error(f"Error in family grouping: {e}")
        
        # Apply universal chain family grouping
        aligned_stacks = self._apply_universal_chain_family_grouping(aligned_stacks, connections)
        logger.info("Chain family grouping complete")
        
        # Update dimension lists with optimized alignment
        left_l4_dimensions = aligned_stacks.get('left_l4_dimensions', [])
        left_l3_dimensions = aligned_stacks.get('left_l3_dimensions', [])
        left_l2_dimensions = aligned_stacks.get('left_l2_dimensions', [])
        left_l1_dimensions = aligned_stacks.get('left_l1_dimensions', [])
        fact_tables = aligned_stacks.get('fact_tables', [])
        right_l1_dimensions = aligned_stacks.get('right_l1_dimensions', [])
        right_l2_dimensions = aligned_stacks.get('right_l2_dimensions', [])
        right_l3_dimensions = aligned_stacks.get('right_l3_dimensions', [])
        right_l4_dimensions = aligned_stacks.get('right_l4_dimensions', [])
        
        logger.info(f"Chain alignment optimization complete: processed {len(aligned_stacks)} table stacks")
            
        # Update categorized for dynamic positioning with SPLIT LISTS
        categorized_with_splits = categorized.copy()
        categorized_with_splits.update({
            'l1_dimensions_left': left_l1_dimensions,
            'l1_dimensions_right': right_l1_dimensions,
            'l2_dimensions_left': left_l2_dimensions,
            'l2_dimensions_right': right_l2_dimensions,
            'l3_dimensions_left': left_l3_dimensions,
            'l3_dimensions_right': right_l3_dimensions,
            'l4_plus_dimensions_left': left_l4_dimensions,
            'l4_plus_dimensions_right': right_l4_dimensions
        })
        
        # Calculate X positions with SPLIT-AWARE spacing (using optimized aligned stacks)
        positions_map = self.position_calculator.calculate_canvas_positions(categorized_with_splits, canvas_width)
        
        # Generate final positions using chain-aware positioning
        logger.info("Generating optimized table positions...")
        
        positions = self.chain_aware_position_generator.generate_chain_aligned_positions(
            aligned_stacks, positions_map, connections, calendar_connected_specials,
            calendar_tables, metrics_tables, parameter_tables, disconnected_tables, calculation_groups
        )
        
        logger.info(f"Position generation complete: {len(positions)} tables positioned")
        logger.info(f"Layout complete: {len(positions)} tables arranged with {len(fact_tables)} fact tables in center")
        
        return positions
        
    def _reposition_extensions(self, extensions: Dict[str, Any], 
                              left_l1: List[str], right_l1: List[str],
                              left_l2: List[str], right_l2: List[str],
                              left_l3: List[str], right_l3: List[str],
                              left_l4: List[str], right_l4: List[str]) -> None:
        """🚀 NEW: Reposition extension tables one level further from facts than their base table"""
        
        for extension_table, extension_info in extensions.items():
            try:
                # Handle both tuple and dict formats for backward compatibility
                if isinstance(extension_info, tuple):
                    base_table = extension_info[0] if len(extension_info) > 0 else None
                    rel_type = extension_info[1] if len(extension_info) > 1 else 'extension'
                elif isinstance(extension_info, dict):
                    base_table = extension_info.get('base_table')
                    rel_type = extension_info.get('type', 'extension')
                else:
                    logger.warning(f"Unknown extension_info format for {extension_table}: {type(extension_info)}")
                    continue
                
                if rel_type != 'extension' or not base_table:
                    continue
                    
                logger.info(f"Repositioning extension: {extension_table} (base: {base_table})")
                
                # Find base table's current category and side
                base_category, base_side = self._find_table_category_and_side(
                    base_table, left_l1, right_l1, left_l2, right_l2, left_l3, right_l3, left_l4, right_l4)
                
                if not base_category:
                    logger.warning(f"Base table {base_table} not found in any category")
                    continue
                    
                # Calculate target category (one level further from facts)
                target_category = self._calculate_target_category(base_category)
                
                if not target_category:
                    logger.warning(f"Could not calculate target category for base {base_category}")
                    continue
                
                # Remove extension from all current positions
                self._remove_table_from_all_categories(
                    extension_table, left_l1, right_l1, left_l2, right_l2, left_l3, right_l3, left_l4, right_l4)
                
                # Add to target category
                self._add_table_to_category(
                    extension_table, target_category, base_side,
                    left_l1, right_l1, left_l2, right_l2, left_l3, right_l3, left_l4, right_l4)
                
                logger.info(f"Extension repositioned: {extension_table} moved from {base_category}_{base_side} → {target_category}_{base_side}")
                
            except Exception as e:
                logger.error(f"Error repositioning extension {extension_table}: {e}")
                continue
    
    def _find_table_category_and_side(self, table_name: str,
                                     left_l1: List[str], right_l1: List[str],
                                     left_l2: List[str], right_l2: List[str],
                                     left_l3: List[str], right_l3: List[str],
                                     left_l4: List[str], right_l4: List[str]) -> tuple:
        """Find which category and side a table belongs to"""
        
        if table_name in left_l1:
            return 'l1', 'left'
        elif table_name in right_l1:
            return 'l1', 'right'
        elif table_name in left_l2:
            return 'l2', 'left'
        elif table_name in right_l2:
            return 'l2', 'right'
        elif table_name in left_l3:
            return 'l3', 'left'
        elif table_name in right_l3:
            return 'l3', 'right'
        elif table_name in left_l4:
            return 'l4', 'left'
        elif table_name in right_l4:
            return 'l4', 'right'
        else:
            return None, None
    
    def _calculate_target_category(self, base_category: str) -> str:
        """Calculate target category (one level further from facts)"""
        
        if base_category == 'l1':
            return 'l2'
        elif base_category == 'l2':
            return 'l3'
        elif base_category == 'l3':
            return 'l4'
        elif base_category == 'l4':
            return 'l4'  # Already at maximum distance
        else:
            return None
    
    def _remove_table_from_all_categories(self, table_name: str,
                                         left_l1: List[str], right_l1: List[str],
                                         left_l2: List[str], right_l2: List[str],
                                         left_l3: List[str], right_l3: List[str],
                                         left_l4: List[str], right_l4: List[str]) -> None:
        """Remove table from all dimension categories"""
        
        for dimension_list in [left_l1, right_l1, left_l2, right_l2, left_l3, right_l3, left_l4, right_l4]:
            if table_name in dimension_list:
                dimension_list.remove(table_name)
                logger.debug(f"Removed {table_name} from dimension list")
    
    def _add_table_to_category(self, table_name: str, category: str, side: str,
                              left_l1: List[str], right_l1: List[str],
                              left_l2: List[str], right_l2: List[str],
                              left_l3: List[str], right_l3: List[str],
                              left_l4: List[str], right_l4: List[str]) -> None:
        """Add table to specified category and side"""
        
        if category == 'l1' and side == 'left':
            left_l1.append(table_name)
        elif category == 'l1' and side == 'right':
            right_l1.append(table_name)
        elif category == 'l2' and side == 'left':
            left_l2.append(table_name)
        elif category == 'l2' and side == 'right':
            right_l2.append(table_name)
        elif category == 'l3' and side == 'left':
            left_l3.append(table_name)
        elif category == 'l3' and side == 'right':
            right_l3.append(table_name)
        elif category == 'l4' and side == 'left':
            left_l4.append(table_name)
        elif category == 'l4' and side == 'right':
            right_l4.append(table_name)
        else:
            logger.warning(f"Unknown category/side combination: {category}_{side}")
            return
            
        logger.debug(f"Added {table_name} to {category}_{side}")
        
    def _apply_universal_chain_family_grouping(self, aligned_stacks: Dict[str, List[str]], 
                                              connections: Dict[str, set]) -> Dict[str, List[str]]:
        """Universal chain family grouping that keeps related table families together within stacks"""
        
        logger.info("Starting universal chain family grouping")
        
        # Identify chain families dynamically based on naming patterns and relationships
        chain_families = self._identify_chain_families(aligned_stacks, connections)
        
        if not chain_families:
            logger.info("🏠 No chain families detected, skipping grouping")
            return aligned_stacks
        
        logger.info(f"Found {len(chain_families)} chain families to group")
        
        # Apply family grouping within each stack
        for stack_name, stack_tables in aligned_stacks.items():
            if not stack_tables:
                continue
                
            logger.info(f"Processing stack: {stack_name} with {len(stack_tables)} tables")
            
            # Group tables by their family membership
            grouped_tables = self._group_tables_by_family(stack_tables, chain_families)
            
            # Reorganize stack to keep families together
            reorganized_stack = self._reorganize_stack_by_families(grouped_tables)
            
            # Update the stack
            aligned_stacks[stack_name] = reorganized_stack
            
            logger.info(f"Reorganized {stack_name}: {reorganized_stack}")
        
        logger.info("Universal chain family grouping complete")
        return aligned_stacks
    
    def _identify_chain_families(self, aligned_stacks: Dict[str, List[str]], 
                               connections: Dict[str, set]) -> Dict[str, List[str]]:
        """Identify chain families based on naming patterns and relationships"""
        
        all_tables = []
        for stack_tables in aligned_stacks.values():
            all_tables.extend(stack_tables)
        
        families = {}
        processed_tables = set()
        
        for table in all_tables:
            if table in processed_tables:
                continue
                
            # Extract potential family base name
            family_base = self._extract_family_base_name(table)
            
            if not family_base:
                continue
                
            # Find all related tables with same family base
            family_members = []
            for candidate in all_tables:
                if candidate in processed_tables:
                    continue
                    
                candidate_base = self._extract_family_base_name(candidate)
                
                # Check if they share the same base name or are connected
                if (candidate_base == family_base or 
                    self._are_tables_in_same_family(table, candidate, connections)):
                    family_members.append(candidate)
                    processed_tables.add(candidate)
            
            if len(family_members) > 1:  # Only consider actual families (2+ members)
                families[family_base] = family_members
                logger.info(f"Identified chain family '{family_base}': {family_members}")
        
        return families
    
    def _extract_family_base_name(self, table_name: str) -> str:
        """Extract the base family name from a table name"""
        
        # Remove common prefixes
        clean_name = table_name
        prefixes = ['Dim_', 'Fact_', 'dim_', 'fact_']
        for prefix in prefixes:
            if clean_name.startswith(prefix):
                clean_name = clean_name[len(prefix):]
                break
        
        # Look for common base patterns
        # Example: Property, PropertyAttributes, PropertyList -> Property
        # Example: Tenant, TenantProgramType -> Tenant
        
        # Find the longest common prefix with other words
        words = clean_name.split('_')
        if len(words) > 1:
            return words[0]  # First word is often the base
        
        # Check for camelCase patterns
        import re
        camel_parts = re.findall(r'[A-Z][a-z]*', clean_name)
        if len(camel_parts) > 1:
            return camel_parts[0]  # First camelCase part
        
        # For simple names, use the first part before any numbers or special suffixes
        base_match = re.match(r'^([A-Za-z]+)', clean_name)
        if base_match:
            return base_match.group(1)
        
        return clean_name
    
    def _are_tables_in_same_family(self, table1: str, table2: str, connections: Dict[str, set]) -> bool:
        """Check if two tables belong to the same family based on relationships"""
        
        # Check direct connection
        if table2 in connections.get(table1, set()) or table1 in connections.get(table2, set()):
            return True
        
        # Check if they share connections to the same base table
        table1_connections = connections.get(table1, set())
        table2_connections = connections.get(table2, set())
        
        # If they both connect to a common table, they might be family
        common_connections = table1_connections.intersection(table2_connections)
        if common_connections:
            return True
        
        return False
    
    def _group_tables_by_family(self, stack_tables: List[str], 
                              chain_families: Dict[str, List[str]]) -> Dict[str, List[str]]:
        """Group tables in a stack by their family membership"""
        
        grouped = {'unassigned': []}
        
        for table in stack_tables:
            assigned = False
            
            # Find which family this table belongs to
            for family_name, family_members in chain_families.items():
                if table in family_members:
                    if family_name not in grouped:
                        grouped[family_name] = []
                    grouped[family_name].append(table)
                    assigned = True
                    break
            
            if not assigned:
                grouped['unassigned'].append(table)
        
        return grouped
    
    def _reorganize_stack_by_families(self, grouped_tables: Dict[str, List[str]]) -> List[str]:
        """Reorganize a stack to keep families together"""
        
        reorganized = []
        
        # First, add all family groups (keeping families together)
        for family_name, family_tables in grouped_tables.items():
            if family_name == 'unassigned':
                continue
            
            # Sort family members to ensure consistent ordering within family
            sorted_family = sorted(family_tables)
            reorganized.extend(sorted_family)
            
            logger.info(f"Grouped family '{family_name}' together: {sorted_family}")
        
        # Then add unassigned tables
        if 'unassigned' in grouped_tables:
            reorganized.extend(sorted(grouped_tables['unassigned']))
        
        return reorganized
        
    def _format_extensions_for_family_grouping(self, extensions: Dict[str, Any]) -> Dict[str, Dict[str, Any]]:
        """Convert extension format to match family grouping expectations"""
        
        formatted = {}
        
        for extension_table, extension_info in extensions.items():
            try:
                # Handle both tuple and dict formats
                if isinstance(extension_info, tuple):
                    base_table = extension_info[0] if len(extension_info) > 0 else None
                    rel_type = extension_info[1] if len(extension_info) > 1 else 'extension'
                elif isinstance(extension_info, dict):
                    base_table = extension_info.get('base_table')
                    rel_type = extension_info.get('type', 'extension')
                else:
                    logger.warning(f"Unknown extension format for {extension_table}: {type(extension_info)} - {extension_info}")
                    continue
                
                if base_table and rel_type == 'extension':
                    formatted[extension_table] = {
                        'base_table': base_table,
                        'type': rel_type
                    }
                    
            except Exception as e:
                logger.error(f"Error formatting extension {extension_table}: {e}")
                logger.error(f"Extension info: {extension_info}")
                continue
        
        logger.info(f"Formatted {len(formatted)} extensions for family grouping")
        return formatted
        
    def apply_middle_out_layout(self, canvas_width: int = 1400, canvas_height: int = 900, 
                              save_changes: bool = True, expand_all: bool = False) -> Dict[str, Any]:
        """Apply universal Haven's middle-out layout to the PBIP model"""
        try:
            if not self.base_engine.semantic_model_path:
                return {
                    'success': False,
                    'error': 'No .SemanticModel folder found',
                    'operation': 'apply_middle_out_layout'
                }
            
            # Generate new positions using universal layout
            new_positions = self.generate_middle_out_layout(canvas_width, canvas_height)
            
            if not new_positions:
                return {
                    'success': False,
                    'error': 'No tables found or failed to generate positions',
                    'operation': 'apply_middle_out_layout'
                }
            
            # Get current layout to preserve structure
            current_layout = self.base_engine.parse_diagram_layout(str(self.pbip_folder))
            
            # Prepare nodes
            enhanced_nodes = []
            for pos in new_positions:
                node = {
                    'location': pos['location'],
                    'nodeIndex': pos['nodeIndex'],
                    'size': pos['size'],
                    'zIndex': pos['zIndex']
                }
                enhanced_nodes.append(node)
            
            # Create new layout data
            if current_layout and 'diagrams' in current_layout:
                new_layout_data = current_layout.copy()
                if new_layout_data['diagrams']:
                    new_layout_data['diagrams'][0]['nodes'] = enhanced_nodes
                    new_layout_data['diagrams'][0]['hideKeyFieldsWhenCollapsed'] = False
            else:
                new_layout_data = {
                    "version": "1.1.0",
                    "diagrams": [
                        {
                            "ordinal": 0,
                            "scrollPosition": {"x": 0, "y": 0},
                            "nodes": enhanced_nodes,
                            "name": "All tables",
                            "zoomValue": 100,
                            "pinKeyFieldsToTop": False,
                            "showExtraHeaderInfo": False,
                            "hideKeyFieldsWhenCollapsed": False,
                            "tablesLocked": False
                        }
                    ],
                    "selectedDiagram": "All tables",
                    "defaultDiagram": "All tables"
                }
            
            # Save if requested
            saved = False
            if save_changes:
                saved = self.base_engine.save_diagram_layout(str(self.pbip_folder), new_layout_data)
            
            # Generate summary with updated categories
            category_summary = {}
            for pos in new_positions:
                category = pos.get('category', 'unknown')
                category_summary[category] = category_summary.get(category, 0) + 1
            
            # Update category names to reflect L1-L2 adjacency improvements
            if 'l1_l2_dimension_left' in category_summary:
                category_summary['l1_l2_combined_left'] = category_summary.pop('l1_l2_dimension_left')
            if 'l1_l2_dimension_right' in category_summary:
                category_summary['l1_l2_combined_right'] = category_summary.pop('l1_l2_dimension_right')
            
            fact_count = category_summary.get('fact', 0)
            
            return {
                'success': True,
                'operation': 'apply_middle_out_layout_universal_modular',
                'pbip_folder': str(self.pbip_folder),
                'semantic_model_path': str(self.base_engine.semantic_model_path),
                'changes_saved': saved,
                'canvas_size': {'width': canvas_width, 'height': canvas_height},
                'tables_arranged': len(new_positions),
                'layout_philosophy': 'middle_out_universal_dynamic_modular',
                'category_breakdown': category_summary,
                'facts_in_center': fact_count,
                'layout_features': {
                    'facts_properly_centered': True,
                    'improved_fact_detection': True,
                    'calendar_at_top': True,
                    'dimensions_left_right': True,
                    'l1_l2_adjacent_positioning': True,
                    'smart_dimension_pairing': True,
                    'snowflake_layout_support': True,
                    'uniform_table_heights': True,
                    'universal_spacing': f'{self.spacing_config["stack_gap"]}px consistent throughout',
                    'dynamic_time_period_positioning': True,
                    'no_hardcoded_table_names': True,
                    'relationship_aware_grouping': True,
                    'modular_architecture': True
                },
                'dynamic_improvements': {
                    'consistent_spacing': f'{self.spacing_config["stack_gap"]}px everywhere',
                    'pattern_based_time_detection': 'period, closed, fiscal, etc.',
                    'relationship_based_calendar_proximity': 'tables connected to calendar',
                    'automatic_connected_pair_grouping': 'adjacent positioning',
                    'l1_l2_pair_detection': 'automatically identifies L1-L2 relationships',
                    'adjacent_l1_l2_positioning': 'L1-L2 pairs placed immediately next to each other',
                    'smart_snowflake_distribution': 'L2/L3+ tables distributed on both sides when needed',
                    'relationship_aware_side_placement': 'L2s placed on same side as their L1 partners',
                    'modular_components': 'relationship_analyzer, table_categorizer, position_calculator, dimension_optimizer'
                },
                'architecture': {
                    'base_engine': 'EnhancedPBIPLayoutCore',
                    'analyzers': ['RelationshipAnalyzer', 'TableCategorizer'],
                    'positioning': ['PositionCalculator', 'DimensionOptimizer'],
                    'spacing_config': self.spacing_config
                },
                'preview_positions': new_positions[:5]
            }
            
        except Exception as e:
            logger.error(f"Error applying universal middle-out layout: {str(e)}")
            return {
                'success': False,
                'error': str(e),
                'operation': 'apply_middle_out_layout_universal_modular'
            }
