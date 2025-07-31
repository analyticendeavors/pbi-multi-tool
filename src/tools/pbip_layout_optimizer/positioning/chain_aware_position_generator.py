# tools/pbip_layout_optimizer/positioning/chain_aware_position_generator.py
"""
Chain-Aware Position Generator
Generates positions that directly honor Universal Chain Alignment optimization
"""

import logging
from typing import Dict, List, Tuple, Set, Any, Optional

logger = logging.getLogger("pbip-tools-mcp")


class ChainAwarePositionGenerator:
    """Generates positions that respect Universal Chain Alignment optimization"""
    
    def __init__(self, spacing_config: Dict[str, Any]):
        self.spacing_config = spacing_config
        
    def generate_chain_aligned_positions(self, aligned_stacks: Dict[str, List[str]], 
                                       positions_map: Dict[str, int], 
                                       connections: Dict[str, set],
                                       calendar_connected_specials: List[str],
                                       calendar_tables: List[str], 
                                       metrics_tables: List[str],
                                       parameter_tables: List[str],
                                       disconnected_tables: List[str],
                                       calculation_groups: List[str]) -> List[Dict[str, Any]]:
        """Generate positions that directly honor chain alignment optimization"""
        
        logger.info("Generating chain-aligned positions...")
        
        positions = []
        
        # Stack to X-position mapping
        stack_x_mapping = {
            'left_l4_dimensions': 'l4_plus_left',
            'left_l3_dimensions': 'l3_left', 
            'left_l2_dimensions': 'l2_left',
            'left_l1_dimensions': 'l1_left',
            'fact_tables': 'center',
            'right_l1_dimensions': 'l1_right',
            'right_l2_dimensions': 'l2_right',
            'right_l3_dimensions': 'l3_right',
            'right_l4_dimensions': 'l4_plus_right'
        }
        
        # Process all aligned stacks in order
        for stack_name, tables in aligned_stacks.items():
            if not tables:
                continue
                
            x_key = stack_x_mapping.get(stack_name)
            if not x_key or x_key not in positions_map:
                logger.warning(f"No X position found for stack {stack_name} (key: {x_key})")
                continue
                
            x_pos = positions_map[x_key]
            
            # Generate positions respecting chain alignment order
            if stack_name == 'fact_tables':
                # Facts go in center with calendar consideration
                stack_positions = self._generate_center_positions(
                    tables, x_pos, connections, calendar_connected_specials, calendar_tables)
            else:
                # Dimension stacks use chain-aligned positioning
                stack_positions = self._generate_dimension_stack_positions(
                    tables, x_pos, stack_name, connections)
            
            positions.extend(stack_positions)
            logger.info(f"Positioned {stack_name}: {len(tables)} tables with chain alignment")
        
        # Add metrics, parameters, etc. (unchanged positioning)
        additional_positions = self._add_additional_table_positions(
            positions_map, metrics_tables, parameter_tables, disconnected_tables, 
            calculation_groups, connections)
        positions.extend(additional_positions)
        
        logger.info(f"Chain-aligned positioning complete: {len(positions)} total positions")
        return positions
    
    def _generate_dimension_stack_positions(self, tables: List[str], x_position: int, 
                                          stack_name: str, connections: Dict[str, set]) -> List[Dict[str, Any]]:
        """Generate positions for dimension stacks RESPECTING RESERVED POSITIONS from chain alignment"""
        
        positions = []
        current_y = 150  # Standard start position for non-reserved tables
        reserved_positions_used = set()  # Track which Y positions are reserved
        
        logger.info(f"Processing {stack_name}: {len(tables)} tables (checking for reserved positions)")
        
        # Check if this generator has access to reserved position data
        alignment_engine = getattr(self, '_alignment_engine', None)
        reserved_positions = set()
        alignment_map = {}
        
        if alignment_engine and hasattr(alignment_engine, '_reserved_positions'):
            reserved_positions = alignment_engine._reserved_positions
            # Try to get alignment map from universal chain alignment
            if hasattr(alignment_engine, '_current_alignment_map'):
                alignment_map = alignment_engine._current_alignment_map
        
        logger.debug(f"Found {len(reserved_positions)} reserved positions")
        logger.debug(f"Alignment map contains {len(alignment_map)} locked tables")
        
        # Phase 1: Position tables with RESERVED positions first (preserve chain alignment)
        reserved_tables = []
        non_reserved_tables = []
        
        for table in tables:
            if table in alignment_map:
                reserved_tables.append((table, alignment_map[table]))
                logger.debug(f"Reserved: {table} → locked to position {alignment_map[table]}")
            else:
                non_reserved_tables.append(table)
                logger.debug(f"Non-reserved: {table} → will use gap filling")
        
        # Sort reserved tables by their locked position
        reserved_tables.sort(key=lambda x: x[1])
        
        # Position reserved tables at their EXACT locked positions (with proper spacing)
        for table, locked_position in reserved_tables:
            # Convert logical position to pixel Y coordinate with PROPER SPACING
            locked_y = 150 + (locked_position * (self.spacing_config['expanded_height'] + self.spacing_config['within_stack_spacing']))
            
            table_height = self.spacing_config['expanded_height']
            
            positions.append({
                'nodeIndex': table,
                'location': {'x': x_position, 'y': locked_y},
                'size': {'width': self.spacing_config['table_width'], 'height': table_height},
                'zIndex': len(positions),
                'category': stack_name.replace('_dimensions', '_dimension'),
                'collapsed': False,
                'connection_count': len(connections.get(table, set()))
            })
            
            reserved_positions_used.add(locked_position)
            logger.debug(f"Chain aligned: {table} at Y={locked_y} (position {locked_position})")
        
        # Phase 2: Position non-reserved tables in available gaps
        gap_position = 0
        for table in non_reserved_tables:
            # Find next available gap position
            while gap_position in reserved_positions_used:
                gap_position += 1
            
            gap_y = 150 + (gap_position * (self.spacing_config['expanded_height'] + self.spacing_config['within_stack_spacing']))
            table_height = self.spacing_config['expanded_height']
            
            positions.append({
                'nodeIndex': table,
                'location': {'x': x_position, 'y': gap_y},
                'size': {'width': self.spacing_config['table_width'], 'height': table_height},
                'zIndex': len(positions),
                'category': stack_name.replace('_dimensions', '_dimension'),
                'collapsed': False,
                'connection_count': len(connections.get(table, set()))
            })
            
            reserved_positions_used.add(gap_position)
            gap_position += 1
            logger.debug(f"Gap positioned: {table} at Y={gap_y} (gap position {gap_position-1})")
        
        return positions
    
    def _generate_center_positions(self, fact_tables: List[str], center_x: int, 
                                 connections: Dict[str, set], calendar_connected_specials: List[str],
                                 calendar_tables: List[str]) -> List[Dict[str, Any]]:
        """Generate center positions for facts with calendar consideration - FIXED: Calendar specials ABOVE calendar"""
        
        positions = []
        center_y = 50  # Start at the top
        
        logger.info("Calendar positioning - Specials above calendar")
        logger.debug(f"Calendar-connected specials: {calendar_connected_specials}")
        logger.debug(f"Calendar tables: {calendar_tables}")
        
        # FIXED: Calendar-connected specials FIRST (above calendar)
        if calendar_connected_specials:
            logger.info(f"Positioning {len(calendar_connected_specials)} calendar specials above calendar")
            for i, table in enumerate(calendar_connected_specials):
                # Use calendar height for visual consistency
                special_height = self.spacing_config['calendar_table_height']
                
                positions.append({
                    'nodeIndex': table,
                    'location': {'x': center_x, 'y': center_y},
                    'size': {'width': self.spacing_config['table_width'], 'height': special_height},
                    'zIndex': len(positions),
                    'category': 'calendar_connected_special',
                    'collapsed': True,  # Collapsed to match calendar appearance
                    'connection_count': len(connections.get(table, set()))
                })
                
                logger.debug(f"Calendar special above: '{table}' at Y={center_y} (position {i})")
                center_y += special_height + self.spacing_config['within_stack_spacing']
            
            # Add extra spacing between specials and calendar
            center_y += 20
        
        # THEN: Calendar tables below the specials
        if calendar_tables:
            logger.info(f"Positioning {len(calendar_tables)} calendar tables below specials")
            for i, table in enumerate(calendar_tables):
                positions.append({
                    'nodeIndex': table,
                    'location': {'x': center_x, 'y': center_y},
                    'size': {'width': self.spacing_config['table_width'], 'height': self.spacing_config['calendar_table_height']},
                    'zIndex': len(positions),
                    'category': 'calendar',
                    'collapsed': True,
                    'connection_count': len(connections.get(table, set()))
                })
                
                logger.debug(f"Calendar: '{table}' at Y={center_y} (position {i})")
                center_y += self.spacing_config['calendar_table_height'] + self.spacing_config['within_stack_spacing']

            # Add spacing after calendar before facts
            center_y += self.spacing_config['calendar_spacing']
        else:
            # No calendar tables, start facts lower
            center_y = 150

        # FINALLY: Facts below calendar section
        if fact_tables:
            logger.info(f"Positioning {len(fact_tables)} facts below calendar section")
            for i, table in enumerate(fact_tables):
                table_height = self.spacing_config['expanded_height']
                
                positions.append({
                    'nodeIndex': table,
                    'location': {'x': center_x, 'y': center_y},
                    'size': {'width': self.spacing_config['table_width'], 'height': table_height},
                    'zIndex': len(positions),
                    'category': 'fact',
                    'collapsed': False,
                    'connection_count': len(connections.get(table, set()))
                })
                
                logger.debug(f"Fact: '{table}' at Y={center_y} (position {i})")
                center_y += table_height + self.spacing_config['within_stack_spacing']
        
        logger.info(f"Center positioning complete: {len(positions)} tables arranged vertically")
        return positions
    
    def _add_additional_table_positions(self, positions_map: Dict[str, int],
                                      metrics_tables: List[str], parameter_tables: List[str],
                                      disconnected_tables: List[str], calculation_groups: List[str],
                                      connections: Dict[str, set]) -> List[Dict[str, Any]]:
        """Add positions for metrics, parameters, etc. (unchanged logic)"""
        
        positions = []
        
        # Position Metrics
        if metrics_tables and 'metrics' in positions_map:
            current_x = positions_map['metrics']
            current_y = 150
            
            for table in metrics_tables:
                table_height = self.spacing_config['expanded_height']
                
                positions.append({
                    'nodeIndex': table,
                    'location': {'x': current_x, 'y': current_y},
                    'size': {'width': self.spacing_config['table_width'], 'height': table_height},
                    'zIndex': len(positions),
                    'category': 'metrics',
                    'collapsed': False,
                    'connection_count': len(connections.get(table, set()))
                })
                
                current_y += table_height + self.spacing_config['within_stack_spacing']
        
        # Position Parameters, Disconnected Tables, and Calc Groups in Grid
        if (parameter_tables or disconnected_tables or calculation_groups) and 'parameters' in positions_map:
            param_start_x = positions_map['parameters']
            param_start_y = 50
            
            # Combine all unconnected tables
            all_unconnected_tables = parameter_tables + disconnected_tables + calculation_groups
            
            if all_unconnected_tables:
                grid_positions = self._generate_parameter_grid_positions(
                    all_unconnected_tables, param_start_x, param_start_y, connections)
                positions.extend(grid_positions)
        
        return positions
    
    def _generate_parameter_grid_positions(self, tables: List[str], start_x: int, start_y: int, 
                                         connections: Dict[str, set]) -> List[Dict[str, Any]]:
        """Generate grid positions for parameter/disconnected tables"""
        
        positions = []
        total_count = len(tables)
        
        # Calculate optimal grid layout
        if total_count <= 4:
            grid_cols = total_count
            col_spacing = 220
            row_spacing = 190
        elif total_count <= 8:
            grid_cols = 4
            col_spacing = 220
            row_spacing = 190
        else:
            grid_cols = 4
            col_spacing = 220
            row_spacing = 190
        
        # Position all tables in the grid
        for i, table in enumerate(tables):
            col = i % grid_cols
            row = i // grid_cols
            
            table_height = self.spacing_config['expanded_height']
            
            positions.append({
                'nodeIndex': table,
                'location': {
                    'x': start_x + col * col_spacing,
                    'y': start_y + row * row_spacing
                },
                'size': {'width': self.spacing_config['table_width'], 'height': table_height},
                'zIndex': len(positions),
                'category': 'parameter_grid',
                'collapsed': False,
                'connection_count': len(connections.get(table, set()))
            })
            
        return positions
