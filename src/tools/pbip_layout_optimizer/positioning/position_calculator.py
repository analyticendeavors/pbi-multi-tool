# tools/pbip_layout_optimizer/positioning/position_calculator.py
"""
Position Calculator for PBIP model diagrams
Calculates optimal X,Y positions for tables with consistent spacing
"""

import logging
import math
from typing import Dict, List, Tuple, Set

logger = logging.getLogger("pbip-tools-mcp")


class PositionCalculator:
    """Calculates optimal positions for tables with consistent spacing"""
    
    def __init__(self, spacing_config):
        self.spacing_config = spacing_config
        
    def calculate_canvas_positions(self, categorized: Dict[str, List[str]], canvas_width: int) -> Dict[str, int]:
        """DYNAMIC: Calculate X positions - only create columns when needed with SPLIT AWARENESS"""
        positions = {}
        
        # Force center column to always exist if we have calendar or facts
        has_center_content = bool(categorized.get('calendar_tables') or categorized.get('fact_tables'))
        
        # DYNAMIC: Determine which stacks are actually needed with SPLIT SUPPORT
        stacks = []
        
        # LEFT SIDE: Add columns only if data exists (furthest to closest)
        if (categorized.get('l4_plus_dimensions_left') or 
            categorized.get('l4_plus_dimensions')):
            stacks.append('l4_plus_left')
            
        if (categorized.get('l3_dimensions_left') or 
            categorized.get('l3_dimensions')):
            stacks.append('l3_left')
            
        if (categorized.get('l2_dimensions_left') or 
            categorized.get('l2_dimensions')):
            stacks.append('l2_left')
            
        if (categorized.get('l1_dimensions_left') or 
            categorized.get('l1_dimensions')):
            stacks.append('l1_left')
            
        # CENTER: Always include if we have calendar or facts
        if has_center_content:
            stacks.append('center')
            
        # RIGHT SIDE: Add columns only if data exists (closest to furthest)
        if (categorized.get('l1_dimensions_right') or 
            categorized.get('l1_dimensions')):
            stacks.append('l1_right')
            
        if (categorized.get('l2_dimensions_right') or 
            categorized.get('l2_dimensions')):
            stacks.append('l2_right')
            
        if (categorized.get('l3_dimensions_right') or 
            categorized.get('l3_dimensions')):
            stacks.append('l3_right')
            
        if (categorized.get('l4_plus_dimensions_right') or 
            categorized.get('l4_plus_dimensions')):
            stacks.append('l4_plus_right')
            
        # UTILITY SECTIONS: Add only if data exists
        if categorized.get('metrics_tables'):
            stacks.append('metrics')
            
        if categorized.get('parameter_tables') or categorized.get('disconnected_tables'):
            stacks.append('parameters')
        
        logger.info(f"DYNAMIC LAYOUT WITH SPLIT SUPPORT: {len(stacks)} columns needed: {stacks}")
        
        # Assign positions with OPTIMIZED spacing
        current_x = self.spacing_config['left_margin']
        
        for i, stack in enumerate(stacks):
            positions[stack] = current_x
            
            # ENHANCED: Smart spacing based on stack relationships
            next_stack = stacks[i + 1] if i + 1 < len(stacks) else None
            
            if stack in ['metrics', 'parameters']:
                # Compact spacing for metrics/parameters
                current_x += self.spacing_config['table_width'] + 100
            elif next_stack in ['metrics', 'parameters']:
                # Reduced gap before metrics/parameters (fix the large gap issue)
                current_x += self.spacing_config['table_width'] + 100
            else:
                # Standard gap for dimension stacks
                current_x += self.spacing_config['table_width'] + self.spacing_config['stack_gap']
        
        logger.info(f"DYNAMIC SPACING WITH SPLIT SUPPORT: Generated {len(stacks)} columns with positions:")
        for stack in stacks:
            logger.info(f"  {stack}: X={positions[stack]}")
        
        return positions
        
    def calculate_table_height(self, table_name: str, connections: Dict[str, Set[str]]) -> Tuple[int, bool]:
        """Calculate uniform table height"""
        # Universal approach: All tables expanded with uniform height
        return self.spacing_config['expanded_height'], False
        
    def generate_positions_for_stack(self, tables: List[str], x_position: int, start_y: int, 
                                   connections: Dict[str, Set[str]], category: str) -> List[Dict]:
        """Generate positions for tables in a vertical stack"""
        positions = []
        current_y = start_y
        
        for table in tables:
            table_height, is_collapsed = self.calculate_table_height(table, connections)
            
            positions.append({
                'nodeIndex': table,
                'location': {'x': x_position, 'y': current_y},
                'size': {'width': self.spacing_config['table_width'], 'height': table_height},
                'zIndex': len(positions),
                'category': category,
                'collapsed': is_collapsed,
                'connection_count': len(connections.get(table, set()))
            })
            
            current_y += table_height + self.spacing_config['within_stack_spacing']
            
        return positions
        
    def generate_parameter_grid_positions(self, tables: List[str], start_x: int, start_y: int, 
                                        connections: Dict[str, Set[str]]) -> List[Dict]:
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
            
            table_height, is_collapsed = self.calculate_table_height(table, connections)
            
            positions.append({
                'nodeIndex': table,
                'location': {
                    'x': start_x + col * col_spacing,
                    'y': start_y + row * row_spacing
                },
                'size': {'width': self.spacing_config['table_width'], 'height': table_height},
                'zIndex': len(positions),
                'category': 'parameter_grid',
                'collapsed': is_collapsed,
                'connection_count': len(connections.get(table, set()))
            })
            
        return positions
