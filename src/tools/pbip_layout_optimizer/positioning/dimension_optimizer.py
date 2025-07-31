# tools/pbip_layout_optimizer/positioning/dimension_optimizer.py
"""
Dimension Optimizer for PBIP model diagrams
Optimizes dimension table placement and grouping
"""

import logging
from typing import Dict, List, Tuple, Set

logger = logging.getLogger("pbip-tools-mcp")


class DimensionOptimizer:
    """Optimizes dimension table placement and grouping"""
    
    def __init__(self, table_categorizer):
        self.table_categorizer = table_categorizer
        
    def detect_parent_child_relationships(self, table_list: List[str], connections: Dict[str, Set[str]]) -> List[Tuple[str, str]]:
        """UNIVERSAL: Detect 1:Many parent-child relationships that should be positioned adjacent"""
        parent_child_pairs = []
        
        for parent_table in table_list:
            parent_connections = connections.get(parent_table, set())
            
            for child_table in table_list:
                if parent_table == child_table:
                    continue
                    
                child_connections = connections.get(child_table, set())
                
                # Check if they're connected
                if child_table in parent_connections:
                    # Determine parent-child relationship by connection counts
                    parent_total_connections = len(parent_connections)
                    child_total_connections = len(child_connections)
                    
                    # Parent typically has MORE connections than child
                    # OR use naming patterns to identify hierarchy
                    is_parent_child = False
                    
                    # Method 1: Connection count heuristic
                    if parent_total_connections > child_total_connections:
                        is_parent_child = True
                        
                    # Method 2: Naming pattern detection for common hierarchies
                    elif self._detect_naming_hierarchy(parent_table, child_table):
                        is_parent_child = True
                        
                    # Method 3: Business logic patterns (Property > Unit, Account > AccountTree, etc.)
                    elif self._detect_business_hierarchy(parent_table, child_table):
                        is_parent_child = True
                        
                    if is_parent_child:
                        pair = (parent_table, child_table)
                        if pair not in parent_child_pairs:
                            parent_child_pairs.append(pair)
                            logger.info(f"UNIVERSAL PARENT-CHILD: '{parent_table}' (parent) → '{child_table}' (child)")
                            
        return parent_child_pairs
        
    def _detect_naming_hierarchy(self, table1: str, table2: str) -> bool:
        """Detect parent-child relationships based on naming patterns"""
        name1_lower = table1.lower()
        name2_lower = table2.lower()
        
        # Pattern 1: Base name + Tree/Detail/Category patterns
        hierarchical_suffixes = ['tree', 'detail', 'category', 'attribute', 'extended', 'child']
        
        for suffix in hierarchical_suffixes:
            # table2 is child if it has hierarchical suffix and shares base name with table1
            if suffix in name2_lower and any(part in name1_lower for part in name2_lower.split(suffix)[0].split('_') if len(part) > 2):
                return True
                
        return False
        
    def _detect_business_hierarchy(self, table1: str, table2: str) -> bool:
        """Detect business domain hierarchies (Property > Unit, Account > AccountCategory, etc.)"""
        name1_lower = table1.lower()
        name2_lower = table2.lower()
        
        # Known business hierarchies
        business_hierarchies = [
            ('property', 'unit'),
            ('account', 'accountcategory'),
            ('job', 'jobdetail'),
            ('tenant', 'tenantprogram'),
            ('unit', 'unittype'),
            ('building', 'unit'),
            ('portfolio', 'property')
        ]
        
        for parent_keyword, child_keyword in business_hierarchies:
            if parent_keyword in name1_lower and child_keyword in name2_lower:
                return True
                
        return False
    
    def sort_tables_by_connections(self, table_list: List[str], connections: Dict[str, Set[str]], 
                                 time_period_tables: List[str] = None) -> List[str]:
        """UNIVERSAL: Sort tables to group parent-child pairs and connected ones together"""
        if len(table_list) <= 1:
            return table_list
            
        time_period_tables = time_period_tables or []
        sorted_tables = []
        remaining_tables = table_list.copy()
        
        # UNIVERSAL: Place time/period tables at TOP of the list (closest to Calendar table)
        for time_table in time_period_tables:
            if time_table in remaining_tables:
                sorted_tables.insert(0, time_table)  # TOP position
                remaining_tables.remove(time_table)
                logger.info(f"POSITIONED: Time/period table '{time_table}' placed at TOP of stack (closest to Calendar)")
        
        # UNIVERSAL: Find and group parent-child pairs FIRST (highest priority)
        parent_child_pairs = self.detect_parent_child_relationships(remaining_tables, connections)
        
        # Add parent-child pairs together (parent first, then child)
        for parent_table, child_table in parent_child_pairs:
            if parent_table in remaining_tables and child_table in remaining_tables:
                # Add parent first
                sorted_tables.append(parent_table)
                remaining_tables.remove(parent_table)
                # Add child immediately after parent
                sorted_tables.append(child_table)
                remaining_tables.remove(child_table)
                logger.info(f"UNIVERSAL ADJACENT PARENT-CHILD: '{parent_table}' → '{child_table}' positioned together")
                
        # FALLBACK: Add any remaining connected pairs that aren't parent-child
        remaining_connected_pairs = self.identify_general_connected_pairs(remaining_tables, connections)
        
        for table1, table2 in remaining_connected_pairs:
            if table1 in remaining_tables and table2 in remaining_tables:
                sorted_tables.append(table1)
                remaining_tables.remove(table1)
                sorted_tables.append(table2)
                remaining_tables.remove(table2)
                logger.info(f"CONNECTED PAIR: '{table1}' + '{table2}' placed together")
                
        # Add remaining unconnected tables
        sorted_tables.extend(remaining_tables)
        
        logger.info(f"UNIVERSAL SORTED ORDER: {sorted_tables}")
        return sorted_tables
        
    def identify_general_connected_pairs(self, table_list: List[str], connections: Dict[str, Set[str]]) -> List[Tuple[str, str]]:
        """Identify general connected pairs (not parent-child) for grouping"""
        pairs = []
        
        for table1 in table_list:
            connected_tables = connections.get(table1, set())
            for table2 in connected_tables:
                if table2 in table_list and table1 != table2:
                    # Avoid duplicates by ordering the pair
                    pair = tuple(sorted([table1, table2]))
                    if pair not in pairs:
                        pairs.append(pair)
                        logger.info(f"GENERAL CONNECTED PAIR: '{pair[0]}' <-> '{pair[1]}'")
        
        return pairs

    def identify_connected_pairs(self, table_list: List[str], connections: Dict[str, Set[str]]) -> List[Tuple[str, str]]:
        """Identify tables that are directly connected for grouping"""
        pairs = []
        processed = set()
        
        for table1 in table_list:
            if table1 in processed:
                continue
                
            connected_tables = connections.get(table1, set())
            for table2 in connected_tables:
                if table2 in table_list and table2 != table1:
                    # Create sorted pair to avoid duplicates
                    pair = tuple(sorted([table1, table2]))
                    if pair not in pairs:
                        pairs.append(pair)
                        processed.add(table1)
                        processed.add(table2)
                        logger.info(f"CONNECTED PAIR: '{table1}' ↔ '{table2}'")
                        break  # Only pair each table once
                        
        return pairs

    def optimize_dimension_placement(self, l1_dimensions: List[str], fact_tables: List[str], 
                                   connections: Dict[str, Set[str]], calendar_tables: List[str] = None) -> Tuple[List[str], List[str]]:
        """DYNAMIC: Split L1 dimensions left/right with smart positioning for connected tables"""
        if not l1_dimensions:
            return [], []
            
        # DYNAMIC: Identify time/period tables that should be near calendar
        calendar_tables = calendar_tables or []
        time_period_tables = self.table_categorizer.identify_time_period_tables(l1_dimensions, calendar_tables, connections)
        
        # Check for special connected pairs that should be grouped together
        special_pairs = self.identify_connected_pairs(l1_dimensions, connections)
        
        # Start with simple split
        mid_point = len(l1_dimensions) // 2
        left_l1_dimensions = l1_dimensions[:mid_point]
        right_l1_dimensions = l1_dimensions[mid_point:]
        
        # DYNAMIC: Move time/period tables to left side AND to top position (closest to Calendar)
        for time_table in time_period_tables:
            if time_table in right_l1_dimensions:
                right_l1_dimensions.remove(time_table)
                # Insert at beginning to be at top of left stack
                if time_table not in left_l1_dimensions:
                    left_l1_dimensions.insert(0, time_table)
                logger.info(f"Moved time/period table '{time_table}' to TOP of left side (closest to Calendar)")
            elif time_table in left_l1_dimensions:
                # If already on left, move to top
                left_l1_dimensions.remove(time_table)
                left_l1_dimensions.insert(0, time_table)
                logger.info(f"Moved time/period table '{time_table}' to TOP of left stack (closest to Calendar)")
        
        # Adjust for special connected pairs
        for table1, table2 in special_pairs:
            if table1 in l1_dimensions and table2 in l1_dimensions:
                # Move both to the same side (prefer left if table1 is there)
                if table1 in left_l1_dimensions:
                    if table2 in right_l1_dimensions:
                        right_l1_dimensions.remove(table2)
                        left_l1_dimensions.append(table2)
                        logger.info(f"Moved '{table2}' to left side to group with '{table1}'")
                elif table1 in right_l1_dimensions:
                    if table2 in left_l1_dimensions:
                        left_l1_dimensions.remove(table2)
                        right_l1_dimensions.append(table2)
                        logger.info(f"Moved '{table2}' to right side to group with '{table1}'")
        
        # Sort tables within each side to group connected ones together
        left_l1_dimensions = self.sort_tables_by_connections(left_l1_dimensions, connections, time_period_tables)
        right_l1_dimensions = self.sort_tables_by_connections(right_l1_dimensions, connections, time_period_tables)
        
        logger.info(f"Left L1 dimensions: {left_l1_dimensions}")
        logger.info(f"Right L1 dimensions: {right_l1_dimensions}")
        
        return left_l1_dimensions, right_l1_dimensions
        
    def place_l2_dimensions_near_l1(self, l2_dimensions: List[str], left_l1_dimensions: List[str], 
                                  right_l1_dimensions: List[str], connections: Dict[str, Set[str]]) -> Tuple[List[str], List[str]]:
        """RELATIONSHIP-FIRST: Place L2 dimensions based on actual TMDL relationships, not names"""
        if not l2_dimensions:
            return [], []
            
        left_l2_dimensions = []
        right_l2_dimensions = []
        
        logger.info(f"RELATIONSHIP-BASED PLACEMENT: Processing {len(l2_dimensions)} L2 dimensions")
        
        # PHASE 1: Pure relationship-based placement (ignore names completely)
        placed_l2s = set()
        
        for l2_table in l2_dimensions:
            if l2_table in placed_l2s:
                continue
                
            # Find direct L1 connections from TMDL relationships
            l2_connections = connections.get(l2_table, set())
            connected_l1s = l2_connections.intersection(set(left_l1_dimensions + right_l1_dimensions))
            
            logger.info(f"L2 '{l2_table}' connects to L1s: {list(connected_l1s)}")
            
            if len(connected_l1s) == 1:
                # SINGLE CONNECTION = STRONG PAIRING (most reliable)
                l1_partner = list(connected_l1s)[0]
                
                if l1_partner in left_l1_dimensions:
                    left_l2_dimensions.append(l2_table)
                    placed_l2s.add(l2_table)
                    logger.info(f"SINGLE CONNECTION: L2 '{l2_table}' → LEFT with L1 '{l1_partner}'")
                elif l1_partner in right_l1_dimensions:
                    right_l2_dimensions.append(l2_table)
                    placed_l2s.add(l2_table)
                    logger.info(f"SINGLE CONNECTION: L2 '{l2_table}' → RIGHT with L1 '{l1_partner}'")
                    
            elif len(connected_l1s) > 1:
                # MULTIPLE CONNECTIONS = Find primary relationship
                # Use connection strength analysis
                connection_strengths = {}
                for l1_table in connected_l1s:
                    # Primary relationship = L1 with fewer total connections (more specific)
                    l1_total_connections = len(connections.get(l1_table, set()))
                    connection_strengths[l1_table] = l1_total_connections
                    
                primary_l1 = min(connection_strengths, key=connection_strengths.get)
                
                if primary_l1 in left_l1_dimensions:
                    left_l2_dimensions.append(l2_table)
                    placed_l2s.add(l2_table)
                    logger.info(f"MULTI CONNECTION: L2 '{l2_table}' → LEFT with primary L1 '{primary_l1}' (of {len(connected_l1s)})")
                elif primary_l1 in right_l1_dimensions:
                    right_l2_dimensions.append(l2_table)
                    placed_l2s.add(l2_table)
                    logger.info(f"MULTI CONNECTION: L2 '{l2_table}' → RIGHT with primary L1 '{primary_l1}' (of {len(connected_l1s)})")
                    
            # If no L1 connections found, will be handled in Phase 2
            
        # PHASE 2: Handle unplaced L2s (no direct L1 connections)
        unplaced_l2s = [l2 for l2 in l2_dimensions if l2 not in placed_l2s]
        
        for l2_table in unplaced_l2s:
            # Check for any connections to dimension tables to infer side preference
            l2_connections = connections.get(l2_table, set())
            
            # Count indirect connections through other dimensions
            left_side_score = 0
            right_side_score = 0
            
            # Check connections to left side tables
            for left_table in left_l1_dimensions + left_l2_dimensions:
                if left_table in l2_connections:
                    left_side_score += 1
                    
            # Check connections to right side tables  
            for right_table in right_l1_dimensions + right_l2_dimensions:
                if right_table in l2_connections:
                    right_side_score += 1
                    
            # Place based on connection preference, or balance if no preference
            if left_side_score > right_side_score:
                left_l2_dimensions.append(l2_table)
                logger.info(f"INDIRECT PREFERENCE: L2 '{l2_table}' → LEFT (score: {left_side_score} vs {right_side_score})")
            elif right_side_score > left_side_score:
                right_l2_dimensions.append(l2_table)
                logger.info(f"INDIRECT PREFERENCE: L2 '{l2_table}' → RIGHT (score: {right_side_score} vs {left_side_score})")
            else:
                # No preference - balance sides
                if len(left_l2_dimensions) <= len(right_l2_dimensions):
                    left_l2_dimensions.append(l2_table)
                    logger.info(f"BALANCE PLACEMENT: L2 '{l2_table}' → LEFT (balancing)")
                else:
                    right_l2_dimensions.append(l2_table)
                    logger.info(f"BALANCE PLACEMENT: L2 '{l2_table}' → RIGHT (balancing)")
        
        logger.info(f"RELATIONSHIP-BASED RESULTS:")
        logger.info(f"  Left L2 dimensions: {left_l2_dimensions}")
        logger.info(f"  Right L2 dimensions: {right_l2_dimensions}")
        
        return left_l2_dimensions, right_l2_dimensions
        
    def place_l3_dimensions_near_l2(self, l3_dimensions: List[str], left_l2_dimensions: List[str], 
                                   right_l2_dimensions: List[str], connections: Dict[str, Set[str]]) -> Tuple[List[str], List[str]]:
        """RELATIONSHIP-FIRST: Place L3+ dimensions based on actual connections to L2 dimensions"""
        if not l3_dimensions:
            return [], []
            
        left_l3_dimensions = []
        right_l3_dimensions = []
        
        logger.info(f"RELATIONSHIP-BASED L3 PLACEMENT: Processing {len(l3_dimensions)} L3+ dimensions")
        
        placed_l3s = set()
        
        for l3_table in l3_dimensions:
            if l3_table in placed_l3s:
                continue
                
            # Find direct L2 connections from TMDL relationships
            l3_connections = connections.get(l3_table, set())
            connected_l2s = l3_connections.intersection(set(left_l2_dimensions + right_l2_dimensions))
            
            logger.info(f"L3 '{l3_table}' connects to L2s: {list(connected_l2s)}")
            
            if len(connected_l2s) == 1:
                # SINGLE L2 CONNECTION = Follow that L2's side
                l2_partner = list(connected_l2s)[0]
                
                if l2_partner in left_l2_dimensions:
                    left_l3_dimensions.append(l3_table)
                    placed_l3s.add(l3_table)
                    logger.info(f"L3 FOLLOWS L2: '{l3_table}' → LEFT with L2 '{l2_partner}'")
                elif l2_partner in right_l2_dimensions:
                    right_l3_dimensions.append(l3_table)
                    placed_l3s.add(l3_table)
                    logger.info(f"L3 FOLLOWS L2: '{l3_table}' → RIGHT with L2 '{l2_partner}'")
                    
            elif len(connected_l2s) > 1:
                # MULTIPLE L2 CONNECTIONS = Find primary based on connection strength
                connection_strengths = {}
                for l2_table in connected_l2s:
                    l2_total_connections = len(connections.get(l2_table, set()))
                    connection_strengths[l2_table] = l2_total_connections
                    
                primary_l2 = min(connection_strengths, key=connection_strengths.get)
                
                if primary_l2 in left_l2_dimensions:
                    left_l3_dimensions.append(l3_table)
                    placed_l3s.add(l3_table)
                    logger.info(f"L3 MULTI L2: '{l3_table}' → LEFT with primary L2 '{primary_l2}' (of {len(connected_l2s)})")
                elif primary_l2 in right_l2_dimensions:
                    right_l3_dimensions.append(l3_table)
                    placed_l3s.add(l3_table)
                    logger.info(f"L3 MULTI L2: '{l3_table}' → RIGHT with primary L2 '{primary_l2}' (of {len(connected_l2s)})")
            
            # If no L2 connections, check for L1 connections as fallback
            elif not connected_l2s:
                # Check L1 connections as fallback
                from typing import TYPE_CHECKING
                if hasattr(self, '_last_l1_dimensions'):
                    all_l1_dimensions = self._last_l1_dimensions
                    connected_l1s = l3_connections.intersection(set(all_l1_dimensions))
                    
                    if connected_l1s:
                        # Use indirect L1 connection to infer side
                        # This is a fallback for L3s that skip L2 level
                        logger.info(f"L3 '{l3_table}' has L1 connections (no L2): {list(connected_l1s)}")
                        # Will be handled in balance phase below
        
        # Handle unplaced L3s (no clear L2 or L1 connections)
        unplaced_l3s = [l3 for l3 in l3_dimensions if l3 not in placed_l3s]
        
        for l3_table in unplaced_l3s:
            # Balance sides for unconnected L3s
            if len(left_l3_dimensions) <= len(right_l3_dimensions):
                left_l3_dimensions.append(l3_table)
                logger.info(f"L3 BALANCE: '{l3_table}' → LEFT (balancing, no clear connections)")
            else:
                right_l3_dimensions.append(l3_table)
                logger.info(f"L3 BALANCE: '{l3_table}' → RIGHT (balancing, no clear connections)")
        
        # =============================================================================
        # UNIVERSAL POSITIONING FIXES - Fix 1: Validation and Correction Loop
        # =============================================================================
        
        logger.info(f"🔧 VALIDATION PHASE: Checking L3 placement against L2 relationships...")
        
        # Validate and correct misplacements
        misplaced_l3s = self._find_misplaced_l3_tables(
            left_l3_dimensions, right_l3_dimensions, 
            left_l2_dimensions, right_l2_dimensions, 
            connections
        )
        
        # Correct misplacements
        for l3_table, correct_side in misplaced_l3s:
            if correct_side == 'left':
                self._move_table_to_left(l3_table, left_l3_dimensions, right_l3_dimensions)
            else:
                self._move_table_to_right(l3_table, left_l3_dimensions, right_l3_dimensions)
        
        # Final validation to confirm fixes worked
        remaining_misplacements = self._find_misplaced_l3_tables(
            left_l3_dimensions, right_l3_dimensions, 
            left_l2_dimensions, right_l2_dimensions, 
            connections
        )
        
        if remaining_misplacements:
            logger.warning(f"🚨 WARNING: {len(remaining_misplacements)} L3 tables still misplaced after correction!")
        else:
            logger.info(f"✅ SUCCESS: All L3 tables correctly positioned relative to their L2 connections")
        
        logger.info(f"RELATIONSHIP-BASED L3 RESULTS (AFTER VALIDATION):")
        logger.info(f"  Left L3+ dimensions: {left_l3_dimensions}")
        logger.info(f"  Right L3+ dimensions: {right_l3_dimensions}")
        
        return left_l3_dimensions, right_l3_dimensions
        
    def place_l4_dimensions_near_l3(self, l4_dimensions: List[str], left_l3_dimensions: List[str], 
                                   right_l3_dimensions: List[str], left_l2_dimensions: List[str], 
                                   right_l2_dimensions: List[str], connections: Dict[str, Set[str]]) -> Tuple[List[str], List[str]]:
        """RELATIONSHIP-FIRST: Place L4+ dimensions based on connections to L3 dimensions (with L2 fallback)"""
        if not l4_dimensions:
            return [], []
            
        left_l4_dimensions = []
        right_l4_dimensions = []
        
        logger.info(f"RELATIONSHIP-BASED L4+ PLACEMENT: Processing {len(l4_dimensions)} L4+ dimensions")
        
        placed_l4s = set()
        
        for l4_table in l4_dimensions:
            if l4_table in placed_l4s:
                continue
                
            l4_connections = connections.get(l4_table, set())
            
            # PRIORITY 1: Check for L3 connections first
            connected_l3s = l4_connections.intersection(set(left_l3_dimensions + right_l3_dimensions))
            
            if connected_l3s:
                logger.info(f"L4+ '{l4_table}' connects to L3s: {list(connected_l3s)}")
                
                if len(connected_l3s) == 1:
                    l3_partner = list(connected_l3s)[0]
                    
                    if l3_partner in left_l3_dimensions:
                        left_l4_dimensions.append(l4_table)
                        placed_l4s.add(l4_table)
                        logger.info(f"L4+ FOLLOWS L3: '{l4_table}' → LEFT with L3 '{l3_partner}'")
                    elif l3_partner in right_l3_dimensions:
                        right_l4_dimensions.append(l4_table)
                        placed_l4s.add(l4_table)
                        logger.info(f"L4+ FOLLOWS L3: '{l4_table}' → RIGHT with L3 '{l3_partner}'")
                        
                else:
                    # Multiple L3 connections - find primary
                    primary_l3 = min(connected_l3s, key=lambda x: len(connections.get(x, set())))
                    
                    if primary_l3 in left_l3_dimensions:
                        left_l4_dimensions.append(l4_table)
                        placed_l4s.add(l4_table)
                        logger.info(f"L4+ MULTI L3: '{l4_table}' → LEFT with primary L3 '{primary_l3}'")
                    elif primary_l3 in right_l3_dimensions:
                        right_l4_dimensions.append(l4_table)
                        placed_l4s.add(l4_table)
                        logger.info(f"L4+ MULTI L3: '{l4_table}' → RIGHT with primary L3 '{primary_l3}'")
                        
                continue
                
            # PRIORITY 2: Fallback to L2 connections if no L3 connections
            connected_l2s = l4_connections.intersection(set(left_l2_dimensions + right_l2_dimensions))
            
            if connected_l2s:
                logger.info(f"L4+ '{l4_table}' fallback to L2s: {list(connected_l2s)}")
                
                if len(connected_l2s) == 1:
                    l2_partner = list(connected_l2s)[0]
                    
                    if l2_partner in left_l2_dimensions:
                        left_l4_dimensions.append(l4_table)
                        placed_l4s.add(l4_table)
                        logger.info(f"L4+ FALLBACK L2: '{l4_table}' → LEFT with L2 '{l2_partner}'")
                    elif l2_partner in right_l2_dimensions:
                        right_l4_dimensions.append(l4_table)
                        placed_l4s.add(l4_table)
                        logger.info(f"L4+ FALLBACK L2: '{l4_table}' → RIGHT with L2 '{l2_partner}'")
                        
                else:
                    # Multiple L2 connections - find primary
                    primary_l2 = min(connected_l2s, key=lambda x: len(connections.get(x, set())))
                    
                    if primary_l2 in left_l2_dimensions:
                        left_l4_dimensions.append(l4_table)
                        placed_l4s.add(l4_table)
                        logger.info(f"L4+ MULTI L2: '{l4_table}' → LEFT with primary L2 '{primary_l2}'")
                    elif primary_l2 in right_l2_dimensions:
                        right_l4_dimensions.append(l4_table)
                        placed_l4s.add(l4_table)
                        logger.info(f"L4+ MULTI L2: '{l4_table}' → RIGHT with primary L2 '{primary_l2}'")
                        
                continue
        
        # Handle unplaced L4+ dimensions (no clear connections)
        unplaced_l4s = [l4 for l4 in l4_dimensions if l4 not in placed_l4s]
        
        for l4_table in unplaced_l4s:
            if len(left_l4_dimensions) <= len(right_l4_dimensions):
                left_l4_dimensions.append(l4_table)
                logger.info(f"L4+ BALANCE: '{l4_table}' → LEFT (balancing, no clear connections)")
            else:
                right_l4_dimensions.append(l4_table)
                logger.info(f"L4+ BALANCE: '{l4_table}' → RIGHT (balancing, no clear connections)")
        
        # Apply universal validation to L4+ dimensions too
        logger.info(f"🔧 L4+ VALIDATION: Checking L4+ placement against L3/L2 relationships...")
        
        # Check L4 against L3 relationships
        l4_misplacements = self._find_misplaced_l4_tables(
            left_l4_dimensions, right_l4_dimensions,
            left_l3_dimensions, right_l3_dimensions,
            left_l2_dimensions, right_l2_dimensions,
            connections
        )
        
        # Correct L4 misplacements
        for l4_table, correct_side in l4_misplacements:
            if correct_side == 'left':
                self._move_table_to_left(l4_table, left_l4_dimensions, right_l4_dimensions)
            else:
                self._move_table_to_right(l4_table, left_l4_dimensions, right_l4_dimensions)
        
        if l4_misplacements:
            logger.info(f"✅ CORRECTED: Fixed {len(l4_misplacements)} L4+ misplacements")
        
        logger.info(f"RELATIONSHIP-BASED L4+ RESULTS (AFTER VALIDATION):")
        logger.info(f"  Left L4+ dimensions: {left_l4_dimensions}")
        logger.info(f"  Right L4+ dimensions: {right_l4_dimensions}")
        
        return left_l4_dimensions, right_l4_dimensions
        
    def optimize_l3_order(self, l3_tables: List[str], target_tables: List[str], connections: Dict[str, Set[str]]) -> List[str]:
        """UNIVERSAL: Optimize table order to minimize relationship line distance and apply parent-child adjacency"""
        if len(l3_tables) <= 1:
            return l3_tables
            
        # UNIVERSAL PATTERN 1: Apply parent-child adjacency detection
        parent_child_pairs = self.detect_parent_child_relationships(l3_tables, connections)
        
        if parent_child_pairs:
            # Use parent-child ordering for better alignment
            ordered_tables = []
            remaining_tables = l3_tables.copy()
            
            # Add parent-child pairs first (parent then child)
            for parent_table, child_table in parent_child_pairs:
                if parent_table in remaining_tables and child_table in remaining_tables:
                    ordered_tables.append(parent_table)
                    remaining_tables.remove(parent_table)
                    ordered_tables.append(child_table)
                    remaining_tables.remove(child_table)
                    logger.info(f"UNIVERSAL L3 PARENT-CHILD: '{parent_table}' → '{child_table}' positioned for better alignment")
            
            # Add remaining tables
            ordered_tables.extend(remaining_tables)
            return ordered_tables
            
        # UNIVERSAL PATTERN 2: Connection-based optimization (fallback)
        table_scores = []
        for table in l3_tables:
            connected_targets = connections.get(table, set()).intersection(set(target_tables))
            score = len(connected_targets)
            table_scores.append((table, score, connected_targets))
            logger.info(f"L3+ table '{table}': {score} connections to target tables → {list(connected_targets)}")
        
        # Sort by connection score (tables with more connections to targets go first)
        table_scores.sort(key=lambda x: x[1], reverse=True)
        
        optimized_order = [table for table, score, connections in table_scores]
        logger.info(f"UNIVERSAL optimized order: {optimized_order}")
        
        return optimized_order
        
    def _detect_l1_l2_pairs(self, l2_dimensions: List[str], l1_dimensions: List[str], 
                          connections: Dict[str, Set[str]]) -> Dict[str, str]:
        """ENHANCED: Detect L1-L2 relationship pairs with better single-connection detection"""
        l1_l2_pairs = {}
        
        # SPECIAL CASE: Property-Portfolio relationship detection
        property_portfolio_pairs = self._detect_property_portfolio_pairs(l2_dimensions, l1_dimensions, connections)
        l1_l2_pairs.update(property_portfolio_pairs)
        
        for l2_table in l2_dimensions:
            connected_l1s = connections.get(l2_table, set()).intersection(set(l1_dimensions))
            
            # Skip if already paired by property-portfolio detection
            if l2_table in l1_l2_pairs:
                continue
            
            # ENHANCED: Any table with exactly one connection to an L1 should be L2
            if len(connected_l1s) == 1:
                l1_partner = list(connected_l1s)[0]
                l1_l2_pairs[l2_table] = l1_partner
                logger.info(f"DETECTED L1-L2 PAIR: '{l2_table}' ↔ '{l1_partner}' (exclusive connection)")
            
            # If L2 connects to multiple L1s, check for dominant connection
            elif len(connected_l1s) > 1:
                # Find L1 with minimum connections (most likely to be the primary pair)
                connection_strengths = {}
                for l1_table in connected_l1s:
                    l1_total_connections = len(connections.get(l1_table, set()))
                    connection_strengths[l1_table] = l1_total_connections
                
                primary_l1 = min(connection_strengths, key=connection_strengths.get)
                l1_l2_pairs[l2_table] = primary_l1
                logger.info(f"DETECTED L1-L2 PAIR: '{l2_table}' ↔ '{primary_l1}' (primary of {len(connected_l1s)} connections)")
        
        return l1_l2_pairs
        
    def _detect_property_portfolio_pairs(self, l2_dimensions: List[str], l1_dimensions: List[str],
                                       connections: Dict[str, Set[str]]) -> Dict[str, str]:
        """Detect Property-Portfolio and similar real estate L1-L2 pairs"""
        property_pairs = {}
        
        # Find property-related L1 tables
        property_l1_tables = []
        for l1_table in l1_dimensions:
            l1_lower = l1_table.lower()
            if any(keyword in l1_lower for keyword in ['property', 'building', 'asset', 'site', 'tenant']):
                property_l1_tables.append(l1_table)
                logger.info(f"PROPERTY L1 DETECTED: '{l1_table}'")
        
        # Find portfolio/unit-related L2 tables that connect to property L1s
        for l2_table in l2_dimensions:
            l2_lower = l2_table.lower()
            l2_connections = connections.get(l2_table, set())
            
            # Check if this L2 is portfolio/unit/tenant related and connects to a property L1
            if any(keyword in l2_lower for keyword in ['portfolio', 'unit', 'propertyunit', 'tenant']):
                # Find which property L1 it connects to
                connected_property_l1s = [p for p in property_l1_tables if p in l2_connections]
                
                if connected_property_l1s:
                    # Use the first/primary connection
                    primary_property = connected_property_l1s[0]
                    property_pairs[l2_table] = primary_property
                    logger.info(f"PROPERTY-PORTFOLIO PAIR: '{l2_table}' ↔ '{primary_property}'")
        
        return property_pairs
        
    def identify_additional_l2_tables(self, table_names: List[str], l1_dimensions: List[str], 
                                     connections: Dict[str, Set[str]], excluded_tables: Set[str]) -> List[str]:
        """ENHANCED: Identify tables that should be L2 based on single connections to L1 tables"""
        additional_l2s = []
        
        for table_name in table_names:
            if table_name in excluded_tables:
                continue
                
            table_connections = connections.get(table_name, set())
            
            # ENHANCED RULE: Single connection to an L1 dimension = should be L2
            l1_connections = table_connections.intersection(set(l1_dimensions))
            
            if len(l1_connections) == 1:
                connected_l1 = list(l1_connections)[0]
                additional_l2s.append(table_name)
                logger.info(f"ADDITIONAL L2: '{table_name}' → L2 (single connection to L1 '{connected_l1}')")
        
        return additional_l2s
        
    def order_dimensions_with_l1_l2_adjacency(self, l1_dimensions: List[str], l2_dimensions: List[str], 
                                            connections: Dict[str, Set[str]]) -> List[str]:
        """Order L1 and L2 dimensions so that L1-L2 pairs are positioned adjacent to each other"""
        if not l1_dimensions and not l2_dimensions:
            return []
        if not l2_dimensions:
            return l1_dimensions
        if not l1_dimensions:
            return l2_dimensions
            
        # Detect L1-L2 pairs
        l1_l2_pairs = self._detect_l1_l2_pairs(l2_dimensions, l1_dimensions, connections)
        
        ordered_dimensions = []
        used_l1s = set()
        used_l2s = set()
        
        # Process L1-L2 pairs first - place them adjacent
        for l2_table, l1_partner in l1_l2_pairs.items():
            if l1_partner not in used_l1s and l2_table not in used_l2s:
                ordered_dimensions.append(l1_partner)  # L1 first
                ordered_dimensions.append(l2_table)    # L2 immediately after
                used_l1s.add(l1_partner)
                used_l2s.add(l2_table)
                logger.info(f"ADJACENT POSITIONING: L1 '{l1_partner}' + L2 '{l2_table}' placed together")
        
        # Add remaining unpaired L1s
        for l1_table in l1_dimensions:
            if l1_table not in used_l1s:
                ordered_dimensions.append(l1_table)
                logger.info(f"Added unpaired L1: '{l1_table}'")
        
        # Add remaining unpaired L2s
        for l2_table in l2_dimensions:
            if l2_table not in used_l2s:
                ordered_dimensions.append(l2_table)
                logger.info(f"Added unpaired L2: '{l2_table}'")
        
        logger.info(f"Final ordered dimensions with L1-L2 adjacency: {ordered_dimensions}")
        return ordered_dimensions
        
    def find_table_category(self, table_name: str, categorized: Dict[str, List[str]]) -> str:
        """Find which category a table belongs to"""
        for category, tables in categorized.items():
            if table_name in tables:
                return category
        return 'unknown'
        
    def same_side(self, category1: str, category2: str) -> bool:
        """Check if two categories are on the same side"""
        left_categories = ['l1_dimension_left', 'l2_dimension_left', 'l3_dimension_left', 'l4_plus_dimension_left']
        right_categories = ['l1_dimension_right', 'l2_dimension_right', 'l3_dimension_right', 'l4_plus_dimension_right']
        
        return ((category1 in left_categories and category2 in left_categories) or
                (category1 in right_categories and category2 in right_categories))
                
    def apply_opposite_side_placement(self, categorized: Dict[str, List[str]], connections: Dict[str, Set[str]]) -> Dict[str, List[str]]:
        """Apply opposite-side placement for 1:1 relationships"""
        logger.info("🔍 DEBUG: Starting opposite-side placement for 1:1 relationships...")
        
        # Detect bidirectional relationships
        bidirectional_pairs = self.table_categorizer.detect_bidirectional_relationships(connections)
        relationship_types = self.table_categorizer.analyze_relationship_cardinality(connections)
        
        logger.info(f"🔍 DEBUG: Found {len(bidirectional_pairs)} bidirectional pairs: {bidirectional_pairs}")
        logger.info(f"🔍 DEBUG: Relationship types: {relationship_types}")
        
        # Apply placement rules
        for table1, table2 in bidirectional_pairs:
            rel_type = relationship_types.get((table1, table2), 'unknown')
            
            if rel_type == "one_to_one":
                # Force 1:1 pairs to opposite sides
                self._force_opposite_sides_for_pair(table1, table2, categorized, connections)
                logger.info(f"🔍 DEBUG: 1:1 OPPOSITE SIDES: {table1} ↔ {table2}")
                
            elif rel_type == "many_to_many":
                # Keep M:M pairs on same side
                self._force_same_side_for_pair(table1, table2, categorized, connections)
                logger.info(f"🔍 DEBUG: M:M SAME SIDE: {table1} ↔ {table2}")
                
        # ENHANCED: Handle 1:1 dimension extensions for adjacent positioning
        if hasattr(self.table_categorizer, 'relationship_analyzer'):
            try:
                # Get dimension extensions from the categorization
                initial_categories = {
                    'l1_dimensions': categorized.get('l1_dimension_left', []) + categorized.get('l1_dimension_right', []),
                    'l2_dimensions': categorized.get('l2_dimension_left', []) + categorized.get('l2_dimension_right', []),
                    'l3_dimensions': categorized.get('l3_dimension_left', []) + categorized.get('l3_dimension_right', []),
                    'l4_plus_dimensions': categorized.get('l4_plus_dimension_left', []) + categorized.get('l4_plus_dimension_right', [])
                }
                
                dimension_extensions = self.table_categorizer.relationship_analyzer.find_dimension_extensions(connections, initial_categories)
                
                # Apply adjacent positioning for dimension extensions
                for extension_table, (base_table, relationship_type) in dimension_extensions.items():
                    if relationship_type == 'extension':
                        self._ensure_extension_adjacency(extension_table, base_table, categorized)
                        logger.info(f"🔗 ADJACENT EXTENSION: {extension_table} positioned next to {base_table}")
                        
            except Exception as e:
                logger.warning(f"Could not apply dimension extension positioning: {e}")
        
        return categorized
        
    def _force_opposite_sides_for_pair(self, table1: str, table2: str, categorized: Dict[str, List[str]], connections: Dict[str, Set[str]]):
        """Force two tables to opposite sides (for 1:1 relationships)"""
        table1_category = self.find_table_category(table1, categorized)
        table2_category = self.find_table_category(table2, categorized)
        
        if table1_category == 'unknown' or table2_category == 'unknown':
            return
            
        # If both on same side, move the one with fewer connections
        if self.same_side(table1_category, table2_category):
            table1_connections = len(connections.get(table1, set()))
            table2_connections = len(connections.get(table2, set()))
            
            # Move the table with fewer connections (more specialized)
            if table1_connections <= table2_connections:
                self._move_to_opposite_side(table1, table1_category, categorized)
                logger.info(f"MOVED TO OPPOSITE: {table1} (fewer connections: {table1_connections} vs {table2_connections})")
            else:
                self._move_to_opposite_side(table2, table2_category, categorized)
                logger.info(f"MOVED TO OPPOSITE: {table2} (fewer connections: {table2_connections} vs {table1_connections})")
                
    def _force_same_side_for_pair(self, table1: str, table2: str, categorized: Dict[str, List[str]], connections: Dict[str, Set[str]]):
        """Force two tables to same side (for M:M relationships)"""
        table1_category = self.find_table_category(table1, categorized)
        table2_category = self.find_table_category(table2, categorized)
        
        if table1_category == 'unknown' or table2_category == 'unknown':
            return
            
        # If on different sides, move one to match the other
        if not self.same_side(table1_category, table2_category):
            # Move table with fewer connections to match the other
            table1_connections = len(connections.get(table1, set()))
            table2_connections = len(connections.get(table2, set()))
            
            if table1_connections <= table2_connections:
                self._move_to_same_side_as(table1, table1_category, table2_category, categorized)
                logger.info(f"MOVED TO SAME SIDE: {table1} to match {table2}")
            else:
                self._move_to_same_side_as(table2, table2_category, table1_category, categorized)
                logger.info(f"MOVED TO SAME SIDE: {table2} to match {table1}")
                
    def _move_to_opposite_side(self, table_name: str, current_category: str, categorized: Dict[str, List[str]]):
        """Move a table to the opposite side"""
        # Remove from current category
        if table_name in categorized[current_category]:
            categorized[current_category].remove(table_name)
            
        # Determine opposite category
        if 'left' in current_category:
            opposite_category = current_category.replace('left', 'right')
        elif 'right' in current_category:
            opposite_category = current_category.replace('right', 'left')
        else:
            return  # Can't determine opposite
            
        # Add to opposite category
        if opposite_category not in categorized:
            categorized[opposite_category] = []
        categorized[opposite_category].append(table_name)
        
    def _move_to_same_side_as(self, table_name: str, current_category: str, target_category: str, categorized: Dict[str, List[str]]):
        """Move a table to the same side as another table"""
        # Remove from current category
        if table_name in categorized[current_category]:
            categorized[current_category].remove(table_name)
            
        # Add to target side (same level, different side)
        if 'left' in target_category:
            same_side_category = current_category.replace('right', 'left')
        elif 'right' in target_category:
            same_side_category = current_category.replace('left', 'right')
        else:
            return
            
        if same_side_category not in categorized:
            categorized[same_side_category] = []
        categorized[same_side_category].append(table_name)
        
    def _ensure_extension_adjacency(self, extension_table: str, base_table: str, categorized: Dict[str, List[str]]):
        """Ensure 1:1 dimension extension is positioned on the same side and adjacent to its base table"""
        # Find which category the base table is in
        base_category = self.find_table_category(base_table, categorized)
        extension_category = self.find_table_category(extension_table, categorized)
        
        if base_category == 'unknown' or extension_category == 'unknown':
            logger.warning(f"Cannot position extension {extension_table} next to {base_table} - category not found")
            return
            
        # If extension is not on the same side as base, move it
        if not self.same_side(base_category, extension_category):
            # Remove extension from current category
            if extension_table in categorized[extension_category]:
                categorized[extension_category].remove(extension_table)
                
            # Add to same side as base table
            if 'left' in base_category:
                target_category = extension_category.replace('right', 'left')
            elif 'right' in base_category:
                target_category = extension_category.replace('left', 'right')
            else:
                target_category = extension_category
                
            if target_category not in categorized:
                categorized[target_category] = []
            categorized[target_category].append(extension_table)
            
            logger.info(f"MOVED EXTENSION: {extension_table} from {extension_category} to {target_category} to match {base_table}")
            
        # Now ensure they're positioned adjacent within the same category list
        target_category = self.find_table_category(base_table, categorized)
        if target_category != 'unknown' and extension_table in categorized[target_category] and base_table in categorized[target_category]:
            # Remove extension from its current position
            categorized[target_category].remove(extension_table)
            
            # Find base table position and insert extension right after it
            base_index = categorized[target_category].index(base_table)
            categorized[target_category].insert(base_index + 1, extension_table)
            
            logger.info(f"ADJACENT POSITIONING: {extension_table} positioned immediately after {base_table} in {target_category}")
        else:
            logger.warning(f"Could not ensure adjacency for {extension_table} and {base_table}")
        
    # =============================================================================
    # UNIVERSAL POSITIONING FIXES - Fix 4: Universal Table Movement Logic
    # =============================================================================
    
    def _move_table_to_left(self, table: str, left_list: List[str], right_list: List[str]):
        """Universal method to move table from right side to left side"""
        if table in right_list:
            right_list.remove(table)
            left_list.append(table)
            logger.info(f"CORRECTED: Moved '{table}' from RIGHT to LEFT")
    
    def _move_table_to_right(self, table: str, left_list: List[str], right_list: List[str]):
        """Universal method to move table from left side to right side"""
        if table in left_list:
            left_list.remove(table)
            right_list.append(table)
            logger.info(f"CORRECTED: Moved '{table}' from LEFT to RIGHT")
    
    # =============================================================================
    # UNIVERSAL POSITIONING FIXES - Fix 3: Universal Primary Connection Logic
    # =============================================================================
    
    def _get_primary_connection(self, connected_tables: Set[str], connections: Dict[str, Set[str]]) -> str:
        """Universal method to find the primary connection when multiple exist"""
        if len(connected_tables) == 1:
            return list(connected_tables)[0]
        
        # Primary = table with FEWER total connections (more specific relationship)
        connection_strengths = {}
        for table in connected_tables:
            total_connections = len(connections.get(table, set()))
            connection_strengths[table] = total_connections
        
        primary = min(connection_strengths, key=connection_strengths.get)
        logger.info(f"PRIMARY CONNECTION: {primary} (fewest connections: {connection_strengths[primary]})")
        return primary
    
    # =============================================================================
    # UNIVERSAL POSITIONING FIXES - Fix 2: Universal Misplacement Detection
    # =============================================================================
    
    def _find_misplaced_l3_tables(self, left_l3: List[str], right_l3: List[str], 
                                 left_l2: List[str], right_l2: List[str], 
                                 connections: Dict[str, Set[str]]) -> List[Tuple[str, str]]:
        """Universal detector for L3 tables on wrong side relative to their L2 connections"""
        misplacements = []
        
        all_l3 = left_l3 + right_l3
        all_l2 = left_l2 + right_l2
        
        for l3_table in all_l3:
            l3_connections = connections.get(l3_table, set())
            connected_l2s = l3_connections.intersection(set(all_l2))
            
            if connected_l2s:
                # Find primary L2 connection
                primary_l2 = self._get_primary_connection(connected_l2s, connections)
                
                # Determine where L3 should be based on L2 location
                should_be_left = primary_l2 in left_l2
                actually_left = l3_table in left_l3
                
                # Detect misplacement
                if should_be_left and not actually_left:
                    misplacements.append((l3_table, 'left'))
                    logger.error(f"MISPLACED: L3 '{l3_table}' on RIGHT but should be LEFT with L2 '{primary_l2}'")
                elif not should_be_left and actually_left:
                    misplacements.append((l3_table, 'right'))
                    logger.error(f"MISPLACED: L3 '{l3_table}' on LEFT but should be RIGHT with L2 '{primary_l2}'")
        
        return misplacements
    
    def _find_misplaced_l4_tables(self, left_l4: List[str], right_l4: List[str],
                                 left_l3: List[str], right_l3: List[str],
                                 left_l2: List[str], right_l2: List[str],
                                 connections: Dict[str, Set[str]]) -> List[Tuple[str, str]]:
        """Universal detector for L4+ tables on wrong side relative to their L3/L2 connections"""
        misplacements = []
        
        all_l4 = left_l4 + right_l4
        all_l3 = left_l3 + right_l3
        all_l2 = left_l2 + right_l2
        
        for l4_table in all_l4:
            l4_connections = connections.get(l4_table, set())
            
            # Check L3 connections first (priority)
            connected_l3s = l4_connections.intersection(set(all_l3))
            
            if connected_l3s:
                primary_l3 = self._get_primary_connection(connected_l3s, connections)
                should_be_left = primary_l3 in left_l3
                actually_left = l4_table in left_l4
                
                if should_be_left and not actually_left:
                    misplacements.append((l4_table, 'left'))
                    logger.error(f"MISPLACED L4+: '{l4_table}' on RIGHT but should be LEFT with L3 '{primary_l3}'")
                elif not should_be_left and actually_left:
                    misplacements.append((l4_table, 'right'))
                    logger.error(f"MISPLACED L4+: '{l4_table}' on LEFT but should be RIGHT with L3 '{primary_l3}'")
                continue
            
            # Fallback to L2 connections if no L3 connections
            connected_l2s = l4_connections.intersection(set(all_l2))
            
            if connected_l2s:
                primary_l2 = self._get_primary_connection(connected_l2s, connections)
                should_be_left = primary_l2 in left_l2
                actually_left = l4_table in left_l4
                
                if should_be_left and not actually_left:
                    misplacements.append((l4_table, 'left'))
                    logger.error(f"MISPLACED L4+: '{l4_table}' on RIGHT but should be LEFT with L2 '{primary_l2}'")
                elif not should_be_left and actually_left:
                    misplacements.append((l4_table, 'right'))
                    logger.error(f"MISPLACED L4+: '{l4_table}' on LEFT but should be RIGHT with L2 '{primary_l2}'")
        
        return misplacements
