# tools/pbip_layout_optimizer/positioning/universal_chain_alignment.py
"""
Universal Chain Alignment System - CHAIN FAMILIES
Group complete relationship chains into families and align entire families across stacks
"""

import logging
from typing import Dict, List, Tuple, Set, Any, Optional
from collections import defaultdict

logger = logging.getLogger("pbip-tools-mcp")


class UniversalChainAlignment:
    """Universal chain family alignment - group complete chains and align families"""
    
    def __init__(self, spacing_config: Dict[str, Any]):
        self.spacing_config = spacing_config
        
    def optimize_universal_stack_alignment(self, all_stacks: Dict[str, List[str]], 
                                         connections: Dict[str, set]) -> Dict[str, List[str]]:
        """🌍 Universal chain family alignment"""
        
        logger.info("🌍 STARTING UNIVERSAL CHAIN FAMILY ALIGNMENT")
        logger.info("=" * 80)
        
        # Log input analysis
        logger.info("📊 INPUT ANALYSIS:")
        total_tables = sum(len(tables) for tables in all_stacks.values())
        logger.info(f"   Total stacks: {len(all_stacks)}")
        logger.info(f"   Total tables: {total_tables}")
        
        for stack_name, tables in all_stacks.items():
            if tables:
                logger.info(f"   {stack_name}: {len(tables)} tables")
        
        # Phase 1: Detect complete relationship chain families
        logger.info("🔗 PHASE 1: DETECTING COMPLETE RELATIONSHIP CHAIN FAMILIES")
        chain_families = self._detect_chain_families(all_stacks, connections)
        logger.info(f"🔗 FOUND {len(chain_families)} CHAIN FAMILIES")
        
        # Phase 2: Assign family positions and organize stacks
        logger.info("🎯 PHASE 2: ORGANIZING STACKS BY CHAIN FAMILIES")
        aligned_stacks = self._organize_stacks_by_families(all_stacks, chain_families)
        
        # Verification phase
        total_output = sum(len(tables) for tables in aligned_stacks.values())
        logger.info(f"📊 VERIFICATION: Input {total_tables} tables → Output {total_output} tables")
        
        if total_output != total_tables:
            logger.error(f"🚨 WARNING: Table count mismatch! Missing {total_tables - total_output} tables")
        else:
            logger.info("✅ PERFECT: All tables preserved")
        
        logger.info("✅ UNIVERSAL CHAIN FAMILY ALIGNMENT COMPLETE")
        logger.info("=" * 80)
        return aligned_stacks
    
    def _detect_chain_families(self, all_stacks: Dict[str, List[str]], 
                              connections: Dict[str, set]) -> List[Dict[str, Any]]:
        """🔗 Detect complete relationship chain families (end-to-end chains)"""
        
        logger.info("🔗 DETECTING COMPLETE CHAIN FAMILIES")
        
        # Build level mappings
        level_mappings = {
            'l3': set(all_stacks.get('left_l3_dimensions', []) + all_stacks.get('right_l3_dimensions', [])),
            'l2': set(all_stacks.get('left_l2_dimensions', []) + all_stacks.get('right_l2_dimensions', [])),
            'l1': set(all_stacks.get('left_l1_dimensions', []) + all_stacks.get('right_l1_dimensions', [])),
            'facts': set(all_stacks.get('fact_tables', []))
        }
        
        chain_families = []
        used_tables = set()
        
        # Start from L3 tables and trace complete end-to-end chains
        for l3_table in level_mappings['l3']:
            if l3_table in used_tables:
                continue
                
            family = self._trace_complete_chain_family(l3_table, level_mappings, connections, all_stacks)
            if family and len(family['chain']) >= 2:  # At least 2 levels
                chain_families.append(family)
                used_tables.update(family['all_tables'])
                logger.info(f"   ✅ FAMILY: {family['name']} - {' → '.join(family['chain'])}")
        
        # Detect L2-start families for tables not in L3 families
        for l2_table in level_mappings['l2']:
            if l2_table in used_tables:
                continue
                
            family = self._trace_complete_chain_family(l2_table, level_mappings, connections, all_stacks)
            if family and len(family['chain']) >= 2:
                chain_families.append(family)
                used_tables.update(family['all_tables'])
                logger.info(f"   ✅ FAMILY: {family['name']} - {' → '.join(family['chain'])}")
        
        logger.info(f"🔗 TOTAL CHAIN FAMILIES: {len(chain_families)}")
        return chain_families
    
    def _trace_complete_chain_family(self, start_table: str, level_mappings: Dict[str, set], 
                                   connections: Dict[str, set], all_stacks: Dict[str, List[str]]) -> Optional[Dict[str, Any]]:
        """🔍 Trace a complete chain family starting from a table"""
        
        family_chain = [start_table]
        family_tables = {start_table}
        current_table = start_table
        
        # Determine starting level
        current_level = None
        for level, tables in level_mappings.items():
            if start_table in tables:
                current_level = level
                break
        
        if not current_level:
            return None
        
        # Trace down the hierarchy: L3 → L2 → L1 → Facts
        level_order = ['l3', 'l2', 'l1', 'facts']
        current_level_index = level_order.index(current_level)
        
        # Follow connections down the levels
        for next_level_index in range(current_level_index + 1, len(level_order)):
            next_level = level_order[next_level_index]
            next_level_tables = level_mappings[next_level]
            
            # Find connections to next level
            current_connections = connections.get(current_table, set())
            connected_next = current_connections.intersection(next_level_tables)
            
            if connected_next:
                # Take the first connection (could be enhanced to handle multiple)
                next_table = list(connected_next)[0]
                family_chain.append(next_table)
                family_tables.add(next_table)
                current_table = next_table
                
                # Also collect any extension tables at this level
                extensions = self._find_extension_tables(next_table, next_level_tables, connections, all_stacks)
                family_tables.update(extensions)
            else:
                # Chain ends here
                break
        
        # Only return families with meaningful chains
        if len(family_chain) >= 2:
            # Generate family name from the start table
            family_name = start_table.replace('Dim_', '').replace('Tree', '').replace('Category', '')
            
            return {
                'name': family_name,
                'chain': family_chain,
                'all_tables': list(family_tables),
                'start_table': start_table,
                'levels_spanned': len(family_chain)
            }
        
        return None
    
    def _find_extension_tables(self, base_table: str, level_tables: set, 
                             connections: Dict[str, set], all_stacks: Dict[str, List[str]]) -> List[str]:
        """🔍 Find extension tables that belong to the same family"""
        
        extensions = []
        base_connections = connections.get(base_table, set())
        
        # Look for other tables at this level connected to the base table
        for table in level_tables:
            if table != base_table:
                table_connections = connections.get(table, set())
                
                # Check for connections between tables at same level (extensions)
                if (base_table in table_connections or table in base_connections):
                    extensions.append(table)
        
        return extensions
    
    def _organize_stacks_by_families(self, all_stacks: Dict[str, List[str]], 
                                   chain_families: List[Dict[str, Any]]) -> Dict[str, List[str]]:
        """🎯 Organize all stacks by chain families"""
        
        logger.info("🎯 ORGANIZING STACKS BY CHAIN FAMILIES")
        
        # Create family position mapping
        family_positions = {}
        for i, family in enumerate(chain_families):
            family_positions[family['name']] = i
            logger.info(f"   Family '{family['name']}' → Position {i}")
        
        # Reorganize each stack
        aligned_stacks = {}
        
        for stack_name, original_tables in all_stacks.items():
            if not original_tables:
                aligned_stacks[stack_name] = []
                continue
            
            logger.info(f"   Organizing stack: {stack_name}")
            
            # Create position-ordered list
            stack_organization = []
            unassigned_tables = list(original_tables)
            
            # Phase 1: Assign family tables to their positions
            for family in sorted(chain_families, key=lambda f: family_positions[f['name']]):
                family_tables_in_stack = [t for t in family['all_tables'] if t in original_tables]
                
                for table in family_tables_in_stack:
                    if table in unassigned_tables:
                        stack_organization.append({
                            'table': table,
                            'family': family['name'],
                            'position': family_positions[family['name']]
                        })
                        unassigned_tables.remove(table)
                        logger.info(f"      ✅ {table} → Family '{family['name']}' (position {family_positions[family['name']]})")
            
            # Phase 2: Add unassigned tables at the end
            for table in unassigned_tables:
                stack_organization.append({
                    'table': table,
                    'family': 'unassigned',
                    'position': len(chain_families) + len(stack_organization)
                })
                logger.info(f"      📋 {table} → Unassigned (position {len(chain_families) + len(stack_organization)})")
            
            # Sort by position and extract table names
            stack_organization.sort(key=lambda x: x['position'])
            aligned_stacks[stack_name] = [item['table'] for item in stack_organization]
            
            logger.info(f"   Stack {stack_name}: {len(original_tables)} → {len(aligned_stacks[stack_name])} tables")
        
        # Store family information for position generator
        self._chain_families = chain_families
        self._family_positions = family_positions
        
        return aligned_stacks
