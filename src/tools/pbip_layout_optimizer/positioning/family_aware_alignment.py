# tools/pbip_layout_optimizer/positioning/family_aware_alignment.py
"""
Universal Family-Aware Alignment Enhancement
Keeps related table families together within their respective L2/L3 stacks
without hardcoding any specific table names.
"""

import logging
from typing import Dict, List, Set, Any, Optional
from collections import defaultdict

logger = logging.getLogger("pbip-tools-mcp")


class FamilyAwareAlignment:
    """Universal enhancement to keep chain families together within stacks"""
    
    def __init__(self):
        self.extension_families = {}  # Maps extension tables to their base tables
        
    def apply_family_grouping(self, aligned_stacks: Dict[str, List[str]], 
                             extensions: Dict[str, Dict[str, Any]],
                             connections: Dict[str, set]) -> Dict[str, List[str]]:
        """
        🏠 UNIVERSAL: Keep extension families together within their respective stacks
        This ensures that extension tables are positioned near their base tables
        """
        
        logger.info("🏠 APPLYING UNIVERSAL FAMILY-AWARE GROUPING")
        logger.info("🎯 Goal: Keep extension families together within each stack")
        
        # Build family mappings from extension detection
        family_groups = self._build_family_groups(extensions, connections)
        
        # Apply family grouping to each stack
        family_organized_stacks = {}
        
        for stack_name, tables in aligned_stacks.items():
            if not tables:
                family_organized_stacks[stack_name] = []
                continue
                
            logger.info(f"🏠 ORGANIZING STACK: {stack_name} ({len(tables)} tables)")
            
            # Group tables by families within this stack
            organized_tables = self._organize_stack_by_families(
                tables, family_groups, stack_name)
            
            family_organized_stacks[stack_name] = organized_tables
            
            # Log the family organization
            self._log_family_organization(stack_name, tables, organized_tables, family_groups)
        
        logger.info("✅ UNIVERSAL FAMILY GROUPING COMPLETE")
        return family_organized_stacks
    
    def _build_family_groups(self, extensions: Dict[str, Dict[str, Any]], 
                           connections: Dict[str, set]) -> Dict[str, List[str]]:
        """🔍 Build family groups from extension relationships"""
        
        family_groups = defaultdict(list)
        base_to_extensions = defaultdict(list)
        
        # Map extensions to their base tables
        for extension_table, ext_info in extensions.items():
            base_table = ext_info['base_table']
            base_to_extensions[base_table].append(extension_table)
            
            logger.info(f"   📋 Extension: {extension_table} → Base: {base_table}")
        
        # Create family groups (base + all its extensions)
        for base_table, extension_list in base_to_extensions.items():
            family_members = [base_table] + extension_list
            family_name = self._generate_family_name(base_table)
            
            family_groups[family_name] = family_members
            logger.info(f"   🏠 Family '{family_name}': {family_members}")
        
        logger.info(f"🔍 BUILT {len(family_groups)} FAMILY GROUPS")
        return dict(family_groups)
    
    def _generate_family_name(self, base_table: str) -> str:
        """🏷️ Generate a universal family name from base table"""
        # Remove common prefixes and suffixes to get core family name
        name = base_table.replace('Dim_', '').replace('Fact_', '')
        name = name.replace('Tree', '').replace('Category', '')
        
        # Take first meaningful part for family grouping
        if '_' in name:
            name = name.split('_')[0]
        
        return name
    
    def _organize_stack_by_families(self, tables: List[str], 
                                   family_groups: Dict[str, List[str]],
                                   stack_name: str) -> List[str]:
        """🎯 Organize a single stack by keeping families together"""
        
        organized_tables = []
        used_tables = set()
        
        # Phase 1: Group tables by their families
        table_to_family = {}
        for family_name, family_members in family_groups.items():
            for table in family_members:
                if table in tables:
                    table_to_family[table] = family_name
        
        # Phase 2: Process families in order, keeping members together
        families_in_stack = defaultdict(list)
        non_family_tables = []
        
        for table in tables:
            if table in table_to_family:
                family_name = table_to_family[table]
                families_in_stack[family_name].append(table)
            else:
                non_family_tables.append(table)
        
        # Phase 3: Order families by their first appearance in original list
        family_order = {}
        for i, table in enumerate(tables):
            if table in table_to_family:
                family_name = table_to_family[table]
                if family_name not in family_order:
                    family_order[family_name] = i
        
        # Phase 4: Build final organized list
        # First, add families in order
        for family_name in sorted(family_order.keys(), key=lambda f: family_order[f]):
            family_tables = families_in_stack[family_name]
            
            # Within each family, sort to put base table first, then extensions
            organized_family = self._sort_family_members(family_tables, family_groups[family_name])
            
            organized_tables.extend(organized_family)
            used_tables.update(organized_family)
            
            logger.info(f"      🏠 Family '{family_name}': {organized_family}")
        
        # Add non-family tables in their original relative positions
        for table in tables:
            if table not in used_tables:
                organized_tables.append(table)
                used_tables.add(table)
        
        return organized_tables
    
    def _sort_family_members(self, family_tables_in_stack: List[str], 
                           all_family_members: List[str]) -> List[str]:
        """🎯 Sort family members to put base table first, then extensions"""
        
        # The first member in all_family_members should be the base table
        base_table = all_family_members[0] if all_family_members else None
        
        sorted_family = []
        
        # Add base table first if it's in this stack
        if base_table and base_table in family_tables_in_stack:
            sorted_family.append(base_table)
        
        # Add extensions in their original order
        for table in family_tables_in_stack:
            if table != base_table:
                sorted_family.append(table)
        
        return sorted_family
    
    def _log_family_organization(self, stack_name: str, original_tables: List[str], 
                               organized_tables: List[str], 
                               family_groups: Dict[str, List[str]]) -> None:
        """📊 Log the family organization changes"""
        
        if original_tables == organized_tables:
            logger.info(f"   ℹ️ {stack_name}: No family reorganization needed")
            return
        
        logger.info(f"   🔄 {stack_name}: Family reorganization applied")
        logger.info(f"      Original:  {original_tables}")
        logger.info(f"      Organized: {organized_tables}")
        
        # Show which families were grouped together
        families_found = set()
        for table in organized_tables:
            for family_name, family_members in family_groups.items():
                if table in family_members:
                    families_found.add(family_name)
                    break
        
        if families_found:
            logger.info(f"      👨‍👩‍👧‍👦 Families in stack: {sorted(families_found)}")


def enhance_alignment_with_family_grouping(aligned_stacks: Dict[str, List[str]], 
                                          extensions: Dict[str, Dict[str, Any]],
                                          connections: Dict[str, set]) -> Dict[str, List[str]]:
    """
    🚀 MAIN ENTRY POINT: Apply universal family-aware grouping enhancement
    
    This function enhances any existing alignment by keeping extension families
    together within their respective stacks, without hardcoding specific names.
    """
    
    family_aligner = FamilyAwareAlignment()
    return family_aligner.apply_family_grouping(aligned_stacks, extensions, connections)
