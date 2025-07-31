# tools/pbip_layout_optimizer/analyzers/relationship_analyzer.py
"""
Relationship Analyzer for PBIP models
Analyzes table connections and builds relationship graphs
"""

import logging
from typing import Dict, List, Any, Set
from pathlib import Path
from collections import defaultdict, deque

logger = logging.getLogger("pbip-tools-mcp")


class RelationshipAnalyzer:
    """Analyzes relationships between tables in PBIP models"""
    
    def __init__(self, base_engine):
        self.base_engine = base_engine
        
    def parse_relationships_from_tmdl(self) -> List[Dict[str, Any]]:
        """Parse relationships from relationships.tmdl file"""
        if not self.base_engine.semantic_model_path:
            return []
            
        definition_path = self.base_engine.semantic_model_path / "definition"
        relationships_file = definition_path / "relationships.tmdl"
        
        if not relationships_file.exists():
            logger.warning(f"No relationships.tmdl file found at {relationships_file}")
            return []
            
        relationships = []
        try:
            with open(relationships_file, 'r', encoding='utf-8') as f:
                content = f.read()
                
            logger.info(f"Parsing relationships from {relationships_file}")
            
            # Parse relationships from TMDL format
            lines = content.split('\n')
            in_relationship = False
            current_rel = {}
            
            for line in lines:
                line = line.strip()
                
                if line.startswith('relationship '):
                    in_relationship = True
                    rel_name = line.replace('relationship ', '').strip()
                    current_rel = {'name': rel_name}
                    
                elif in_relationship and 'fromColumn:' in line:
                    column_ref = line.split('fromColumn:')[1].strip()
                    from_table = column_ref.split('.')[0].strip().strip("'").strip('"')
                    from_table = self.base_engine._normalize_table_name(from_table)
                    current_rel['fromTable'] = from_table
                    
                elif in_relationship and 'toColumn:' in line:
                    column_ref = line.split('toColumn:')[1].strip()
                    to_table = column_ref.split('.')[0].strip().strip("'").strip('"')
                    to_table = self.base_engine._normalize_table_name(to_table)
                    current_rel['toTable'] = to_table
                    
                elif in_relationship and 'fromCardinality:' in line:
                    cardinality = line.split('fromCardinality:')[1].strip()
                    current_rel['fromCardinality'] = cardinality
                    
                elif in_relationship and 'toCardinality:' in line:
                    cardinality = line.split('toCardinality:')[1].strip()
                    current_rel['toCardinality'] = cardinality
                    
                elif in_relationship and 'crossFilteringBehavior:' in line:
                    behavior = line.split('crossFilteringBehavior:')[1].strip()
                    current_rel['crossFilteringBehavior'] = behavior
                    
                elif in_relationship and line == '':
                    in_relationship = False
                    if 'fromTable' in current_rel and 'toTable' in current_rel:
                        relationships.append(current_rel)
                        logger.debug(f"Found relationship: {current_rel['fromTable']} -> {current_rel['toTable']}")
                    current_rel = {}
                    
            # Handle last relationship if file doesn't end with blank line
            if in_relationship and current_rel and 'fromTable' in current_rel and 'toTable' in current_rel:
                relationships.append(current_rel)
                logger.debug(f"Found relationship: {current_rel['fromTable']} -> {current_rel['toTable']}")
                
            logger.info(f"Parsed {len(relationships)} relationships from TMDL")
                    
        except Exception as e:
            logger.error(f"Error parsing relationships from TMDL: {e}")
            
        return relationships
        
    def build_relationship_graph(self) -> Dict[str, Set[str]]:
        """Build a graph of table connections"""
        relationships = self.parse_relationships_from_tmdl()
        connections = defaultdict(set)
        
        # Get all table names for validation
        all_tables = set(self.base_engine._get_table_names_from_tmdl())
        
        for rel in relationships:
            from_table = rel.get('fromTable', '')
            to_table = rel.get('toTable', '')
            
            if from_table and to_table:
                # Validate that both tables exist
                if from_table in all_tables and to_table in all_tables:
                    connections[from_table].add(to_table)
                    connections[to_table].add(from_table)
                    logger.debug(f"Added connection: {from_table} <-> {to_table}")
                else:
                    logger.warning(f"Relationship references unknown table(s): {from_table} -> {to_table}")
        
        # Log connection counts
        for table in sorted(all_tables):
            count = len(connections.get(table, set()))
            logger.info(f"Table '{table}': {count} connections")
        
        return dict(connections)
        
    def calculate_distance_to_facts(self, table_name: str, connections: Dict[str, Set[str]], fact_tables: Set[str]) -> int:
        """Calculate shortest distance from table to any fact table"""
        if table_name in fact_tables:
            return 0
            
        visited = set()
        queue = deque([(table_name, 0)])
        
        while queue:
            current_table, distance = queue.popleft()
            
            if current_table in visited:
                continue
            visited.add(current_table)
            
            # Check direct connections
            for connected_table in connections.get(current_table, set()):
                if connected_table in fact_tables:
                    return distance + 1
                    
                if connected_table not in visited:
                    queue.append((connected_table, distance + 1))
                    
        return 999  # No path to facts found
        
    def get_downstream_connections(self, table_name: str, connections: Dict[str, Set[str]]) -> Set[str]:
        """Get tables that connect TO this table (downstream in snowflake)"""
        downstream = set()
        for other_table, connected_tables in connections.items():
            if table_name in connected_tables and other_table != table_name:
                downstream.add(other_table)
        return downstream
        
    def get_upstream_connections(self, table_name: str, connections: Dict[str, Set[str]]) -> Set[str]:
        """Get tables that this table connects TO (upstream in snowflake)"""
        return connections.get(table_name, set())
        
    def is_star_schema_table(self, table_name: str, connections: Dict[str, Set[str]], fact_tables: Set[str]) -> bool:
        """Check if table is pure star schema (connects to facts, no downstream tables)"""
        upstream = self.get_upstream_connections(table_name, connections)
        downstream = self.get_downstream_connections(table_name, connections)
        
        # Must connect to at least one fact table
        connects_to_fact = bool(upstream.intersection(fact_tables))
        
        # Must have no downstream tables
        has_no_downstream = len(downstream) == 0
        
        return connects_to_fact and has_no_downstream
        
    def identify_extension_tables_universal(self) -> Dict[str, Dict[str, Any]]:
        """🚀 UNIVERSAL: Generic detection of extension tables across any model"""
        relationships = self.parse_relationships_from_tmdl()
        connections = self.build_relationship_graph()
        extensions = {}
        
        logger.info("🔍 UNIVERSAL EXTENSION DETECTION: Starting comprehensive analysis...")
        logger.info(f"🔍 DEBUG: Looking specifically for Portfolio relationships...")
        
        # Step 1: Find all 1:1 relationship candidates
        one_to_one_candidates = []
        
        for rel in relationships:
            from_table = rel.get('fromTable', '')
            to_table = rel.get('toTable', '')
            
            # DEBUG: Log all Portfolio relationships
            if 'portfolio' in from_table.lower() or 'portfolio' in to_table.lower():
                logger.info(f"🎯 PORTFOLIO RELATIONSHIP FOUND: {from_table} ↔ {to_table}")
                logger.info(f"   fromCardinality: {rel.get('fromCardinality', 'none')}")
                logger.info(f"   toCardinality: {rel.get('toCardinality', 'none')}")
                logger.info(f"   crossFilteringBehavior: {rel.get('crossFilteringBehavior', 'none')}")
            
            if not from_table or not to_table:
                continue
                
            # Strong 1:1 indicators
            is_bidirectional = rel.get('crossFilteringBehavior') == 'bothDirections'
            has_one_cardinality = (rel.get('fromCardinality') == 'one' or 
                                  rel.get('toCardinality') == 'one')
            
            # Universal detection criteria (much more inclusive)
            is_extension_candidate = False
            detection_reasons = []
            
            # Pattern 1: Bidirectional + One Cardinality (STRONGEST)
            if is_bidirectional and has_one_cardinality:
                is_extension_candidate = True
                detection_reasons.append('bidirectional_and_one_cardinality')
                
            # Pattern 2: Pure Bidirectional (STRONG)
            elif is_bidirectional:
                is_extension_candidate = True
                detection_reasons.append('bidirectional_filtering')
                
            # Pattern 3: One Cardinality Only (MEDIUM)
            elif has_one_cardinality:
                is_extension_candidate = True
                detection_reasons.append('one_cardinality_only')
            
            if is_extension_candidate:
                strength = 3 if (is_bidirectional and has_one_cardinality) else 2 if is_bidirectional else 1
                one_to_one_candidates.append({
                    'table_a': from_table,
                    'table_b': to_table,
                    'strength': strength,
                    'bidirectional': is_bidirectional,
                    'one_cardinality': has_one_cardinality,
                    'detection_reasons': detection_reasons,
                    'relationship_id': rel.get('name', 'unknown')
                })
                
                # PORTFOLIO DEBUG: Special logging for Portfolio candidates
                if 'portfolio' in from_table.lower() or 'portfolio' in to_table.lower():
                    logger.info(f"🎯 PORTFOLIO EXTENSION CANDIDATE ADDED: {from_table} ↔ {to_table} ")
                    logger.info(f"   strength={strength}, reasons={detection_reasons}")
                    logger.info(f"   bidirectional={is_bidirectional}, one_cardinality={has_one_cardinality}")
                
                logger.info(f"🎯 EXTENSION CANDIDATE: {from_table} ↔ {to_table} "
                           f"(strength={strength}, reasons={detection_reasons})")
        
        logger.info(f"🔍 Found {len(one_to_one_candidates)} extension candidates")
        
        # PORTFOLIO DEBUG: Check if Portfolio was found in candidates
        portfolio_candidates = [c for c in one_to_one_candidates 
                              if 'portfolio' in c['table_a'].lower() or 'portfolio' in c['table_b'].lower()]
        logger.info(f"🎯 PORTFOLIO CANDIDATES FOUND: {len(portfolio_candidates)}")
        for candidate in portfolio_candidates:
            logger.info(f"   📋 Portfolio candidate: {candidate['table_a']} ↔ {candidate['table_b']}")
        
        # DEBUG: Print first few candidates for inspection
        for i, candidate in enumerate(one_to_one_candidates[:5]):
            logger.info(f"🔍 Candidate {i+1}: {candidate['table_a']} ↔ {candidate['table_b']} "
                       f"(strength={candidate['strength']}, reasons={candidate['detection_reasons']})")
        
        # Step 2: For each 1:1 pair, determine which is the extension
        for candidate in one_to_one_candidates:
            table_a = candidate['table_a']
            table_b = candidate['table_b']
            
            # Count total connections for each table
            connections_a = len(connections.get(table_a, set()))
            connections_b = len(connections.get(table_b, set()))
            
            logger.info(f"🔍 ANALYZING PAIR: {table_a} ({connections_a} connections) ↔ {table_b} ({connections_b} connections)")
            
            # PORTFOLIO DEBUG: Special attention to Portfolio analysis
            if 'portfolio' in table_a.lower() or 'portfolio' in table_b.lower():
                logger.info(f"🎯 PORTFOLIO ANALYSIS: {table_a} ({connections_a} conn) ↔ {table_b} ({connections_b} conn)")
                logger.info(f"   🔍 Portfolio connections: {connections.get('Portfolio', set())}")
                logger.info(f"   🔍 Dim_Property connections: {connections.get('Dim_Property', set())}")
            
            # Determine extension vs base table
            extension_table = None
            base_table = None
            determination_method = ""
            
            # Method 1: Connection count difference
            if connections_a < connections_b:
                extension_table = table_a
                base_table = table_b
                determination_method = "fewer_connections"
            elif connections_b < connections_a:
                extension_table = table_b
                base_table = table_a
                determination_method = "fewer_connections"
            else:
                # Method 2: Universal naming heuristics (no hardcoded names)
                ext_table, base_table_temp = self._determine_extension_by_universal_naming(table_a, table_b)
                extension_table = ext_table
                base_table = base_table_temp
                determination_method = "naming_heuristics"
            
            # PORTFOLIO DEBUG: Log Portfolio determination
            if 'portfolio' in table_a.lower() or 'portfolio' in table_b.lower():
                logger.info(f"🎯 PORTFOLIO DETERMINATION: extension={extension_table}, base={base_table}, method={determination_method}")
            
            # Step 3: Validate this looks like a real extension
            logger.info(f"🔍 VALIDATION: About to validate {extension_table} -> {base_table}")
            if self._validate_extension_pattern(extension_table, base_table, connections):
                extensions[extension_table] = {
                    'base_table': base_table,
                    'type': 'extension',
                    'detection_method': candidate['detection_reasons'],
                    'determination_method': determination_method,
                    'connection_count_ext': connections_a if extension_table == table_a else connections_b,
                    'connection_count_base': connections_b if extension_table == table_a else connections_a,
                    'confidence': candidate['strength'],
                    'relationship_id': candidate['relationship_id']
                }
                
                # PORTFOLIO DEBUG: Special logging for Portfolio extensions
                if extension_table == 'Portfolio' or base_table == 'Portfolio':
                    logger.info(f"🎯 PORTFOLIO EXTENSION DETECTED! {extension_table} extends {base_table}")
                    logger.info(f"   📋 Extension info: {extensions[extension_table]}")
                
                logger.info(f"✅ UNIVERSAL EXTENSION CONFIRMED: {extension_table} extends {base_table} "
                           f"(method={determination_method}, confidence={candidate['strength']})")
            else:
                # PORTFOLIO DEBUG: Log Portfolio validation failures
                if 'portfolio' in table_a.lower() or 'portfolio' in table_b.lower():
                    logger.info(f"🎯 PORTFOLIO VALIDATION FAILED: {extension_table} -> {base_table}")
                logger.info(f"❌ EXTENSION REJECTED: {extension_table} -> {base_table} (failed validation)")
        
        logger.info(f"🎯 FINAL UNIVERSAL EXTENSIONS DETECTED: {len(extensions)}")
        for ext_table, ext_info in extensions.items():
            logger.info(f"   📋 {ext_table} extends {ext_info['base_table']} "
                       f"(confidence={ext_info['confidence']}, method={ext_info['determination_method']})")
        
        # PORTFOLIO DEBUG: Final check for Portfolio in extensions
        if 'Portfolio' in extensions:
            logger.info(f"🎯 PORTFOLIO FINAL RESULT: Portfolio detected as extension of {extensions['Portfolio']['base_table']}")
        else:
            logger.info(f"🎯 PORTFOLIO FINAL RESULT: Portfolio NOT detected as extension")
        
        return extensions

    def _determine_extension_by_universal_naming(self, table_a: str, table_b: str) -> tuple:
        """🌐 UNIVERSAL: Generic naming heuristics to identify likely extension table"""
        
        # Universal extension indicators (not domain-specific)
        extension_indicators = [
            'list', 'attribute', 'detail', 'extended', 'profile', 'program', 
            'portfolio', 'category', 'type', 'option', 'meta', 'extra',
            'additional', 'supplemental', 'auxiliary', 'secondary', 'properties',
            'config', 'configuration', 'settings', 'parameters', 'tags',
            'classifications', 'hierarchy', 'tree', 'node', 'leaf'
        ]
        
        # Convert to lowercase for analysis
        a_lower = table_a.lower()
        b_lower = table_b.lower()
        
        # Count extension indicators in each table name
        a_score = sum(1 for indicator in extension_indicators if indicator in a_lower)
        b_score = sum(1 for indicator in extension_indicators if indicator in b_lower)
        
        logger.info(f"🔍 NAMING ANALYSIS: {table_a} score={a_score}, {table_b} score={b_score}")
        
        # Additional heuristics
        # Tables with underscores often indicate more specific/detailed tables
        a_underscore_count = table_a.count('_')
        b_underscore_count = table_b.count('_')
        
        # Tables with longer names often indicate extensions
        a_length = len(table_a)
        b_length = len(table_b)
        
        # Composite scoring
        a_total_score = a_score + (a_underscore_count * 0.5) + (a_length * 0.01)
        b_total_score = b_score + (b_underscore_count * 0.5) + (b_length * 0.01)
        
        if a_total_score > b_total_score:
            logger.info(f"🎯 NAMING DECISION: {table_a} is extension (score={a_total_score:.2f} > {b_total_score:.2f})")
            return table_a, table_b  # A is extension
        elif b_total_score > a_total_score:
            logger.info(f"🎯 NAMING DECISION: {table_b} is extension (score={b_total_score:.2f} > {a_total_score:.2f})")
            return table_b, table_a  # B is extension
        else:
            # Final fallback: alphabetical (consistent but arbitrary)
            if table_a > table_b:  # Later alphabetically = extension
                logger.info(f"🎯 NAMING DECISION: {table_a} is extension (alphabetical fallback)")
                return table_a, table_b
            else:
                logger.info(f"🎯 NAMING DECISION: {table_b} is extension (alphabetical fallback)")
                return table_b, table_a

    def _validate_extension_pattern(self, extension_table: str, base_table: str, 
                                   connections: Dict[str, Set[str]]) -> bool:
        """🛡️ UNIVERSAL: Validate that this looks like a real extension relationship"""
        
        # Extension table should have very few connections (1-4 typically for real extensions)
        ext_connections = len(connections.get(extension_table, set()))
        base_connections = len(connections.get(base_table, set()))
        
        logger.info(f"🛡️ VALIDATING: {extension_table} ({ext_connections} conn) extends {base_table} ({base_connections} conn)")
        
        # Rule 1: Extension should have fewer or equal connections than base
        if ext_connections > base_connections:
            logger.info(f"❌ VALIDATION FAILED: Extension has more connections than base")
            return False
        
        # Rule 2: Extension should connect to the base
        if extension_table not in connections.get(base_table, set()):
            logger.info(f"❌ VALIDATION FAILED: Extension doesn't connect to base")
            return False
        
        # Rule 3: Extension shouldn't have too many connections (reasonable limit)
        if ext_connections > 6:  # Extensions rarely have many connections
            logger.info(f"❌ VALIDATION FAILED: Extension has too many connections ({ext_connections})")
            return False
        
        # Rule 4: Base should have at least 1 connection
        if base_connections == 0:
            logger.info(f"❌ VALIDATION FAILED: Base table has no connections")
            return False
        
        logger.info(f"✅ VALIDATION PASSED: Extension pattern confirmed")
        return True

    def identify_one_to_one_relationships(self) -> Dict[str, str]:
        """Legacy method - now calls universal detection for backward compatibility"""
        universal_extensions = self.identify_extension_tables_universal()
        one_to_one_pairs = {}
        
        # Convert to old format for backward compatibility
        for extension_table, extension_info in universal_extensions.items():
            base_table = extension_info['base_table']
            one_to_one_pairs[extension_table] = base_table
            one_to_one_pairs[base_table] = extension_table
            
        logger.info(f"Total 1:1 relationships found (legacy format): {len(one_to_one_pairs) // 2}")
        return one_to_one_pairs
        
    def find_dimension_extensions(self, connections: Dict[str, Set[str]], categorized_tables: Dict[str, List[str]]) -> Dict[str, tuple]:
        """🚀 UNIVERSAL: Find dimension extension relationships using enhanced detection"""
        
        # Use the new universal detection method
        universal_extensions = self.identify_extension_tables_universal()
        dimension_extensions = {}
        
        # Get all dimension tables by level
        all_dimensions = (
            categorized_tables.get('l1_dimensions', []) +
            categorized_tables.get('l2_dimensions', []) +
            categorized_tables.get('l3_dimensions', []) +
            categorized_tables.get('l4_plus_dimensions', [])
        )
        
        # Create level mapping for each table
        level_mapping = {}
        for table in categorized_tables.get('l1_dimensions', []):
            level_mapping[table] = 1
        for table in categorized_tables.get('l2_dimensions', []):
            level_mapping[table] = 2
        for table in categorized_tables.get('l3_dimensions', []):
            level_mapping[table] = 3
        for table in categorized_tables.get('l4_plus_dimensions', []):
            level_mapping[table] = 4
        
        # Process universal extensions that involve dimension tables
        logger.info(f"🔍 DEBUG: Starting extension processing for dimension tables...")
        logger.info(f"🔍 all_dimensions includes: {all_dimensions}")
        logger.info(f"🔍 Portfolio in all_dimensions: {'Portfolio' in all_dimensions}")
        logger.info(f"🔍 Dim_Property in all_dimensions: {'Dim_Property' in all_dimensions}")
        
        for extension_table, extension_info in universal_extensions.items():
            base_table = extension_info['base_table']
            
            logger.info(f"🔍 PROCESSING EXTENSION: {extension_table} -> {base_table}")
            logger.info(f"   extension_table in all_dimensions: {extension_table in all_dimensions}")
            logger.info(f"   base_table in all_dimensions: {base_table in all_dimensions}")
            
            # Only process if both tables are dimensions
            if extension_table in all_dimensions and base_table in all_dimensions:
                base_level = level_mapping.get(base_table, 1)
                extension_level = level_mapping.get(extension_table, 1)
                
                # Store extension with enhanced info
                dimension_extensions[extension_table] = (
                    base_table, 
                    'extension', 
                    base_level, 
                    extension_level,
                    extension_info['confidence'],
                    extension_info['detection_method']
                )
                
                logger.info(f"🎯 DIMENSION EXTENSION CONFIRMED: {extension_table} (L{extension_level}) "
                           f"extends {base_table} (L{base_level}) - confidence={extension_info['confidence']}")
                
        return dimension_extensions
