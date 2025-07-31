# tools/pbip_layout_optimizer/analyzers/table_categorizer.py
"""
Universal Table Categorizer for PBIP models
Dynamically categorizes tables into facts, dimensions, calendar, etc.
NO HARDCODED TABLE NAMES OR MODEL-SPECIFIC LOGIC
"""

import logging
from typing import Dict, List, Set, Any
from pathlib import Path

logger = logging.getLogger("pbip-tools-mcp")


class TableCategorizer:
    """Universal table categorizer based on patterns and relationships only"""
    
    def __init__(self, base_engine, relationship_analyzer):
        self.base_engine = base_engine
        self.relationship_analyzer = relationship_analyzer
    
    def identify_auto_date_tables(self, table_names: List[str]) -> List[str]:
        """UNIVERSAL: Identify Power BI auto date/time tables that should be excluded from layout"""
        auto_date_tables = []
        tmdl_files = self.base_engine._find_tmdl_files()
        
        for table_name in table_names:
            name_lower = table_name.lower()
            
            # Check naming patterns for auto date tables
            if (name_lower.startswith('datetabletemplate_') or 
                name_lower.startswith('localdatetable_')):
                auto_date_tables.append(table_name)
                logger.info(f"AUTO DATE TABLE BY NAME: '{table_name}'")
                continue
                
            # Check TMDL content for auto date table markers
            tmdl_file = tmdl_files.get(table_name)
            if tmdl_file and tmdl_file.exists():
                try:
                    content = tmdl_file.read_text(encoding='utf-8')
                    
                    # Check for Power BI auto date table annotations
                    auto_date_patterns = [
                        '__PBI_TemplateDateTable = true',
                        '__PBI_LocalDateTable = true',
                        'showAsVariationsOnly',  # Often found in auto date tables
                    ]
                    
                    # Also check if table is hidden AND has date hierarchy pattern
                    is_hidden = 'isHidden' in content
                    has_date_hierarchy = "'Date Hierarchy'" in content or '"Date Hierarchy"' in content
                    has_calendar_source = 'Calendar(' in content
                    
                    # Auto date table if has annotations OR is hidden date table with hierarchy
                    if (any(pattern in content for pattern in auto_date_patterns) or
                        (is_hidden and has_date_hierarchy and has_calendar_source)):
                        auto_date_tables.append(table_name)
                        logger.info(f"AUTO DATE TABLE BY TMDL: '{table_name}' (hidden={is_hidden}, hierarchy={has_date_hierarchy}, calendar={has_calendar_source})")
                        continue
                        
                except Exception as e:
                    logger.warning(f"Could not read TMDL file for {table_name}: {e}")
                    
        logger.info(f"🗓️ IDENTIFIED {len(auto_date_tables)} AUTO DATE TABLES TO EXCLUDE: {auto_date_tables[:5]}...")
        return auto_date_tables
        
    def identify_time_period_tables(self, table_names: List[str], calendar_tables: List[str], connections: Dict[str, Set[str]]) -> List[str]:
        """UNIVERSAL: Identify time/period-related tables based on naming patterns and calendar connections"""
        time_period_tables = []
        
        # Universal time/period indicators (language-agnostic patterns)
        time_period_indicators = [
            'period', 'time', 'date', 'calendar', 'fiscal', 'quarter', 'month', 'year',
            'cycle', 'season', 'week', 'day', 'hour', 'minute', 'interval',
            'duration', 'span', 'range', 'tempo', 'zeit', 'temps', 'hora'  # Multi-language support
        ]
        
        for table_name in table_names:
            name_lower = table_name.lower()
            
            # Check for time/period patterns in name
            is_time_period = any(indicator in name_lower for indicator in time_period_indicators)
            
            # Check if table connects to calendar tables (relationship-based)
            connects_to_calendar = False
            table_connections = connections.get(table_name, set())
            if any(cal_table in table_connections for cal_table in calendar_tables):
                connects_to_calendar = True
                
            # Include if either pattern-based or relationship-based match
            if is_time_period or connects_to_calendar:
                time_period_tables.append(table_name)
                logger.info(f"TIME/PERIOD TABLE: '{table_name}' (pattern={is_time_period}, calendar_connected={connects_to_calendar})")
                
        return time_period_tables
        
    def identify_calendar_connected_specials(self, table_names: List[str], calendar_tables: List[str], 
                                           connections: Dict[str, Set[str]], excluded_tables: Set[str]) -> List[str]:
        """UNIVERSAL: Identify tables ONLY connected to calendar for special positioning"""
        calendar_connected_specials = []
        
        for table_name in table_names:
            if table_name in excluded_tables:
                continue
                
            table_connections = connections.get(table_name, set())
            
            # UNIVERSAL LOGIC: Only tables EXCLUSIVELY connected to calendar
            calendar_connections = table_connections.intersection(set(calendar_tables))
            non_calendar_connections = table_connections - set(calendar_tables)
            
            # Must connect ONLY to calendar (no other connections)
            if len(calendar_connections) > 0 and len(non_calendar_connections) == 0:
                calendar_connected_specials.append(table_name)
                logger.info(f"CALENDAR-CONNECTED SPECIAL: '{table_name}' (only connects to calendar)")
                
        return calendar_connected_specials
        
    def identify_parameter_tables(self, table_names: List[str]) -> List[str]:
        """UNIVERSAL: Identify parameter tables using TMDL properties and universal naming patterns"""
        parameter_tables = []
        tmdl_files = self.base_engine._find_tmdl_files()
        
        # Universal parameter indicators
        parameter_indicators = [
            'param', 'parameter', 'config', 'setting', 'option', 'preference',
            'variable', 'constant', 'lookup', 'reference', 'master',
            'código', 'parametro', 'configuración'  # Multi-language support
        ]
        
        for table_name in table_names:
            name_lower = table_name.lower()
            
            # Check universal naming patterns first
            if (table_name.startswith('.') or  # Hidden/system tables
                any(indicator in name_lower for indicator in parameter_indicators)):
                parameter_tables.append(table_name)
                continue
                
            # Check TMDL content for parameter metadata (universal approach)
            tmdl_file = tmdl_files.get(table_name)
            if tmdl_file and tmdl_file.exists():
                try:
                    content = tmdl_file.read_text(encoding='utf-8')
                    # Look for parameter-related TMDL properties
                    parameter_patterns = [
                        'extendedProperty ParameterMetadata',
                        'isHidden = true',
                        'type = Parameter'
                    ]
                    if any(pattern in content for pattern in parameter_patterns):
                        parameter_tables.append(table_name)
                        continue
                except Exception as e:
                    logger.warning(f"Could not read TMDL file for {table_name}: {e}")
                    
        logger.info(f"Identified parameter tables: {parameter_tables}")
        return parameter_tables
        
    def identify_calculation_groups(self, table_names: List[str]) -> List[str]:
        """UNIVERSAL: Identify calculation group tables using TMDL properties"""
        calc_group_tables = []
        tmdl_files = self.base_engine._find_tmdl_files()
        
        for table_name in table_names:
            tmdl_file = tmdl_files.get(table_name)
            if tmdl_file and tmdl_file.exists():
                try:
                    content = tmdl_file.read_text(encoding='utf-8')
                    # Universal TMDL pattern for calculation groups
                    if 'calculationGroup' in content:
                        calc_group_tables.append(table_name)
                except Exception as e:
                    logger.warning(f"Could not read TMDL file for {table_name}: {e}")
                    
        logger.info(f"Identified calculation group tables: {calc_group_tables}")
        return calc_group_tables
        
    def identify_metrics_tables(self, table_names: List[str], connections: Dict[str, Set[str]]) -> List[str]:
        """UNIVERSAL: Enhanced measure table detection with strict requirements"""
        metrics_tables = []
        tmdl_files = self.base_engine._find_tmdl_files()
        
        for table_name in table_names:
            name_lower = table_name.lower()
            connection_count = len(connections.get(table_name, set()))
            
            # STRICT REQUIREMENT 1: Must be disconnected (no relationships)
            if connection_count != 0:
                continue
                
            # Check TMDL content to count regular vs calculated columns
            tmdl_file = tmdl_files.get(table_name)
            if not tmdl_file or not tmdl_file.exists():
                continue
                
            try:
                content = tmdl_file.read_text(encoding='utf-8')
                
                # Count regular columns (non-calculated)
                regular_column_count = 0
                calculated_column_count = 0
                
                # Simple parsing for column types
                lines = content.split('\n')
                
                for line in lines:
                    line = line.strip()
                    
                    # Detect column definitions
                    if line.startswith('column ') and not line.startswith('column\t'):
                        # This is a column definition
                        if 'dataType:' in line or 'type:' in line:
                            # Look ahead for calculated column indicators
                            if 'expression:' in content[content.find(line):content.find(line) + 200]:
                                calculated_column_count += 1
                            else:
                                regular_column_count += 1
                        else:
                            regular_column_count += 1
                    elif 'expression =' in line or 'expression:' in line:
                        # This indicates a calculated column
                        calculated_column_count += 1
                        
                # STRICT REQUIREMENT 2: Must have only 1 regular column OR be a pure measures table (0 regular columns)
                if regular_column_count > 1:
                    continue
                    
                # For pure measures tables (0 regular columns), require that it has measures
                if regular_column_count == 0:
                    # Check if table has measures (pure measures table)
                    has_measures = 'measure ' in content and ('lineageTag:' in content or 'formatString:' in content)
                    if not has_measures:
                        continue
                    
                # NOW check naming indicators for enhanced detection
                naming_match = False
                
                # Check for non-alphabetical prefix + "measure" pattern
                if (table_name.startswith(('_', '.', '*', '-', '#')) and 
                    'measure' in name_lower):
                    naming_match = True
                    logger.info(f"ENHANCED METRICS TABLE DETECTED: '{table_name}' (disconnected + 1 column + naming pattern)")
                    
                # Also check for traditional metrics patterns
                elif any(keyword in name_lower for keyword in ['metric', 'measure', 'kpi', 'score']):
                    naming_match = True
                    logger.info(f"TRADITIONAL METRICS TABLE DETECTED: '{table_name}' (disconnected + 1 column + keyword)")
                    
                if naming_match:
                    metrics_tables.append(table_name)
                    logger.info(f"METRICS TABLE: '{table_name}' (connections=0, regular_columns={regular_column_count}, calculated_columns={calculated_column_count})")
                    
            except Exception as e:
                logger.warning(f"Could not analyze TMDL file for {table_name}: {e}")
                
        logger.info(f"Enhanced metrics table detection found {len(metrics_tables)} tables: {metrics_tables}")
        return metrics_tables
        
    def identify_special_disconnected_tables(self, table_names: List[str]) -> List[str]:
        """UNIVERSAL: Identify special disconnected tables that should be in parameter grid"""
        special_tables = []
        tmdl_files = self.base_engine._find_tmdl_files()
        
        # Universal disconnected table indicators
        disconnected_indicators = [
            'parameter', 'config', 'setting', 'lookup', 'reference', 'master',
            'static', 'constant', 'readonly', 'system'
        ]
        
        for table_name in table_names:
            name_lower = table_name.lower()
            
            # Universal patterns for special disconnected tables
            if (table_name.startswith('.') or  # System/hidden tables
                any(indicator in name_lower for indicator in disconnected_indicators)):
                special_tables.append(table_name)
                continue
                
            # Check TMDL content for parameter-like properties
            tmdl_file = tmdl_files.get(table_name)
            if tmdl_file and tmdl_file.exists():
                try:
                    content = tmdl_file.read_text(encoding='utf-8')
                    special_patterns = [
                        'extendedProperty ParameterMetadata',
                        'calculationGroup',
                        'isHidden = true'
                    ]
                    if any(pattern in content for pattern in special_patterns):
                        special_tables.append(table_name)
                except Exception as e:
                    logger.warning(f"Could not read TMDL file for {table_name}: {e}")
                    
        logger.info(f"Identified special disconnected tables for parameter grid: {special_tables}")
        return special_tables
        
    def detect_bidirectional_relationships(self, connections: Dict[str, Set[str]]) -> List[tuple]:
        """UNIVERSAL: Detect 1:1 bidirectional pairs for opposite-side placement"""
        bidirectional_pairs = []
        processed_pairs = set()
        
        for table1, connected_tables in connections.items():
            for table2 in connected_tables:
                # Check if connection exists in both directions
                if table1 in connections.get(table2, set()):
                    # Create sorted pair to avoid duplicates
                    pair = tuple(sorted([table1, table2]))
                    if pair not in processed_pairs:
                        bidirectional_pairs.append(pair)
                        processed_pairs.add(pair)
                        logger.info(f"BIDIRECTIONAL RELATIONSHIP: {table1} ↔ {table2}")
        
        return bidirectional_pairs
        
    def analyze_relationship_cardinality(self, connections: Dict[str, Set[str]]) -> Dict[tuple, str]:
        """UNIVERSAL: Analyze relationship types for better placement"""
        relationship_types = {}
        
        for table1, connected_tables in connections.items():
            for table2 in connected_tables:
                # Check cardinality patterns
                if table1 in connections.get(table2, set()):
                    # Bidirectional = potential 1:1 or M:M
                    connection_strength_1 = len(connections[table1])
                    connection_strength_2 = len(connections[table2])
                    
                    if connection_strength_1 == 1 and connection_strength_2 == 1:
                        relationship_types[(table1, table2)] = "one_to_one"
                        logger.info(f"1:1 RELATIONSHIP: {table1} ↔ {table2}")
                    elif connection_strength_1 > 3 and connection_strength_2 > 3:
                        relationship_types[(table1, table2)] = "many_to_many"
                        logger.info(f"M:M RELATIONSHIP: {table1} ↔ {table2}")
                    else:
                        relationship_types[(table1, table2)] = "one_to_many"
                        logger.info(f"1:M RELATIONSHIP: {table1} ↔ {table2}")
                else:
                    relationship_types[(table1, table2)] = "many_to_one"
        
        return relationship_types
        
    def calculate_table_score(self, table_name: str, connections: Dict[str, Set[str]]) -> Dict[str, Any]:
        """UNIVERSAL: Calculate hybrid score for table classification without domain-specific logic"""
        name_lower = table_name.lower()
        connection_count = len(connections.get(table_name, set()))
        
        # Check if this is a special disconnected table FIRST
        special_disconnected = self.identify_special_disconnected_tables([table_name])
        if table_name in special_disconnected:
            return {
                'table_name': table_name,
                'fact_score': 0,
                'dim_score': 0,
                'connection_count': connection_count,
                'is_calendar': False,
                'classification': 'disconnected'
            }
        
        # UNIVERSAL naming pattern scores
        fact_name_score = 0
        dim_name_score = 0
        
        # Universal fact indicators (cross-industry patterns)
        fact_indicators = [
            'fact', 'trans', 'event', 'activity', 'record', 'log', 'history',
            'operation', 'process', 'action', 'movement', 'entry', 'line',
            'detail', 'item', 'occurrence', 'instance', 'measure'
        ]
        
        # Universal dimension indicators
        dim_indicators = [
            'dim', 'dimension', 'master', 'lookup', 'reference', 'category',
            'type', 'class', 'group', 'entity', 'object', 'subject'
        ]
        
        # Strong fact patterns (universal)
        strong_fact_patterns = ['fact_', '_fact', 'fact ', ' fact']
        for pattern in strong_fact_patterns:
            if pattern in name_lower:
                fact_name_score += 25
                break
        
        # Check for universal fact indicators
        for indicator in fact_indicators:
            if indicator in name_lower:
                fact_name_score += 15
                break
        
        # Universal dimension indicators
        for indicator in dim_indicators:
            if indicator in name_lower:
                dim_name_score += 25
                break
        
        # Universal dimension prefixes
        dim_prefixes = ['d_', 'dim_', 'dim-', 'dim ', 'master_', 'ref_', 'lookup_']
        for prefix in dim_prefixes:
            if name_lower.startswith(prefix):
                dim_name_score += 20
                break
                
        # Universal calendar indicators
        calendar_indicators = ['calendar', 'date', 'time', 'period']
        is_calendar = any(cal in name_lower for cal in calendar_indicators)
        
        # Connection-based scoring (universal logic)
        connection_score = 0
        if connection_count >= 5:
            connection_score += 15
        elif connection_count >= 3:
            connection_score += 10
        elif connection_count >= 2:
            connection_score += 5
        elif connection_count == 1:
            connection_score -= 3
        elif connection_count == 0:
            connection_score -= 5
            
        # Calculate final scores
        total_fact_score = fact_name_score + max(0, connection_score)
        total_dim_score = dim_name_score + max(0, -connection_score)
        
        # UNIVERSAL CLASSIFICATION RULES
        # Priority 1: Explicit fact naming
        if any(pattern in name_lower for pattern in ['fact_', '_fact', 'fact ', ' fact']):
            classification = 'fact'
            total_fact_score = max(total_fact_score, total_dim_score + 50)
            logger.info(f"FORCED FACT: '{table_name}' due to explicit fact naming")
        # Priority 2: Explicit dimension naming
        elif any(name_lower.startswith(prefix) for prefix in ['dim_', 'dim-', 'dim ', 'd_']):
            classification = 'dimension'
            total_dim_score = max(total_dim_score, total_fact_score + 10)
            logger.info(f"FORCED DIMENSION: '{table_name}' due to explicit dimension naming")
        # Priority 3: Connection count rule (universal threshold)
        elif connection_count >= 3:
            classification = 'fact'
            total_fact_score += 20
            logger.info(f"FACT BY CONNECTIONS: '{table_name}' has {connection_count} connections")
        else:
            # Bias adjustment for ambiguous cases
            if total_fact_score > 0 and connection_count >= 2:
                total_fact_score += 5
            
            classification = 'fact' if total_fact_score > total_dim_score else 'dimension'
        
        logger.info(f"Table '{table_name}': fact_score={total_fact_score}, dim_score={total_dim_score}, "
                   f"connections={connection_count}, classification={classification}")
        
        return {
            'table_name': table_name,
            'fact_score': total_fact_score,
            'dim_score': total_dim_score,
            'connection_count': connection_count,
            'is_calendar': is_calendar,
            'classification': classification
        }
        
    def categorize_tables(self, table_names: List[str], connections: Dict[str, Set[str]]) -> Dict[str, List[str]]:
        """UNIVERSAL: Categorize tables using universal principles only"""
        
        logger.info("🌍 UNIVERSAL TABLE CATEGORIZATION - NO HARDCODED NAMES")
        logger.info(f"📊 INPUT TABLES: {len(table_names)} total")
        
        # PHASE 0: Filter out Power BI auto date/time tables
        logger.info("🔍 PHASE 0: FILTERING AUTO DATE TABLES")
        auto_date_tables = self.identify_auto_date_tables(table_names)
        filtered_table_names = [name for name in table_names if name not in auto_date_tables]
        
        logger.info(f"🗓️ FILTERED OUT {len(auto_date_tables)} AUTO DATE TABLES")
        logger.info(f"📊 REMAINING TABLES FOR LAYOUT: {len(filtered_table_names)}")
        
        # Use filtered table names for all further processing
        table_names = filtered_table_names
        
        fact_tables = []
        l1_dimensions = []
        l2_dimensions = []
        l3_dimensions = []
        l4_plus_dimensions = []
        calendar_tables = []
        parameter_tables = []
        calculation_groups = []
        disconnected_tables = []
        dimension_extensions = {}
        
        processed_tables = set()
        
        # PHASE 1: Identify special table types using universal patterns
        logger.info("🔍 PHASE 1: UNIVERSAL SPECIAL TABLE IDENTIFICATION")
        
        parameter_tables = self.identify_parameter_tables(table_names)
        calculation_groups = self.identify_calculation_groups(table_names)
        metrics_tables = self.identify_metrics_tables(table_names, connections)
        
        processed_tables.update(parameter_tables)
        processed_tables.update(calculation_groups)
        processed_tables.update(metrics_tables)
        
        # PHASE 2: Universal fact identification
        logger.info("🔍 PHASE 2: UNIVERSAL FACT IDENTIFICATION")
        potential_facts = set()
        
        # Universal hybrid scoring - SKIP DISCONNECTED TABLES (they can't be facts)
        for table_name in table_names:
            if table_name in parameter_tables or table_name in calculation_groups or table_name in metrics_tables:
                continue
                
            # CRITICAL: Skip disconnected tables from fact detection
            connection_count = len(connections.get(table_name, set()))
            if connection_count == 0:
                logger.info(f"SKIPPING DISCONNECTED TABLE FROM FACT DETECTION: '{table_name}' (0 connections)")
                continue
                
            score = self.calculate_table_score(table_name, connections)
            if score['classification'] == 'fact':
                fact_tables.append(table_name)
                potential_facts.add(table_name)
                logger.info(f"FACT identified: {table_name}")
        
        # Universal connection-based fallback - EXCLUDE DISCONNECTED TABLES
        if not fact_tables:
            logger.warning("No facts found by scoring, using universal connection-based approach")
            connection_counts = [(name, len(connections.get(name, set()))) for name in table_names 
                               if name not in parameter_tables and name not in calculation_groups and name not in metrics_tables]
            connection_counts.sort(key=lambda x: x[1], reverse=True)
            
            for name, count in connection_counts[:10]:  # Top 10 most connected
                if count >= 2:  # Must have connections to be a fact
                    fact_tables.append(name)
                    potential_facts.add(name)
                    logger.info(f"FACT by connections: {name} ({count} connections)")
                    
        processed_tables.update(fact_tables)
        
        # PHASE 3: Universal table categorization
        logger.info("🔍 PHASE 3: UNIVERSAL TABLE CATEGORIZATION")
        
        special_disconnected = self.identify_special_disconnected_tables(table_names)
        
        for table_name in table_names:
            if table_name in processed_tables:
                continue
                
            name_lower = table_name.lower()
            connection_count = len(connections.get(table_name, set()))
            
            # Universal calendar detection
            if any(cal in name_lower for cal in ['calendar', 'date', 'time', 'period']):
                calendar_tables.append(table_name)
                processed_tables.add(table_name)
                continue
                
            # Skip if already identified as metrics table in Phase 1
            if table_name in metrics_tables:
                continue
                
            # Special disconnected tables
            if table_name in special_disconnected:
                parameter_tables.append(table_name)
                processed_tables.add(table_name)
                continue
                
            # Disconnected tables
            if connection_count == 0:
                disconnected_tables.append(table_name)
                processed_tables.add(table_name)
                continue
                
            # Universal dimension level classification
            if self.relationship_analyzer.is_star_schema_table(table_name, connections, potential_facts):
                l1_dimensions.append(table_name)
                processed_tables.add(table_name)
                continue
                
            # Universal snowflake detection
            distance = self.relationship_analyzer.calculate_distance_to_facts(table_name, connections, potential_facts)
            
            if distance == 1:
                l1_dimensions.append(table_name)
            elif distance == 2:
                l2_dimensions.append(table_name)
            elif distance == 3:
                l3_dimensions.append(table_name)
            elif distance >= 4:
                l4_plus_dimensions.append(table_name)
            else:
                l1_dimensions.append(table_name)  # Default fallback
                
            processed_tables.add(table_name)
        
        # PHASE 4: Universal 1:1 extension handling
        logger.info("🔍 PHASE 4: UNIVERSAL 1:1 EXTENSION PROCESSING")
        
        initial_categories = {
            'fact_tables': fact_tables,
            'l1_dimensions': l1_dimensions,
            'l2_dimensions': l2_dimensions,
            'l3_dimensions': l3_dimensions,
            'l4_plus_dimensions': l4_plus_dimensions,
            'calendar_tables': calendar_tables,
            'metrics_tables': metrics_tables,
            'parameter_tables': parameter_tables,
            'calculation_groups': calculation_groups,
            'disconnected_tables': disconnected_tables
        }
        
        dimension_extensions = self.relationship_analyzer.find_dimension_extensions(connections, initial_categories)
        
        # Universal extension repositioning (one level further from facts)
        for extension_table, extension_info in dimension_extensions.items():
            if isinstance(extension_info, dict):
                base_table = extension_info.get('base_table')
                relationship_type = extension_info.get('type', 'extension')
            else:
                # Handle tuple format
                base_table = extension_info[0] if extension_info else None
                relationship_type = extension_info[1] if len(extension_info) > 1 else 'extension'
                
            if relationship_type == 'extension' and base_table:
                # Remove from current position
                for level_list in [l1_dimensions, l2_dimensions, l3_dimensions, l4_plus_dimensions]:
                    if extension_table in level_list:
                        level_list.remove(extension_table)
                
                # Place one level further from facts than base
                if base_table in l1_dimensions:
                    l2_dimensions.append(extension_table)
                    logger.info(f"Extension {extension_table} moved to L2 (base {base_table} in L1)")
                elif base_table in l2_dimensions:
                    l3_dimensions.append(extension_table)
                    logger.info(f"Extension {extension_table} moved to L3 (base {base_table} in L2)")
                elif base_table in l3_dimensions:
                    l4_plus_dimensions.append(extension_table)
                    logger.info(f"Extension {extension_table} moved to L4+ (base {base_table} in L3)")
                else:
                    # Default: assume base is L1, move extension to L2
                    l2_dimensions.append(extension_table)
                    logger.info(f"Extension {extension_table} moved to L2 (default - base {base_table})")
        
        # Final verification
        all_categorized = (
            fact_tables + l1_dimensions + l2_dimensions + l3_dimensions + l4_plus_dimensions + 
            calendar_tables + metrics_tables + parameter_tables + calculation_groups + disconnected_tables
        )
        
        missing_tables = set(table_names) - set(all_categorized)
        if missing_tables:
            logger.warning(f"🚨 MISSING TABLES: {missing_tables} - adding to disconnected")
            disconnected_tables.extend(missing_tables)
        
        logger.info(f"🌍 UNIVERSAL CATEGORIZATION COMPLETE")
        logger.info(f"  Facts: {len(fact_tables)}")
        logger.info(f"  L1-L4+ Dimensions: {len(l1_dimensions + l2_dimensions + l3_dimensions + l4_plus_dimensions)}")
        logger.info(f"  Special tables: {len(calendar_tables + metrics_tables + parameter_tables + calculation_groups)}")
        logger.info(f"  Extensions: {len(dimension_extensions)}")
        logger.info(f"  🗓️ AUTO DATE TABLES EXCLUDED: {len(auto_date_tables)}")
        
        return {
            'fact_tables': fact_tables,
            'l1_dimensions': l1_dimensions,
            'l2_dimensions': l2_dimensions,
            'l3_dimensions': l3_dimensions,
            'l4_plus_dimensions': l4_plus_dimensions,
            'calendar_tables': calendar_tables,
            'metrics_tables': metrics_tables,
            'parameter_tables': parameter_tables,
            'calculation_groups': calculation_groups,
            'disconnected_tables': disconnected_tables,
            'dimension_tables': l1_dimensions + l2_dimensions + l3_dimensions + l4_plus_dimensions,
            'dimension_extensions': dimension_extensions,
            'auto_date_tables': auto_date_tables  # Track excluded tables
        }
