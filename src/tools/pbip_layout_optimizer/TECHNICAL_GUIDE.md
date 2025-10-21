# 🎯 PBIP Layout Optimizer - Technical Architecture & How It Works

**Version**: v2.0.0  
**Built by**: Reid Havens of Analytic Endeavors  
**Last Updated**: October 21, 2025

---

## 📚 Table of Contents

1. [Overview](#overview)
2. [Haven's Middle-Out Design Philosophy](#havens-middle-out-design-philosophy)
3. [Architecture](#architecture)
4. [Core Components](#core-components)
5. [Analysis Engines](#analysis-engines)
6. [Layout Engines](#layout-engines)
7. [Positioning Systems](#positioning-systems)
8. [Universal Table Categorization](#universal-table-categorization)
9. [Data Flow](#data-flow)
10. [File Structure](#file-structure)
11. [Key Algorithms](#key-algorithms)
12. [Future Enhancements](#future-enhancements)

---

## 🎯 Overview

The **PBIP Layout Optimizer** automatically organizes Power BI relationship diagrams using **Haven's Middle-Out Design Philosophy**. It analyzes your data model's structure, categorizes tables intelligently, and positions them for optimal clarity and understanding.

### What It Does

1. **Analyzes Relationships** - Scans TMDL files to understand table connections
2. **Categorizes Tables** - Identifies facts, dimensions, calendars, parameters dynamically
3. **Optimizes Layout** - Positions tables using middle-out philosophy
4. **Handles Complex Models** - Supports snowflake schemas, extensions, calculation groups
5. **No Hardcoding** - Works with any data model, any naming convention

### Key Innovations

- **Universal Logic** - Zero hardcoded table names or model-specific logic
- **Middle-Out Philosophy** - Facts in center, dimensions radiating outward
- **Relationship-Aware** - Uses actual TMDL relationships for intelligent placement
- **Modular Architecture** - Composable analyzers and positioning engines
- **Chain Alignment** - Keeps related table families together
- **Family Grouping** - Groups extensions with their base tables

---

## 💡 Haven's Middle-Out Design Philosophy

### The Philosophy

Traditional BI tools arrange tables randomly or in simple grids. Haven's Middle-Out design organizes models **the way humans think about data**:

1. **Facts in the Center** - Transaction/event tables are the focal point
2. **Dimensions Radiate Outward** - Context tables surround facts by distance
3. **Levels Indicate Distance** - L1 (closest) to L4+ (furthest from facts)
4. **Calendars at Top** - Special time dimensions get prominence
5. **Metrics Separate** - Pure measure tables in their own zone

### Visual Structure

```
                      📅 CALENDAR TABLES
                            (Top)
                              
L4+    L3      L2      L1     FACTS    L1      L2      L3    L4+
Dim    Dim     Dim     Dim    🎯🎯🎯    Dim     Dim     Dim   Dim
Dim    Dim     Dim     Dim    🎯🎯🎯    Dim     Dim     Dim   Dim
       Dim     Dim     Dim    🎯🎯🎯    Dim     Dim     Dim   
              Dim     Dim              Dim     Dim
                                              
                    📊 METRICS / PARAMETERS
                         (Bottom)
```

### Why It Works

- **Cognitive Alignment** - Matches mental model of fact-dimension relationships
- **Distance = Abstraction** - Further from facts = more abstracted/generic
- **Visual Hierarchy** - Important tables (facts) are visually central
- **Relationship Clarity** - Connections are shorter and clearer
- **Scalability** - Works for 10 tables or 1,000 tables

---

## 🏗️ Architecture

### Design Pattern: **Modular Composition with Analyzers & Engines**

The tool uses a sophisticated multi-layer architecture:

```
PBIPLayoutOptimizerTool (BaseTool)
    └── PBIPLayoutOptimizerTab (UI)
            └── EnhancedPBIPLayoutCore (Orchestrator)
                    ├── BaseLayoutEngine (Common Functionality)
                    ├── Analyzers/
                    │   ├── RelationshipAnalyzer (Relationship Graphs)
                    │   └── TableCategorizer (Universal Classification)
                    ├── Engines/
                    │   └── MiddleOutLayoutEngine (Middle-Out Layout)
                    └── Positioning/
                        ├── PositionCalculator (Canvas Positioning)
                        ├── DimensionOptimizer (Side Placement)
                        ├── UniversalChainAlignment (Chain Optimization)
                        ├── ChainAwarePositionGenerator (Final Positions)
                        └── FamilyAwareAlignment (Extension Grouping)
```

### Key Principles

- **Separation of Concerns**: Each component has one clear responsibility
- **Composition Over Inheritance**: Components are composed, not inherited
- **Universal Logic**: No hardcoded table names anywhere
- **Relationship-First**: Uses actual TMDL relationships, not assumptions
- **Modular Positioning**: Swappable positioning strategies

---

## 🧩 Core Components

### 1. **EnhancedPBIPLayoutCore** (`enhanced_layout_core.py`)

**Role**: Main orchestrator that integrates all advanced components

**Key Responsibilities**:
- PBIP folder validation
- Component initialization and lifecycle
- Bridge methods for tool compatibility
- Operation coordination (analyze, optimize)
- Schema validation and persistence

**Critical Methods**:

```python
validate_pbip_folder(pbip_folder) -> Dict[str, Any]
    """Validate PBIP structure and find SemanticModel"""

analyze_table_categorization(pbip_folder) -> Dict[str, Any]
    """Analyze and categorize all tables"""

optimize_layout_with_advanced(...) -> Dict[str, Any]
    """Optimize using middle-out engine"""

analyze_layout_quality(pbip_folder) -> Dict[str, Any]
    """Assess current layout quality and provide recommendations"""
```

**Component Initialization**:

```python
def _initialize_advanced_components(self):
    """Initialize advanced components if available"""
    from .analyzers.table_categorizer import TableCategorizer
    from .analyzers.relationship_analyzer import RelationshipAnalyzer
    from .engines.middle_out_layout_engine import MiddleOutLayoutEngine
    
    # Components created on-demand per operation
    self.mcp_available = True
```

---

### 2. **BaseLayoutEngine** (`base_layout_engine.py`)

**Role**: Base class providing common PBIP file operations

**Key Capabilities**:
- Find SemanticModel path
- Parse TMDL files
- Read/write diagramLayout.json
- Table name normalization
- File system utilities

**Critical Methods**:

```python
_find_semantic_model_path() -> Optional[Path]
    """Find .SemanticModel folder in PBIP"""

_find_tmdl_files() -> Dict[str, Path]
    """Find all TMDL files (tables and model)"""

_get_table_names_from_tmdl() -> List[str]
    """Extract normalized table names"""

_parse_diagram_layout() -> Dict[str, Any]
    """Read diagramLayout.json"""

_save_diagram_layout(layout_data) -> bool
    """Write diagramLayout.json"""
```

**Name Normalization**:

```python
def _normalize_table_name(self, table_name: str) -> str:
    """Handle special characters and encoding"""
    decoded = urllib.parse.unquote(table_name)
    normalized = decoded.strip().strip("'").strip('"')
    return normalized
```

---

## 🔍 Analysis Engines

### 1. **RelationshipAnalyzer** (`analyzers/relationship_analyzer.py`)

**Role**: Build and analyze the relationship graph from TMDL files

**Key Capabilities**:
- Parse relationships from model.tmdl
- Build bidirectional relationship graph
- Calculate table distances from facts
- Identify star vs snowflake patterns
- Detect dimension extensions (1:1 relationships)

**Critical Methods**:

```python
build_relationship_graph() -> Dict[str, Set[str]]
    """Build complete relationship graph from TMDL"""

calculate_distance_to_facts(table, connections, facts) -> int
    """Calculate minimum distance from table to any fact"""

is_star_schema_table(table, connections, facts) -> bool
    """Check if table connects directly to facts"""

find_dimension_extensions(connections, categories) -> Dict
    """Identify 1:1 extension tables"""
```

**Relationship Parsing**:

```python
def _parse_relationships_from_model_tmdl(self) -> List[Dict]:
    """Parse relationship definitions from model.tmdl"""
    
    # Parse relationship blocks:
    # relationship <name> {
    #     fromColumn: Table1[Column]
    #     toColumn: Table2[Column]
    #     ...
    # }
```

**Distance Calculation** (BFS Algorithm):

```python
def calculate_distance_to_facts(table, connections, facts):
    """Breadth-first search to find shortest path to any fact"""
    if table in facts:
        return 0
    
    visited = {table}
    queue = [(table, 0)]
    
    while queue:
        current, distance = queue.pop(0)
        neighbors = connections.get(current, set())
        
        for neighbor in neighbors:
            if neighbor in visited:
                continue
            if neighbor in facts:
                return distance + 1
            
            visited.add(neighbor)
            queue.append((neighbor, distance + 1))
    
    return 99  # Disconnected
```

---

### 2. **TableCategorizer** (`analyzers/table_categorizer.py`)

**Role**: Universal table categorization with zero hardcoding

**Key Capabilities**:
- Identify auto-date tables (Power BI generated)
- Classify tables as fact vs dimension
- Detect calendar, metrics, parameter tables
- Identify calculation groups
- Handle disconnected tables
- Snowflake level assignment (L1-L4+)

**Universal Principles**:

1. **Pattern-Based Detection** - Uses naming patterns, not specific names
2. **Relationship-Based Logic** - Uses connection counts and structure
3. **TMDL Content Analysis** - Reads file structure and properties
4. **Multi-Language Support** - Works with any language's naming conventions
5. **No Model Assumptions** - Works with finance, healthcare, retail, etc.

**Critical Methods**:

```python
identify_auto_date_tables(table_names) -> List[str]
    """Identify Power BI auto date/time tables to exclude"""

calculate_table_score(table, connections) -> Dict[str, Any]
    """Hybrid scoring: naming + connection patterns"""

categorize_tables(table_names, connections) -> Dict[str, List[str]]
    """Universal categorization into all categories"""

identify_metrics_tables(table_names, connections) -> List[str]
    """Detect pure measure tables (disconnected, 0-1 columns)"""

identify_parameter_tables(table_names) -> List[str]
    """Detect parameter/config tables"""
```

**Fact vs Dimension Scoring**:

```python
def calculate_table_score(table, connections):
    """Hybrid scoring algorithm"""
    
    # Universal fact indicators
    fact_indicators = [
        'fact', 'trans', 'event', 'activity', 'record', 
        'log', 'history', 'operation', 'process'
    ]
    
    # Universal dimension indicators
    dim_indicators = [
        'dim', 'dimension', 'master', 'lookup', 
        'reference', 'category', 'type'
    ]
    
    # Connection-based scoring
    if connection_count >= 5: fact_score += 15
    elif connection_count >= 3: fact_score += 10
    elif connection_count == 0: classification = 'disconnected'
    
    # Priority rules:
    # 1. Explicit naming (fact_, dim_)
    # 2. Connection count >= 3 → fact
    # 3. Hybrid score comparison
```

**Auto-Date Detection**:

```python
def identify_auto_date_tables(table_names):
    """Universal auto-date detection"""
    
    # Naming patterns
    if name.startswith('DateTableTemplate_') or 
       name.startswith('LocalDateTable_'):
        return True
    
    # TMDL content checks
    if '__PBI_TemplateDateTable = true' in content:
        return True
    
    # Hidden + date hierarchy + calendar source
    if is_hidden and has_date_hierarchy and has_calendar_source:
        return True
```

---

## 🎨 Layout Engines

### 1. **MiddleOutLayoutEngine** (`engines/middle_out_layout_engine.py`)

**Role**: Implement Haven's middle-out design philosophy

**Key Capabilities**:
- Generate middle-out layout positions
- Handle L1-L4+ dimension levels
- Position facts in center
- Place calendar tables at top
- Manage special tables (metrics, parameters)
- Apply chain alignment
- Handle family grouping for extensions

**Spacing Configuration**:

```python
spacing_config = {
    'table_width': 200,
    'base_collapsed_height': 104,
    'height_per_relationship': 24,
    'expanded_height': 180,
    'calendar_table_height': 140,
    'within_stack_spacing': 15,      # Within same category
    'table_grid_height': 210,
    'stack_gap': 150,                # Between categories (CONSISTENT)
    'calendar_spacing': 80,
    'left_margin': 50
}
```

**Layout Generation Process**:

```python
def generate_middle_out_layout(canvas_width, canvas_height):
    """Main layout generation algorithm"""
    
    # 1. Get table names from TMDL
    table_names = base_engine._get_table_names_from_tmdl()
    
    # 2. Build relationship graph
    connections = relationship_analyzer.build_relationship_graph()
    
    # 3. Categorize tables (universal logic)
    categorized = table_categorizer.categorize_tables(
        table_names, connections
    )
    
    # 4. Optimize dimension placement (left vs right)
    left_l1, right_l1 = dimension_optimizer.optimize_dimension_placement(
        l1_dimensions, fact_tables, connections, calendar_tables
    )
    
    # 5. Place L2-L4+ dimensions near their parents
    left_l2, right_l2 = dimension_optimizer.place_l2_dimensions_near_l1(...)
    left_l3, right_l3 = dimension_optimizer.place_l3_dimensions_near_l2(...)
    left_l4, right_l4 = dimension_optimizer.place_l4_dimensions_near_l3(...)
    
    # 6. Reposition extensions one level further
    _reposition_extensions(extensions, l1, l2, l3, l4)
    
    # 7. Apply opposite-side placement for 1:1 relationships
    categorized = dimension_optimizer.apply_opposite_side_placement(...)
    
    # 8. Apply universal chain alignment
    aligned_stacks = universal_chain_alignment.optimize_universal_stack_alignment(
        all_dimension_stacks, connections
    )
    
    # 9. Apply family grouping for extensions
    aligned_stacks = enhance_alignment_with_family_grouping(
        aligned_stacks, extensions, connections
    )
    
    # 10. Calculate X positions on canvas
    positions_map = position_calculator.calculate_canvas_positions(
        categorized_with_splits, canvas_width
    )
    
    # 11. Generate final Y positions with chain awareness
    positions = chain_aware_position_generator.generate_chain_aligned_positions(
        aligned_stacks, positions_map, connections, ...
    )
    
    return positions
```

---

## 📐 Positioning Systems

### 1. **PositionCalculator** (`positioning/position_calculator.py`)

**Role**: Calculate X positions for table categories on canvas

**Algorithm**:

```python
def calculate_canvas_positions(categorized, canvas_width):
    """Calculate X position for each category"""
    
    # Category order (left to right):
    # L4+ Left → L3 Left → L2 Left → L1 Left → Facts → 
    # L1 Right → L2 Right → L3 Right → L4+ Right
    
    # Start position
    current_x = left_margin
    
    # For each category:
    positions_map[category] = current_x
    current_x += table_width + stack_gap
    
    return positions_map
```

**Consistent Spacing**: All categories separated by 150px (stack_gap)

---

### 2. **DimensionOptimizer** (`positioning/dimension_optimizer.py`)

**Role**: Optimize dimension placement (left vs right side)

**Key Capabilities**:
- Split L1 dimensions between left and right
- Place L2 dimensions near their L1 parents
- Place L3 dimensions near their L2 parents
- Place L4+ dimensions near their L3 parents
- Handle opposite-side placement for 1:1 relationships
- Identify additional L2 tables from connections

**L1 Placement Algorithm**:

```python
def optimize_dimension_placement(l1_dims, facts, connections, calendars):
    """Split L1 dimensions optimally"""
    
    # Score each dimension
    scores = []
    for dim in l1_dims:
        left_score = 0
        right_score = 0
        
        # Prefer calendar-connected on left
        if connects_to_calendar(dim, calendars, connections):
            left_score += 10
        
        # Prefer high-connectivity on sides
        conn_count = len(connections.get(dim, set()))
        if conn_count >= 5:
            right_score += 5  # High connectivity → right
        
        scores.append((dim, left_score, right_score))
    
    # Sort by left_score - right_score
    scores.sort(key=lambda x: x[2] - x[1])
    
    # Split: top half → left, bottom half → right
    midpoint = len(scores) // 2
    left_dims = [s[0] for s in scores[:midpoint]]
    right_dims = [s[0] for s in scores[midpoint:]]
    
    return left_dims, right_dims
```

**L2 Parent-Based Placement**:

```python
def place_l2_dimensions_near_l1(l2_dims, left_l1, right_l1, connections):
    """Place L2s on same side as their L1 parents"""
    
    left_l2 = []
    right_l2 = []
    
    for l2_table in l2_dims:
        l2_connections = connections.get(l2_table, set())
        
        # Count connections to left vs right L1s
        left_l1_connections = l2_connections.intersection(set(left_l1))
        right_l1_connections = l2_connections.intersection(set(right_l1))
        
        # Place with majority connections
        if len(left_l1_connections) > len(right_l1_connections):
            left_l2.append(l2_table)
        elif len(right_l1_connections) > len(left_l1_connections):
            right_l2.append(l2_table)
        else:
            # Tie: default to left
            left_l2.append(l2_table)
    
    return left_l2, right_l2
```

---

### 3. **UniversalChainAlignment** (`positioning/universal_chain_alignment.py`)

**Role**: Optimize vertical alignment of related table chains

**Purpose**: Keep related tables (connected tables) vertically aligned across stacks

**Algorithm**:

```python
def optimize_universal_stack_alignment(all_stacks, connections):
    """Optimize vertical alignment across all stacks"""
    
    # 1. Identify chains (sequences of connected tables)
    chains = identify_relationship_chains(all_stacks, connections)
    
    # 2. For each chain, align tables across stacks
    for chain in chains:
        align_chain_across_stacks(chain, all_stacks)
    
    # 3. Reserve positions for aligned tables
    reserved_positions = {}
    
    # 4. Reorganize each stack to respect reservations
    for stack_name, stack_tables in all_stacks.items():
        reorganized = apply_reserved_positions(
            stack_tables, reserved_positions
        )
        all_stacks[stack_name] = reorganized
    
    return all_stacks
```

**Chain Identification**:

```python
def identify_relationship_chains(all_stacks, connections):
    """Find sequences of connected tables across stacks"""
    
    chains = []
    processed_tables = set()
    
    # For each table, trace its connections
    for stack_name, tables in all_stacks.items():
        for table in tables:
            if table in processed_tables:
                continue
            
            # Build chain by following connections
            chain = trace_chain(table, connections, all_stacks)
            
            if len(chain) > 1:
                chains.append(chain)
                processed_tables.update(chain)
    
    return chains
```

---

### 4. **ChainAwarePositionGenerator** (`positioning/chain_aware_position_generator.py`)

**Role**: Generate final Y positions with chain alignment awareness

**Key Capabilities**:
- Generate Y positions for all tables
- Respect reserved positions from chain alignment
- Handle within-stack spacing
- Position special tables (calendar, metrics, parameters)
- Apply consistent spacing

**Position Generation**:

```python
def generate_chain_aligned_positions(aligned_stacks, positions_map, 
                                   connections, ...):
    """Generate final positions with chain awareness"""
    
    positions = []
    
    # 1. Position main dimension/fact stacks
    for stack_name in stack_order:
        x_position = positions_map[stack_name]
        y_position = top_margin
        
        tables = aligned_stacks[stack_name]
        
        for table in tables:
            # Check if position is reserved by chain alignment
            if table in reserved_positions:
                y_position = reserved_positions[table]
            
            # Create position entry
            positions.append({
                'nodeIndex': table,
                'location': {'x': x_position, 'y': y_position},
                'size': {'width': 200, 'height': 104},
                'zIndex': len(positions)
            })
            
            # Advance Y position
            y_position += table_height + within_stack_spacing
    
    # 2. Position calendar tables (top)
    position_calendar_tables(...)
    
    # 3. Position metrics tables (bottom right)
    position_metrics_tables(...)
    
    # 4. Position parameter grid (bottom left)
    position_parameter_grid(...)
    
    return positions
```

---

### 5. **FamilyAwareAlignment** (`positioning/family_aware_alignment.py`)

**Role**: Group extension tables with their base tables

**Purpose**: Keep 1:1 extension tables adjacent to their parent tables

**Algorithm**:

```python
def enhance_alignment_with_family_grouping(aligned_stacks, extensions, 
                                          connections):
    """Group extensions with base tables"""
    
    for stack_name, tables in aligned_stacks.items():
        # Find extensions in this stack
        extensions_in_stack = [t for t in tables if t in extensions]
        
        if not extensions_in_stack:
            continue
        
        # Group by family
        families = {}
        for ext_table in extensions_in_stack:
            base_table = extensions[ext_table]['base_table']
            if base_table not in families:
                families[base_table] = []
            families[base_table].append(ext_table)
        
        # Reorganize stack to keep families together
        reorganized = []
        for table in tables:
            reorganized.append(table)
            
            # Add extensions immediately after base
            if table in families:
                reorganized.extend(sorted(families[table]))
        
        aligned_stacks[stack_name] = reorganized
    
    return aligned_stacks
```

---

## 🌍 Universal Table Categorization

### Categorization Process

**Four-Phase Approach**:

1. **Phase 0: Filter Auto-Date Tables**
   - Remove Power BI-generated tables
   - Exclude from layout entirely

2. **Phase 1: Special Table Identification**
   - Parameters (hidden, config, settings)
   - Calculation Groups (TMDL property)
   - Metrics Tables (disconnected, 0-1 columns, measure-focused)

3. **Phase 2: Fact Identification**
   - Hybrid scoring (naming + connections)
   - Connection count >= 3 → likely fact
   - Disconnected tables excluded

4. **Phase 3: Dimension Classification**
   - Calculate distance to facts (BFS)
   - L1 = distance 1 (star schema)
   - L2 = distance 2 (snowflake)
   - L3 = distance 3 (deeper snowflake)
   - L4+ = distance 4+ (very deep)

5. **Phase 4: Extension Handling**
   - Identify 1:1 relationships
   - Reposition extensions one level further
   - Group extensions with base tables

### Category Definitions

| Category | Definition | Placement |
|----------|-----------|-----------|
| **Fact Tables** | High connectivity (3+), transaction-focused | Center |
| **L1 Dimensions** | Direct connection to facts (distance 1) | Left/Right of Facts |
| **L2 Dimensions** | One hop from facts (distance 2) | Outside L1 |
| **L3 Dimensions** | Two hops from facts (distance 3) | Outside L2 |
| **L4+ Dimensions** | Three+ hops from facts (distance 4+) | Outermost |
| **Calendar Tables** | Date/time/period focused | Top of canvas |
| **Metrics Tables** | Disconnected, pure measures | Bottom right |
| **Parameter Tables** | Config/lookup, disconnected | Bottom left grid |
| **Calculation Groups** | TMDL calculation group property | Bottom left grid |
| **Extensions** | 1:1 relationships, one level further than base | Adjacent to base |

---

## 📊 Data Flow

### Complete Workflow

```
┌─────────────────────┐
│   User Selects      │
│   PBIP Folder       │
└──────────┬──────────┘
           │
           ▼
┌─────────────────────┐
│   Validate PBIP     │
│   Structure         │ ──► Check .SemanticModel exists
└──────────┬──────────┘     Verify definition/ folder
           │                Find diagramLayout.json
           ▼
┌─────────────────────┐
│   Find TMDL Files   │ ──► Scan tables/ directory
└──────────┬──────────┘     Read model.tmdl
           │                Normalize table names
           ▼
┌─────────────────────┐
│   Build             │
│   Relationship      │ ──► Parse relationship blocks
│   Graph             │     Create bidirectional graph
└──────────┬──────────┘     Map table connections
           │
           ▼
┌─────────────────────┐
│   Categorize        │
│   Tables            │ ──► Phase 0: Filter auto-date
└──────────┬──────────┘     Phase 1: Special tables
           │                Phase 2: Facts
           ▼                Phase 3: Dimensions
┌─────────────────────┐     Phase 4: Extensions
│   Calculate Hybrid  │
│   Scores            │ ──► Naming pattern analysis
└──────────┬──────────┘     Connection count analysis
           │                TMDL content analysis
           ▼
┌─────────────────────┐
│   Optimize          │
│   Dimension         │ ──► Split L1 (left/right)
│   Placement         │     Place L2 near L1 parents
└──────────┬──────────┘     Place L3 near L2 parents
           │                Place L4 near L3 parents
           ▼
┌─────────────────────┐
│   Identify & Handle │
│   Extensions        │ ──► Find 1:1 relationships
└──────────┬──────────┘     Reposition one level further
           │                Group with base tables
           ▼
┌─────────────────────┐
│   Apply Opposite-   │
│   Side Placement    │ ──► Find bidirectional pairs
└──────────┬──────────┘     Place on opposite sides
           │
           ▼
┌─────────────────────┐
│   Universal Chain   │
│   Alignment         │ ──► Identify chains
└──────────┬──────────┘     Align vertically
           │                Reserve positions
           ▼
┌─────────────────────┐
│   Family-Aware      │
│   Grouping          │ ──► Group extensions with base
└──────────┬──────────┘     Keep families adjacent
           │
           ▼
┌─────────────────────┐
│   Calculate X       │
│   Positions         │ ──► Consistent 150px spacing
└──────────┬──────────┘     L4+ → L3 → L2 → L1 → Facts → L1 → L2 → L3 → L4+
           │
           ▼
┌─────────────────────┐
│   Generate Y        │
│   Positions         │ ──► Chain-aware positioning
└──────────┬──────────┘     Within-stack spacing
           │                Special table zones
           ▼
┌─────────────────────┐
│   Create            │
│   diagramLayout     │ ──► Format positions as nodes
│   JSON              │     Preserve existing structure
└──────────┬──────────┘     Set hideKeyFieldsWhenCollapsed
           │
           ▼
┌─────────────────────┐
│   Save to           │
│   diagramLayout.    │ ──► Write to .SemanticModel/
│   json              │
└─────────────────────┘
```

---

## 📁 File Structure

### Tool Directory Layout

```
pbip_layout_optimizer/
├── tool.py                          # BaseTool implementation
├── base_layout_engine.py            # Base class with common operations
├── enhanced_layout_core.py          # Main orchestrator
├── layout_ui.py                     # User interface
├── analyzers/                       # Analysis components
│   ├── __init__.py
│   ├── relationship_analyzer.py    # Relationship graph building
│   └── table_categorizer.py        # Universal table classification
├── engines/                         # Layout engines
│   ├── __init__.py
│   └── middle_out_layout_engine.py # Middle-out layout implementation
├── positioning/                     # Positioning systems
│   ├── __init__.py
│   ├── position_calculator.py      # X position calculation
│   ├── dimension_optimizer.py      # Dimension placement optimization
│   ├── universal_chain_alignment.py # Chain alignment
│   ├── chain_aware_position_generator.py # Y position generation
│   └── family_aware_alignment.py   # Extension family grouping
└── TECHNICAL_GUIDE.md              # This document
```

### Separation of Concerns

- **`tool.py`**: Tool registration and BaseTool interface
- **`base_layout_engine.py`**: Common PBIP file operations
- **`enhanced_layout_core.py`**: Orchestration and integration
- **`layout_ui.py`**: User interface (no business logic)
- **`analyzers/`**: Analysis logic (relationships, categorization)
- **`engines/`**: Layout generation strategies
- **`positioning/`**: Position calculation algorithms

---

## 🔧 Key Algorithms

### Algorithm 1: Relationship Graph Building

**Purpose**: Create bidirectional graph of all table relationships

**Process**:

```python
def build_relationship_graph():
    """Parse model.tmdl and build graph"""
    
    graph = defaultdict(set)
    
    # 1. Read model.tmdl
    model_content = read_file('model.tmdl')
    
    # 2. Parse relationship blocks
    for relationship_block in parse_relationships(model_content):
        from_table = extract_table_name(relationship_block['fromColumn'])
        to_table = extract_table_name(relationship_block['toColumn'])
        
        # 3. Add bidirectional connections
        graph[from_table].add(to_table)
        graph[to_table].add(from_table)
    
    return dict(graph)
```

---

### Algorithm 2: Hybrid Table Scoring

**Purpose**: Classify tables as fact or dimension using multiple signals

**Scoring Factors**:

1. **Naming Patterns** (25-50 points)
   - Explicit "fact_" prefix → +25 fact points
   - Explicit "dim_" prefix → +25 dim points
   - Keywords (transaction, event, etc.) → +15 fact points
   - Keywords (master, lookup, etc.) → +15 dim points

2. **Connection Count** (10-15 points)
   - 5+ connections → +15 fact points
   - 3-4 connections → +10 fact points
   - 2 connections → +5 fact points
   - 1 connection → -3 points
   - 0 connections → disconnected (special handling)

3. **Priority Rules**:
   - Explicit naming overrides everything
   - Connection count >= 3 → likely fact
   - Ties resolved by score comparison

**Example Scoring**:

```
Table: "FactSales"
  - Explicit fact naming: +25
  - 5 connections: +15
  - Total: 40 points → FACT

Table: "DimProduct"
  - Explicit dim naming: +25
  - 2 connections: +5
  - Total: 25 dim points → DIMENSION

Table: "Customer"
  - No explicit naming: 0
  - 4 connections: +10 fact
  - Total: 10 fact points → FACT (borderline)
```

---

### Algorithm 3: BFS Distance Calculation

**Purpose**: Calculate minimum distance from table to any fact table

**Process**:

```python
def calculate_distance_to_facts(table, connections, facts):
    """Breadth-first search to find shortest path"""
    
    if table in facts:
        return 0  # Table IS a fact
    
    visited = {table}
    queue = [(table, 0)]  # (table, distance)
    
    while queue:
        current_table, current_distance = queue.pop(0)
        
        # Get all connected tables
        neighbors = connections.get(current_table, set())
        
        for neighbor in neighbors:
            if neighbor in visited:
                continue  # Already processed
            
            if neighbor in facts:
                return current_distance + 1  # Found a fact!
            
            visited.add(neighbor)
            queue.append((neighbor, current_distance + 1))
    
    return 99  # Disconnected from all facts
```

**Complexity**: O(V + E) where V = tables, E = relationships

---

### Algorithm 4: L1-L2-L3-L4 Dimension Placement

**Purpose**: Place dimension levels on appropriate sides based on parent relationships

**L2 Placement** (near L1 parents):

```python
def place_l2_dimensions_near_l1(l2_dims, left_l1, right_l1, connections):
    """Place L2s on same side as their L1 connections"""
    
    left_l2 = []
    right_l2 = []
    
    for l2_table in l2_dims:
        # Count connections to each side
        left_connections = count_connections_to(l2_table, left_l1, connections)
        right_connections = count_connections_to(l2_table, right_l1, connections)
        
        # Place with majority
        if left_connections > right_connections:
            left_l2.append(l2_table)
        elif right_connections > left_connections:
            right_l2.append(l2_table)
        else:
            # Tie: default to left
            left_l2.append(l2_table)
    
    return left_l2, right_l2
```

**L3 Placement** (near L2 parents):
- Same algorithm, but looks at L2 connections instead of L1

**L4+ Placement** (near L3 parents):
- Same algorithm, looks at L3 connections (fallback to L2 if no L3 connections)

---

### Algorithm 5: Chain Alignment

**Purpose**: Keep related tables vertically aligned across stacks

**Process**:

```python
def optimize_universal_stack_alignment(all_stacks, connections):
    """Align related tables vertically"""
    
    # 1. Identify chains
    chains = []
    for stack_name, tables in all_stacks.items():
        for table in tables:
            chain = trace_chain_forward_and_backward(
                table, connections, all_stacks
            )
            if len(chain) >= 2:
                chains.append(chain)
    
    # 2. Reserve positions for chain alignment
    reserved_positions = {}
    
    for chain in chains:
        # Find "anchor" table (usually the fact or L1)
        anchor = find_anchor_table(chain)
        anchor_position = get_position_in_stack(anchor, all_stacks)
        
        # Align other tables in chain to same Y position
        for table in chain:
            if table != anchor:
                reserved_positions[table] = anchor_position
    
    # 3. Reorganize stacks to respect reservations
    for stack_name, tables in all_stacks.items():
        reorganized = []
        
        # Sort by reserved position (if any), then original order
        sorted_tables = sort_by_reserved_position(
            tables, reserved_positions
        )
        
        all_stacks[stack_name] = sorted_tables
    
    return all_stacks
```

---

## 🚀 Future Enhancements

### Potential Improvements

1. **Interactive Layout Editor**:
   - Drag-and-drop table positioning
   - Manual override of auto-categorization
   - Real-time relationship preview
   - Layout templates and presets

2. **Advanced Optimization**:
   - Minimize relationship line crossings
   - Optimize for specific relationship patterns
   - Multi-objective optimization (clarity + compactness)
   - Genetic algorithm for large models

3. **Layout Quality Metrics**:
   - Quantitative clarity score
   - Relationship crossing analysis
   - Density metrics
   - Cognitive load estimation

4. **Multiple Layout Modes**:
   - Circular layout
   - Hierarchical tree layout
   - Force-directed graph layout
   - Hybrid layouts for different model types

5. **Model Analysis Reports**:
   - Relationship complexity analysis
   - Schema pattern detection
   - Optimization recommendations
   - Model health metrics

6. **Integration Features**:
   - Export layouts as images/PDFs
   - Share layout configurations
   - Version control for layouts
   - Collaborative layout editing

7. **Performance Optimization**:
   - Parallel processing for large models
   - Incremental layout updates
   - Caching and memoization
   - Optimized graph algorithms

8. **Enhanced Categorization**:
   - Machine learning for classification
   - User feedback incorporation
   - Industry-specific rules
   - Custom categorization plugins

---

## 📝 Code Quality Notes

### Strengths

- ✅ **Modular architecture** - Highly composable components
- ✅ **Universal logic** - Zero hardcoded names
- ✅ **Comprehensive docstrings**
- ✅ **Type hints throughout**
- ✅ **Relationship-first approach**
- ✅ **Clean separation of concerns**
- ✅ **Sophisticated positioning algorithms**
- ✅ **Proper error handling**
- ✅ **Logging and debugging support**

### Standards Followed

- **PEP 8**: Python style guide compliance
- **Type Hints**: Full coverage
- **Docstrings**: NumPy/Google style
- **Composition**: Over inheritance
- **Graph Algorithms**: BFS, DFS for analysis
- **Modular Design**: Swappable components

---

## 🎓 Learning Resources

### Understanding PBIP Structure

- **PBIP Format**: Folder-based Power BI project format
- **.SemanticModel/**: Contains the data model
- **definition/**: Model definition files
- **tables/**: Individual table TMDL files
- **model.tmdl**: Relationships and model-level properties
- **diagramLayout.json**: Visual positioning data

### Key Power BI Concepts

- **Fact Tables**: Transaction/event tables with measures
- **Dimension Tables**: Descriptive/context tables
- **Star Schema**: Facts connect directly to dimensions
- **Snowflake Schema**: Dimensions connect to other dimensions
- **Relationships**: Foreign key connections between tables
- **Cardinality**: 1:1, 1:Many, Many:Many relationships

### Graph Theory

- **Directed Graph**: Edges have direction
- **Bidirectional Graph**: Edges work both ways
- **BFS**: Breadth-first search for shortest paths
- **Connected Components**: Groups of connected nodes
- **Distance**: Minimum hops between nodes

---

## 📞 Support

For questions about this tool's architecture or implementation:

- **Documentation**: This file and inline code comments
- **Issues**: Report bugs via GitHub Issues
- **Professional Support**: [Analytic Endeavors](https://www.analyticendeavors.com)

---

**Document Version**: 1.0  
**Tool Version**: v2.0.0  
**Last Updated**: October 21, 2025  
**Author**: Reid Havens, Analytic Endeavors
