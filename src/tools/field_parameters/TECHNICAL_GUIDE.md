# Field Parameters Tool - Technical Guide

**Tool Version:** 1.0.0
**Last Updated:** January 2026
**Built by:** Reid Havens of Analytic Endeavors

---

## Table of Contents

1. [Overview](#overview)
2. [Architecture](#architecture)
3. [Core Components](#core-components)
4. [Data Models](#data-models)
5. [TMDL Generation](#tmdl-generation)
6. [Category System](#category-system)
7. [Parser Logic](#parser-logic)
8. [UI Panels](#ui-panels)
9. [Usage Workflow](#usage-workflow)
10. [Best Practices](#best-practices)
11. [Troubleshooting](#troubleshooting)
12. [Limitations](#limitations)

---

## Overview

### What is the Field Parameters Tool?

The **Field Parameters Tool** provides a comprehensive visual interface for creating and editing Power BI Field Parameters. Instead of manually writing DAX code, users can drag-and-drop measures and columns, assign them to multi-level categories, and automatically generate the proper TMDL/Tabular script code.

### Why This Tool Exists

Field Parameters are a powerful Power BI feature that allows users to dynamically switch between different measures or columns in visuals. However, the native UI for creating them is limited:

- Only accessible once during creation
- No editing UI after initial setup
- Single-level categories only
- Manual DAX coding required for modifications

This tool solves these problems by:
- Providing a persistent, intuitive UI for parameter creation
- Supporting multi-level category hierarchies
- Enabling drag-and-drop reordering for sort order control
- Generating clean, properly-formatted TMDL code
- Preserving extra calculated columns when editing existing parameters

### Key Features

| Feature | Description |
|---------|-------------|
| Model Connection | Auto-connects from Power BI Desktop or manual selection |
| Create/Edit Modes | Start fresh or modify existing parameters |
| Drag-and-Drop | Add fields and reorder with visual feedback |
| Multi-Level Categories | Unlimited category hierarchy levels |
| Real-Time Preview | See generated TMDL as you work |
| Lineage Tag Support | Preserve version control metadata |
| Extra Column Preservation | Maintains custom calculated columns |

---

## Architecture

### File Structure

```
tools/field_parameters/
├── __init__.py                      # Package initialization
├── field_parameters_tool.py         # BaseTool implementation
├── field_parameters_core.py         # Parser & generator logic
├── field_parameters_ui.py           # Main tab orchestrator
├── field_parameters_connector.py    # Shim for backward compat
├── field_parameters_builder.py      # Parameter builder panel
├── models.py                        # Data model re-exports
├── panels/                          # UI panel components
│   ├── model_connection_panel.py    # Model selection & connection
│   ├── parameter_config_panel.py    # New/edit mode configuration
│   ├── available_fields_panel.py    # Browse model fields
│   ├── category_manager.py          # Category hierarchy management
│   └── tmdl_preview_panel.py        # Generated code display
├── dialogs/                         # Dialog components
│   ├── connection_dialog.py         # Advanced connection options
│   ├── add_label.py                 # Add category label dialog
│   ├── category_label_editor.py     # Edit labels within category
│   └── field_category_editor.py     # Bulk category assignment editor
├── widgets/                         # Reusable UI widgets
│   └── scrollbar.py                 # Auto-hiding scrollbar
├── FIELD_PARAMETERS_TOOL.md         # Extended documentation
└── TECHNICAL_GUIDE.md               # This file
```

### Component Hierarchy

```
FieldParametersTab (orchestrator)
├── ModelConnectionPanel       # Model selection & connection
├── ParameterConfigPanel       # New/edit mode, naming, options
├── AvailableFieldsPanel       # Browse model fields
├── ParameterBuilderPanel      # Drag-drop field list
├── CategoryManagerPanel       # Category hierarchy management
└── TmdlPreviewPanel           # Generated code display
```

### Data Flow

```
User Action              Internal Process                    Output
-----------              ----------------                    ------
Connect to Model    →    Load tables/fields            →    Enable UI
                         from semantic model

Create/Load Param   →    Initialize FieldParameter     →    Populate builder
                         object

Add Fields          →    Create FieldItem objects      →    Update builder list
                         with NAMEOF references

Assign Categories   →    Update CategoryLevel list     →    Refresh dropdowns
                         for each field

Reorder Items       →    Update order_within_group     →    Regenerate TMDL

Any Change          →    FieldParameterGenerator       →    Update preview
                         regenerates code

Copy Code           →    Clipboard                     →    Paste into
                                                            Tabular Editor
```

---

## Core Components

### 1. FieldParameterParser (field_parameters_core.py)

**Role**: Parses existing TMDL content to extract field parameter structure.

**Key Methods**:
```python
parse_tmdl(self, tmdl_content: str) -> Optional[FieldParameter]
    # Main entry point - extracts parameter from TMDL

_extract_table_name(self, tmdl_content: str) -> Optional[str]
    # Gets table name from 'table "Name"' declaration

_extract_partition_content(self, tmdl_content: str) -> Optional[str]
    # Extracts the calculated partition DAX expression

_parse_tuple_items(self, partition_content: str) -> List[FieldItem]
    # Parses individual field items from tuples

_extract_column_definitions(self, tmdl_content: str) -> Dict[str, Dict]
    # Gets column metadata including lineage tags
```

**Parsing Logic**:
1. Extract table name from TMDL header
2. Find partition source block
3. Parse each tuple into FieldItem
4. Extract column definitions for lineage tags
5. Detect category levels from column count

### 2. FieldParameterGenerator (field_parameters_core.py)

**Role**: Generates TMDL code from FieldParameter objects.

**Key Methods**:
```python
generate_tmdl(self, parameter: FieldParameter) -> str
    # Main entry point - generates complete TMDL

_generate_columns(self, parameter: FieldParameter) -> str
    # Creates column definitions with proper metadata

_generate_partition(self, parameter: FieldParameter) -> str
    # Creates partition source with field tuples

_format_tuple(self, item: FieldItem, categories: List[CategoryLevel]) -> str
    # Formats single field item as TMDL tuple
```

**Generation Order**:
1. `createOrReplace` header
2. Table declaration with annotation
3. Column definitions (display, fields, order, categories)
4. Extra columns (if preserving)
5. Partition with calculated source
6. Field tuples with category assignments

### 3. FieldParametersTab (field_parameters_ui.py)

**Role**: Main UI orchestrator coordinating all panels.

**Key Responsibilities**:
- Initialize and layout panels
- Manage state between panels
- Coordinate events (field added, category changed, etc.)
- Trigger TMDL regeneration on changes
- Handle model connection lifecycle

---

## Data Models

### FieldParameter

The main container for all parameter data:

```python
@dataclass(slots=True)
class FieldParameter:
    table_name: str                          # '.Parameter - Sales Metrics'
    parameter_name: str                      # 'Sales Metrics'
    fields: List[FieldItem]                  # All field items
    category_levels: List[CategoryLevel]     # Category hierarchy
    keep_lineage_tags: bool = True           # Preserve IDs for editing
    lineage_tags: Dict[str, str] = field(default_factory=dict)
    pbi_id: Optional[str] = None             # Power BI table annotation
    extra_columns: List[Dict[str, Any]] = field(default_factory=list)
```

### FieldItem

Represents a single field in the parameter:

```python
@dataclass(slots=True)
class FieldItem:
    display_name: str                        # 'Total Sales'
    field_reference: str                     # "NAMEOF('Sales'[Amount])"
    table_name: str                          # 'Sales'
    field_name: str                          # 'Amount'
    order_within_group: int = 1              # Sort order in category
    original_order_within_group: int = 1     # Original order (for reset)
    categories: List[Tuple[int, str]] = field(default_factory=list)
                                             # [(10, 'Revenue'), (20, 'Actual')]
```

### CategoryLevel

Defines a category hierarchy level:

```python
@dataclass(slots=True)
class CategoryLevel:
    name: str                                # 'Metric Type'
    sort_order: int                          # 1, 2, 3...
    column_name: str                         # 'Metric Type Category'
```

### Model Connection Types (from core/pbi_connector.py)

```python
@dataclass
class ModelConnection:
    port: int
    model_name: str
    is_connected: bool
    connection_string: str

@dataclass
class TableInfo:
    name: str
    is_hidden: bool
    table_type: str  # 'Regular', 'Calculated', etc.

@dataclass
class FieldInfo:
    name: str
    field_type: str  # 'measure', 'column'
    data_type: str
    is_hidden: bool
    expression: Optional[str] = None

@dataclass
class TableFieldsInfo:
    table_name: str
    fields: List[FieldInfo]
```

---

## TMDL Generation

### Basic Parameter Structure (No Categories)

```tmdl
createOrReplace
    table '.Parameter - Sales Metrics'
        annotation PBI_Id = "guid-here"

        column 'Sales Metrics'
            dataType: string
            lineageTag: "guid-1"
            summarizeBy: none
            sourceColumn: "Sales Metrics"
            sortByColumn: 'Sales Metrics Order'

        column 'Sales Metrics Fields'
            dataType: string
            isHidden
            lineageTag: "guid-2"
            summarizeBy: none
            sourceColumn: "Sales Metrics Fields"

            extendedProperty ParameterMetadata =
                {
                    "version": 3,
                    "kind": 2
                }

        column 'Sales Metrics Order'
            dataType: int64
            isHidden
            lineageTag: "guid-3"
            summarizeBy: none
            sourceColumn: "Sales Metrics Order"

        partition '.Parameter - Sales Metrics' = calculated
            mode: import
            source =
                {
                    ("Total Sales", NAMEOF('Sales'[Total Sales]), 0),
                    ("Total Cost", NAMEOF('Sales'[Total Cost]), 1),
                    ("Profit", NAMEOF('Sales'[Profit]), 2)
                }
```

### With Single-Level Category

Adds category columns to the structure:

```tmdl
        column 'Metric Type Sort'
            dataType: int64
            isHidden
            lineageTag: "guid-4"
            summarizeBy: none
            sourceColumn: "Metric Type Sort"

        column 'Metric Type'
            dataType: string
            lineageTag: "guid-5"
            summarizeBy: none
            sourceColumn: "Metric Type"
            sortByColumn: 'Metric Type Sort'
```

Tuple format expands to include category:

```tmdl
                {
                    ("Total Sales", NAMEOF('Sales'[Total Sales]), 0, 10, "Revenue"),
                    ("Total Cost", NAMEOF('Sales'[Total Cost]), 1, 20, "Cost"),
                    ("Profit", NAMEOF('Sales'[Profit]), 2, 10, "Revenue")
                }
```

### With Multi-Level Categories

For each additional level, adds:
- `[Level Name] Sort` column (hidden)
- `[Level Name]` column (visible)

Tuple format grows accordingly:

```tmdl
    (DisplayName, NAMEOF(...), ItemOrder,
     Cat1Sort, "Cat1Name",
     Cat2Sort, "Cat2Name",
     Cat3Sort, "Cat3Name")
```

### NAMEOF Syntax

The tool uses modern DAX NAMEOF syntax for field references:

```dax
NAMEOF('Table Name'[Field Name])
```

This ensures:
- Proper escaping of special characters
- Automatic update if field is renamed
- Clear reference to source table and column

---

## Category System

### Category Hierarchy Design

Categories create groupings within the parameter slicer. The system supports unlimited levels:

```
Level 1: Department
├── Finance
│   Level 2: Metric Type
│   ├── Revenue
│   └── Cost
├── Operations
│   Level 2: Metric Type
│   ├── Efficiency
│   └── Volume
```

### Category Assignment Flow

1. **Create Category Level** in Category Manager
2. **Set Column Name** for TMDL output
3. **Assign to Fields** via dropdown in builder
4. **Order Categories** by dragging in Category Manager

### Sort Order Handling

Each field has:
- `order_within_group`: Position within its category (0, 1, 2...)
- Category sort values: Position of category itself (10, 20, 30...)

Using gaps (10, 20, 30) allows inserting items later without renumbering.

### Category Storage in FieldItem

```python
# Example: Field in "Revenue" category at level 1, "Actual" at level 2
field_item.categories = [
    (10, "Revenue"),   # Level 1: sort=10, name="Revenue"
    (20, "Actual")     # Level 2: sort=20, name="Actual"
]
```

---

## Parser Logic

### TMDL Parsing Pipeline

```
Input TMDL Text
      │
      ▼
Extract table name ───────► "'.Parameter - Name'"
      │
      ▼
Find partition block ─────► "source = { ... }"
      │
      ▼
Parse each tuple ─────────► FieldItem objects
      │
      ▼
Extract column defs ──────► Lineage tags, metadata
      │
      ▼
Detect category levels ───► CategoryLevel objects
      │
      ▼
Assemble FieldParameter
```

### Tuple Parsing Regex

The parser uses regex to extract tuple components:

```python
# Basic tuple pattern
r'\("([^"]+)",\s*NAMEOF\(([^)]+)\),\s*(\d+)'

# Extended pattern with categories
r'\("([^"]+)",\s*NAMEOF\(([^)]+)\),\s*(\d+)(?:,\s*(\d+),\s*"([^"]+)")*'
```

### Handling Edge Cases

| Case | Handling |
|------|----------|
| Missing NAMEOF | Fall back to literal string parsing |
| Extra columns | Preserve in `extra_columns` list |
| Missing lineage tags | Generate new GUIDs if needed |
| Malformed tuples | Skip and log warning |

---

## UI Panels

### ModelConnectionPanel

**Purpose**: Connect to Power BI Desktop semantic models.

**Features**:
- Dropdown of available models
- Manual port entry option
- Connection status indicator
- Refresh button for model list

### ParameterConfigPanel

**Purpose**: Configure parameter creation/editing mode.

**Features**:
- New/Edit radio buttons
- Parameter name entry
- Existing parameter dropdown (edit mode)
- Lineage tag checkbox
- Create/Load button

### AvailableFieldsPanel

**Purpose**: Browse and add fields from connected model.

**Features**:
- Treeview of tables and fields
- Search/filter box
- Add button (or double-click)
- Field type indicators (measure/column)

### ParameterBuilderPanel

**Purpose**: Manage fields in the parameter.

**Features**:
- List of added fields with order numbers
- Drag-and-drop reordering
- Inline display name editing
- Category assignment dropdowns
- Remove field button
- Revert name buttons (individual/all)

### CategoryManagerPanel

**Purpose**: Create and manage category hierarchy.

**Features**:
- List of category levels
- Add/Edit/Remove buttons
- Drag-and-drop level reordering
- Column name configuration

### TmdlPreviewPanel

**Purpose**: Show generated TMDL code.

**Features**:
- Syntax-highlighted code view
- Real-time updates
- Copy to clipboard button
- Line numbers

---

## Usage Workflow

### Creating a New Parameter

1. **Connect to Model**
   - Launch from Power BI Desktop External Tools
   - Or manually select semantic model
   - Click "Connect"

2. **Configure Parameter**
   - Select "Create New" mode
   - Enter parameter name
   - Uncheck "Keep Lineage Tags"
   - Click "Create"

3. **Add Fields**
   - Browse available fields
   - Double-click or use Add button
   - Repeat for all desired fields

4. **Set Display Names** (optional)
   - Double-click names to edit
   - Use "Revert All" to reset

5. **Create Categories** (optional)
   - Click "Add Level"
   - Enter level name and column name
   - Repeat for additional levels

6. **Assign Categories**
   - Use dropdowns in builder
   - Drag to reorder

7. **Copy Code**
   - Review in preview panel
   - Click "Copy to Clipboard"
   - Paste into Tabular Editor

### Editing an Existing Parameter

1. **Connect to Model**
2. **Select "Edit Existing" mode**
3. **Choose parameter from dropdown**
4. **Check "Keep Lineage Tags"**
5. **Click "Load"**
6. **Make modifications**
7. **Copy updated code**

---

## Best Practices

### Parameter Naming

| Convention | Example |
|------------|---------|
| Table name prefix | `.Parameter - Sales Metrics` |
| Display column | `Sales Metrics` (no prefix) |
| Category columns | `Metric Type`, `Time Period` |

### Category Design

- Limit to 2-3 levels for usability
- Use logical hierarchies (Geography -> Country -> City)
- Assign consistent sort orders with gaps (10, 20, 30)
- Keep category names concise

### Field Organization

- Group related measures logically
- Use display names matching report terminology
- Order most common fields first within categories
- Avoid duplicate display names

### Version Control

| Scenario | Lineage Tags |
|----------|--------------|
| New parameter | Uncheck (let Power BI generate) |
| Editing existing | Check (preserve IDs) |
| Cloning parameter | Uncheck (new IDs needed) |

---

## Troubleshooting

### Model Won't Connect

**Symptoms**: No models in dropdown, connection fails.

**Solutions**:
- Ensure Power BI Desktop is running with a model open
- Check that model is in PBIP format
- Try manual port entry (find in PBI task manager)
- Verify external tool configuration

### Fields Not Appearing

**Symptoms**: Empty field list after connection.

**Solutions**:
- Confirm model connection is active (green status)
- Check that table contains measures or columns
- Refresh model list
- Verify table is not hidden

### Generated Code Errors

**Symptoms**: TMDL paste fails in Tabular Editor.

**Solutions**:
- Verify all field references exist in model
- Check for duplicate display names
- Ensure category assignments are complete
- Validate NAMEOF syntax in preview

### Lineage Tag Issues

**Symptoms**: Conflicts or errors after pasting.

**Solutions**:
- For existing parameters: Always keep lineage tags checked
- For new parameters: Always uncheck
- Don't mix modes within same parameter

---

## Limitations

### Current Version (1.0.0)

| Limitation | Description |
|------------|-------------|
| Model connection | Requires Power BI Desktop running |
| Field parameters only | Does not support other TMDL objects |
| Manual code transfer | Copy/paste required (no direct write) |
| Single model | Cannot reference fields across models |

### Power BI Limitations

| Limitation | Description |
|------------|-------------|
| Format requirement | Requires PBIP (TMDL) format |
| NAMEOF syntax | Requires Power BI 2020 or later |
| No calculated columns | Cannot use calculated columns as items |
| Performance | Practical limit ~100 fields per parameter |

---

## Dependencies

- Python 3.8+
- tkinter (standard library)
- Core PBI Connector (shared)
- Core UI Base classes

---

**Built by Reid Havens of Analytic Endeavors**
**Website**: https://www.analyticendeavors.com
**Email**: reid@analyticendeavors.com
