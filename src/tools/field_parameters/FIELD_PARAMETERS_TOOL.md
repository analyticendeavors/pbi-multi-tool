# Field Parameters Tool

**Version:** 1.0.0  
**Built by:** Reid Havens of Analytic Endeavors

## Overview

The Field Parameters Tool provides a comprehensive visual interface for creating and editing Power BI Field Parameters. Instead of manually writing DAX code, users can drag-and-drop measures and columns, assign them to multi-level categories, and automatically generate the proper TMDL/Tabular script code.

## Purpose

Field Parameters are a powerful Power BI feature that allows users to dynamically switch between different measures or columns in visuals. However, the native UI for creating them is limited and only accessible once. Editing existing parameters requires manual DAX coding, which is error-prone and tedious.

This tool solves these problems by:
- Providing a persistent, intuitive UI for parameter creation
- Supporting multi-level category hierarchies (not just single-level)
- Enabling drag-and-drop reordering for sort order control
- Generating clean, properly-formatted TMDL code
- Preserving extra calculated columns when editing existing parameters

## Features

### 1. Model Connection
- **External Tool Integration**: Automatically connects when launched from Power BI Desktop
- **Manual Selection**: Browse and select from available semantic models
- **Live Validation**: Only allows selection of fields from the connected model
- **Parameter Detection**: Auto-identifies existing field parameters via ParameterMetadata property

### 2. Parameter Configuration
- **Create New**: Start fresh with a blank parameter
- **Edit Existing**: Load and modify existing field parameters
- **Lineage Tags**: Optional preservation for version control
- **Custom Naming**: Flexible table and column naming

### 3. Field Management
- **Available Fields Panel**: Browse tables and fields from the model
- **Search & Filter**: Quickly find specific measures or columns
- **Drag & Drop**: Add fields by double-clicking or using Add button
- **Display Name Editing**: Customize friendly names (different from measure names)
- **Bulk Revert**: Reset all display names to original field names with one click
- **Individual Revert**: Reset single field names as needed

### 4. Parameter Builder
- **Visual List**: See all fields in the parameter with order numbers
- **Drag-to-Reorder**: Change sort order by dragging fields up/down
- **Inline Editing**: Double-click display names to edit
- **Category Assignment**: Assign each field to category levels
- **Remove Fields**: Easy deletion with confirmation

### 5. Category Manager
- **Multi-Level Support**: Create unlimited category hierarchies
- **Drag-to-Reorder**: Control category sort order visually
- **Custom Column Names**: Specify exact column names for each level
- **Add/Edit/Remove**: Full category lifecycle management

### 6. Code Generation
- **Real-Time Preview**: See generated TMDL code as you work
- **Proper Formatting**: Clean, readable code with section headers
- **Comment Blocks**: Auto-generated category section comments
- **Copy to Clipboard**: One-click copying for paste into Tabular Editor
- **Extra Column Preservation**: Maintains non-standard calculated columns

## Architecture

### File Structure
```
tools/field_parameters/
├── __init__.py                      # Package initialization
├── field_parameters_tool.py         # Tool registration & help content
├── field_parameters_core.py         # Parser & generator logic (FieldItem, CategoryLevel, etc.)
├── field_parameters_ui.py           # Main tab orchestrator
├── field_parameters_connector.py    # DEPRECATED - shim for backward compat (see core/pbi_connector.py)
├── field_parameters_builder.py      # Parameter builder panel with virtual scrolling
├── models.py                        # Data classes (ModelConnection, TableInfo, etc.)
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
└── FIELD_PARAMETERS_TOOL.md         # This file
```

### Component Hierarchy
```
FieldParametersTab (orchestrator)
├── ModelConnectionPanel (model selection & connection)
├── ParameterConfigPanel (new/edit mode, naming, options)
├── AvailableFieldsPanel (browse model fields)
├── ParameterBuilderPanel (drag-drop field list)
├── CategoryManagerPanel (category hierarchy management)
└── TmdlPreviewPanel (generated code display)
```

### Data Flow
1. **Model Connection**: User connects → loads tables/fields → enables UI
2. **Parameter Setup**: User creates new or loads existing → initializes FieldParameter object
3. **Field Addition**: User adds fields → creates FieldItem objects → updates builder
4. **Category Assignment**: User creates categories → updates CategoryLevel list → refreshes combos
5. **Code Generation**: Any change → regenerates TMDL → updates preview
6. **Output**: User copies TMDL → pastes into Tabular Editor or model file

## Usage Workflow

### Creating a New Parameter

1. **Connect to Model**
   - Launch tool from Power BI Desktop (External Tool)
   - Or manually select a semantic model from dropdown
   - Click "Connect"

2. **Configure Parameter**
   - Select "Create New" mode
   - Enter parameter name (e.g., "My Measures")
   - Uncheck "Keep Lineage Tags" (for new parameters)
   - Click "Create"

3. **Add Fields**
   - Browse available fields in left panel
   - Double-click or select and click "Add Selected Field"
   - Repeat for all desired fields

4. **Set Display Names** (Optional)
   - Double-click display names to edit
   - Or keep original field names
   - Use "Revert All" to reset if needed

5. **Create Categories** (Optional)
   - Click "Add Level" in Category Manager
   - Enter level name (e.g., "Geography")
   - Specify column name (e.g., "Region Category")
   - Repeat for additional levels

6. **Assign Categories**
   - In Parameter Builder, use category dropdowns for each field
   - Drag fields to reorder within categories
   - Drag categories in Category Manager to set group order

7. **Copy Code**
   - Review generated TMDL in preview panel
   - Click "Copy to Clipboard"
   - Paste into Tabular Editor or model definition file

### Editing an Existing Parameter

1. **Connect to Model** (same as above)

2. **Load Parameter**
   - Select "Edit Existing" mode
   - Choose parameter from dropdown
   - Ensure "Keep Lineage Tags" is checked
   - Click "Load"

3. **Modify Parameter**
   - Add/remove fields as needed
   - Edit display names
   - Adjust categories or reorder items

4. **Copy Updated Code**
   - Review changes in preview
   - Copy and paste back into model

## TMDL Structure

### Basic Parameter (No Categories)
```
createOrReplace
	table 'Parameter Name'
		column 'Display Column'
		column 'Display Column Fields' (hidden, has ParameterMetadata)
		column 'Display Column Order' (hidden, sort control)
		
		partition 'Parameter Name' = calculated
			source = {
				("Display", NAMEOF('Table'[Field]), 0),
				("Display 2", NAMEOF('Table'[Field2]), 1)
			}
```

### With Single Category
```
Adds two columns:
- 'Category Sort' (hidden, controls category order)
- 'Category Name' (display column for categories)

Tuple format:
("Display", NAMEOF('Table'[Field]), ItemOrder, CategorySort, "Category Name")
```

### With Multi-Level Categories
```
For each level, adds:
- '[Level] Sort' (hidden)
- '[Level Name]' (display)

Tuple format:
("Display", NAMEOF('Table'[Field]), ItemOrder, 
 Cat1Sort, "Cat1Name", 
 Cat2Sort, "Cat2Name", 
 ...)
```

## Technical Details

### FieldParameter Data Structure
```python
@dataclass
class FieldParameter:
    table_name: str                              # '.Parameter - Name'
    parameter_name: str                          # Display column name
    fields: List[FieldItem]                      # All field items
    category_levels: List[CategoryLevel]         # Category hierarchy
    keep_lineage_tags: bool                      # Preserve IDs?
    lineage_tags: Dict[str, str]                 # column → lineage tag
    pbi_id: Optional[str]                        # Table annotation
    extra_columns: List[Dict[str, Any]]          # Non-standard columns
```

### FieldItem Structure
```python
@dataclass
class FieldItem:
    display_name: str                            # User-facing name
    field_reference: str                         # NAMEOF(...) syntax
    table_name: str                              # Source table
    field_name: str                              # Source field
    order_within_group: int                      # Sort order (usually 1)
    categories: List[Tuple[int, str]]            # [(sort, name), ...]
```

### CategoryLevel Structure
```python
@dataclass
class CategoryLevel:
    name: str                                    # Level identifier
    sort_order: int                              # Position in hierarchy
    column_name: str                             # TMDL column name
```

## Integration Points

### Power BI Modeling MCP (Future)
- Connect to open semantic models
- Read table and measure definitions
- Identify existing field parameters
- Load parameter TMDL for editing

### PBIP Tools MCP (Future)
- Navigate PBIP file structure
- Read/write TMDL files directly
- Validate model structure

## Best Practices

### Parameter Naming
- Use `.Parameter - [Name]` format for calculated tables
- Keep display column names concise
- Use descriptive category names

### Category Design
- Limit to 2-3 levels for usability
- Use logical hierarchies (Geography → Country → City)
- Assign consistent sort orders (10, 20, 30 for flexibility)

### Field Organization
- Group related measures logically
- Use display names that match report terminology
- Order most common fields first within categories

### Version Control
- **New parameters**: Uncheck "Keep Lineage Tags"
- **Editing**: Check "Keep Lineage Tags" to preserve IDs
- Test changes in development environment first

## Troubleshooting

### Model Won't Connect
- Ensure Power BI Desktop is running
- Check that model is in PBIP format (not .pbix)
- Try manual model selection
- Verify external tool configuration

### Fields Not Appearing
- Confirm model connection is active (green status)
- Check that table contains measures or columns
- Refresh model list if recently modified

### Generated Code Errors
- Verify all field references exist in model
- Check for duplicate display names
- Ensure category assignments are complete
- Validate NAMEOF syntax in preview

### Lineage Tag Issues
- For existing parameters: Always keep lineage tags checked
- For new parameters: Always uncheck (let Power BI generate)
- Don't mix modes - use consistently

## Limitations

### Current Version
- Manual model connection (no auto-detection yet)
- Mock data for available fields (pending MCP integration)
- Simplified drag-drop (visual feedback pending)
- Single-level category UI (multi-level supported in code)

### Power BI Limitations
- Requires TMDL format (PBIP) for editing existing parameters
- NAMEOF syntax requires Power BI 2020 or later
- Field parameters don't support calculated columns as items
- Maximum practical limit ~100 fields per parameter

## Future Enhancements

### Planned Features
- [ ] Full MCP integration for model connection
- [ ] Real-time validation against model schema
- [ ] Advanced drag-drop with visual feedback
- [ ] Multi-level category UI (tree view)
- [ ] Template library for common patterns
- [ ] Bulk import from CSV/Excel
- [ ] Direct model application (without copy/paste)
- [ ] Parameter cloning and templating

### Nice-to-Have
- [ ] Preview of how parameter will look in slicer
- [ ] Usage analysis (which parameters used in visuals)
- [ ] Unused field detection
- [ ] Category optimization suggestions
- [ ] Integration with Git for version tracking

## Support & Feedback

**Author:** Reid Havens  
**Company:** Analytic Endeavors  
**Tool Version:** 1.0.0  
**Last Updated:** November 2025

For issues, questions, or feature requests, please contact Reid at Analytic Endeavors.

---

*This tool is not officially supported by Microsoft. Use at your own discretion and always backup your work before applying changes.*
