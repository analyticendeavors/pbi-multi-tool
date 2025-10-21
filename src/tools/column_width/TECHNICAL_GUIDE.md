# 📊 Table Column Widths Tool - Technical Architecture & How It Works

**Version**: v2.0.0  
**Built by**: Reid Havens of Analytic Endeavors  
**Last Updated**: October 21, 2025

---

## 📚 Table of Contents

1. [Overview](#overview)
2. [Architecture](#architecture)
3. [Core Components](#core-components)
4. [Width Calculation Algorithms](#width-calculation-algorithms)
5. [Content-Aware Intelligence](#content-aware-intelligence)
6. [Matrix-Specific Handling](#matrix-specific-handling)
7. [Data Flow](#data-flow)
8. [File Structure](#file-structure)
9. [Per-Visual Configuration](#per-visual-configuration)
10. [Key Operations](#key-operations)
11. [Font Intelligence](#font-intelligence)
12. [Future Enhancements](#future-enhancements)

---

## 🎯 Overview

The **Table Column Widths Tool** provides intelligent, content-aware column width standardization for Power BI Table and Matrix visuals. It analyzes field types, content patterns, and visual configurations to calculate optimal widths that prevent text wrapping while maintaining visual consistency.

### What Makes It Intelligent

1. **Content-Aware Analysis** - Detects dates, currencies, hierarchies, and data types
2. **Auto-Fit Algorithm** - Calculates optimal width based on actual content
3. **Matrix Intelligence** - Handles Compact, Outline, and Tabular layouts differently
4. **Hierarchy Detection** - Automatically adjusts for nested hierarchies
5. **Font-Aware Calculations** - Reads actual font settings from visuals
6. **Per-Visual Configuration** - Allows custom settings per visual or global defaults
7. **Fit to Totals Mode** - Optimizes measure columns for total/subtotal display

---

## 🏗️ Architecture

### Design Pattern: **Core Engine + UI Separation**

The tool uses a clean separation between business logic and user interface:

```
TableColumnWidthsTool (BaseTool)
    └── TableColumnWidthsTab (UI)
            ├── TableColumnWidthsEngine (Core Logic)
            │       ├── Visual Scanning
            │       ├── Field Analysis
            │       ├── Width Calculation
            │       └── File Updates
            └── UI Components
                    ├── File Selection
                    ├── Configuration Controls
                    ├── Visual Selection Tree
                    └── Per-Visual Dialogs
```

### Key Principles

- **Separation of Concerns**: Core logic has no UI dependencies
- **Composition**: UI composes the engine rather than inheriting
- **Stateless Engine**: Engine doesn't maintain operation-specific state
- **Configuration-Driven**: All settings passed via configuration objects

---

## 🧩 Core Components

### 1. **TableColumnWidthsEngine** (`column_width_core.py`)

**Role**: Core logic engine for all width calculations and operations

**Key Responsibilities**:
- Visual scanning and field extraction
- Content-aware width calculation
- Font information extraction
- Schema updates and validation
- Backup creation

**Core Data Structures**:

```python
@dataclass
class FieldInfo:
    """Information about a field"""
    name: str
    display_name: str
    field_type: FieldType  # CATEGORICAL or MEASURE
    metadata_key: str
    current_width: Optional[float]
    suggested_width: Optional[float]
    is_overridden: bool
    scale_config: Optional[ScaleConfiguration]

@dataclass
class VisualInfo:
    """Information about a visual"""
    visual_id: str
    visual_name: str
    visual_type: VisualType  # TABLE or MATRIX
    page_name: str
    page_id: str
    fields: List[FieldInfo]
    font_info: FontInfo
    layout_type: Optional[str]  # Compact, Outline, Tabular
    current_width: int
    current_height: int

@dataclass
class FontInfo:
    """Font configuration"""
    family: str = "Segoe UI"
    size: int = 11
    weight: str = "normal"
    
    def get_char_width(self) -> float:
        """Calculate average character width in pixels"""
```

**Critical Methods**:

```python
scan_visuals() -> List[VisualInfo]
    """Scan report for all Table and Matrix visuals"""

calculate_optimal_widths(visual_info, config)
    """Calculate optimal widths for all fields in a visual"""

apply_width_changes(visual_ids, configs) -> Dict[str, Any]
    """Apply width changes to specified visuals"""

create_backup() -> str
    """Create timestamped backup of PBIP file"""
```

---

### 2. **TableColumnWidthsTab** (`column_width_ui.py`)

**Role**: User interface for the tool

**Key Capabilities**:
- File selection and validation
- Visual scanning controls
- Global and per-visual configuration
- Visual selection tree with filtering
- Progress tracking and logging

**UI State Management**:

```python
# Global configuration (applies to all visuals by default)
self.global_categorical_preset_var = tk.StringVar(value=WidthPreset.AUTO_FIT.value)
self.global_measure_preset_var = tk.StringVar(value=WidthPreset.FIT_TO_TOTALS.value)

# Per-visual configurations (overrides global for specific visuals)
self.visual_config_vars: Dict[str, Dict[str, tk.Variable]] = {}

# Visual selection tracking
self.visual_selection_vars: Dict[str, tk.BooleanVar] = {}
```

**Critical UI Components**:

1. **File Input Section** - PBIP file selection with validation
2. **Scanning Section** - Visual scanning with summary display
3. **Configuration Section** - Global width settings (categorical + measures)
4. **Visual Selection Tree** - Multi-select tree with filtering
5. **Per-Visual Config Dialog** - Double-click opens custom config per visual
6. **Action Buttons** - Preview and Apply operations

---

### 3. **Width Presets** (`WidthPreset` enum)

**Available Presets**:

```python
class WidthPreset(Enum):
    NARROW = "narrow"           # Categorical: 60px, Measure: 70px
    MEDIUM = "medium"           # Categorical: 105px, Measure: 95px
    WIDE = "wide"               # Categorical: 165px, Measure: 145px
    AUTO_FIT = "auto_fit"       # Calculate optimal width (default)
    FIT_TO_TOTALS = "fit_to_totals"  # Optimize for totals/subtotals
    CUSTOM = "custom"           # User-specified pixel value
```

**Default Behavior**:
- **Categorical columns**: Auto-fit (intelligent header-based sizing)
- **Measure columns**: Fit to Totals (optimized for total/subtotal display)

---

## 📐 Width Calculation Algorithms

### 1. **Auto-Fit Algorithm** (Primary)

**Purpose**: Calculate optimal width to prevent text wrapping beyond 3 lines

**Process**:

```python
def _calculate_auto_fit_width(text, font_info, max_width, min_width, 
                             field_type, visual_type, scale_config, 
                             layout_type, hierarchy_levels) -> float:
    """
    Intelligent width calculation with content awareness
    
    Steps:
    1. Calculate character width from font settings
    2. Analyze content type (date, currency, hierarchy, etc.)
    3. Calculate base width for single line
    4. Apply content-specific adjustments
    5. Apply visual-type adjustments (matrix vs table)
    6. Add padding (≈20px)
    7. Calculate smart minimum based on content
    8. Cap at max width (3-line wrap allowed)
    """
```

**Key Features**:

- **Font-Aware**: Uses actual font size and weight from visual
- **Content-Aware**: Detects and adjusts for different data types
- **Matrix-Aware**: Applies layout-specific adjustments
- **Hierarchy-Aware**: Adds width per detected hierarchy level

**Width Calculation Formula**:

```
Base Width = Text Length × Character Width
Content Adjustment = Base × Content Multiplier
Matrix Adjustment = Content Adjusted × Matrix Multiplier
Final Width = Matrix Adjusted + Padding
Constrained Width = max(Smart Min, min(Final, Max Width))
```

---

### 2. **Fit to Totals Algorithm** (For Measures)

**Purpose**: Optimize measure columns to properly display totals and subtotals

**Why It's Needed**: Totals are often 2-4x larger than detail values due to aggregation

**Process**:

```python
def _calculate_fit_to_totals_width(field, visual_info, font_info, 
                                  max_width, min_width) -> float:
    """
    Calculate width optimized for total/subtotal display
    
    Steps:
    1. Detect hierarchy levels in visual
    2. Estimate total magnitude multiplier:
       - 4+ levels: 3.5x detail value
       - 3 levels: 2.8x detail value
       - 2 levels: 2.2x detail value
       - 1 level: 1.8x detail value
    3. Adjust for content type (currency, percentage, etc.)
    4. Calculate estimated total display length
    5. Add extra buffer for totals (15px)
    6. Apply matrix adjustments if applicable
    """
```

**Example Calculation**:

```
Detail Value: $861,060 (9 chars)
Hierarchy Levels: 4 (Year > Quarter > Month > Week)
Total Multiplier: 3.5x
Estimated Total: $19,932,899 (12 chars)
Character Width: 7px (11pt font)
Base Width: 12 × 7 = 84px
Content Padding: +25px (currency)
Total Buffer: +15px
Final Width: 124px
```

---

### 3. **Smart Minimum Algorithm**

**Purpose**: Calculate intelligent minimum width based on content and context

**Base Minimums by Category**:

**Categorical Fields**:
- Short categorical (≤4 chars): 50-75px
- Hierarchy parent: 55-90px  
- Hierarchy child: 45-80px
- Date fields: 80-100px
- Standard: 60-85px

**Measure Fields**:
- Currency: 75-100px
- Large numbers: 80-105px
- Standard: 65-90px

**Matrix Adjustments** (Compact layout with hierarchy levels):

```python
# Base minimum + (hierarchy_levels × per_level_bonus)
compact_categorical_min = 75 + (levels × 15)  # Max 150px
compact_measure_min = 90 + (levels × 4)       # Max 120px
```

---

## 🧠 Content-Aware Intelligence

### Content Type Detection

**Purpose**: Identify content patterns to apply specialized width logic

**Detected Types**:

```python
def _analyze_content_type(text: str) -> str:
    """Analyze field name to detect content type"""
    
    # Detection patterns:
    "complex_date"      # EOW:, BOW:, End of Week, etc.
    "date"              # Date, Time, YYYY-MM-DD patterns
    "currency"          # $, USD, Amount, Revenue, Cost
    "percentage"        # %, Percent, Rate, Ratio
    "large_number"      # Total, Sum, Count, Quantity
    "short_categorical" # ≤4 characters, no digits
    "hierarchy_parent"  # Year, Quarter, Month, Week, Day
    "hierarchy_child"   # Q1-Q4, Jan-Dec abbreviations
    "standard"          # Default catch-all
```

### Content-Specific Adjustments

**Applied Width Modifiers**:

| Content Type | Adjustment | Reason |
|-------------|-----------|--------|
| Complex Date | +35px | Extra space for "EOW: 01/07/24" format |
| Standard Date | +15px | Space for date separators |
| Currency | +15-20px | $ symbol + comma separators |
| Percentage | +10px | % symbol |
| Large Number | +15px | Comma separators |
| Short Categorical | min 45px | Ensure readability |
| Hierarchy Parent | +15px | Indentation space |
| Hierarchy Child | +8px | Moderate indentation |

---

## 🎯 Matrix-Specific Handling

### Matrix Layout Types

Power BI matrices have three layout types with different visual requirements:

1. **Compact** - Nested hierarchies, most space-efficient
2. **Tabular** - Flat structure, stepped indentation
3. **Outline** - Hierarchical with outline format

### Hierarchy Level Detection

**Purpose**: Automatically detect hierarchy depth for intelligent width calculation

**Detection Methods**:

```python
def _detect_hierarchy_levels(visual_info: VisualInfo) -> int:
    """
    Detect number of hierarchy levels in matrix
    
    Methods:
    1. Count unique hierarchy fields
    2. Detect level indicators (Year, Quarter, Month, etc.)
    3. Analyze field patterns for depth
    
    Returns: Estimated hierarchy depth (0-5+)
    """
```

**Level Indicators**:

```python
level_indicators = {
    "year": 1,      # Top level
    "quarter": 2,   # Second level
    "month": 3,     # Third level
    "week": 4,      # Fourth level
    "day": 5,       # Deepest level
    "eow": 4,       # End of Week
    "bow": 4        # Beginning of Week
}
```

### Compact Matrix Adjustments

**Categorical Columns** (Need extra width for indentation):

```python
base_adjustment = 1.4  # 40% wider baseline

# Add width per hierarchy level
level_bonus = min(hierarchy_levels × 0.3, 1.5)  # 30% per level, max 150%
base_adjustment += level_bonus

# Content-specific bonuses
if content_type == "hierarchy_parent":
    base_adjustment += 0.2  # Additional 20%
elif content_type == "hierarchy_child":
    base_adjustment += 0.1  # Additional 10%
elif content_type in ["date", "complex_date"]:
    base_adjustment += 0.3  # Additional 30%

# Cap maximum adjustment
base_adjustment = min(base_adjustment, 3.0)  # Max 300% of original

final_width = base_width × base_adjustment
```

**Example Compact Matrix Width**:

```
Base Width: 60px (from "Month" text)
Hierarchy Levels: 4 (Year > Quarter > Month > Week)
Base Adjustment: 1.4 (40% wider)
Level Bonus: 4 × 0.3 = 1.2 (120% bonus)
Content Type: hierarchy_child (+0.1)
Total Adjustment: 1.4 + 1.2 + 0.1 = 2.7 (capped at 3.0)
Final Width: 60 × 2.7 = 162px
```

**Measure Columns** (Need space for subtotals):

```python
base_adjustment = 1.15  # 15% wider for subtotals

# Bonus for complex hierarchies (more subtotal levels)
if hierarchy_levels > 3:
    base_adjustment += 0.05  # Additional 5%

final_width = base_width × base_adjustment
```

### Tabular Matrix Adjustments

**Less aggressive adjustments than Compact**:

```python
# Categorical
base_adjustment = 1.2  # 20% wider
level_bonus = min(hierarchy_levels × 0.15, 0.6)  # 15% per level, max 60%

# Measure
base_adjustment = 1.1  # 10% wider
if hierarchy_levels > 3:
    base_adjustment += 0.03  # Additional 3%
```

### Outline Matrix Adjustments

**Minimal adjustments** (original logic):

```python
# Categorical
if content_type == "hierarchy_parent":
    adjustment = 1.0  # No compression
else:
    adjustment = 0.95  # Slight compression

# Measure
adjustment = 1.05  # 5% wider
```

---

## 📊 Data Flow

### Complete Workflow

```
┌─────────────────────┐
│   User Selects      │
│   PBIP File         │
└──────────┬──────────┘
           │
           ▼
┌─────────────────────┐
│   Click "SCAN       │
│   VISUALS"          │
└──────────┬──────────┘
           │
           ▼
┌─────────────────────┐
│   Engine Scans      │
│   Report Structure  │ ──► Find pages directories
└──────────┬──────────┘     For each page:
           │                  - Read page.json
           ▼                  - Find visuals directory
┌─────────────────────┐       - Process each visual
│   Analyze Each      │
│   Visual            │
└──────────┬──────────┘
           │
           ▼
┌─────────────────────┐
│   Extract Visual    │
│   Information       │ ──► Type (Table/Matrix)
└──────────┬──────────┘     Layout (Compact/Outline/Tabular)
           │                Fields (Categorical/Measure)
           ▼                Font info (family, size, weight)
┌─────────────────────┐     Current widths
│   Extract Field     │
│   Information       │
└──────────┬──────────┘
           │
           ▼
┌─────────────────────┐
│   Display in        │
│   Visual Tree       │ ──► Checkboxes for selection
└──────────┬──────────┘     Double-click for per-visual config
           │
           ▼
┌─────────────────────┐
│   User Configures   │
│   Width Settings    │ ──► Global settings (default)
└──────────┬──────────┘     OR Per-visual settings
           │
           ▼
┌─────────────────────┐
│   User Selects      │
│   Visuals           │ ──► Check/uncheck visuals
└──────────┬──────────┘     Filter by type
           │
           ▼
┌─────────────────────┐
│   Click "APPLY      │
│   CHANGES"          │
└──────────┬──────────┘
           │
           ▼
┌─────────────────────┐
│   Create Backup     │ ──► Timestamped PBIP copy
└──────────┬──────────┘
           │
           ▼
     ┌─────┴─────┐
     │   Loop    │ ◄────────┐
     │  Visuals  │          │
     └─────┬─────┘          │
           │                │
           ▼                │
┌─────────────────────┐     │
│   Get Config for    │     │
│   Visual            │ ──► Per-visual OR Global
└──────────┬──────────┘     │
           │                │
           ▼                │
┌─────────────────────┐     │
│   Calculate Widths  │     │
│   for All Fields    │ ──► Auto-fit, Fit to Totals,
└──────────┬──────────┘     Preset, or Custom
           │                │
           ▼                │
┌─────────────────────┐     │
│   Update visual.json│     │
│   with New Widths   │ ──► Update columnWidth objects
└──────────┬──────────┘     Disable auto-size
           │                │
           └────────────────┘
           │
           ▼
┌─────────────────────┐
│   Validate Schema   │
└──────────┬──────────┘
           │
           ▼
┌─────────────────────┐
│   Report Success    │ ──► Show summary
└─────────────────────┘     Log results
```

---

## 📁 File Structure

### Tool Directory Layout

```
column_width/
├── column_width_tool.py         # BaseTool implementation
├── column_width_core.py         # Core engine (business logic)
├── column_width_ui.py           # Main UI tab
├── column_width_ui_extended.py  # Extended UI features
├── __init__.py
└── TECHNICAL_GUIDE.md           # This document
```

### Separation of Concerns

- **`column_width_tool.py`**: Tool registration and BaseTool interface
- **`column_width_core.py`**: Pure business logic, no UI imports
- **`column_width_ui.py`**: Main UI implementation, delegates logic to engine
- **`column_width_ui_extended.py`**: Additional UI features (per-visual config)

---

## ⚙️ Per-Visual Configuration

### Configuration Hierarchy

**Two-Tier System**:

1. **Global Configuration** - Default settings for all visuals
2. **Per-Visual Configuration** - Custom settings for specific visuals

**Priority**: Per-visual > Global

### Implementation

**Global Configuration Storage**:

```python
# Global settings (apply to all by default)
self.global_categorical_preset_var = tk.StringVar(value=WidthPreset.AUTO_FIT.value)
self.global_measure_preset_var = tk.StringVar(value=WidthPreset.FIT_TO_TOTALS.value)
self.categorical_custom_var = tk.StringVar(value="105")
self.measure_custom_var = tk.StringVar(value="95")
```

**Per-Visual Configuration Storage**:

```python
# Dictionary mapping visual_id to its custom configuration
self.visual_config_vars: Dict[str, Dict[str, tk.Variable]] = {
    'visual_id_123': {
        'categorical_preset': tk.StringVar(value='auto_fit'),
        'measure_preset': tk.StringVar(value='wide'),
        'categorical_custom': tk.StringVar(value="120"),
        'measure_custom': tk.StringVar(value="110")
    }
}
```

### User Workflow

1. **Scan Visuals** - All visuals use global settings by default
2. **Double-Click Visual** - Opens per-visual configuration dialog
3. **Configure Settings** - Set custom widths for that visual
4. **Apply Changes** - Visual uses its custom configuration
5. **Visual Tree Shows "Custom"** - Indicates per-visual config active

### Configuration Dialog

**Triggered by**: Double-clicking a visual in the tree

**Features**:
- Copy of global settings as starting point
- Same configuration options as global
- Three action buttons:
  - **Apply to This Visual**: Save custom config for this visual only
  - **Copy to Global**: Copy these settings to global defaults
  - **Reset to Global**: Remove custom config, use global
  - **Cancel**: Close without saving

### Application Logic

```python
def _get_selected_visual_configs() -> Dict[str, WidthConfiguration]:
    """Get configuration for each selected visual"""
    configs = {}
    
    for visual_id in selected_visual_ids:
        if visual_id in self.visual_config_vars:
            # Use per-visual configuration
            configs[visual_id] = create_config_from_vars(
                self.visual_config_vars[visual_id]
            )
        else:
            # Use global configuration
            configs[visual_id] = self._get_global_config()
    
    return configs
```

---

## 🔧 Key Operations

### Operation 1: Visual Scanning

**User Action**: Click "SCAN VISUALS"

**Backend Flow**:

1. **Validate PBIP File**:
   - Check file exists
   - Verify .Report directory exists
   - Validate structure

2. **Scan Pages**:
   ```python
   pages_dir = report_dir / "definition" / "pages"
   for page_dir in pages_dir.iterdir():
       read_page_info(page_dir)
       scan_page_visuals(page_dir)
   ```

3. **Process Each Visual**:
   - Read visual.json
   - Check visual type (tableEx or pivotTable)
   - Extract visual name
   - Extract font information
   - Extract layout type (for matrices)
   - Parse field information from query
   - Extract current column widths
   - Store in VisualInfo object

4. **Update UI**:
   - Populate visual selection tree
   - Show scan summary
   - Enable configuration controls

**Result**: List of VisualInfo objects ready for width calculation

---

### Operation 2: Width Calculation

**User Action**: Click "PREVIEW" or "APPLY CHANGES"

**Backend Flow**:

1. **For Each Selected Visual**:
   
   ```python
   # Get configuration (per-visual or global)
   config = get_config_for_visual(visual_id)
   
   # Detect hierarchy levels (for matrices)
   hierarchy_levels = detect_hierarchy_levels(visual_info)
   
   # For each field in visual:
   for field in visual_info.fields:
       # Determine preset (categorical vs measure)
       preset = get_preset_for_field(field, config)
       
       # Calculate based on preset
       if preset == AUTO_FIT:
           width = calculate_auto_fit_width(
               field.display_name,
               visual_info.font_info,
               config.max_width,
               config.min_width,
               field.field_type,
               visual_info.visual_type,
               field.scale_config,
               visual_info.layout_type,
               hierarchy_levels
           )
       elif preset == FIT_TO_TOTALS:
           width = calculate_fit_to_totals_width(
               field,
               visual_info,
               visual_info.font_info,
               config.max_width,
               config.min_width
           )
       # ... other presets
       
       field.suggested_width = width
   ```

2. **Preview Mode**:
   - Show calculated widths in dialog
   - Allow user to review before applying

3. **Apply Mode**:
   - Create backup of PBIP file
   - Update visual.json files
   - Validate schemas
   - Report results

---

### Operation 3: Schema Update

**Purpose**: Write calculated widths to visual.json files

**Process**:

```python
def _update_visual_file(visual_info: VisualInfo) -> int:
    """Update visual.json with new column widths"""
    
    # 1. Read current visual.json
    visual_file = find_visual_json(visual_info)
    visual_data = json.load(visual_file)
    
    # 2. Get/create columnWidth objects array
    objects = visual_data["visual"]["objects"]
    column_widths = objects.setdefault("columnWidth", [])
    
    # 3. Disable auto-size (critical for preserving widths)
    disable_auto_size_columns(objects)
    
    # 4. Update/create width entries
    for field in visual_info.fields:
        if field.suggested_width:
            width_obj = {
                "properties": {
                    "value": {
                        "expr": {
                            "Literal": {
                                "Value": f"{field.suggested_width}D"  # 'D' suffix
                            }
                        }
                    }
                },
                "selector": {
                    "metadata": field.metadata_key
                }
            }
            
            # Update or append
            if field.metadata_key in existing_widths:
                column_widths[index] = width_obj
            else:
                column_widths.append(width_obj)
    
    # 5. Write back to file
    json.dump(visual_data, visual_file, indent=2)
```

**Critical Detail**: Auto-size must be disabled, otherwise Power BI will override custom widths

```python
def _disable_auto_size_columns(objects):
    """Disable auto-size to preserve custom widths"""
    column_headers = objects.setdefault("columnHeaders", [])
    
    # Find or create auto-size setting
    auto_size_obj = find_or_create_auto_size_obj(column_headers)
    
    # Set to false
    auto_size_obj["properties"]["autoSizeColumnWidth"] = {
        "expr": {
            "Literal": {
                "Value": "false"
            }
        }
    }
```

---

## 🎨 Font Intelligence

### Font Information Extraction

**Purpose**: Use actual font settings from visuals for precise width calculation

**Extraction Locations**:

Power BI stores font settings in multiple places within visual.json:

```python
def _extract_font_info(visual_config) -> FontInfo:
    """Extract font from visual objects"""
    
    # Check these object types (in order of priority):
    font_objects = ["columnHeaders", "rowHeaders", "values", "general"]
    
    for obj_type in font_objects:
        if obj_type in objects:
            for obj in objects[obj_type]:
                properties = obj["properties"]
                
                # Extract font family
                if "fontFamily" in properties:
                    family = extract_literal_value(properties["fontFamily"])
                
                # Extract font size
                if "fontSize" in properties:
                    size = int(extract_literal_value(properties["fontSize"]))
                
                # Extract font weight
                if "fontWeight" in properties:
                    weight = extract_literal_value(properties["fontWeight"])
    
    return FontInfo(family, size, weight)
```

### Character Width Calculation

**Purpose**: Convert font size to average character width in pixels

**Algorithm**:

```python
def get_char_width(self) -> float:
    """Calculate average character width"""
    
    # Base character widths for Segoe UI at different sizes
    base_char_width = {
        8: 5.5, 9: 6.0, 10: 6.5, 11: 7.0, 12: 7.5,
        13: 8.0, 14: 8.5, 15: 9.0, 16: 9.5
    }
    
    char_width = base_char_width.get(self.size, 7.0)
    
    # Adjust for bold text (15% wider)
    if "bold" in self.weight.lower():
        char_width *= 1.15
    
    return char_width
```

**Example**:
```
Font: Segoe UI, 11pt, Normal
Character Width: 7.0px
Text: "Total Revenue"
Estimated Width: 13 chars × 7.0px = 91px (before adjustments)
```

---

## 🚀 Future Enhancements

### Potential Improvements

1. **Machine Learning Width Prediction**:
   - Train on actual Power BI reports
   - Learn optimal widths from user patterns
   - Predict ideal widths based on content distribution

2. **Data Preview Integration**:
   - Sample actual data from dataset
   - Calculate widths based on real values
   - Handle extreme outliers intelligently

3. **Dynamic Width Profiles**:
   - Save/load width configuration profiles
   - Share profiles across reports
   - Industry-specific presets (Finance, Healthcare, etc.)

4. **Visual-to-Visual Copy**:
   - Copy width settings from one visual to another
   - Bulk apply settings to similar visuals
   - Smart matching by field names

5. **Undo/Redo Support**:
   - Track width change history
   - Allow reverting to previous states
   - Compare before/after states

6. **Advanced Validation**:
   - Warn about potential wrapping issues
   - Suggest optimal max_width values
   - Detect and fix common configuration mistakes

7. **Batch Processing**:
   - Process multiple PBIP files at once
   - Apply consistent standards across reports
   - Generate width standardization reports

8. **Integration with Other Tools**:
   - Coordinate with Layout Optimizer
   - Share font information with other tools
   - Unified configuration management

---

## 📝 Code Quality Notes

### Strengths

- ✅ **Clean separation of UI and logic**
- ✅ **Comprehensive docstrings**
- ✅ **Type hints throughout**
- ✅ **Content-aware intelligence**
- ✅ **Per-visual configuration support**
- ✅ **Matrix-specific handling**
- ✅ **Hierarchy detection**
- ✅ **Font-aware calculations**
- ✅ **Proper error handling**

### Standards Followed

- **PEP 8**: Python style guide compliance
- **Type Hints**: Full coverage (Python 3.8+)
- **Docstrings**: NumPy/Google style
- **Dataclasses**: For structured data
- **Enums**: For type-safe constants
- **Composition**: Over inheritance

---

## 🎓 Learning Resources

### Understanding Power BI Visual Structure

- **visual.json**: Contains all visual configuration
- **objects**: Visual properties and formatting
- **columnWidth**: Array of width configurations
- **query**: Field definitions and data model
- **columnHeaders/rowHeaders**: Font and display settings

### Key Power BI Concepts

- **Table (tableEx)**: Flat table visual
- **Matrix (pivotTable)**: Hierarchical matrix visual
- **Layout Types**: Compact, Outline, Tabular
- **Field Types**: Categorical (dimensions) vs Measures (calculations)
- **Auto-Size**: Power BI feature that auto-adjusts widths

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
