# 🧹 Report Cleanup Tool - Technical Architecture & How It Works

**Version**: v1.0.0  
**Built by**: Reid Havens of Analytic Endeavors  
**Last Updated**: October 21, 2025

---

## 📚 Table of Contents

1. [Overview](#overview)
2. [Architecture](#architecture)
3. [Core Components](#core-components)
4. [Detection Algorithms](#detection-algorithms)
5. [Custom Visual Intelligence](#custom-visual-intelligence)
6. [Bookmark Analysis System](#bookmark-analysis-system)
7. [Visual Filter Analysis](#visual-filter-analysis)
8. [Data Flow](#data-flow)
9. [File Structure](#file-structure)
10. [Key Operations](#key-operations)
11. [Safety & Backup System](#safety--backup-system)
12. [Future Enhancements](#future-enhancements)

---

## 🎯 Overview

The **Report Cleanup Tool** provides comprehensive analysis and safe removal of unused resources from Power BI reports. It intelligently scans for themes, custom visuals, bookmarks, and visual filters that are no longer needed, helping optimize report file size and maintainability.

### What Makes It Intelligent

1. **Comprehensive Scanning** - Analyzes all pages and visuals to determine actual usage
2. **Multi-Type Detection** - Distinguishes between AppSource visuals, build pane visuals, and hidden visuals
3. **Smart Usage Detection** - Uses multiple detection methods to identify what's truly in use
4. **Bookmark Intelligence** - Validates page references and detects navigation patterns
5. **Safety-First Approach** - Creates automatic backups before any changes
6. **Hierarchical Awareness** - Handles bookmark parent/child groups intelligently
7. **Visual Filter Analysis** - Identifies and optionally hides all visual-level filters

---

## 🏗️ Architecture

### Design Pattern: **Analyzer + Engine + UI Separation**

The tool uses a clean three-tier architecture separating analysis, execution, and interface:

```
ReportCleanupTool (BaseTool)
    └── ReportCleanupTab (UI)
            ├── ReportAnalyzer (Analysis Logic)
            │       ├── Theme Detection
            │       ├── Custom Visual Scanning
            │       ├── Bookmark Validation
            │       └── Visual Filter Detection
            ├── ReportCleanupEngine (Execution Logic)
            │       ├── Theme Removal
            │       ├── Visual Removal
            │       ├── Bookmark Removal
            │       ├── Filter Hiding
            │       └── Reference Cleanup
            └── UI Components
                    ├── File Selection
                    ├── Analysis Controls
                    ├── Opportunity Display
                    └── Removal Actions
```

### Key Principles

- **Separation of Concerns**: Analysis, execution, and UI are independent
- **Read-Only Analysis**: Analyzer never modifies files, only Engine does
- **Safety by Default**: All operations create backups automatically
- **Comprehensive Detection**: Multiple methods to identify usage
- **Atomic Operations**: Each removal is independent and traceable

---

## 🧩 Core Components

### 1. **ReportAnalyzer** (`report_analyzer.py`)

**Role**: Comprehensive analysis engine for detecting cleanup opportunities

**Key Responsibilities**:
- Theme usage detection and comparison
- Custom visual type identification and usage scanning
- Bookmark validation and navigation pattern detection
- Visual filter analysis across all pages
- Hierarchical relationship tracking

**Core Data Structures**:

```python
@dataclass
class CleanupOpportunity:
    """Represents a cleanup opportunity found in the report"""
    item_type: str  # 'theme', 'custom_visual_build_pane', 
                    # 'custom_visual_hidden', 'bookmark_guaranteed_unused', 
                    # 'bookmark_likely_unused', 'bookmark_empty_group',
                    # 'visual_filter'
    item_name: str
    location: str   # File path or location description
    reason: str     # Why this can be removed
    safety_level: str = 'safe'  # 'safe', 'warning', 'risky'
    size_bytes: int = 0  # Estimated size savings
    visual_id: str = ""  # For custom visuals
    bookmark_id: str = ""  # For bookmarks
    filter_count: int = 0  # For visual filters
```

**Analysis Results Structure**:

```python
analysis_data = {
    'report_path': str,
    'used_visual_info': {
        'visual_types': Set[str],      # All visualType values found
        'custom_guids': Set[str],      # All customVisualGuid values
        'visual_names': Set[str],      # All visual identifiers
    },
    'themes': {
        'active_theme': (name, type),  # Single active theme
        'available_themes': {          # All theme files found
            'theme_name': {
                'type': 'BaseTheme' | 'CustomTheme',
                'location': 'SharedResources' | 'RegisteredResources',
                'path': Path,
                'size': int
            }
        },
        'theme_files': Dict[str, Path]
    },
    'custom_visuals': {
        'used_visuals': Set[str],      # All visual IDs actually used
        'build_pane_visuals': {        # Visuals in build pane
            'visual_id': {
                'display_name': str,
                'location': str,
                'used': bool,
                'size': int
            }
        },
        'hidden_visuals': {            # Hidden in CustomVisuals folder
            'visual_id': {
                'display_name': str,
                'path': Path,
                'used': False,  # Always false
                'size': int
            }
        },
        'appsource_visuals': {         # AppSource/publicCustomVisuals
            'visual_id': {
                'display_name': str,
                'used': bool,
                'size': 0  # No local space
            }
        }
    },
    'bookmarks': {
        'bookmarks': {
            'bookmark_id': {
                'name': str,
                'display_name': str,
                'page_id': str,
                'page_exists': bool,
                'is_parent': bool,      # Parent group
                'parent_id': str,       # For children
                'bookmark_data': Dict
            }
        },
        'existing_pages': Set[str],
        'bookmark_usage': {
            'used_bookmarks': Set[str],
            'pages_with_navigation': Set[str],
            'bookmark_navigators': List[Tuple],
            'bookmark_buttons': List[Tuple]
        }
    },
    'visual_filters': {
        'pages_with_visual_filters': {
            'page_name': {
                'visuals_with_filters': List[Dict],
                'total_filters': int,
                'visible_filters': int
            }
        },
        'total_visual_filters': int,
        'total_visuals_with_filters': int
    }
}
```

**Critical Methods**:

```python
analyze_pbip_report(pbip_path) -> (Dict[str, Any], List[CleanupOpportunity])
    """Complete analysis returning data and opportunities"""

_scan_pages_for_all_visuals(report_dir) -> Dict[str, Set[str]]
    """Comprehensive scan finding ALL visual usage"""

_analyze_themes(report_data, report_dir) -> Dict[str, Any]
    """Analyze active theme vs available themes"""

_analyze_custom_visuals(report_data, report_dir, used_info) -> Dict[str, Any]
    """Analyze three types of custom visuals"""

_analyze_bookmarks(report_data, report_dir) -> Dict[str, Any]
    """Analyze bookmarks with parent/child relationships"""

_analyze_visual_filters(report_dir) -> Dict[str, Any]
    """Analyze visual-level filters across all pages"""
```

---

### 2. **ReportCleanupEngine** (`cleanup_engine.py`)

**Role**: Execution engine for safe removal operations

**Key Responsibilities**:
- Creating timestamped backups
- Removing unused themes from file system
- Removing custom visuals (all three types)
- Removing bookmarks and updating parent groups
- Hiding visual-level filters
- Cleaning up report.json references

**Core Data Structures**:

```python
@dataclass
class RemovalResult:
    """Result of a removal operation"""
    item_name: str
    item_type: str  # 'theme', 'custom_visual', 'bookmark', 'visual_filter'
    success: bool
    error_message: str = ""
    bytes_freed: int = 0
    filters_hidden: int = 0  # For visual filters
```

**Critical Methods**:

```python
remove_unused_items(pbip_path, themes, visuals, bookmarks, 
                   hide_filters, create_backup) -> List[RemovalResult]
    """Master removal method handling all item types"""

_create_backup(pbip_path) -> None
    """Create timestamped backup of PBIP and .Report directory"""

_remove_themes(report_dir, theme_names) -> List[RemovalResult]
    """Remove theme files from SharedResources/RegisteredResources"""

_remove_custom_visuals(report_dir, visual_opportunities) -> List[RemovalResult]
    """Remove custom visuals handling three different types"""

_remove_bookmarks(report_dir, bookmark_opportunities) -> List[RemovalResult]
    """Remove bookmarks and update parent group structures"""

_hide_visual_filters(report_dir) -> List[RemovalResult]
    """Hide all visual-level filters by setting isHiddenInViewMode"""

_cleanup_report_references(report_dir, removed_themes, 
                          removed_visuals, removed_bookmarks) -> None
    """Update report.json to remove references to deleted items"""
```

---

### 3. **ReportCleanupTab** (`cleanup_ui.py`)

**Role**: User interface for the cleanup tool

**Key Capabilities**:
- File selection and validation
- Analysis triggering and progress display
- Opportunity categorization and display
- Selective removal with checkboxes
- Safety confirmations
- Result logging and summary

**UI State Management**:

```python
# Analysis state
self.analysis_data: Optional[Dict[str, Any]] = None
self.cleanup_opportunities: List[CleanupOpportunity] = []

# Selection state
self.opportunity_vars: Dict[int, tk.BooleanVar] = {}  # Checkbox states
self.hide_visual_filters_var = tk.BooleanVar(value=False)

# UI components
self.opportunities_frame: Optional[tk.Frame] = None
self.summary_label: Optional[tk.Label] = None
```

**Critical UI Methods**:

```python
_analyze_report() -> None
    """Trigger analysis and display opportunities"""

_display_opportunities(opportunities) -> None
    """Build UI with categorized opportunities"""

_remove_selected_items() -> None
    """Execute removal for selected items with confirmation"""

_create_opportunity_section(parent, category, items) -> None
    """Create collapsible section for each opportunity type"""
```

---

## 🔍 Detection Algorithms

### Theme Detection Algorithm

**Purpose**: Identify the single active theme and find all unused theme files

**Key Concept**: Only ONE theme can be active at a time in Power BI

**Detection Process**:

```python
def _analyze_themes(report_data, report_dir):
    """
    Theme detection algorithm
    
    Steps:
    1. Find THE active theme from themeCollection:
       - Check customTheme first (takes precedence)
       - Fall back to baseTheme if no custom theme
       - Default to "Default" if neither specified
    
    2. Scan for ALL available theme files:
       - BaseThemes in SharedResources/BaseThemes/*.json
       - CustomThemes in RegisteredResources/*.json
    
    3. Compare active vs available:
       - Any theme file NOT matching active theme is unused
    
    Returns: Theme analysis with active_theme and available_themes
    """
    
    # Step 1: Find active theme
    theme_collection = report_data.get('themeCollection', {})
    
    # Custom theme takes precedence
    custom_theme = theme_collection.get('customTheme', {})
    if custom_theme and custom_theme.get('name'):
        active_theme = (custom_theme['name'], custom_theme.get('type'))
    else:
        # Fall back to base theme
        base_theme = theme_collection.get('baseTheme', {})
        if base_theme and base_theme.get('name'):
            active_theme = (base_theme['name'], base_theme.get('type'))
        else:
            active_theme = ('Default', 'Built-in')
    
    # Step 2: Scan file system for theme files
    available_themes = {}
    
    # Scan SharedResources/BaseThemes
    base_themes_dir = report_dir / "StaticResources" / "SharedResources" / "BaseThemes"
    for theme_file in base_themes_dir.glob("*.json"):
        available_themes[theme_file.stem] = {
            'type': 'BaseTheme',
            'location': 'SharedResources',
            'path': theme_file,
            'size': theme_file.stat().st_size
        }
    
    # Scan RegisteredResources for custom themes
    reg_resources_dir = report_dir / "StaticResources" / "RegisteredResources"
    for theme_file in reg_resources_dir.glob("*.json"):
        # Validate it's actually a theme file
        if is_theme_file(theme_file):
            available_themes[theme_file.name] = {
                'type': 'CustomTheme',
                'location': 'RegisteredResources',
                'path': theme_file,
                'size': theme_file.stat().st_size
            }
    
    # Step 3: Find unused themes
    active_name, active_type = active_theme
    unused_themes = []
    
    for theme_name, theme_info in available_themes.items():
        # Check if this theme is active
        is_active = (theme_name == active_name or 
                    theme_name.replace('.json', '') == active_name or
                    active_name in theme_name)
        
        if not is_active:
            unused_themes.append(theme_name)
    
    return {
        'active_theme': active_theme,
        'available_themes': available_themes,
        'unused_themes': unused_themes
    }
```

**Cleanup Opportunity Creation**:

```python
for theme_name in unused_themes:
    opportunity = CleanupOpportunity(
        item_type='theme',
        item_name=theme_name,
        location=f"{theme_info['location']}/{theme_info['type']}",
        reason="Theme is not currently active",
        safety_level='safe',
        size_bytes=theme_info['size']
    )
```

---

### Visual Scanning Algorithm

**Purpose**: Comprehensive scan to find ALL visual types used in pages

**Why It's Critical**: Must catch every possible way a visual can be referenced

**Detection Process**:

```python
def _scan_pages_for_all_visuals(report_dir):
    """
    Comprehensive visual usage detection
    
    Detection Methods:
    1. visualType field - Standard visual type identifier
    2. customVisualGuid - GUID for custom visuals
    3. visual names/IDs - Any identifier in visual data
    
    Scans:
    - All pages in definition/pages/
    - All visuals in each page's visuals/ directory
    - All visual.json files for configuration
    
    Returns: Dictionary with three sets of used visual identifiers
    """
    
    visual_usage = {
        'visual_types': set(),      # All visualType values
        'custom_guids': set(),      # All customVisualGuid values
        'visual_names': set(),      # All identifiers
    }
    
    pages_dir = report_dir / "definition" / "pages"
    
    for page_dir in pages_dir.iterdir():
        if not page_dir.is_dir():
            continue
        
        visuals_dir = page_dir / "visuals"
        if not visuals_dir.exists():
            continue
        
        for visual_dir in visuals_dir.iterdir():
            visual_json = visual_dir / "visual.json"
            if not visual_json.exists():
                continue
            
            visual_data = read_json(visual_json)
            visual_config = visual_data.get('visual', {})
            
            # Method 1: Get visualType
            visual_type = visual_config.get('visualType', '')
            if visual_type:
                visual_usage['visual_types'].add(visual_type)
            
            # Method 2: Get customVisualGuid
            custom_guid = visual_config.get('customVisualGuid', '')
            if custom_guid:
                visual_usage['custom_guids'].add(custom_guid)
            
            # Method 3: Get all possible identifiers
            for field in ['name', 'id', 'visualId', 'guid']:
                value = visual_config.get(field, '') or visual_data.get(field, '')
                if value:
                    visual_usage['visual_names'].add(value)
    
    return visual_usage
```

**Usage Matching Algorithm**:

```python
def _is_visual_used(visual_id, all_used_visuals):
    """
    Multi-method visual usage detection
    
    Matching strategies:
    1. Direct match - visual_id in used_visuals
    2. Substring match - Either string contains the other
    3. GUID pattern match - Compare normalized GUIDs
    """
    
    # Direct match
    if visual_id in all_used_visuals:
        return True
    
    # Fuzzy matching for different naming patterns
    for used_visual in all_used_visuals:
        # Substring matching
        if (visual_id.lower() in used_visual.lower() or 
            used_visual.lower() in visual_id.lower()):
            return True
        
        # GUID pattern matching (if both look like GUIDs)
        if len(visual_id) > 10 and len(used_visual) > 10:
            visual_parts = normalize_guid(visual_id)
            used_parts = normalize_guid(used_visual)
            if visual_parts in used_parts or used_parts in visual_parts:
                return True
    
    return False
```

---

### Custom Visual Type Detection

**Purpose**: Distinguish between built-in visuals and custom visuals

**Built-in Visual Types** (Not custom visuals):

```python
BUILTIN_VISUAL_TYPES = {
    # Core built-in visuals
    'clusteredColumnChart', 'table', 'slicer', 'card', 'lineChart', 
    'pieChart', 'map', 'clusteredBarChart', 'scatterChart', 'gauge', 
    'multiRowCard', 'kpi', 'donutChart', 'matrix', 'waterfallChart', 
    'funnelChart', 'treemap', 'ribbonChart', 'histogram', 'filledMap',
    
    # Stacked variants
    'stackedColumnChart', 'stackedBarChart', 'lineStackedColumnChart',
    'lineClusteredColumnChart', 'hundredPercentStackedBarChart',
    'hundredPercentStackedColumnChart',
    
    # Other built-ins
    'shape', 'textbox', 'image', 'actionButton', 'decompositionTreeVisual',
    'smartNarrativeVisual', 'keyInfluencersVisual', 'qnaVisual', 'paginator',
    'advancedSlicerVisual', 'barChart', 'pivotTable', 'basicShape',
    'areaChart', 'columnChart', 'pieDonutChart', 'lineAreaChart',
    'comboChart', 'tableEx', 'matrixEx'
}
```

**Detection Algorithm**:

```python
def _is_custom_visual_type(visual_type):
    """
    Determine if a visual type is custom (not built-in)
    
    Custom visual patterns:
    - Contains '_CV_' or 'PBI_CV_' prefix
    - Known custom visual names (ChordChart, wordCloud, etc.)
    - Not in BUILTIN_VISUAL_TYPES set
    """
    
    if not visual_type:
        return False
    
    # Check for custom visual patterns
    custom_patterns = [
        '_CV_', 'PBI_CV_', 'ChordChart', 'sparklineChart',
        'waterCup', 'searchVisual', 'calendarVisual',
        'hierarchySlicer', 'timelineStoryteller', 'wordCloud'
    ]
    
    for pattern in custom_patterns:
        if pattern.lower() in visual_type.lower():
            return True
    
    # If not in built-in types, it's likely custom
    return visual_type not in BUILTIN_VISUAL_TYPES
```

---

## 🔮 Custom Visual Intelligence

### Three Types of Custom Visuals

Power BI custom visuals can exist in three different states:

#### **Type 1: AppSource Visuals** (`custom_visual_build_pane` from publicCustomVisuals)

**Location**: `report.json` → `publicCustomVisuals` array

**Characteristics**:
- Downloaded from AppSource
- Listed in publicCustomVisuals
- No local files (streamed from cloud)
- Appear in build pane
- Take 0 bytes locally

**Detection**:

```python
# From report.json
public_visuals = report_data.get('publicCustomVisuals', [])

for visual_id in public_visuals:
    is_used = _is_visual_used(visual_id, all_used_visuals)
    
    appsource_visuals[visual_id] = {
        'visual_id': visual_id,
        'display_name': get_display_name(visual_id),
        'location': 'AppSource/publicCustomVisuals',
        'used': is_used,
        'size': 0  # No local space
    }
```

**Removal Process**:
1. Remove from `publicCustomVisuals` array in report.json
2. Clean up references in `resourcePackages`
3. No file system changes (no local files)

---

#### **Type 2: Build Pane Visuals** (`custom_visual_build_pane` from resourcePackages)

**Location**: `report.json` → `resourcePackages` → CustomVisual packages + `CustomVisuals/` directory

**Characteristics**:
- Registered in resourcePackages
- Visible in build pane
- Have local files in CustomVisuals folder
- May be custom imports or older AppSource visuals

**Detection**:

```python
# From report.json resourcePackages
resource_packages = report_data.get('resourcePackages', [])

for package in resource_packages:
    if package.get('type') == 'CustomVisual':
        visual_id = package.get('name', '')
        
        is_used = _is_visual_used(visual_id, all_used_visuals)
        size = estimate_visual_size(report_dir, visual_id)
        
        build_pane_visuals[visual_id] = {
            'visual_id': visual_id,
            'display_name': get_display_name(visual_id),
            'location': 'Build Pane (resourcePackages)',
            'used': is_used,
            'size': size
        }
```

**Removal Process**:
1. Remove from `resourcePackages` in report.json
2. Delete CustomVisuals/{visual_id} directory
3. Clean up references

---

#### **Type 3: Hidden Visuals** (`custom_visual_hidden`)

**Location**: `CustomVisuals/` directory only (not in report.json)

**Characteristics**:
- Files exist in CustomVisuals folder
- NOT registered in resourcePackages
- NOT visible in build pane
- Cannot be used (hidden/orphaned)
- Take up disk space unnecessarily

**Detection**:

```python
# Scan CustomVisuals directory
custom_visuals_dir = report_dir / "CustomVisuals"
registered_visuals = set(build_pane_visuals.keys())

for visual_dir in custom_visuals_dir.iterdir():
    visual_id = visual_dir.name
    
    # Check if already registered in build pane
    if visual_id not in registered_visuals:
        # This is a hidden visual!
        display_name = get_display_name_from_folder(visual_dir)
        size = calculate_directory_size(visual_dir)
        
        hidden_visuals[visual_id] = {
            'visual_id': visual_id,
            'display_name': display_name,
            'location': 'CustomVisuals (hidden)',
            'used': False,  # Hidden visuals can't be used
            'size': size
        }
```

**Removal Process**:
1. Delete CustomVisuals/{visual_id} directory only
2. No report.json changes (not registered)

---

### Custom Visual Analysis Flow

```
┌─────────────────────────────────────────────┐
│   Scan All Pages for Visual Usage          │
│   (Collect visual_types, GUIDs, names)     │
└────────────────┬────────────────────────────┘
                 │
                 ▼
┌─────────────────────────────────────────────┐
│   Analyze publicCustomVisuals Array         │
│   - Extract visual IDs                      │
│   - Check each against used_visuals         │
│   - Mark as used/unused                     │
└────────────────┬────────────────────────────┘
                 │
                 ▼
┌─────────────────────────────────────────────┐
│   Analyze resourcePackages                  │
│   - Find CustomVisual packages              │
│   - Check each against used_visuals         │
│   - Calculate local file sizes              │
│   - Mark as used/unused                     │
└────────────────┬────────────────────────────┘
                 │
                 ▼
┌─────────────────────────────────────────────┐
│   Scan CustomVisuals Directory              │
│   - List all subdirectories                 │
│   - Exclude already-registered visuals      │
│   - Remaining = Hidden visuals              │
│   - Calculate sizes                         │
└────────────────┬────────────────────────────┘
                 │
                 ▼
┌─────────────────────────────────────────────┐
│   Create Cleanup Opportunities              │
│   - Unused AppSource visuals                │
│   - Unused build pane visuals               │
│   - All hidden visuals (always unused)      │
└─────────────────────────────────────────────┘
```

---

## 📖 Bookmark Analysis System

### Bookmark Hierarchical Structure

Power BI bookmarks can be organized in parent/child groups:

```
Bookmark Structure:
├── Standalone Bookmark A
├── Parent Group 1
│   ├── Child Bookmark B
│   ├── Child Bookmark C
│   └── Child Bookmark D
├── Standalone Bookmark E
└── Parent Group 2
    ├── Child Bookmark F
    └── Child Bookmark G
```

### Bookmark Data Sources

**Source 1: bookmarks.json** (Metadata)
```json
{
  "items": [
    {
      "name": "bookmark_id_123",
      "displayName": "Parent Group 1",
      "children": ["child_id_456", "child_id_789"]
    },
    {
      "name": "child_id_456",
      "displayName": "Child Bookmark B"
    }
  ]
}
```

**Source 2: Individual .bookmark.json Files** (Configuration)
```json
{
  "displayName": "Sales Overview",
  "explorationState": {
    "activeSection": "page_id_abc123"  // Page reference
  }
}
```

### Bookmark Analysis Algorithm

```python
def _analyze_bookmarks(report_data, report_dir):
    """
    Comprehensive bookmark analysis
    
    Steps:
    1. Get existing pages (for validation)
    2. Read bookmarks.json for metadata
    3. Process parent groups first
    4. Process child bookmarks
    5. Scan for bookmark usage in visuals
    6. Categorize bookmarks by status
    """
    
    # Step 1: Get existing pages
    existing_pages = set()
    pages_dir = report_dir / "definition" / "pages"
    for page_dir in pages_dir.iterdir():
        # Get both internal ID and display name
        page_json = page_dir / "page.json"
        page_data = read_json(page_json)
        existing_pages.add(page_dir.name)  # Internal ID
        existing_pages.add(page_data.get('displayName'))  # Display name
    
    # Step 2: Read bookmarks metadata
    bookmarks_json = bookmarks_dir / "bookmarks.json"
    bookmarks_meta = read_json(bookmarks_json)
    
    bookmarks = {}
    
    # Step 3: Process each bookmark item
    for bookmark_meta in bookmarks_meta.get('items', []):
        bookmark_id = bookmark_meta.get('name')
        children = bookmark_meta.get('children', [])
        
        if children:
            # This is a parent group
            bookmarks[bookmark_id] = {
                'name': bookmark_id,
                'display_name': bookmark_meta.get('displayName'),
                'page_id': '',  # Parent groups don't have pages
                'page_exists': True,  # Always valid
                'is_parent': True,
                'bookmark_data': bookmark_meta
            }
            
            # Process children
            for child_id in children:
                child_file = bookmarks_dir / f"{child_id}.bookmark.json"
                child_data = read_json(child_file)
                
                # Get page reference from child
                exploration_state = child_data.get('explorationState', {})
                active_section = exploration_state.get('activeSection', '')
                page_exists = active_section in existing_pages
                
                bookmarks[child_id] = {
                    'name': child_id,
                    'display_name': child_data.get('displayName'),
                    'page_id': active_section,
                    'page_exists': page_exists,
                    'parent_id': bookmark_id  # Link to parent
                }
        else:
            # Standalone bookmark
            bookmark_file = bookmarks_dir / f"{bookmark_id}.bookmark.json"
            bookmark_data = read_json(bookmark_file)
            
            exploration_state = bookmark_data.get('explorationState', {})
            active_section = exploration_state.get('activeSection', '')
            page_exists = active_section in existing_pages
            
            bookmarks[bookmark_id] = {
                'name': bookmark_id,
                'display_name': bookmark_data.get('displayName'),
                'page_id': active_section,
                'page_exists': page_exists,
                'bookmark_data': bookmark_data
            }
    
    # Step 4: Scan for bookmark usage
    bookmark_usage = _scan_pages_for_bookmark_usage(report_dir)
    
    return {
        'bookmarks': bookmarks,
        'existing_pages': existing_pages,
        'bookmark_usage': bookmark_usage
    }
```

### Bookmark Usage Detection

**Purpose**: Find bookmarks referenced by navigation buttons or bookmark navigators

**Detection Locations**:

```python
def _scan_pages_for_bookmark_usage(report_dir):
    """
    Scan all visuals for bookmark references
    
    Looks for:
    1. Bookmark navigators (list all bookmarks)
    2. Action buttons (navigate to specific bookmark)
    3. Other visuals with bookmark actions
    """
    
    bookmark_usage = {
        'used_bookmarks': set(),
        'bookmark_navigators': [],
        'bookmark_buttons': []
    }
    
    for page in all_pages:
        for visual in page_visuals:
            visual_data = read_visual_json(visual)
            
            # Search recursively through visual data
            bookmark_refs = _extract_bookmark_references(visual_data)
            
            if bookmark_refs:
                bookmark_usage['used_bookmarks'].update(bookmark_refs)
                
                visual_type = get_visual_type(visual_data)
                if 'bookmarkNavigator' in visual_type:
                    bookmark_usage['bookmark_navigators'].append(
                        (page_name, visual_type, bookmark_refs)
                    )
                elif 'button' in visual_type.lower():
                    bookmark_usage['bookmark_buttons'].append(
                        (page_name, visual_type, bookmark_refs)
                    )
    
    return bookmark_usage
```

**Recursive Bookmark Reference Extraction**:

```python
def _extract_bookmark_references(visual_data):
    """
    Recursively search for bookmark references
    
    Looks for:
    - Any key containing 'bookmark'
    - Navigation actions
    - Bookmark state configurations
    - Power BI's expr.Literal.Value structures
    """
    
    bookmark_refs = set()
    
    def search_dict(obj, path=""):
        if isinstance(obj, dict):
            for key, value in obj.items():
                # Check for bookmark-related keys
                if 'bookmark' in key.lower():
                    if isinstance(value, str) and value:
                        bookmark_refs.add(value)
                    elif isinstance(value, dict):
                        # Handle expr.Literal.Value structure
                        literal_value = extract_literal_value(value)
                        if literal_value:
                            bookmark_refs.add(literal_value)
                
                # Recursively search nested objects
                if isinstance(value, (dict, list)):
                    search_dict(value, f"{path}.{key}")
        
        elif isinstance(obj, list):
            for item in obj:
                search_dict(item, path)
    
    search_dict(visual_data)
    return bookmark_refs
```

### Bookmark Usage Checking

**Includes Parent Group Intelligence**:

```python
def _is_bookmark_used_including_children(bookmark_id, bookmark_name, 
                                        bookmark_info, bookmark_analysis):
    """
    Check if bookmark is used (including parent group logic)
    
    Usage scenarios:
    1. Direct usage - bookmark ID/name in used_bookmarks set
    2. Parent group - bookmark navigator references the parent
    3. Child bookmark - parent group is used by navigator
    """
    
    bookmark_usage = bookmark_analysis['bookmark_usage']
    used_bookmarks = bookmark_usage['used_bookmarks']
    
    # Check direct usage (ID or name)
    if bookmark_id in used_bookmarks or bookmark_name in used_bookmarks:
        return True
    
    # For parent groups
    if bookmark_info.get('is_parent'):
        # Check if any navigator uses this parent group
        for page, visual_type, bookmark_refs in bookmark_usage['bookmark_navigators']:
            if bookmark_id in bookmark_refs:
                return True
        
        # Check if any children are used
        all_bookmarks = bookmark_analysis['bookmarks']
        for child_id, child_info in all_bookmarks.items():
            if child_info.get('parent_id') == bookmark_id:
                child_name = child_info['display_name']
                if child_id in used_bookmarks or child_name in used_bookmarks:
                    return True
    
    # For child bookmarks
    parent_id = bookmark_info.get('parent_id')
    if parent_id:
        # Check if parent group is used
        for page, visual_type, bookmark_refs in bookmark_usage['bookmark_navigators']:
            if parent_id in bookmark_refs:
                return True
    
    return False
```

### Bookmark Cleanup Opportunities

**Four Categories**:

#### **Category 1: Guaranteed Unused** (Page doesn't exist)

```python
if page_id and not page_exists:
    opportunity = CleanupOpportunity(
        item_type='bookmark_guaranteed_unused',
        item_name=bookmark_name,
        location='Bookmarks',
        reason=f"Target page '{page_id[:12]}...' no longer exists",
        safety_level='safe',
        bookmark_id=bookmark_id
    )
```

#### **Category 2: Likely Unused** (No navigation found)

```python
if not is_parent and not is_bookmark_used(bookmark_id):
    opportunity = CleanupOpportunity(
        item_type='bookmark_likely_unused',
        item_name=bookmark_name,
        location='Bookmarks',
        reason="No navigation buttons found, but could be used via bookmark pane",
        safety_level='warning',  # Warning level
        bookmark_id=bookmark_id
    )
```

#### **Category 3: Empty Parent Groups**

```python
# After marking children for removal
if all_children_being_removed:
    opportunity = CleanupOpportunity(
        item_type='bookmark_empty_group',
        item_name=parent_name,
        location='Bookmarks',
        reason=f"Parent group empty after removing {child_count} bookmarks",
        safety_level='safe',
        bookmark_id=parent_id
    )
```

#### **Category 4: Unused Parent Groups**

```python
if is_parent and not is_bookmark_used(parent_id):
    opportunity = CleanupOpportunity(
        item_type='bookmark_empty_group',
        item_name=parent_name,
        location='Bookmarks',
        reason="Group not used by navigators and has no active bookmarks",
        safety_level='safe',
        bookmark_id=parent_id
    )
```

---

## 🎯 Visual Filter Analysis

### Visual Filter Structure

Visual-level filters in Power BI:

```json
// In visual.json
{
  "filterConfig": {
    "filters": [
      {
        "name": "Filter_Product",
        "type": "Categorical",
        "isHiddenInViewMode": false,  // Visible to users
        "filter": {
          // Filter expression
        }
      },
      {
        "name": "Filter_Date",
        "type": "Advanced",
        "isHiddenInViewMode": true,  // Hidden from users
        "filter": {
          // Filter expression
        }
      }
    ]
  }
}
```

### Visual Filter Detection

```python
def _analyze_visual_filters(report_dir):
    """
    Scan all pages for visual-level filters
    
    Returns:
    - Count of total filters
    - Count of visible filters (candidates for hiding)
    - Count of already hidden filters
    - Page-by-page breakdown
    """
    
    filter_analysis = {
        'pages_with_visual_filters': {},
        'total_visual_filters': 0,
        'total_visuals_with_filters': 0
    }
    
    pages_dir = report_dir / "definition" / "pages"
    
    for page_dir in pages_dir.iterdir():
        page_name = get_page_display_name(page_dir)
        page_filters = []
        
        visuals_dir = page_dir / "visuals"
        for visual_dir in visuals_dir.iterdir():
            visual_json = visual_dir / "visual.json"
            visual_data = read_json(visual_json)
            
            # Get filter configuration
            filter_config = visual_data.get('filterConfig', {})
            filters = filter_config.get('filters', [])
            
            if filters:
                # Analyze filters
                total = len(filters)
                visible = sum(1 for f in filters 
                             if not f.get('isHiddenInViewMode', False))
                hidden = total - visible
                
                page_filters.append({
                    'visual_id': visual_dir.name,
                    'visual_type': get_visual_type(visual_data),
                    'total_filters': total,
                    'visible_filters': visible,
                    'hidden_filters': hidden
                })
                
                filter_analysis['total_visual_filters'] += total
                filter_analysis['total_visuals_with_filters'] += 1
        
        if page_filters:
            filter_analysis['pages_with_visual_filters'][page_name] = {
                'visuals_with_filters': page_filters,
                'total_filters': sum(v['total_filters'] for v in page_filters),
                'visible_filters': sum(v['visible_filters'] for v in page_filters)
            }
    
    return filter_analysis
```

### Filter Hiding Opportunity

```python
def _find_visual_filter_opportunities(filter_analysis):
    """
    Create cleanup opportunity for hiding visible filters
    
    Only creates opportunity if there are visible filters
    """
    
    total_visible = sum(
        page_info['visible_filters'] 
        for page_info in filter_analysis['pages_with_visual_filters'].values()
    )
    
    if total_visible > 0:
        opportunity = CleanupOpportunity(
            item_type='visual_filter',
            item_name=f"Hide {total_visible} visible visual-level filters",
            location='All Pages',
            reason="Visual filters can clutter the interface and confuse users",
            safety_level='safe',
            filter_count=total_visible
        )
        return [opportunity]
    
    return []
```

### Filter Hiding Process

```python
def _hide_visual_filters(report_dir):
    """
    Hide all visible visual-level filters
    
    Process:
    1. Scan all pages and visuals
    2. Find filters with isHiddenInViewMode = false
    3. Set isHiddenInViewMode = true
    4. Save modified visual.json
    """
    
    results = []
    total_hidden = 0
    
    for page in all_pages:
        for visual in page_visuals:
            visual_data = read_json(visual_json)
            
            filter_config = visual_data.get('filterConfig', {})
            filters = filter_config.get('filters', [])
            
            filters_hidden_in_visual = 0
            
            for filter_obj in filters:
                if not filter_obj.get('isHiddenInViewMode', False):
                    # Hide this filter
                    filter_obj['isHiddenInViewMode'] = True
                    filters_hidden_in_visual += 1
            
            if filters_hidden_in_visual > 0:
                # Save modified visual.json
                write_json(visual_json, visual_data)
                total_hidden += filters_hidden_in_visual
    
    if total_hidden > 0:
        result = RemovalResult(
            item_name=f"{total_visuals} visuals ({total_hidden} filters)",
            item_type='visual_filter',
            success=True,
            filters_hidden=total_hidden
        )
        results.append(result)
    
    return results
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
│   Click "ANALYZE    │
│   REPORT"           │
└──────────┬──────────┘
           │
           ▼
┌─────────────────────────────────────────────┐
│   ReportAnalyzer.analyze_pbip_report()      │
│   ┌──────────────────────────────────────┐  │
│   │ Step 1: Comprehensive Visual Scan    │  │
│   │ - Scan all pages                     │  │
│   │ - Extract all visual types           │  │
│   │ - Collect all identifiers            │  │
│   └──────────────────────────────────────┘  │
│   ┌──────────────────────────────────────┐  │
│   │ Step 2: Analyze Themes               │  │
│   │ - Find active theme                  │  │
│   │ - Scan theme files                   │  │
│   │ - Compare active vs available        │  │
│   └──────────────────────────────────────┘  │
│   ┌──────────────────────────────────────┐  │
│   │ Step 3: Analyze Custom Visuals       │  │
│   │ - Scan publicCustomVisuals           │  │
│   │ - Scan resourcePackages              │  │
│   │ - Scan CustomVisuals directory       │  │
│   │ - Match usage with three types       │  │
│   └──────────────────────────────────────┘  │
│   ┌──────────────────────────────────────┐  │
│   │ Step 4: Analyze Bookmarks            │  │
│   │ - Read bookmarks.json                │  │
│   │ - Process parent/child groups        │  │
│   │ - Validate page references           │  │
│   │ - Scan for navigation usage          │  │
│   └──────────────────────────────────────┘  │
│   ┌──────────────────────────────────────┐  │
│   │ Step 5: Analyze Visual Filters       │  │
│   │ - Scan all visual.json files         │  │
│   │ - Count visible/hidden filters       │  │
│   │ - Page-by-page breakdown             │  │
│   └──────────────────────────────────────┘  │
└────────────────┬────────────────────────────┘
                 │
                 ▼
┌─────────────────────────────────────────────┐
│   Create Cleanup Opportunities              │
│   - Unused themes                           │
│   - Unused custom visuals (3 types)         │
│   - Unused bookmarks (4 categories)         │
│   - Visual filter hiding option             │
└────────────────┬────────────────────────────┘
                 │
                 ▼
┌─────────────────────────────────────────────┐
│   Display Opportunities in UI               │
│   ┌──────────────────────────────────────┐  │
│   │ Categorize by Type                   │  │
│   │ - Themes section                     │  │
│   │ - Custom Visuals section             │  │
│   │ - Bookmarks section                  │  │
│   │ - Visual Filters option              │  │
│   └──────────────────────────────────────┘  │
│   ┌──────────────────────────────────────┐  │
│   │ Show Details                         │  │
│   │ - Item name                          │  │
│   │ - Reason for removal                 │  │
│   │ - Safety level                       │  │
│   │ - Size savings estimate              │  │
│   └──────────────────────────────────────┘  │
│   ┌──────────────────────────────────────┐  │
│   │ Enable Selection                     │  │
│   │ - Checkbox per opportunity           │  │
│   │ - "Hide Visual Filters" checkbox     │  │
│   │ - "REMOVE SELECTED" button           │  │
│   └──────────────────────────────────────┘  │
└────────────────┬────────────────────────────┘
                 │
                 ▼
┌─────────────────────┐
│   User Selects      │
│   Items to Remove   │
└──────────┬──────────┘
           │
           ▼
┌─────────────────────┐
│   Click "REMOVE     │
│   SELECTED"         │
└──────────┬──────────┘
           │
           ▼
┌─────────────────────┐
│   Confirmation      │
│   Dialog            │
└──────────┬──────────┘
           │ User confirms
           ▼
┌─────────────────────────────────────────────┐
│   ReportCleanupEngine.remove_unused_items() │
│   ┌──────────────────────────────────────┐  │
│   │ Step 1: Create Backup                │  │
│   │ - Timestamped PBIP copy              │  │
│   │ - Timestamped .Report copy           │  │
│   └──────────────────────────────────────┘  │
│   ┌──────────────────────────────────────┐  │
│   │ Step 2: Remove Themes                │  │
│   │ - Delete from SharedResources        │  │
│   │ - Delete from RegisteredResources    │  │
│   │ - Track bytes freed                  │  │
│   └──────────────────────────────────────┘  │
│   ┌──────────────────────────────────────┐  │
│   │ Step 3: Remove Custom Visuals        │  │
│   │ For each visual:                     │  │
│   │   - Delete CustomVisuals directory   │  │
│   │   - Track bytes freed                │  │
│   └──────────────────────────────────────┘  │
│   ┌──────────────────────────────────────┐  │
│   │ Step 4: Remove Bookmarks             │  │
│   │ - Delete .bookmark.json files        │  │
│   │ - Update bookmarks.json              │  │
│   │ - Update parent group children       │  │
│   └──────────────────────────────────────┘  │
│   ┌──────────────────────────────────────┐  │
│   │ Step 5: Hide Visual Filters          │  │
│   │ (if selected)                        │  │
│   │ - Update all visual.json files       │  │
│   │ - Set isHiddenInViewMode = true      │  │
│   └──────────────────────────────────────┘  │
│   ┌──────────────────────────────────────┐  │
│   │ Step 6: Clean Up report.json         │  │
│   │ - Remove from publicCustomVisuals    │  │
│   │ - Remove from resourcePackages       │  │
│   │ - Update bookmarks references        │  │
│   └──────────────────────────────────────┘  │
└────────────────┬────────────────────────────┘
                 │
                 ▼
┌─────────────────────┐
│   Display Results   │
│   - Success count   │
│   - Bytes freed     │
│   - Filters hidden  │
│   - Error summary   │
└─────────────────────┘
```

---

## 📁 File Structure

### Tool Directory Layout

```
report_cleanup/
├── tool.py                  # BaseTool implementation
├── report_analyzer.py       # Analysis engine (read-only)
├── cleanup_engine.py        # Execution engine (writes/deletes)
├── cleanup_ui.py            # Main UI tab
├── shared_types.py          # Shared data classes
├── __init__.py
└── TECHNICAL_GUIDE.md       # This document
```

### Separation of Concerns

- **`tool.py`**: Tool registration and BaseTool interface
- **`report_analyzer.py`**: Pure analysis logic, never modifies files
- **`cleanup_engine.py`**: Pure execution logic, handles all file operations
- **`cleanup_ui.py`**: UI implementation, delegates to analyzer and engine
- **`shared_types.py`**: Shared dataclasses for type safety

---

## 🔧 Key Operations

### Operation 1: Report Analysis

**User Action**: Click "ANALYZE REPORT"

**Backend Flow**:

1. **Validate PBIP Structure**:
   ```python
   pbip_file = Path(pbip_path)
   report_dir = pbip_file.parent / f"{pbip_file.stem}.Report"
   
   # Validate structure
   if not report_dir.exists():
       raise FileNotFoundError("Report directory not found")
   
   report_json = report_dir / "definition" / "report.json"
   if not report_json.exists():
       raise FileNotFoundError("report.json not found")
   ```

2. **Comprehensive Visual Scan** (First Pass):
   - Scan all pages in definition/pages/
   - Read each visual.json
   - Extract ALL visual identifiers
   - Build master set of used visuals

3. **Analyze Each Resource Type**:
   - **Themes**: Compare active vs available
   - **Custom Visuals**: Check three types against usage
   - **Bookmarks**: Validate pages and scan for navigation
   - **Visual Filters**: Count visible filters

4. **Create Opportunities**:
   - Generate CleanupOpportunity for each unused item
   - Categorize by safety level
   - Calculate size savings

5. **Update UI**:
   - Display categorized opportunities
   - Show totals and summaries
   - Enable removal button

**Result**: Populated opportunity list ready for user review

---

### Operation 2: Theme Removal

**User Action**: Select theme opportunities and click "REMOVE SELECTED"

**Backend Flow**:

```python
def _remove_themes(report_dir, themes_to_remove):
    """Remove unused theme files"""
    
    results = []
    
    for theme_name in themes_to_remove:
        bytes_freed = 0
        success = False
        
        try:
            # Check SharedResources/BaseThemes
            base_theme = report_dir / "StaticResources" / "SharedResources" / "BaseThemes" / f"{theme_name}.json"
            if base_theme.exists():
                bytes_freed += base_theme.stat().st_size
                base_theme.unlink()  # Delete file
                success = True
            
            # Check RegisteredResources for custom themes
            reg_resources = report_dir / "StaticResources" / "RegisteredResources"
            for theme_file in reg_resources.glob("*.json"):
                if theme_file.name == theme_name or theme_file.stem == theme_name:
                    bytes_freed += theme_file.stat().st_size
                    theme_file.unlink()  # Delete file
                    success = True
        
        except Exception as e:
            success = False
            error_message = str(e)
        
        results.append(RemovalResult(
            item_name=theme_name,
            item_type='theme',
            success=success,
            bytes_freed=bytes_freed
        ))
    
    return results
```

---

### Operation 3: Custom Visual Removal

**Handles Three Different Types**:

```python
def _remove_custom_visuals(report_dir, visuals_to_remove):
    """Remove custom visuals handling three types"""
    
    results = []
    custom_visuals_dir = report_dir / "CustomVisuals"
    
    for visual_opportunity in visuals_to_remove:
        visual_id = visual_opportunity.visual_id
        item_type = visual_opportunity.item_type
        
        bytes_freed = 0
        success = False
        
        try:
            visual_dir = custom_visuals_dir / visual_id
            
            # Calculate size before deletion
            if visual_dir.exists():
                bytes_freed = sum(f.stat().st_size 
                                for f in visual_dir.rglob('*') 
                                if f.is_file())
            
            if item_type == 'custom_visual_hidden':
                # Hidden visual: Only delete directory
                if visual_dir.exists():
                    force_remove_directory(visual_dir)
                    success = True
            
            elif item_type == 'custom_visual_build_pane':
                # Build pane visual: Delete directory if exists
                # (report.json cleanup happens later)
                if visual_dir.exists():
                    force_remove_directory(visual_dir)
                success = True  # Success even without directory
        
        except Exception as e:
            error_message = str(e)
        
        results.append(RemovalResult(
            item_name=visual_opportunity.item_name,
            item_type='custom_visual',
            success=success,
            bytes_freed=bytes_freed
        ))
    
    return results
```

**Windows-Specific Directory Removal**:

```python
def _force_remove_directory(directory_path):
    """
    Force remove directory handling Windows permission issues
    
    Windows can lock files, so we need aggressive handling:
    1. Try normal removal
    2. If fails, try with error handler for read-only files
    3. If fails, manually walk and delete each file
    """
    
    try:
        shutil.rmtree(directory_path)
    except PermissionError:
        # Try with error handler
        shutil.rmtree(directory_path, onerror=handle_remove_readonly)
    except Exception:
        # Last resort: manual cleanup
        _manual_directory_cleanup(directory_path)

def handle_remove_readonly(func, path, exc):
    """Clear read-only bit and retry"""
    os.chmod(path, stat.S_IWRITE)
    func(path)
```

---

### Operation 4: Bookmark Removal

**Handles Parent/Child Relationships**:

```python
def _remove_bookmarks(report_dir, bookmarks_to_remove):
    """
    Remove bookmarks handling parent/child groups
    
    Process:
    1. Mark bookmarks for removal
    2. Actual removal happens in _cleanup_report_references
    """
    
    results = []
    
    for bookmark_opportunity in bookmarks_to_remove:
        bookmark_id = bookmark_opportunity.bookmark_id
        display_name = bookmark_opportunity.item_name
        
        # Just mark success here
        # Actual removal in _cleanup_report_references
        results.append(RemovalResult(
            item_name=display_name,
            item_type='bookmark',
            success=True,
            bytes_freed=0
        ))
    
    return results
```

**report.json Cleanup** (The actual removal):

```python
def _cleanup_report_references(report_dir, removed_themes, 
                              removed_visuals, removed_bookmarks):
    """
    Update report.json and bookmark files
    
    Critical for:
    - Removing visual references from publicCustomVisuals
    - Removing visual packages from resourcePackages
    - Removing bookmark references from bookmarks.json
    - Deleting individual .bookmark.json files
    - Updating parent group children arrays
    """
    
    # Clean up publicCustomVisuals
    if removed_visuals:
        visual_ids = [v.visual_id for v in removed_visuals]
        report_data['publicCustomVisuals'] = [
            vid for vid in report_data['publicCustomVisuals'] 
            if vid not in visual_ids
        ]
    
    # Clean up resourcePackages
    updated_packages = []
    for package in report_data.get('resourcePackages', []):
        if package.get('type') == 'CustomVisual':
            if package.get('name') not in visual_ids:
                updated_packages.append(package)
        else:
            # Keep non-visual packages
            updated_packages.append(package)
    
    report_data['resourcePackages'] = updated_packages
    
    # Clean up bookmarks
    if removed_bookmarks:
        bookmark_ids = [b.bookmark_id for b in removed_bookmarks]
        
        # Update bookmarks.json
        bookmarks_json = bookmarks_dir / "bookmarks.json"
        bookmarks_data = read_json(bookmarks_json)
        
        updated_items = []
        for item in bookmarks_data.get('items', []):
            item_id = item.get('name')
            
            # Skip items being removed completely
            if item_id in bookmark_ids:
                continue
            
            # Update parent group children arrays
            if 'children' in item:
                original_children = item.get('children', [])
                updated_children = [
                    child_id for child_id in original_children 
                    if child_id not in bookmark_ids
                ]
                item['children'] = updated_children
            
            updated_items.append(item)
        
        bookmarks_data['items'] = updated_items
        write_json(bookmarks_json, bookmarks_data)
        
        # Delete individual bookmark files
        for bookmark_id in bookmark_ids:
            bookmark_file = bookmarks_dir / f"{bookmark_id}.bookmark.json"
            if bookmark_file.exists():
                bookmark_file.unlink()
    
    # Write updated report.json
    write_json(report_json, report_data)
```

---

## 🔒 Safety & Backup System

### Automatic Backup Creation

**Triggered**: Before ANY removal operations

**Process**:

```python
def _create_backup(pbip_path):
    """
    Create timestamped backup of PBIP file and report directory
    
    Creates:
    - {filename}_backup_{timestamp}.pbip
    - {filename}_backup_{timestamp}.Report/
    """
    
    pbip_file = Path(pbip_path)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_suffix = f"_backup_{timestamp}"
    
    # Backup PBIP file
    backup_pbip = pbip_file.parent / f"{pbip_file.stem}{backup_suffix}.pbip"
    shutil.copy2(pbip_file, backup_pbip)
    
    # Backup Report directory
    report_dir = pbip_file.parent / f"{pbip_file.stem}.Report"
    backup_report_dir = pbip_file.parent / f"{pbip_file.stem}{backup_suffix}.Report"
    shutil.copytree(report_dir, backup_report_dir)
    
    log(f"✅ Backup created: {backup_pbip.name}")
    log(f"✅ Backup created: {backup_report_dir.name}")
```

**Example Backup**:
```
Before:
- MyReport.pbip
- MyReport.Report/

After:
- MyReport.pbip
- MyReport.Report/
- MyReport_backup_20251021_143022.pbip
- MyReport_backup_20251021_143022.Report/
```

### Safety Levels

**Three Safety Levels**:

```python
class SafetyLevel:
    SAFE = 'safe'       # Confirmed unused, safe to remove
    WARNING = 'warning' # Likely unused, but verify first
    RISKY = 'risky'     # Could be used, review carefully
```

**Assignment Logic**:

- **SAFE**:
  - Themes not matching active theme
  - Custom visuals with 0 usage found
  - Hidden custom visuals (can't be used)
  - Bookmarks with missing pages
  - Empty parent bookmark groups
  - Visual filters (just hiding, not deleting)

- **WARNING**:
  - Bookmarks with no navigation found (could use bookmark pane in service)

- **RISKY**:
  - Currently not used (reserved for future)

### Validation Checks

**Pre-Operation Validation**:

```python
def validate_pbip_structure(pbip_path):
    """Validate PBIP before operations"""
    
    # Check file exists
    if not Path(pbip_path).exists():
        raise FileNotFoundError("PBIP file not found")
    
    # Check .Report directory exists
    report_dir = Path(pbip_path).parent / f"{Path(pbip_path).stem}.Report"
    if not report_dir.exists():
        raise FileNotFoundError("Report directory not found")
    
    # Check critical files
    required_files = [
        report_dir / "definition" / "report.json",
        report_dir / "definition" / "pages",
    ]
    
    for required_file in required_files:
        if not required_file.exists():
            raise FileNotFoundError(f"Required file not found: {required_file}")
    
    return True
```

---

## 🚀 Future Enhancements

### Potential Improvements

1. **Advanced Usage Detection**:
   - Scan for visual usage in DAX measures
   - Detect theme references in code
   - Find bookmark usage in URL parameters
   - Track custom visual usage in calculated columns

2. **Dry Run Mode**:
   - Preview exact changes before applying
   - Show before/after file diffs
   - Estimate total size reduction
   - Generate removal report without executing

3. **Selective Backup**:
   - Option to skip backup for trusted operations
   - Backup only affected files (not full directory)
   - Compressed backup archives
   - Automatic cleanup of old backups

4. **Undo Capability**:
   - Track all operations in session
   - One-click undo to restore from backup
   - Partial undo for specific items
   - Operation history log

5. **Batch Processing**:
   - Analyze multiple PBIP files at once
   - Apply consistent cleanup rules across reports
   - Generate cleanup summary reports
   - Export/import cleanup configurations

6. **Smart Recommendations**:
   - ML-based usage prediction
   - Historical analysis of removals
   - Learn from user decisions
   - Suggest cleanup schedules

7. **Integration Features**:
   - Coordinate with Report Merger (cleanup before merge)
   - Coordinate with Layout Optimizer
   - Export cleanup opportunities to CSV
   - Import removal lists from external tools

8. **Enhanced Bookmark Analysis**:
   - Detect bookmark usage in report URL parameters
   - Find bookmarks used via API/embedding
   - Track bookmark usage in Power BI Service
   - Identify bookmark chains/dependencies

9. **Advanced Visual Filter Management**:
   - Selective filter hiding by page/visual
   - Filter hiding presets (hide categorical, keep measures)
   - Export filter configurations
   - Filter usage analytics

10. **Performance Optimization**:
    - Parallel page scanning
    - Incremental analysis (only changed pages)
    - Cached analysis results
    - Background analysis while user reviews

---

## 📝 Code Quality Notes

### Strengths

- ✅ **Clean three-tier architecture** (Analyzer, Engine, UI)
- ✅ **Read-only analysis** (no side effects during scanning)
- ✅ **Safety-first approach** (automatic backups)
- ✅ **Comprehensive detection** (multiple methods for each type)
- ✅ **Hierarchical awareness** (parent/child bookmark groups)
- ✅ **Type safety** (dataclasses throughout)
- ✅ **Detailed logging** (user-friendly progress messages)
- ✅ **Error handling** (graceful degradation)
- ✅ **Windows compatibility** (handles permission issues)

### Standards Followed

- **PEP 8**: Python style guide compliance
- **Type Hints**: Full coverage (Python 3.8+)
- **Docstrings**: Comprehensive documentation
- **Dataclasses**: For structured data
- **Separation of Concerns**: Clear boundaries between components
- **Immutability**: Analysis data never modified by engine

---

## 📞 Support

For questions about this tool's architecture or implementation:

- **Documentation**: This file and inline code comments
- **Issues**: Report bugs via GitHub Issues
- **Professional Support**: [Analytic Endeavors](https://www.analyticendeavors.com)

---

**Document Version**: 1.0  
**Tool Version**: v1.0.0  
**Last Updated**: October 21, 2025  
**Author**: Reid Havens, Analytic Endeavors
