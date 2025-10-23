# 🎯 Advanced Copy Tool - Technical Architecture & How It Works

**Version**: v1.1.1  
**Built by**: Reid Havens of Analytic Endeavors  
**Last Updated**: October 22, 2025

---

## 📚 Table of Contents

1. [Overview](#overview)
2. [Architecture](#architecture)
3. [Core Components](#core-components)
4. [Operation Modes](#operation-modes)
5. [Data Flow](#data-flow)
6. [Key Algorithms](#key-algorithms)
7. [File Structure](#file-structure)
8. [How Operations Work](#how-operations-work)
9. [Error Handling](#error-handling)
10. [Future Enhancements](#future-enhancements)

---

## 🎯 Overview

The **Advanced Copy Tool** enables sophisticated copying of Power BI report pages, bookmarks, and visual elements within a single PBIP file or across multiple PBIP files. It handles complex dependencies like bookmark groups, visual actions, and page references.

### What Makes It "Advanced"

1. **Dependency Tracking** - Automatically maps and updates all cross-references
2. **Smart ID Generation** - Creates unique IDs that avoid conflicts
3. **Action Reassignment** - Updates visual actions to point to copied bookmarks
4. **Group Preservation** - Maintains bookmark group structures
5. **Cross-PBIP Support** - Works within a report or between different reports

---

## 🏗️ Architecture

### Design Pattern: **Composition Over Inheritance**

The tool uses a **composition-based architecture** with specialized modules:

```
AdvancedCopyEngine (Orchestrator)
    ├── BookmarkAnalyzer      → Analyzes bookmarks & visual relationships
    ├── PageOperations        → Handles page/bookmark copying
    ├── BookmarkGroupManager  → Manages bookmark groups & navigators
    └── VisualActionUpdater   → Updates visual actions & schema validation
```

### Key Principles

- **Single Responsibility**: Each module has one clear job
- **Shared State**: Critical mappings passed via `__init__` for consistency
- **Stateless Operations**: Modules don't hold operation-specific state
- **Composition**: Engine composes helpers rather than inheriting

---

## 🧩 Core Components

### 1. **AdvancedCopyEngine** (`advanced_copy_core.py`)

**Role**: Orchestrator that coordinates all copy operations

**Responsibilities**:
- Public API for all copy operations
- Manages shared state (page/bookmark mappings)
- Validates PBIP structure
- Schema validation and fixing
- Progress logging

**Shared State Managed**:
```python
self._page_id_mapping = {}              # old_page_id → new_page_id
self._bookmark_copy_mapping = {}        # new_bookmark_id → old_bookmark_id
self._original_bookmark_names = []      # Original bookmark order
self._copied_bookmarks_order = []       # Copied bookmark order
self._bookmark_group_mapping = {}       # old_group_id → new_group_id
self._perpage_bookmark_tracking = {}    # Per-page bookmark tracking
```

---

### 2. **BookmarkAnalyzer** (`advanced_copy_bookmark_analyzer.py`)

**Role**: Analyzes bookmarks to extract dependencies and relationships

**Key Capabilities**:
- **Visual Extraction**: Finds visuals controlled by bookmarks
- **Capture Mode Detection**: Determines "Selected visuals" vs "All visuals"
- **Page-Bookmark Mapping**: Creates relationship maps
- **Link Detection**: Finds visuals that link to bookmarks
- **Group Analysis**: Extracts bookmark group structures

**Critical Methods**:
```python
extract_bookmark_visuals(bookmark_data, page_dir)  # Extract & categorize visuals (groups vs individuals)
get_all_group_members(page_dir, group_id)          # Get current group members at runtime
find_visuals_by_ids(page_dir, visual_ids)          # Find visual directories
find_visuals_linking_to_bookmarks(...)             # Find shapes/buttons with links
analyze_pages_with_bookmarks(report_dir)           # Map pages to bookmarks
get_bookmark_groups(bookmarks_dir)                 # Extract group structure
```

**Recent Enhancements (v1.1.1)**:
- **Group vs Individual Detection**: `extract_bookmark_visuals()` now categorizes visuals as group containers or individuals
- **Dynamic Group Membership**: `get_all_group_members()` gets current members at copy-time, not bookmark snapshot
- **Parent Group Hierarchy**: Automatically detects and includes parent group containers

---

### 3. **PageOperations** (`advanced_copy_operations.py`)

**Role**: Handles core page and bookmark copying with ID management

**Key Capabilities**:
- **Page Duplication**: Creates new pages with unique IDs
- **Bookmark Copying**: Duplicates bookmarks with proper references
- **ID Generation**: Creates unique Power BI-compatible IDs
- **Reference Updates**: Updates page references in bookmarks
- **Cross-PBIP Support**: Handles copying between reports

**Critical Methods**:
```python
copy_single_page(pages_dir, page_data)               # Copy page within PBIP
copy_single_page_cross_pbip(...)                     # Copy page across PBIPs
copy_page_bookmarks(...)                             # Copy page's bookmarks
copy_bookmarks_for_page(...)                         # Copy bookmarks to specific page
update_bookmark_page_reference(bookmark_data, ...)   # Update page refs
```

**ID Generation Strategy**:
- Uses `uuid.uuid4()` for globally unique IDs
- Truncates to 20 characters (Power BI format)
- Removes hyphens for compatibility
- Ensures uniqueness through mapping tracking

---

### 4. **BookmarkGroupManager** (`advanced_copy_bookmark_groups.py`)

**Role**: Manages bookmark groups and navigator visuals

**Key Capabilities**:
- **Group Creation**: Creates/updates bookmark groups in metadata
- **Group Naming**: Intelligent group name generation
- **Navigator Copying**: Copies bookmark navigator visuals
- **Order Preservation**: Maintains bookmark display order
- **Per-Page Groups**: Creates page-specific groups in per-page mode

**Critical Methods**:
```python
update_bookmarks_json_after_copy(...)      # Updates bookmarks metadata
copy_bookmark_navigators(...)              # Copies navigator visuals
capture_original_bookmarks(...)            # Records original order
_create_bookmark_group(...)                # Creates group structure
```

**Group Naming Strategy**:
- **Same-PBIP Mode**: Uses copy suffix (e.g., "My Group (Copy)")
- **Per-Page Mode**: Uses page names (e.g., "My Bookmarks (Page 2)")
- **Collision Detection**: Increments numbers if names exist

---

### 5. **VisualActionUpdater** (`advanced_copy_visual_actions.py`)

**Role**: Updates visual actions and validates schemas

**Key Capabilities**:
- **Action Scanning**: Finds all visuals with bookmark actions
- **Reference Updates**: Updates actions to point to copied bookmarks
- **Schema Validation**: Ensures correct Power BI schemas
- **Cross-Reference**: Handles complex action dependencies

**Critical Methods**:
```python
update_visual_actions_in_page(...)         # Update actions in copied page
_find_copied_bookmark(original_id)         # Maps original to copied bookmark
validate_and_fix_schemas(report_dir)       # Fixes schema references
final_schema_safeguard(report_dir)         # Final validation pass
```

**Action Update Process**:
1. Scan all visuals on copied page
2. Find `visualLink` objects with bookmark references
3. Look up copied bookmark ID from mapping
4. Update `Literal.Value` to new bookmark ID
5. Write back to `visual.json`

---

## 🔄 Operation Modes

### Mode 1: **Full Page Copy (Same PBIP)**

**Use Case**: Duplicate pages within the same report

**Process**:
1. Select source pages with bookmarks
2. Generate new unique page IDs
3. Copy page directories with new IDs
4. Duplicate all bookmarks with new IDs
5. Create bookmark groups with copy suffixes
6. Update visual actions to reference copied bookmarks
7. Update metadata files (pages.json, bookmarks.json)

**ID Strategy**: All new UUIDs, "(Copy)" suffixes on display names

---

### Mode 2: **Full Page Copy (Cross-PBIP)**

**Use Case**: Copy pages from one report to another

**Process**:
1. Validate both source and target reports
2. Generate new unique page IDs for target
3. Copy page directories to target report
4. Copy bookmarks to target with new IDs
5. Update page references to target page IDs
6. Create bookmark groups in target
7. Update visual actions in target
8. Update target metadata only (source untouched)

**ID Strategy**: All new UUIDs, no suffixes (different reports)

**Key Difference**: No collision detection needed across reports

---

### Mode 3: **Bookmark + Visual Copy (Same PBIP)**

**Use Case**: Copy specific bookmarks and their visuals to multiple pages

**Process**:
1. Analyze selected bookmarks to find controlled visuals
2. Find visual directories and visuals with bookmark links
3. For each target page:
   - Copy visual directories with **SAME IDs**
   - Create **NEW bookmarks** with page-specific names
   - Update bookmark page references
   - Copy bookmark navigators
4. Create per-page bookmark groups
5. Update visual actions with page-specific mappings

**ID Strategy**: Visuals keep same IDs, bookmarks get new IDs per page

**Why Same Visual IDs?**: Bookmarks reference visuals by ID - keeping IDs consistent means bookmarks work immediately without remapping

---

### Mode 4: **Bookmark + Visual Copy (Cross-PBIP)**

**Use Case**: Copy bookmarks and visuals from one report to another

**Process**:
1. Analyze bookmarks in source report
2. Find visuals in source page
3. For each target page in target report:
   - Copy visual directories from source to target with **SAME IDs**
   - Create **NEW bookmarks** in target
   - Update bookmark page references to target pages
   - Copy navigators from source to target
4. Create per-page groups in target metadata
5. Update actions in target pages

**ID Strategy**: Visuals keep same IDs across reports, bookmarks get new IDs

**Key Difference**: Reads from source, writes to target completely separately

---

## 📊 Data Flow

### Full Page Copy Flow

```
┌─────────────────┐
│  User Selects   │
│  Source Pages   │
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│   Generate      │
│   New Page IDs  │ ──► Store in _page_id_mapping
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│  Copy Page      │
│  Directories    │
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│  Copy Page      │
│  Bookmarks      │ ──► Store in _bookmark_copy_mapping
└────────┬────────┘                 _copied_bookmarks_order
         │
         ▼
┌─────────────────┐
│  Create Groups  │ ──► Uses _original_bookmark_names
└────────┬────────┘       and _copied_bookmarks_order
         │
         ▼
┌─────────────────┐
│  Update Visual  │ ──► Uses _bookmark_copy_mapping
│    Actions      │       and _page_id_mapping
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│  Update Report  │
│   Metadata      │
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│  Validate       │
│   Schemas       │
└─────────────────┘
```

### Bookmark + Visual Copy Flow

```
┌─────────────────┐
│  User Selects   │
│  Bookmarks &    │
│  Target Pages   │
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│  Analyze        │
│  Bookmarks      │ ──► Extract visual IDs
└────────┬────────┘       Find linked visuals
         │
         ▼
┌─────────────────┐
│  Find Visual    │
│  Directories    │
└────────┬────────┘
         │
         ▼
     ┌───┴───┐
     │ Loop  │ ◄───────┐
     │ Pages │         │
     └───┬───┘         │
         │             │
         ▼             │
┌─────────────────┐    │
│  Copy Visuals   │    │
│  (Same IDs!)    │    │
└────────┬────────┘    │
         │             │
         ▼             │
┌─────────────────┐    │
│  Copy Bookmarks │    │
│  (New IDs per   │ ───┼─► Store in _bookmark_copy_mapping
│   page!)        │    │    _copied_bookmarks_order
└────────┬────────┘    │    _perpage_bookmark_tracking
         │             │
         ▼             │
┌─────────────────┐    │
│  Copy Navigators│    │
└────────┬────────┘    │
         │             │
         └─────────────┘
         │
         ▼
┌─────────────────┐
│  Create Per-    │ ──► Uses _perpage_bookmark_tracking
│  Page Groups    │
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│  Update Visual  │ ──► Uses page-specific mappings
│    Actions      │
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│  Update Report  │
│   Metadata      │
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│  Validate       │
│   Schemas       │
└─────────────────┘
```

---

## 🔍 Key Algorithms

### 1. Bookmark Analysis Algorithm (Enhanced v1.1.1)

**Purpose**: Extract and categorize all visuals controlled by a bookmark

**Steps**:
1. **Check targetVisualNames** (preferred method for "Selected visuals"):
   ```python
   if "options" in bookmark and "targetVisualNames" in bookmark["options"]:
       visual_ids = bookmark["options"]["targetVisualNames"]
   ```

2. **Check visualContainerGroups** (identifies GROUP toggles):
   ```python
   if "visualContainerGroups" in exploration_state["sections"]:
       group_ids = list(visualContainerGroups.keys())
       # These are group toggle bookmarks
   ```

3. **Categorize each visual** by reading its visual.json:
   ```python
   for visual_id in visual_ids:
       visual_data = read_visual_json(page_dir / "visuals" / visual_id / "visual.json")
       
       if 'visualGroup' in visual_data:
           # This is a group container
           result['group_ids'].append(visual_id)
       else:
           # This is an individual visual
           result['individual_visual_ids'].append(visual_id)
           
           # Check for parent groups
           if 'parentGroupName' in visual_data:
               # Recursively add parent hierarchy
               add_parent_groups_recursively()
   ```

4. **For group toggles, get current members**:
   ```python
   for group_id in result['group_ids']:
       current_members = get_all_group_members(page_dir, group_id)
       # Returns ALL visuals where parentGroupName == group_id
       # This is current state, NOT bookmark snapshot!
   ```

**Result**: Dictionary with categorized visual IDs:
```python
{
    'group_ids': [list of group container IDs],
    'individual_visual_ids': [list of individual visual IDs],
    'all_visual_ids': [combined list for backwards compatibility]
}
```

**Key Innovation**: Group toggles now copy ALL current group members, not just what was in the bookmark snapshot. This handles visuals added to groups after bookmark creation.

---

### 2. Visual Action Update Algorithm

**Purpose**: Update visual actions in copied pages to reference copied bookmarks

**Steps**:
1. **Scan all visuals** on the copied page
2. **For each visual**, check for `visualContainerObjects.visualLink`
3. **Extract bookmark reference**:
   ```python
   bookmark_id = props['bookmark']['expr']['Literal']['Value'].strip("'")
   ```
4. **Lookup copied bookmark**:
   ```python
   # Reverse lookup in mapping (new → old, need old → new)
   for new_id, old_id in bookmark_copy_mapping.items():
       if old_id == original_bookmark_id:
           return new_id
   ```
5. **Update reference**:
   ```python
   bookmark_expr['Literal']['Value'] = f"'{new_bookmark_id}'"
   ```
6. **Write back** to visual.json

**Skipping Logic**: If bookmark wasn't copied, skip the action (expected behavior)

---

### 3. Group Creation Algorithm

**Purpose**: Create bookmark groups with intelligent naming

**Steps**:
1. **Load bookmarks.json** structure
2. **For each unique group** in original bookmarks:
   - Determine group name based on mode:
     - **Same-PBIP**: Add copy suffix
     - **Per-Page**: Add page name
   - Check for name collisions, increment if needed
   - Create group structure:
     ```python
     {
         "name": group_id,
         "displayName": group_name,
         "children": [list of copied bookmark IDs]
     }
     ```
3. **Append copied bookmarks** not in groups (flat structure)
4. **Write back** to bookmarks.json with correct schema

**Order Preservation**: Maintains original bookmark order using `_copied_bookmarks_order`

---

### 4. ID Generation Algorithm

**Purpose**: Generate unique Power BI-compatible IDs

**Process**:
```python
# Generate UUID
new_id = str(uuid.uuid4())

# Remove hyphens
new_id = new_id.replace('-', '')

# Truncate to 20 characters (Power BI format)
new_id = new_id[:20]

# Result: "a1b2c3d4e5f6g7h8i9j0" (20 chars, no hyphens)
```

**Why This Works**:
- UUID provides global uniqueness
- 20 characters matches Power BI's ID format
- No hyphens ensures compatibility
- Collision probability: Astronomically low

---

## 📁 File Structure

### Tool Directory Layout

```
advanced_copy/
├── advanced_copy_tool.py           # BaseTool implementation
├── logic/                          # Business logic (no UI)
│   ├── __init__.py
│   ├── advanced_copy_core.py       # Main orchestrator
│   ├── advanced_copy_operations.py # Page/bookmark operations
│   ├── advanced_copy_bookmark_analyzer.py  # Bookmark analysis
│   ├── advanced_copy_bookmark_groups.py    # Group management
│   └── advanced_copy_visual_actions.py     # Action updates
├── ui/                             # User interface (no logic)
│   ├── __init__.py
│   ├── advanced_copy_tab.py        # Main UI tab (composition)
│   ├── ui_data_source.py           # Data source selection UI
│   ├── ui_page_selection.py        # Page selection UI
│   ├── ui_bookmark_selection.py    # Bookmark selection UI
│   ├── ui_event_handlers.py        # Event handling
│   └── ui_helpers.py               # Helper methods
├── CLEANUP_RECOMMENDATIONS.md      # Code quality analysis
└── HOW_IT_WORKS.md                 # This document
```

### Separation of Concerns

- **`logic/`**: Pure business logic, no tkinter imports
- **`ui/`**: Pure UI code, delegates logic to engine
- **Mixins**: UI composition for reusability
- **Engine**: Orchestrates logic modules via composition

---

## 🔧 How Operations Work

### Operation 1: Full Page Copy (Same PBIP)

**User Action**: Select pages → Click "Execute Copy"

**Backend Flow**:

1. **`AdvancedCopyEngine.copy_selected_pages()`** is called
2. **Validation**:
   - Check PBIP structure
   - Validate selected pages exist
3. **Execute Copy**:
   - Call `_execute_page_copy_operations()`
4. **For each selected page**:
   - `PageOperations.copy_single_page()`:
     - Generate new page ID (UUID)
     - Copy page directory
     - Update page.json with new ID and display name
     - Store mapping in `_page_id_mapping`
   - `PageOperations.copy_page_bookmarks()`:
     - For each bookmark on page:
       - Generate new bookmark ID
       - Copy bookmark file
       - Update page reference
       - Store mapping in `_bookmark_copy_mapping`
       - Append to `_copied_bookmarks_order`
5. **Create Groups**:
   - `BookmarkGroupManager.update_bookmarks_json_after_copy()`:
     - Read bookmarks.json
     - Create groups with copy suffixes
     - Preserve order using `_copied_bookmarks_order`
     - Write back with correct schema
6. **Update Actions**:
   - `VisualActionUpdater.update_visual_actions_in_page()`:
     - Scan visuals for bookmark actions
     - Update references using `_bookmark_copy_mapping`
7. **Update Metadata**:
   - Rebuild pages.json with new pages
   - Update bookmarks.json with groups
8. **Validate Schemas**:
   - Force correct schemas on metadata files
   - Final safeguard check

**Result**: Copied pages with working bookmarks, groups, and actions

---

### Operation 2: Bookmark + Visual Copy (Same PBIP)

**User Action**: Select bookmarks → Select target pages → Click "Execute"

**Backend Flow**:

1. **`AdvancedCopyEngine.copy_bookmarks_with_visuals()`** is called
2. **Analyze Bookmarks**:
   - `BookmarkAnalyzer.extract_bookmark_visuals()`:
     - Extract visual IDs from `targetVisualNames`
   - `BookmarkAnalyzer.find_visuals_by_ids()`:
     - Find visual directories
   - `BookmarkAnalyzer.find_visuals_linking_to_bookmarks()`:
     - Find shapes/buttons with bookmark links
3. **For each target page**:
   - Copy visual directories with **same IDs**:
     - `shutil.copytree(source_visual_dir, target_visual_dir)`
   - `PageOperations.copy_bookmarks_for_page()`:
     - Generate new bookmark IDs
     - Copy bookmarks with updated page references
     - NO visual ID remapping (visuals keep same IDs!)
     - Track in `_perpage_bookmark_tracking`
   - `BookmarkGroupManager.copy_bookmark_navigators()`:
     - Copy navigator visuals to target page
4. **Create Per-Page Groups**:
   - `BookmarkGroupManager.update_bookmarks_json_after_copy()` with `mode='perpage'`:
     - Create groups named after pages
     - Use `_perpage_bookmark_tracking` for grouping
5. **Update Actions**:
   - `VisualActionUpdater.update_visual_actions_in_page()` with page-specific mapping:
     - Use per-page bookmark mappings
     - Update actions in each target page
6. **Update Metadata & Validate**

**Result**: Bookmarks and visuals copied to multiple pages with proper grouping

---

### Operation 3: Cross-PBIP Operations

**Key Differences**:
- Reads from **source** PBIP directories
- Writes to **target** PBIP directories
- No collision detection needed (different reports)
- No copy suffixes on display names
- Updates only **target** metadata

**Flow**: Similar to same-PBIP but with separate source/target handling

---

## 🛡️ Error Handling

### Validation Strategy

1. **Pre-Flight Checks**:
   - PBIP structure validation
   - Required directory existence
   - JSON file validity

2. **Operation Safeguards**:
   - Try-catch around each file operation
   - Continue on non-critical errors
   - Log warnings for skipped items

3. **Schema Enforcement**:
   - Multiple schema fixing passes
   - Force correct schemas regardless of input
   - Final safeguard before completion

### Error Recovery

```python
try:
    # Attempt operation
    copy_bookmark(...)
except Exception as e:
    # Log warning
    self.log_callback(f"⚠️ Warning: Could not copy bookmark: {e}")
    # Continue with next bookmark
    continue
```

**Philosophy**: Graceful degradation - complete as much as possible, report issues

---

## 🚀 Future Enhancements

### Potential Improvements

1. **Unit Testing**:
   - Add pytest tests for each module
   - Mock file operations for testing
   - Integration tests for full workflows

2. **Configuration File**:
   - Make copy suffixes configurable
   - Allow custom schema URLs
   - User preferences for naming

3. **Logging Levels**:
   - Add DEBUG/INFO/WARNING levels
   - User-configurable verbosity
   - Separate log files for debugging

4. **Performance**:
   - Parallel copying for large operations
   - Progress callbacks for long operations
   - Memory optimization for huge reports

5. **Advanced Features**:
   - Selective visual copying (choose specific visuals)
   - Custom bookmark renaming patterns
   - Merge conflict resolution for cross-PBIP

---

## 📝 Code Quality Notes

### Strengths

- ✅ **Excellent composition architecture**
- ✅ **Complete docstrings throughout**
- ✅ **Proper error handling**
- ✅ **Clean UI/logic separation**
- ✅ **No debug code or print statements**
- ✅ **Consistent naming conventions**
- ✅ **Professional logging**

### Standards Followed

- **PEP 8**: Python style guide compliance
- **Type Hints**: Good coverage (Python 3.8+ compatible)
- **Docstrings**: NumPy/Google style
- **Error Messages**: Clear and actionable
- **Comments**: Explain "why", not "what"

---

## 📝 Recent Changes & Bug Fixes

### v1.1.1 - Group Toggle Detection (October 22, 2025)

**Problem Solved**: Bookmarks store `targetVisualNames` as a snapshot at creation time. If a group gains new members after bookmark creation, those visuals weren't being copied, breaking the bookmark.

**Solution**: 
1. **Detect GROUP toggles** vs individual visual selections by checking `visualContainerGroups`
2. **For groups**: Call `get_all_group_members()` at copy-time to get current members
3. **For individuals**: Use `targetVisualNames` directly for precision

**Benefits**:
- ✅ Groups stay current - new members automatically included
- ✅ No breaking changes - backwards compatible
- ✅ Robust - handles removed members without errors
- ✅ Smart detection - distinguishes groups vs individuals

See `logic/BOOKMARK_GROUP_CHANGES_SUMMARY.md` for detailed implementation notes.

---

## 🎓 Learning Resources

### Understanding PBIP Structure

- **PBIP Format**: Folder-based Power BI project format
- **Key Files**:
  - `pages/page_id/page.json` - Page metadata
  - `bookmarks/bookmark_id.bookmark.json` - Bookmark data
  - `pages/pages.json` - Page order and metadata
  - `bookmarks/bookmarks.json` - Bookmark groups and order

### Power BI Concepts

- **Bookmarks**: Capture report state (filters, visuals, pages)
- **Visual Actions**: Click actions that navigate to bookmarks
- **Bookmark Groups**: Organize bookmarks in hierarchies
- **Navigator Visuals**: Bookmarks visual type for navigation

---

## 📞 Support

For questions about this tool's architecture or implementation:

- **Documentation**: This file and inline code comments
- **Code Review**: See `CLEANUP_RECOMMENDATIONS.md`
- **Issues**: Report bugs via GitHub Issues
- **Professional Support**: [Analytic Endeavors](https://www.analyticendeavors.com)

---

**Document Version**: 1.1  
**Tool Version**: v1.1.1  
**Last Updated**: October 22, 2025  
**Author**: Reid Havens, Analytic Endeavors
