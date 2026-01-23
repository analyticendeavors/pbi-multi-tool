# AE Multi-Tool - Architecture & Framework Guide

**Built by Reid Havens of Analytic Endeavors**

**Version:** 1.0.0
**Last Updated:** January 2026

---

## Table of Contents

1. [Overview](#1-overview)
2. [Core Framework Architecture](#2-core-framework-architecture)
3. [Application Structure](#3-application-structure)
4. [Tool Lifecycle](#4-tool-lifecycle)
5. [Cross-Cutting Concerns](#5-cross-cutting-concerns)
6. [MCP Server Integration](#6-mcp-server-integration)
7. [Adding New Tools](#7-adding-new-tools)
8. [File Organization](#8-file-organization)
9. [Design Patterns Used](#9-design-patterns-used)
10. [Key Technical Decisions](#10-key-technical-decisions)
11. [Future Architecture Enhancements](#11-future-architecture-enhancements)
12. [Appendix](#12-appendix)

---

## 1. Overview

### 1.1 What is the AE Multi-Tool?

The **AE Multi-Tool** is a professional-grade Power BI report management suite that consolidates five specialized tools into a unified application with a plugin-like architecture. It provides intelligent automation for working with Power BI's enhanced PBIP (PBIR) format files.

**Current Tools:**
1. **Report Merger** - Consolidates multiple PBIP reports into one
2. **Layout Optimizer** - Intelligently arranges data model tables
3. **Report Cleanup** - Removes empty/unused pages and bookmarks
4. **Column Width Tool** - Standardizes table column widths across visuals
5. **Advanced Copy** - Duplicates pages with bookmarks

### 1.2 Architecture Philosophy

The AE Multi-Tool follows these core principles:

**🎯 Plugin Architecture**
- Tools are discovered and registered automatically
- New tools can be added without modifying the core framework
- Clean separation between framework and tool implementations

**🧩 Composition Over Inheritance**
- Heavy use of mixins for reusable UI patterns
- BaseToolTab + Mixins = Flexible, maintainable components
- Avoid deep inheritance hierarchies

**🔒 Security & Reliability**
- Enhanced logging and error handling
- Input validation at multiple levels
- Secure file operations with proper path handling

**🎨 Professional UI/UX**
- Consistent theming via AppConstants
- Tabbed interface with dynamic sizing
- Progress tracking and real-time feedback

### 1.3 Design Principles

1. **Framework Abstraction** - Core framework provides reusable patterns
2. **Tool Independence** - Tools can evolve independently
3. **Mixin Composition** - Compose functionality from small, focused mixins
4. **Explicit Contracts** - Clear interfaces via abstract base classes
5. **Discoverable** - Tools register themselves automatically

---

## 2. Core Framework Architecture

### 2.1 BaseTool Pattern

**Location:** `src/core/tool_manager.py`

The `BaseTool` abstract base class defines the contract that all tools must implement.

#### Purpose & Responsibilities

```
BaseTool (ABC)
├── Defines tool metadata (id, name, version, description)
├── Provides logging infrastructure
├── Enforces UI creation interface
├── Supports dependency checking
└── Enables help content integration
```

#### Key Methods

```python
class BaseTool(ABC):
    """Base class for all Power BI tools"""
    
    def __init__(self, tool_id: str, name: str, description: str, version: str):
        self.tool_id = tool_id          # Unique identifier
        self.name = name                # Display name
        self.description = description  # Tool description
        self.version = version          # Semantic version
        self.enabled = True             # Can be disabled
    
    @abstractmethod
    def create_ui_tab(self, parent, main_app) -> 'BaseToolTab':
        """Create the UI tab for this tool"""
        pass
    
    @abstractmethod
    def get_tab_title(self) -> str:
        """Get the display title (with emoji)"""
        pass
    
    @abstractmethod
    def get_help_content(self) -> Dict[str, Any]:
        """Get help content specific to this tool"""
        pass
    
    def can_run(self) -> bool:
        """Check if tool can run (dependencies, etc.)"""
        return True
```

#### Tool Registration Process

1. **Discovery** - ToolManager scans `tools/` directory
2. **Import** - Each tool module is dynamically imported
3. **Instantiation** - Tool class is instantiated
4. **Validation** - `can_run()` checks dependencies
5. **Registration** - Tool added to ToolManager registry

#### Example Tool Implementation

```python
# src/tools/advanced_copy/advanced_copy_tool.py
class AdvancedCopyTool(BaseTool):
    def __init__(self):
        super().__init__(
            tool_id="advanced_copy",
            name="Advanced Copy",
            description="Duplicate pages with bookmarks",
            version="1.1.0"
        )
    
    def create_ui_tab(self, parent, main_app) -> BaseToolTab:
        return AdvancedCopyTab(parent, main_app)
    
    def get_tab_title(self) -> str:
        return "📋 Advanced Copy"
    
    def get_help_content(self) -> Dict[str, Any]:
        return {...}  # Help documentation
    
    def can_run(self) -> bool:
        try:
            from tools.advanced_copy.logic.advanced_copy_core import AdvancedCopyEngine
            return True
        except ImportError:
            return False
```

---

### 2.2 BaseToolTab Pattern

**Location:** `src/core/ui_base.py`

The `BaseToolTab` abstract base class provides the UI framework for all tool tabs.

#### Purpose & Responsibilities

```
BaseToolTab (ABC)
├── Creates main frame for tab
├── Provides common UI components (logs, progress, buttons)
├── Handles background thread execution
├── Manages validation and error handling
├── Enforces UI lifecycle methods
└── Provides styling utilities
```

#### Key Methods

```python
class BaseToolTab(ABC):
    """Base class for tool UI tabs"""
    
    def __init__(self, parent, main_app, tool_id: str, tool_name: str):
        self.parent = parent
        self.main_app = main_app
        self.tool_id = tool_id
        self.tool_name = tool_name
        self.frame = ttk.Frame(parent, padding="20")
        self.is_busy = False
        self.progress_bar = None
        self.log_text = None
    
    @abstractmethod
    def setup_ui(self) -> None:
        """Setup the UI for this tab - MUST implement"""
        pass
    
    @abstractmethod
    def reset_tab(self) -> None:
        """Reset tab to initial state - MUST implement"""
        pass
    
    @abstractmethod
    def show_help_dialog(self) -> None:
        """Show help dialog - MUST implement"""
        pass
    
    # Common UI component creators
    def create_file_input_section(self, parent, title, file_types, guide_text):
        """Creates standardized file input with browse button"""
        
    def create_log_section(self, parent, title):
        """Creates standardized log area with export/clear"""
        
    def create_progress_bar(self, parent):
        """Creates standardized progress bar"""
        
    def create_action_buttons(self, parent, buttons):
        """Creates standardized action buttons"""
    
    # Utility methods
    def log_message(self, message: str):
        """Log message to UI"""
        
    def update_progress(self, progress_percent: int, message: str):
        """Update progress bar"""
        
    def run_in_background(self, target_func, success_callback, error_callback):
        """Execute function in background thread"""
```

#### UI Component Creators

BaseToolTab provides factory methods for common UI components:

| Method | Creates | Returns |
|--------|---------|---------|
| `create_file_input_section()` | File path input with browse button | Dict with frame, path_var, entry, browse_button |
| `create_log_section()` | Log text area with export/clear | Dict with frame, text_widget, buttons |
| `create_progress_bar()` | Progress bar with status label | Dict with frame, progress_bar, label |
| `create_action_buttons()` | Styled button set | Dict mapping text to button widgets |

---

### 2.3 Mixins

Mixins provide reusable functionality that can be composed into tool tabs. They follow the principle of **single responsibility** - each mixin does one thing well.

#### Core Mixins

**FileInputMixin** (`core/ui_base.py`)
```python
class FileInputMixin:
    """Mixin for tabs that need file input functionality"""
    
    def clean_file_path(self, path: str) -> str:
        """Remove quotes, normalize separators"""
        
    def setup_path_cleaning(self, path_var: tk.StringVar):
        """Auto-clean paths on StringVar changes"""
```

**ValidationMixin** (`core/ui_base.py`)
```python
class ValidationMixin:
    """Mixin for tabs that need validation functionality"""
    
    def validate_file_exists(self, file_path: str, file_description: str):
        """Validate file exists and is accessible"""
        
    def validate_pbip_file(self, file_path: str, file_description: str):
        """Validate PBIP file structure (.pbip + .Report directory)"""
```

#### Tool-Specific Mixins (Example: Advanced Copy)

Advanced Copy demonstrates advanced mixin usage:

```python
# src/tools/advanced_copy/ui/advanced_copy_tab.py
class AdvancedCopyTab(
    BaseToolTab,           # Core UI framework
    FileInputMixin,        # File handling
    ValidationMixin,       # Input validation
    DataSourceMixin,       # Data source UI
    EventHandlersMixin,    # Event handlers
    HelpersMixin,          # Helper methods
    PageSelectionMixin,    # Page selection UI
    BookmarkSelectionMixin # Bookmark selection UI
):
    """Advanced Copy tab with full mixin composition"""
```

**Mixin Responsibilities:**
- **DataSourceMixin** - Renders file input sections
- **PageSelectionMixin** - Renders page selection listbox
- **BookmarkSelectionMixin** - Renders bookmark tree view
- **EventHandlersMixin** - Handles user events (browse, analyze, copy)
- **HelpersMixin** - Background operations and help dialogs

#### How Mixins Extend Base Classes

```
┌─────────────────────────────────────────────────────┐
│                    Python MRO                       │
│  (Method Resolution Order - Left to Right)          │
└─────────────────────────────────────────────────────┘

AdvancedCopyTab.__init__()
    ↓
1. Check AdvancedCopyTab
2. Check BookmarkSelectionMixin
3. Check PageSelectionMixin
4. Check HelpersMixin
5. Check EventHandlersMixin
6. Check DataSourceMixin
7. Check ValidationMixin
8. Check FileInputMixin
9. Check BaseToolTab ← First abstract method found
10. Check object

When calling self.log_message():
    AdvancedCopyTab → ... → BaseToolTab.log_message()

When calling self.clean_file_path():
    AdvancedCopyTab → ... → FileInputMixin.clean_file_path()
```

#### Mixin Best Practices

✅ **DO:**
- Keep mixins focused on one responsibility
- Make mixins independent (avoid dependencies on other mixins)
- Document mixin requirements clearly
- Use descriptive method names to avoid conflicts

❌ **DON'T:**
- Create deep mixin hierarchies
- Make mixins depend on each other
- Override methods without calling super()
- Use generic method names that might conflict

---

## 3. Application Structure

### 3.1 Main Window (main.py)

**Location:** `src/main.py`

The main application class orchestrates the entire tool suite.

#### Responsibilities

```
EnhancedPowerBIReportToolsApp
├── Initializes ToolManager
├── Creates main window
├── Discovers and registers tools
├── Creates tabbed notebook
├── Manages tab state and sizing
├── Provides shared utilities (help, about, website)
└── Handles application lifecycle
```

#### Key Components

```python
class EnhancedPowerBIReportToolsApp(EnhancedBaseExternalTool):
    def __init__(self):
        # Get tool manager singleton
        self.tool_manager = get_tool_manager()
        
        # Application state
        self.notebook = None               # ttk.Notebook widget
        self.tool_tabs = {}               # tool_id -> tab instance
        self.tab_states = {}              # tool_id -> {custom_size, is_modified}
        self.current_tab_id = None        # Current active tab
        
        # Initialize tools
        self._initialize_tools()
    
    def _initialize_tools(self):
        """Discover and register all tools"""
        registered_count = self.tool_manager.discover_and_register_tools("tools")
    
    def create_ui(self) -> tk.Tk:
        """Create main tabbed interface"""
        root = self.create_secure_ui_base()
        self._setup_main_interface(root)
        return root
    
    def _setup_main_interface(self, root):
        """Setup header, notebook, and tabs"""
        self._setup_header(main_frame)
        self._setup_notebook(main_frame)
        self._setup_tabs_with_tool_manager()
```

#### Window Sizing Strategy

The app uses **smart tab state tracking** to remember custom window sizes:

```python
# Each tab has a state:
tab_states[tool_id] = {
    'custom_size': (width, height),  # User-modified size
    'default_size': (width, height), # Tool's default size
    'is_modified': False             # Has user interacted?
}

# On tab change:
if tab_is_modified and has_custom_size:
    use_custom_size()
else:
    use_default_size()
```

**Default Sizes by Tool:**
- Report Merger: 1150×950
- Advanced Copy: 1175×820
- Layout Optimizer: 1130×850
- Report Cleanup: 1100×1095
- Column Width: 1200×1035

---

### 3.2 Content Frame (Integrated into Main Window)

Unlike some architectures with separate content managers, the AE Multi-Tool integrates tab management directly into the main window for simplicity.

#### Tab Management

```python
def _setup_notebook(self, main_frame):
    """Create and style the notebook"""
    self.notebook = ttk.Notebook(main_frame)
    
    # Style notebook tabs
    style.configure('TNotebook.Tab',
                   font=('Segoe UI', 11, 'bold'),
                   padding=(25, 12))
    
    # Bind tab change event
    self.notebook.bind('<<NotebookTabChanged>>', self._on_tab_changed)

def _setup_tabs_with_tool_manager(self):
    """Create tabs for all enabled tools"""
    self.tool_tabs = self.tool_manager.create_tool_tabs(
        self.notebook, 
        self  # Pass main_app reference
    )
    
    # Reorder tabs for better UX
    self._reorder_tabs_for_better_ux()
```

#### Tab Reordering

Tools are displayed in a specific order for optimal user experience:

```python
desired_order = [
    "report_merger",           # Most common use case first
    "pbip_layout_optimizer",   # Layout optimization second
    "report_cleanup",          # Cleanup third
    "column_width",           # Column width fourth
    "advanced_copy"            # Utility last
]
```

---

### 3.3 Header Section

The header provides consistent branding and actions across all tabs.

#### Header Layout

```
┌────────────────────────────────────────────────────────────┐
│  📊 ANALYTIC ENDEAVORS                    [❓] [ℹ️] [🌐]    │
│  Enhanced Power BI Report Tools — Professional suite...    │
│  Built by Reid Havens of Analytic Endeavors                │
└────────────────────────────────────────────────────────────┘
```

#### Header Components

```python
def _setup_header(self, main_frame):
    # Left: Branding
    ttk.Label(text="📊 ANALYTIC ENDEAVORS", style='Brand.TLabel')
    ttk.Label(text="Enhanced Power BI Report Tools", style='Title.TLabel')
    ttk.Label(text="Built by Reid Havens...", style='Subtitle.TLabel')
    
    # Right: Actions
    ttk.Button(text="❓ HELP", command=self.show_help_dialog)
    ttk.Button(text="ℹ️ ABOUT", command=self.show_about_dialog)
    ttk.Button(text="🌐 WEBSITE", command=self.open_company_website)
```

---

## 4. Tool Lifecycle

### 4.1 Registration

**Phase 1: Discovery**

```python
# ToolManager.discover_and_register_tools()
tools_package = importlib.import_module("tools")
for finder, name, ispkg in pkgutil.iter_modules([tools_package_path]):
    if ispkg:  # Each tool is a package
        tool_module = importlib.import_module(f"tools.{name}")
        tool_class = find_tool_class(tool_module)  # Ends with 'Tool'
        tool_instance = tool_class()
        register_tool(tool_instance)
```

**Phase 2: Validation**

```python
def register_tool(self, tool: BaseTool):
    # 1. Type check
    if not isinstance(tool, BaseTool):
        raise ToolRegistrationError()
    
    # 2. Duplicate check
    if tool.tool_id in self._tools:
        raise ToolRegistrationError()
    
    # 3. Dependency check
    if not tool.can_run():
        raise ToolRegistrationError()
    
    # 4. Register
    self._tools[tool.tool_id] = tool
```

**Expected Tool Structure:**

```
tools/
└── my_new_tool/
    ├── __init__.py
    ├── my_new_tool_tool.py    ← Tool class (ends with 'Tool')
    ├── my_new_tool_ui.py      ← UI tab class
    ├── my_new_tool_core.py    ← Business logic
    └── TECHNICAL_GUIDE.md     ← Documentation
```

---

### 4.2 Initialization

**Tool Instance Creation:**

```python
# 1. Tool instantiation
tool = AdvancedCopyTool()  # Calls __init__

# 2. Metadata setup
tool.tool_id = "advanced_copy"
tool.name = "Advanced Copy"
tool.version = "1.1.0"

# 3. Logger creation
tool.logger = logging.getLogger(f"tool.{tool.tool_id}")
```

**Dependency Checking:**

```python
def can_run(self) -> bool:
    """Check dependencies before running"""
    try:
        from tools.advanced_copy.logic.advanced_copy_core import AdvancedCopyEngine
        from tools.report_merger.merger_core import ValidationService
        return True
    except ImportError:
        return False
```

**Resource Setup:**

Tools should initialize resources lazily (when needed) rather than eagerly (at startup):

```python
# ❌ Bad: Eager initialization
def __init__(self):
    self.engine = HeavyEngine()  # Expensive!
    self.data = load_large_dataset()  # Memory intensive!

# ✅ Good: Lazy initialization
def __init__(self):
    self._engine = None
    self._data = None

@property
def engine(self):
    if self._engine is None:
        self._engine = HeavyEngine()
    return self._engine
```

---

### 4.3 UI Creation

**Tab Creation Flow:**

```python
# 1. ToolManager creates tabs
tool_tabs = tool_manager.create_tool_tabs(notebook, main_app)

# 2. For each enabled tool:
tab = tool.create_ui_tab(notebook, main_app)  # Tool creates tab instance

# 3. Tool tab setup
tab.setup_ui()  # Tab builds its UI

# 4. Add to notebook
notebook.add(tab.get_frame(), text=tool.get_tab_title())
```

**UI Component Initialization:**

```python
class MyToolTab(BaseToolTab):
    def setup_ui(self):
        # 1. Create file input
        self.file_input = self.create_file_input_section(
            parent=self.frame,
            title="📁 DATA SOURCE",
            file_types=[("PBIP Files", "*.pbip")],
            guide_text=["Step 1:", "• Select file", "• Click Browse"]
        )
        
        # 2. Create log section
        self.log = self.create_log_section(self.frame)
        
        # 3. Create action buttons
        self.buttons = self.create_action_buttons(
            self.frame,
            buttons=[
                {'text': '🚀 EXECUTE', 'command': self.execute},
                {'text': '🔄 RESET', 'command': self.reset_tab}
            ]
        )
        
        # 4. Create progress bar
        self.create_progress_bar(self.frame)
```

---

### 4.4 Activation/Deactivation

**Tab Switching:**

```python
def _on_tab_changed(self, event=None):
    """Handle tab changes with smart state tracking"""
    
    # 1. Save previous tab's size if modified
    if self.current_tab_id and was_modified:
        save_custom_size()
    
    # 2. Get new tab's tool ID
    tool_id = get_tool_id_from_tab(current_tab_id)
    
    # 3. Check if new tab has been modified
    is_modified = check_if_tab_modified(tool_id)
    
    # 4. Apply appropriate window size
    if is_modified and has_custom_size:
        apply_custom_size()
    else:
        apply_default_size()
    
    # 5. Update current tab tracking
    self.current_tab_id = current_tab_id
```

**State Preservation:**

Each tab maintains its own state:
- File paths selected
- Options/checkboxes
- Log content
- Progress state

**Resource Cleanup:**

```python
# Tools should clean up resources when deactivated
def on_deactivate(self):
    # Cancel ongoing operations
    if self.background_thread:
        self.background_thread.cancel()
    
    # Clear large data structures
    self.cached_data = None
    
    # Release file handles
    if self.file_handle:
        self.file_handle.close()
```

---

## 5. Cross-Cutting Concerns

### 5.1 Logging System

**Framework Logging:**

```python
# Core logging (ToolManager, BaseTool)
logger = logging.getLogger("ToolManager")
logger = logging.getLogger(f"tool.{tool_id}")

# Tool-specific logging
self.logger.info("Operation started")
self.logger.warning("Potential issue detected")
self.logger.error("Operation failed")
```

**UI Logging:**

```python
# BaseToolTab provides log_message()
self.log_message("✅ Operation complete")
self.log_message("⏳ Processing...")
self.log_message("❌ Error occurred")

# Emojis for visual feedback:
# ✅ Success
# ⏳ In Progress
# ❌ Error
# ⚠️ Warning
# 📊 Analysis/Data
# 🔍 Discovery
# 🚀 Action
```

**UI Log Display:**

```python
# Log widget setup
log_text = scrolledtext.ScrolledText(
    height=12,
    font=('Consolas', 9),  # Monospace font
    state=tk.DISABLED,      # Read-only
    wrap=tk.NONE           # No word wrap
)

# Log export feature
def _export_log(self):
    content = self.log_text.get(1.0, tk.END)
    save_to_file(content)
```

---

### 5.2 Validation Framework

**Input Validation Patterns:**

```python
# FileInputMixin: Path cleaning
def clean_file_path(self, path: str) -> str:
    # Remove quotes
    # Normalize separators
    # Return cleaned path

# ValidationMixin: Existence checks
def validate_file_exists(self, path: str):
    if not Path(path).exists():
        raise ValueError("File not found")

def validate_pbip_file(self, path: str):
    if not path.endswith('.pbip'):
        raise ValueError("Not a PBIP file")
    
    report_dir = Path(path).parent / f"{Path(path).stem}.Report"
    if not report_dir.exists():
        raise ValueError("Missing .Report directory")
```

**Error Reporting:**

```python
# Background operation with error handling
def execute_operation(self):
    try:
        # Validate inputs
        self.validate_file_exists(self.file_path.get())
        self.validate_pbip_file(self.file_path.get())
        
        # Perform operation
        result = self.engine.process()
        
        # Success callback
        self.on_success(result)
        
    except ValueError as e:
        # User input error
        self.log_message(f"❌ Validation Error: {e}")
        self.show_error("Invalid Input", str(e))
        
    except Exception as e:
        # Unexpected error
        self.log_message(f"❌ Error: {e}")
        self.log_message(f"📋 Traceback: {traceback.format_exc()}")
        self.show_error("Operation Failed", str(e))
```

**User Feedback:**

```python
# Dialog boxes for user interaction
self.show_error(title, message)     # Error (red X)
self.show_warning(title, message)   # Warning (yellow triangle)
self.show_info(title, message)      # Info (blue i)
self.ask_yes_no(title, message)     # Question (returns bool)
```

---

### 5.3 Progress Tracking

**Progress System:**

```python
def update_progress(self, progress_percent: int, message: str, 
                   show: bool, persist: bool):
    """
    Universal progress update method
    
    Args:
        progress_percent: 0-100
        message: Status message (logged, not shown in UI)
        show: True to show, False to hide
        persist: True to keep visible after 100%
    """
    if show:
        self.progress_bar['value'] = progress_percent
        self.log_message(f"⏳ {progress_percent}% - {message}")
    else:
        if not self.progress_persist:
            self.progress_bar['value'] = 0
            self.progress_frame.grid_remove()
```

**Background Task Handling:**

```python
def run_in_background(self, target_func, success_callback, 
                     error_callback, progress_steps):
    """Execute function in background thread"""
    
    def thread_target():
        try:
            # Show progress
            for step_msg, step_pct in progress_steps:
                self.update_progress(step_pct, step_msg, True, False)
            
            # Execute operation
            result = target_func()
            
            # Schedule success callback on main thread
            self.frame.after(0, lambda: success_callback(result))
            
        except Exception as e:
            # Schedule error callback on main thread
            self.frame.after(0, lambda: error_callback(e))
        
        finally:
            # Hide progress
            self.frame.after(0, lambda: self.update_progress(0, "", False, False))
    
    thread = threading.Thread(target=thread_target, daemon=True)
    thread.start()
```

**Status Updates:**

```python
# Progress during operation
self.update_progress(0, "Starting...", show=True)
self.update_progress(25, "Loading files...", show=True)
self.update_progress(50, "Processing data...", show=True)
self.update_progress(75, "Saving results...", show=True)
self.update_progress(100, "Complete!", show=True)
self.update_progress(0, "", show=False)  # Hide
```

---

### 5.4 Theming System

**Color Constants:**

```python
# core/constants.py
COLORS = {
    # Primary colors
    'primary': '#066c7c',      # Teal (AE brand)
    'secondary': '#0891a5',    # Light teal
    'accent': '#10b9d1',       # Bright teal
    
    # Status colors
    'success': '#059669',      # Green
    'warning': '#d97706',      # Orange
    'error': '#dc2626',        # Red
    'info': '#2563eb',         # Blue
    
    # Interface colors
    'background': '#f8fafc',   # Light gray
    'surface': '#ffffff',      # White
    'border': '#e2e8f0',       # Border gray
    'text_primary': '#1e293b', # Dark gray
    'text_secondary': '#64748b' # Medium gray
}
```

**Style Management:**

```python
def _setup_common_styling(self):
    """Setup professional styling"""
    style = ttk.Style()
    style.theme_use('clam')
    
    # Configure styles
    style.configure('Section.TLabelframe',
                   background=colors['background'],
                   borderwidth=1,
                   relief='solid')
    
    style.configure('Action.TButton',
                   background=colors['primary'],
                   foreground=colors['surface'],
                   font=('Segoe UI', 10, 'bold'),
                   padding=(20, 10))
```

**Brand Consistency:**

All tools use:
- **Segoe UI** font family
- **Teal (#066c7c)** as primary brand color
- **Consistent spacing** (padding="20" for frames)
- **Emojis** for visual hierarchy
- **Professional language** (no slang)

---

## 6. MCP Server Integration

### 6.1 Architecture

The AE Multi-Tool integrates with **Model Context Protocol (MCP)** servers to leverage external capabilities for documentation, file operations, and automation.

**MCP Server Types:**
1. **Microsoft Docs MCP** - Documentation search and code samples
2. **Filesystem MCP** - Advanced file operations
3. **Windows MCP** - Windows automation and UI interaction
4. **PBIP MCP** - PBIP-specific file operations

**Integration Pattern:**

```
Tool
  ↓
  ├─→ MCP Client (Claude Desktop)
  │     ↓
  │     ├─→ Microsoft Docs Server
  │     ├─→ Filesystem Server
  │     ├─→ Windows MCP Server
  │     └─→ PBIP Server
  │
  └─→ Direct File Operations (fallback)
```

---

### 6.2 Microsoft Docs MCP

**Purpose:** Search Microsoft documentation and retrieve code samples

**Usage Pattern:**
```python
# Search for documentation
results = microsoft_docs_search(query="Power BI PBIP format")

# Fetch code samples
code = microsoft_code_sample_search(
    query="read PBIP file JSON",
    language="python"
)

# Fetch full documentation
content = microsoft_docs_fetch(url="https://docs.microsoft.com/...")
```

**Tools Using Microsoft Docs MCP:**
- **Report Merger** - PBIP format documentation
- **Layout Optimizer** - Data model best practices
- **Column Width** - Visual formatting documentation

**Benefits:**
- Up-to-date official documentation
- Code sample validation
- Best practices guidance

---

### 6.3 Filesystem MCP

**Purpose:** Advanced file system operations with validation

**Usage Pattern:**
```python
# Read files with encoding detection
content = read_file(path="report.pbip")

# Read multiple files efficiently
contents = read_multiple_files(paths=[...])

# Write files with atomic operations
write_file(path="output.json", content=data)

# Edit files with line-based replacements
edit_file(path="config.json", edits=[...])

# Search file system
results = search_files(path=".", pattern="*.pbip")
```

**Tools Using Filesystem MCP:**
- **All Tools** - File reading and writing operations
- **Advanced Copy** - Complex file structure manipulation
- **Layout Optimizer** - Multi-file coordination

**Benefits:**
- Atomic file operations (all-or-nothing)
- Better error handling
- Encoding auto-detection
- Path validation

---

### 6.4 Windows MCP

**Purpose:** Windows automation and UI interaction

**Usage Pattern:**
```python
# Launch applications
launch_tool(name="notepad")

# Execute PowerShell commands
result = powershell_tool(command="Get-Process")

# Capture desktop state
state = state_tool(use_vision=True)

# Clipboard operations
clipboard_tool(mode="copy", text="data")
clipboard_tool(mode="paste")

# UI automation
click_tool(loc=[x, y], button="left")
type_tool(loc=[x, y], text="input")
scroll_tool(direction="down", wheel_times=3)
```

**Tools Using Windows MCP:**
- **Report Cleanup** - Potential for automated Power BI interaction
- **Future Tools** - UI automation possibilities

**Benefits:**
- Desktop automation capabilities
- Screenshot and visual analysis
- Clipboard integration
- PowerShell scripting

---

### 6.5 PBIP MCP

**Purpose:** PBIP-specific file operations and validation

**Usage Pattern:**
```python
# Validate PBIP structure
validate_pbip(path="report.pbip")

# Read PBIP metadata
metadata = read_pbip_metadata(path="report.pbip")

# Extract PBIP components
pages = extract_pages(path="report.pbip")
bookmarks = extract_bookmarks(path="report.pbip")

# Modify PBIP files
update_pbip(path="report.pbip", changes={...})
```

**Tools Using PBIP MCP:**
- **All Tools** - PBIP file validation
- **Report Merger** - PBIP structure merging
- **Advanced Copy** - Page and bookmark manipulation

**Benefits:**
- PBIP-specific validation
- Structure awareness
- Schema validation
- Atomic updates

---

### 6.6 Tool-Specific MCP Usage Patterns

| Tool | Microsoft Docs | Filesystem | Windows | PBIP |
|------|---------------|-----------|---------|------|
| **Report Merger** | Documentation | ✅ Read/Write | ❌ | ✅ Validation |
| **Advanced Copy** | Documentation | ✅ Complex Ops | ❌ | ✅ Manipulation |
| **Layout Optimizer** | Best Practices | ✅ Read/Write | ❌ | ✅ Structure |
| **Report Cleanup** | Documentation | ✅ Read/Write | Potential | ✅ Analysis |
| **Column Width** | Visual Docs | ✅ Read/Write | ❌ | ✅ Visual Data |

**Legend:**
- ✅ Active usage
- ❌ Not used
- Potential - Future enhancement

---

## 7. Adding New Tools

### 7.1 Step-by-Step Guide

**Step 1: Create Tool Directory**

```bash
tools/
└── my_new_tool/
    ├── __init__.py
    ├── my_new_tool_tool.py      # BaseTool implementation
    ├── my_new_tool_ui.py         # BaseToolTab implementation
    ├── my_new_tool_core.py       # Business logic
    └── TECHNICAL_GUIDE.md        # Documentation
```

**Step 2: Implement BaseTool**

```python
# tools/my_new_tool/my_new_tool_tool.py
from core.tool_manager import BaseTool
from tools.my_new_tool.my_new_tool_ui import MyNewToolTab

class MyNewTool(BaseTool):
    def __init__(self):
        super().__init__(
            tool_id="my_new_tool",          # Unique ID (lowercase_underscore)
            name="My New Tool",             # Display name
            description="What this tool does",
            version="1.0.0"
        )
    
    def create_ui_tab(self, parent, main_app) -> 'BaseToolTab':
        return MyNewToolTab(parent, main_app)
    
    def get_tab_title(self) -> str:
        return "🆕 My New Tool"  # With emoji
    
    def get_help_content(self) -> Dict[str, Any]:
        return {
            "title": "My New Tool - Help",
            "sections": [
                {
                    "title": "🚀 Quick Start",
                    "items": ["Step 1...", "Step 2..."]
                }
            ]
        }
    
    def can_run(self) -> bool:
        """Check dependencies"""
        try:
            from tools.my_new_tool.my_new_tool_core import MyEngine
            return True
        except ImportError:
            return False
```

**Step 3: Implement BaseToolTab**

```python
# tools/my_new_tool/my_new_tool_ui.py
from core.ui_base import BaseToolTab, FileInputMixin, ValidationMixin

class MyNewToolTab(BaseToolTab, FileInputMixin, ValidationMixin):
    def __init__(self, parent, main_app):
        super().__init__(parent, main_app, "my_new_tool", "My New Tool")
        
        # UI Variables
        self.input_path = tk.StringVar()
        
        # Core components
        self.engine = None
        
        # Setup
        self.setup_ui()
    
    def setup_ui(self) -> None:
        """Setup UI components"""
        # Create file input
        self.file_input = self.create_file_input_section(
            parent=self.frame,
            title="📁 INPUT FILE",
            file_types=[("PBIP Files", "*.pbip")],
            guide_text=["Select a PBIP file", "Click Browse"]
        )
        self.file_input['frame'].grid(row=0, column=0, sticky=(tk.W, tk.E))
        
        # Create log section
        log_components = self.create_log_section(self.frame)
        log_components['frame'].grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Create action buttons
        button_frame = ttk.Frame(self.frame)
        button_frame.grid(row=2, column=0, pady=(20, 0))
        
        ttk.Button(button_frame, text="🚀 EXECUTE",
                  command=self.execute, style='Action.TButton').pack(side=tk.LEFT)
        ttk.Button(button_frame, text="🔄 RESET",
                  command=self.reset_tab, style='Secondary.TButton').pack(side=tk.LEFT)
        
        # Create progress bar
        self.create_progress_bar(self.frame)
        
        # Welcome message
        self.log_message("🎉 Welcome to My New Tool!")
    
    def _position_progress_frame(self):
        """Position progress frame for this layout"""
        if self.progress_frame:
            self.progress_frame.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=(10, 5))
    
    def execute(self):
        """Execute the tool's main operation"""
        try:
            # Validate inputs
            file_path = self.clean_file_path(self.input_path.get())
            self.validate_file_exists(file_path, "Input file")
            
            # Run in background
            self.run_in_background(
                target_func=lambda: self._execute_operation(file_path),
                success_callback=self._on_success,
                error_callback=self._on_error
            )
            
        except Exception as e:
            self.log_message(f"❌ Error: {e}")
            self.show_error("Execution Error", str(e))
    
    def _execute_operation(self, file_path: str):
        """Background operation"""
        self.update_progress(0, "Starting...", show=True)
        
        # Initialize engine
        if self.engine is None:
            from tools.my_new_tool.my_new_tool_core import MyEngine
            self.engine = MyEngine(logger_callback=self.log_message)
        
        # Process
        self.update_progress(50, "Processing...", show=True)
        result = self.engine.process(file_path)
        
        self.update_progress(100, "Complete!", show=True)
        return result
    
    def _on_success(self, result):
        """Success callback"""
        self.log_message("✅ Operation completed successfully!")
        self.show_info("Success", "Operation completed!")
    
    def _on_error(self, error: Exception):
        """Error callback"""
        self.log_message(f"❌ Error: {error}")
        self.show_error("Operation Failed", str(error))
    
    def reset_tab(self) -> None:
        """Reset to initial state"""
        self.input_path.set("")
        self._clear_log(self.log_text)
        self.log_message("🎉 Welcome to My New Tool!")
    
    def show_help_dialog(self) -> None:
        """Show help dialog"""
        help_content = """
My New Tool - Help

🚀 Quick Start:
1. Select a PBIP file
2. Click EXECUTE
3. Review results in the log

📋 What This Tool Does:
- Feature 1
- Feature 2
- Feature 3

⚠️ Important Notes:
- Always backup your files
- Only works with PBIP files
- NOT officially supported by Microsoft
        """
        self.show_scrollable_info("My New Tool - Help", help_content)
```

**Step 4: Implement Business Logic**

```python
# tools/my_new_tool/my_new_tool_core.py
from pathlib import Path
import json

class MyEngine:
    def __init__(self, logger_callback=None):
        self.logger_callback = logger_callback or print
    
    def process(self, file_path: str):
        """Process the file"""
        self.log("Starting processing...")
        
        # Read file
        with open(file_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        # Process data
        result = self._process_data(data)
        
        # Save result
        output_path = Path(file_path).parent / "output.json"
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(result, f, indent=2)
        
        self.log("Processing complete!")
        return result
    
    def _process_data(self, data):
        """Process data logic"""
        # Your logic here
        return data
    
    def log(self, message: str):
        """Log message"""
        if self.logger_callback:
            self.logger_callback(message)
```

**Step 5: Register Tool (Automatic)**

The tool will be discovered and registered automatically on application startup. No manual registration needed!

**Step 6: Test Integration**

```python
# Test the tool
if __name__ == "__main__":
    from tools.my_new_tool.my_new_tool_tool import MyNewTool
    
    tool = MyNewTool()
    print(f"Tool ID: {tool.tool_id}")
    print(f"Can run: {tool.can_run()}")
    print(f"Tab title: {tool.get_tab_title()}")
```

---

### 7.2 Code Templates

#### Minimal Tool Template

```python
# tools/minimal_tool/minimal_tool_tool.py
from core.tool_manager import BaseTool
from core.ui_base import BaseToolTab
import tkinter as tk
from tkinter import ttk

class MinimalTool(BaseTool):
    def __init__(self):
        super().__init__(
            tool_id="minimal_tool",
            name="Minimal Tool",
            description="Minimal tool example",
            version="1.0.0"
        )
    
    def create_ui_tab(self, parent, main_app):
        return MinimalToolTab(parent, main_app)
    
    def get_tab_title(self):
        return "🔧 Minimal Tool"
    
    def get_help_content(self):
        return {"title": "Minimal Tool", "sections": []}

class MinimalToolTab(BaseToolTab):
    def __init__(self, parent, main_app):
        super().__init__(parent, main_app, "minimal_tool", "Minimal Tool")
        self.setup_ui()
    
    def setup_ui(self):
        ttk.Label(self.frame, text="Hello from Minimal Tool!").pack()
    
    def reset_tab(self):
        pass
    
    def show_help_dialog(self):
        self.show_info("Help", "This is a minimal tool")
```

#### Full-Featured Tool Template

See **Step 2** and **Step 3** above for complete examples with:
- File input handling
- Log section
- Progress tracking
- Background operations
- Error handling
- Help dialog

---

### 7.3 Best Practices

#### When to Use Mixins

**Use FileInputMixin if:**
- ✅ Tool needs file/folder selection
- ✅ Tool needs path cleaning (quotes, separators)
- ✅ Tool has browse buttons

**Use ValidationMixin if:**
- ✅ Tool validates file existence
- ✅ Tool validates PBIP structure
- ✅ Tool validates user inputs

**Create Custom Mixins if:**
- ✅ Functionality is reusable across tools
- ✅ Logic is self-contained
- ✅ It follows single responsibility principle

**Don't Use Mixins if:**
- ❌ Logic is specific to one tool
- ❌ It creates dependencies between mixins
- ❌ It complicates the inheritance chain

#### Error Handling Patterns

```python
# Pattern 1: Validation Errors (User Input)
try:
    self.validate_file_exists(path)
except ValueError as e:
    self.log_message(f"❌ Validation Error: {e}")
    self.show_error("Invalid Input", str(e))
    return

# Pattern 2: Processing Errors (Background)
def run_in_background(
    target_func=self._process,
    success_callback=self._on_success,
    error_callback=self._on_error  # Handles exceptions
)

# Pattern 3: Unexpected Errors (Catchall)
except Exception as e:
    self.log_message(f"❌ Unexpected Error: {e}")
    self.log_message(f"📋 Traceback: {traceback.format_exc()}")
    self.show_error("Operation Failed", "An unexpected error occurred")
```

#### UI Layout Guidelines

**Layout Principles:**
1. **Top-to-Bottom Flow** - User reads top to bottom
2. **Left-to-Right Sections** - Guide on left, input on right
3. **Fixed Action Area** - Buttons in consistent location
4. **Expandable Log** - Log section gets extra vertical space

**Grid Layout Example:**
```python
# Configure grid weights
self.frame.columnconfigure(0, weight=1)  # Stretch horizontally
self.frame.rowconfigure(2, weight=1)     # Log section gets extra vertical space

# Row 0: File input section
file_section.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 15))

# Row 1: Options section (if needed)
options_section.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 15))

# Row 2: Log section (expandable)
log_section.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 15))

# Row 3: Progress bar
progress_frame.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=(10, 5))

# Row 4: Action buttons
button_frame.grid(row=4, column=0, pady=(20, 0))
```

**Responsive Progress Bar Positioning:**

Override `_position_progress_frame()` to control progress bar placement:

```python
def _position_progress_frame(self):
    """Position progress frame for this specific layout"""
    if self.progress_frame:
        # Position between log (row 2) and buttons (row 4)
        self.progress_frame.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=(10, 5))
```

---

## 8. File Organization

### 8.1 Directory Structure

```
ae-multi-tool/
├── src/
│   ├── main.py                    # Application entry point
│   ├── run_ae_multi_tool.bat      # Windows launcher
│   │
│   ├── assets/                    # Application resources
│   │   ├── favicon.ico
│   │   └── website_icon.png
│   │
│   ├── core/                      # Core framework
│   │   ├── __init__.py
│   │   ├── tool_manager.py        # BaseTool, ToolManager
│   │   ├── ui_base.py            # BaseToolTab, mixins
│   │   ├── constants.py          # AppConstants, colors, themes
│   │   ├── base_tool.py          # Legacy base tool
│   │   └── enhanced_base_tool.py # Enhanced tool wrapper
│   │
│   └── tools/                     # Tool implementations
│       ├── __init__.py
│       │
│       ├── report_merger/         # Tool 1: Report Merger
│       │   ├── __init__.py
│       │   ├── merger_tool.py     # BaseTool implementation
│       │   ├── merger_ui.py       # BaseToolTab implementation
│       │   ├── merger_core.py     # Business logic
│       │   └── TECHNICAL_GUIDE.md # Documentation
│       │
│       ├── advanced_copy/         # Tool 2: Advanced Copy
│       │   ├── __init__.py
│       │   ├── advanced_copy_tool.py
│       │   ├── ui/                # UI components (modular)
│       │   │   ├── advanced_copy_tab.py
│       │   │   ├── ui_data_source.py
│       │   │   ├── ui_page_selection.py
│       │   │   ├── ui_bookmark_selection.py
│       │   │   ├── ui_event_handlers.py
│       │   │   └── ui_helpers.py
│       │   ├── logic/             # Business logic (modular)
│       │   │   ├── advanced_copy_core.py
│       │   │   ├── advanced_copy_operations.py
│       │   │   ├── advanced_copy_bookmark_analyzer.py
│       │   │   └── advanced_copy_visual_actions.py
│       │   └── TECHNICAL_GUIDE.md
│       │
│       ├── pbip_layout_optimizer/ # Tool 3: Layout Optimizer
│       │   ├── __init__.py
│       │   ├── tool.py            # BaseTool implementation
│       │   ├── layout_ui.py       # BaseToolTab implementation
│       │   ├── enhanced_layout_core.py
│       │   ├── base_layout_engine.py
│       │   ├── analyzers/         # Analysis modules
│       │   │   ├── relationship_analyzer.py
│       │   │   └── table_categorizer.py
│       │   ├── engines/           # Layout engines
│       │   │   └── middle_out_layout_engine.py
│       │   ├── positioning/       # Positioning algorithms
│       │   │   ├── position_calculator.py
│       │   │   ├── dimension_optimizer.py
│       │   │   └── chain_aware_position_generator.py
│       │   └── TECHNICAL_GUIDE.md
│       │
│       ├── report_cleanup/        # Tool 4: Report Cleanup
│       │   ├── __init__.py
│       │   ├── tool.py
│       │   ├── cleanup_ui.py
│       │   ├── cleanup_engine.py
│       │   ├── report_analyzer.py
│       │   ├── shared_types.py
│       │   └── TECHNICAL_GUIDE.md
│       │
│       └── column_width/          # Tool 5: Column Width
│           ├── __init__.py
│           ├── column_width_tool.py
│           ├── column_width_ui.py
│           ├── column_width_core.py
│           └── TECHNICAL_GUIDE.md
│
├── ARCHITECTURE.md                # This file (parent-level guide)
└── README.md                      # User-facing documentation
```

---

### 8.2 File Naming Conventions

**Tool Files:**
- `{tool_name}_tool.py` - BaseTool implementation (e.g., `merger_tool.py`)
- `{tool_name}_ui.py` - BaseToolTab implementation (e.g., `merger_ui.py`)
- `{tool_name}_core.py` - Business logic (e.g., `merger_core.py`)

**UI Component Files:**
- `ui_{component}.py` - UI-specific mixins/components
- Example: `ui_data_source.py`, `ui_page_selection.py`

**Logic Files:**
- `{tool_name}_{module}.py` - Logic modules
- Example: `advanced_copy_operations.py`, `cleanup_engine.py`

**Class Names:**
- Tool class: `{ToolName}Tool` (e.g., `AdvancedCopyTool`)
- UI class: `{ToolName}Tab` (e.g., `AdvancedCopyTab`)
- Logic class: `{ToolName}Engine` (e.g., `MergerEngine`)
- Mixin class: `{Function}Mixin` (e.g., `FileInputMixin`)

---

### 8.3 Import Patterns

**Core Framework Imports:**

```python
# Tool implementations import from core
from core.tool_manager import BaseTool
from core.ui_base import BaseToolTab, FileInputMixin, ValidationMixin
from core.constants import AppConstants
```

**Tool-Specific Imports:**

```python
# UI imports from logic
from tools.advanced_copy.logic.advanced_copy_core import AdvancedCopyEngine

# Logic imports from logic (internal)
from tools.advanced_copy.logic.advanced_copy_operations import OperationsManager
```

**Circular Import Avoidance:**

```python
# ❌ Bad: Circular dependency
# ui.py imports core.py, core.py imports ui.py

# ✅ Good: One-directional dependency
# ui.py imports core.py only
# core.py has no knowledge of ui.py
```

**Lazy Imports:**

```python
# Lazy import for optional dependencies
def can_run(self) -> bool:
    try:
        from tools.my_tool.my_tool_core import MyEngine
        return True
    except ImportError:
        return False
```

---

### 8.4 Dependencies

**Internal Dependencies:**

```
tool_manager.py
  └── (no dependencies - pure Python)

ui_base.py
  ├── tkinter (standard library)
  └── constants.py

{tool}_tool.py
  ├── core.tool_manager
  └── tools.{tool}.{tool}_ui

{tool}_ui.py
  ├── core.ui_base
  ├── core.constants
  └── tools.{tool}.{tool}_core

{tool}_core.py
  ├── pathlib (standard library)
  ├── json (standard library)
  └── (minimal external dependencies)
```

**External Packages:**

```python
# Standard Library (Always Available)
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, scrolledtext
from pathlib import Path
import json
import threading
import logging
import importlib
import pkgutil

# No External Dependencies Required!
# Pure Python + Standard Library = Maximum Compatibility
```

**MCP Dependencies:**

When MCP servers are available:
- `microsoft_docs_search()` - Microsoft Docs MCP
- `read_file()`, `write_file()` - Filesystem MCP
- `powershell_tool()` - Windows MCP
- (Future) PBIP-specific MCP functions

---

## 9. Design Patterns Used

### 9.1 Composition Over Inheritance

**Pattern:** Use mixins to compose functionality instead of deep inheritance hierarchies.

**Why:**
- ✅ More flexible - add/remove features easily
- ✅ Avoids the "fragile base class" problem
- ✅ Reusable across unrelated classes
- ✅ Explicit about what functionality each class has

**Example:**

```python
# ❌ Bad: Deep inheritance hierarchy
class BaseTool:
    pass

class FileBasedTool(BaseTool):
    def handle_files(self): pass

class ValidatingTool(FileBasedTool):
    def validate(self): pass

class AdvancedTool(ValidatingTool):
    # Too deep! Changes to BaseTool affect everything
    pass

# ✅ Good: Composition with mixins
class AdvancedTool(BaseTool, FileInputMixin, ValidationMixin):
    # Clear, explicit, composable
    pass
```

**Mixin Guidelines:**

1. **Single Responsibility** - Each mixin does one thing
2. **No Mixin Dependencies** - Mixins don't depend on each other
3. **Explicit Naming** - Names describe what they do
4. **Documented Requirements** - Clear docs on what mixin needs

---

### 9.2 Factory Pattern

**Pattern:** ToolManager acts as a factory for creating tool instances and UI tabs.

**Why:**
- ✅ Centralized tool creation
- ✅ Consistent initialization
- ✅ Easy to add new tools
- ✅ Dependency injection

**Example:**

```python
# Factory: ToolManager
class ToolManager:
    def register_tool(self, tool: BaseTool):
        """Factory method for registering tools"""
        self._tools[tool.tool_id] = tool
    
    def create_tool_tabs(self, notebook, main_app):
        """Factory method for creating UI tabs"""
        for tool in self.get_enabled_tools():
            tab = tool.create_ui_tab(notebook, main_app)  # Delegation
            self._tool_tabs[tool.tool_id] = tab
        return self._tool_tabs

# Usage
tool_manager.register_tool(AdvancedCopyTool())
tabs = tool_manager.create_tool_tabs(notebook, app)
```

**Benefits:**
- Tools don't need to know about the registration process
- UI creation is delegated to tools themselves
- Easy to swap tool implementations

---

### 9.3 Observer Pattern

**Pattern:** Event handling and state updates follow observer pattern.

**Why:**
- ✅ Decouples UI from business logic
- ✅ Multiple observers can react to same event
- ✅ Easy to add new event handlers

**Example:**

```python
# Observable: StringVar (Tkinter built-in observer)
path_var = tk.StringVar()

# Observer: trace callback
def on_path_change(*args):
    cleaned = self.clean_file_path(path_var.get())
    path_var.set(cleaned)

path_var.trace('w', on_path_change)  # Register observer

# Observable: Notebook tab changes
notebook.bind('<<NotebookTabChanged>>', self._on_tab_changed)

# Observer: Tab change handler
def _on_tab_changed(self, event):
    # React to tab change
    self._adjust_window_size()
```

**Event Types:**
- **UI Events** - Button clicks, tab changes, text input
- **Progress Events** - Background operation updates
- **State Events** - Tool activation/deactivation
- **File Events** - File selection, file modifications

---

### 9.4 Template Method Pattern

**Pattern:** BaseToolTab defines the template for tool UI creation.

**Why:**
- ✅ Enforces consistent structure
- ✅ Provides reusable components
- ✅ Subclasses fill in specifics
- ✅ Common patterns extracted to base class

**Example:**

```python
# Template: BaseToolTab
class BaseToolTab(ABC):
    """Template for tool UI"""
    
    def __init__(self, parent, main_app, tool_id, tool_name):
        # Common initialization
        self.frame = ttk.Frame(parent, padding="20")
        self._setup_common_styling()
    
    @abstractmethod
    def setup_ui(self) -> None:
        """Subclasses must implement"""
        pass
    
    def create_file_input_section(self, ...):
        """Template provides this"""
        pass
    
    def create_log_section(self, ...):
        """Template provides this"""
        pass

# Concrete Implementation: AdvancedCopyTab
class AdvancedCopyTab(BaseToolTab):
    def setup_ui(self):
        # Use template methods
        self.file_input = self.create_file_input_section(...)
        self.log = self.create_log_section(...)
```

**Template Methods Provided:**
- `create_file_input_section()` - File input UI
- `create_log_section()` - Log UI
- `create_progress_bar()` - Progress bar UI
- `create_action_buttons()` - Button set UI
- `run_in_background()` - Background execution
- `update_progress()` - Progress updates

---

### 9.5 Singleton Pattern

**Pattern:** ToolManager is a singleton (single global instance).

**Why:**
- ✅ Only one tool registry needed
- ✅ Consistent state across application
- ✅ Easy access from anywhere

**Example:**

```python
# Singleton implementation
_tool_manager: Optional[ToolManager] = None

def get_tool_manager() -> ToolManager:
    """Get the global tool manager instance"""
    global _tool_manager
    if _tool_manager is None:
        _tool_manager = ToolManager()
    return _tool_manager

# Usage
tool_manager = get_tool_manager()  # Always returns same instance
```

**Benefits:**
- No duplicate registrations
- Consistent tool state
- Easy testing (can reset singleton)

---

## 10. Key Technical Decisions

### 10.1 Why These Patterns?

#### BaseTool Abstraction

**Decision:** Use abstract base class with strict interface

**Rationale:**
- ✅ **Enforces Contract** - All tools must implement required methods
- ✅ **Plugin Architecture** - Easy to add new tools
- ✅ **Type Safety** - Tools can be type-checked
- ✅ **Discoverability** - Clear interface for developers

**Alternative Considered:** Protocol-based (duck typing)
- ❌ Less explicit about requirements
- ❌ Harder to validate at registration time

**Trade-off:** Slightly more boilerplate, but much clearer contracts

---

#### Mixin Composition

**Decision:** Use mixins instead of deep inheritance

**Rationale:**
- ✅ **Flexibility** - Pick and choose features
- ✅ **Reusability** - Share code across unrelated tools
- ✅ **Maintainability** - Changes to mixin don't cascade
- ✅ **Clarity** - Explicit about what each tool does

**Alternative Considered:** Deep inheritance hierarchy
- ❌ Fragile base class problem
- ❌ Difficult to add features selectively
- ❌ Tight coupling

**Trade-off:** More classes to understand, but much more maintainable

---

#### MCP Integration

**Decision:** Integrate MCP servers for external capabilities

**Rationale:**
- ✅ **Official Documentation** - Direct access to Microsoft docs
- ✅ **File Safety** - Atomic operations via Filesystem MCP
- ✅ **Automation** - Windows automation via Windows MCP
- ✅ **Future-Proof** - Easy to add new MCP capabilities

**Alternative Considered:** Direct API calls
- ❌ More complex error handling
- ❌ Less standardized
- ❌ Harder to add new capabilities

**Trade-off:** Dependency on MCP infrastructure, but huge capability gains

---

### 10.2 Trade-offs

#### Framework Complexity vs Flexibility

**Trade-off:** More framework code upfront for easier tool development later

**Complexity Added:**
- BaseTool abstract class
- BaseToolTab template
- ToolManager singleton
- Mixin system

**Flexibility Gained:**
- Add new tools without modifying core
- Reuse UI components across tools
- Consistent user experience
- Easy testing and maintenance

**Verdict:** Worth it - initial complexity pays off long-term

---

#### Performance vs Features

**Trade-off:** Background threading for responsiveness vs complexity

**Features:**
- Non-blocking UI during operations
- Real-time progress updates
- Cancel operations (future)
- Parallel processing (future)

**Complexity:**
- Thread safety concerns
- Callback orchestration
- State management

**Verdict:** Worth it - responsive UI is critical for professional tools

---

#### Maintainability vs Simplicity

**Trade-off:** More files and structure vs everything in one file

**Structure:**
```
# Complex but maintainable
advanced_copy/
  ├── ui/
  │   ├── ui_data_source.py
  │   ├── ui_page_selection.py
  │   └── ui_bookmark_selection.py
  └── logic/
      ├── advanced_copy_core.py
      └── advanced_copy_operations.py

# Simple but unmaintainable
advanced_copy/
  └── advanced_copy.py  # 2000+ lines!
```

**Verdict:** Worth it - organized code is easier to maintain

---

## 11. Future Architecture Enhancements

### 11.1 Plugin System

**Enhancement:** External plugin support

**Vision:**
```
plugins/
├── community_tool_1/
│   ├── plugin.json       # Plugin metadata
│   └── tool.py
└── community_tool_2/
    ├── plugin.json
    └── tool.py
```

**Benefits:**
- Community contributions
- No core code changes
- Dynamic loading
- Version management

**Challenges:**
- Security (untrusted code)
- API stability
- Plugin dependencies
- Update mechanism

---

### 11.2 Dynamic Tool Loading

**Enhancement:** Load/unload tools at runtime

**Vision:**
```python
# Enable/disable tools without restart
tool_manager.enable_tool("advanced_copy")
tool_manager.disable_tool("advanced_copy")
tool_manager.reload_tool("advanced_copy")

# Download new tools
tool_manager.install_tool_from_url("https://...")
```

**Benefits:**
- No restart needed
- Faster development
- Resource management
- User customization

**Challenges:**
- State management
- Resource cleanup
- UI updates
- Error handling

---

### 11.3 Enhanced MCP Integration

**Enhancement:** Deeper MCP server usage

**Vision:**
- **AI-Assisted Operations** - Claude suggests optimizations
- **Documentation Integration** - Context-aware help
- **Automated Testing** - AI-generated test cases
- **Code Generation** - Template-based tool creation

**Example:**
```python
# AI-assisted optimization
suggestions = claude_suggest_optimizations(report_path)
apply_suggestions(suggestions)

# Context-aware help
help_content = generate_help_for_current_state(ui_state)
```

---

### 11.4 Improved State Management

**Enhancement:** Centralized state management

**Vision:**
```python
# Redux-style state management
class AppState:
    tools: Dict[str, ToolState]
    ui: UIState
    preferences: UserPreferences

# Reducers
def tool_state_reducer(state, action):
    if action.type == "TOOL_ENABLED":
        return enable_tool(state, action.tool_id)
    # ...

# Store
store = Store(initial_state, reducer)
store.subscribe(lambda: update_ui(store.state))
```

**Benefits:**
- Predictable state changes
- Time-travel debugging
- Undo/redo functionality
- State persistence

---

## 12. Appendix

### 12.1 Tool Comparison Matrix

| Feature | Report Merger | Advanced Copy | Layout Optimizer | Report Cleanup | Column Width |
|---------|--------------|---------------|------------------|----------------|--------------|
| **Uses MCP** | ✅ | ✅ | ✅ | ✅ | ✅ |
| **Uses Mixins** | ✅ | ✅✅✅ | ✅ | ✅ | ✅ |
| **Background Tasks** | ✅ | ✅ | ✅ | ✅ | ✅ |
| **File I/O** | ✅✅ | ✅✅ | ✅✅ | ✅✅ | ✅✅ |
| **Progress Tracking** | ✅ | ✅ | ✅ | ✅ | ✅ |
| **Validation** | ✅✅ | ✅✅ | ✅ | ✅ | ✅ |
| **Help Dialog** | ✅ | ✅ | ✅ | ✅ | ✅ |
| **Complex UI** | Medium | High | Medium | High | Medium |
| **Code Modules** | 3 | 11 | 9 | 5 | 4 |

**Legend:**
- ✅ Basic usage
- ✅✅ Heavy usage
- ✅✅✅ Very heavy usage

---

### 12.2 Framework Class Hierarchy

```
┌─────────────────────────────────────────┐
│         CORE FRAMEWORK CLASSES          │
└─────────────────────────────────────────┘

BaseTool (ABC)
├── AdvancedCopyTool
├── ColumnWidthTool
├── LayoutOptimizerTool
├── ReportCleanupTool
└── ReportMergerTool

BaseToolTab (ABC)
├── (mixin composable)
└── Used by all tool tabs

FileInputMixin
├── Used by: All tools
└── Provides: Path cleaning, file selection

ValidationMixin
├── Used by: All tools
└── Provides: File/PBIP validation

┌─────────────────────────────────────────┐
│         TOOL-SPECIFIC MIXINS            │
└─────────────────────────────────────────┘

AdvancedCopyTab Composition:
    BaseToolTab
    + FileInputMixin
    + ValidationMixin
    + DataSourceMixin          (tool-specific)
    + EventHandlersMixin       (tool-specific)
    + HelpersMixin            (tool-specific)
    + PageSelectionMixin      (tool-specific)
    + BookmarkSelectionMixin  (tool-specific)

Other Tool Tabs:
    BaseToolTab
    + FileInputMixin
    + ValidationMixin
    + (tool-specific mixins as needed)
```

---

### 12.3 Method Resolution Order (MRO) Example

**For AdvancedCopyTab:**

```python
# Define class with multiple inheritance
class AdvancedCopyTab(
    BaseToolTab,              # 1
    FileInputMixin,           # 2
    ValidationMixin,          # 3
    DataSourceMixin,          # 4
    EventHandlersMixin,       # 5
    HelpersMixin,             # 6
    PageSelectionMixin,       # 7
    BookmarkSelectionMixin    # 8
):
    pass

# MRO (left to right, depth-first)
AdvancedCopyTab.__mro__ = (
    AdvancedCopyTab,          # 0: Self
    BaseToolTab,              # 1
    FileInputMixin,           # 2
    ValidationMixin,          # 3
    DataSourceMixin,          # 4
    EventHandlersMixin,       # 5
    HelpersMixin,             # 6
    PageSelectionMixin,       # 7
    BookmarkSelectionMixin,   # 8
    ABC,                      # 9: Abstract base
    object                    # 10: Root
)

# Method lookup:
self.log_message()
    → AdvancedCopyTab?        No
    → BaseToolTab?           Yes! ✓ Found

self.clean_file_path()
    → AdvancedCopyTab?        No
    → BaseToolTab?            No
    → FileInputMixin?         Yes! ✓ Found

self.validate_file_exists()
    → AdvancedCopyTab?        No
    → BaseToolTab?            No
    → FileInputMixin?         No
    → ValidationMixin?        Yes! ✓ Found
```

---

### 12.4 Common Patterns Cheat Sheet

#### Create a New Tool

1. Create directory: `tools/my_tool/`
2. Create `my_tool_tool.py` (BaseTool)
3. Create `my_tool_ui.py` (BaseToolTab)
4. Create `my_tool_core.py` (Business logic)
5. Run app - tool auto-discovered!

#### Add File Input to Tool

```python
self.file_input = self.create_file_input_section(
    parent=self.frame,
    title="📁 FILE INPUT",
    file_types=[("PBIP Files", "*.pbip")],
    guide_text=["Instructions here"]
)
self.file_input['frame'].grid(row=0, column=0)
```

#### Add Background Operation

```python
self.run_in_background(
    target_func=self._do_work,
    success_callback=self._on_success,
    error_callback=self._on_error
)
```

#### Update Progress

```python
self.update_progress(50, "Processing...", show=True)
```

#### Log Message

```python
self.log_message("✅ Success!")
```

#### Show Dialog

```python
self.show_error("Error", "Something went wrong")
self.show_info("Info", "Operation complete")
self.show_warning("Warning", "Proceed with caution")
```

---

### 12.5 Tool Development Checklist

When creating a new tool, use this checklist:

**Planning:**
- [ ] Define tool purpose and capabilities
- [ ] Identify required inputs and outputs
- [ ] List dependencies and requirements
- [ ] Design UI layout on paper

**Implementation:**
- [ ] Create tool directory
- [ ] Implement `BaseTool` subclass
- [ ] Implement `BaseToolTab` subclass
- [ ] Implement business logic
- [ ] Add proper error handling
- [ ] Add progress tracking
- [ ] Add logging

**Testing:**
- [ ] Test with valid inputs
- [ ] Test with invalid inputs
- [ ] Test error cases
- [ ] Test background operations
- [ ] Test UI responsiveness
- [ ] Test with large files

**Documentation:**
- [ ] Write TECHNICAL_GUIDE.md
- [ ] Document all public methods
- [ ] Add inline comments for complex logic
- [ ] Create help dialog content
- [ ] Update this ARCHITECTURE.md if needed

**Integration:**
- [ ] Test tool discovery and registration
- [ ] Test tab creation
- [ ] Test with other tools
- [ ] Verify window sizing
- [ ] Test help dialog
- [ ] Test reset functionality

**Polish:**
- [ ] Apply consistent styling
- [ ] Use appropriate emojis
- [ ] Add welcome message
- [ ] Test all button states
- [ ] Verify progress bar positioning
- [ ] Final UX review

---

## Conclusion

The AE Multi-Tool architecture demonstrates a mature, extensible framework for building professional Power BI management tools. Key strengths:

**✅ Maintainable** - Clean separation of concerns
**✅ Extensible** - Easy to add new tools
**✅ Composable** - Mixins provide flexible functionality
**✅ Professional** - Consistent UI/UX and error handling
**✅ Future-Ready** - Plugin system and MCP integration

For detailed implementation guides for specific tools, refer to their individual `TECHNICAL_GUIDE.md` files.

---

**Built by Reid Havens of Analytic Endeavors**  
**Website:** https://www.analyticendeavors.com

---