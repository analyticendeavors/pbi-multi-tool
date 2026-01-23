# AE Multi-Tool - UI Design Patterns & Component Reference

**Built by Reid Havens of Analytic Endeavors**

**Last Updated:** January 2026

This guide documents the design patterns, color schemes, and UI components used across all tools in the AE Multi-Tool suite. Following these patterns ensures visual consistency and a professional appearance.

---

## Table of Contents

1. [Quick Reference](#quick-reference)
2. [Critical Patterns & Common Mistakes](#critical-patterns--common-mistakes)
3. [Theme System](#theme-system)
4. [Section Header Pattern](#section-header-pattern)
5. [Template Widgets](#template-widgets)
6. [Button Patterns](#button-patterns)
7. [Checkbox & Radio Patterns](#checkbox--radio-patterns)
8. [Treeview & List Patterns](#treeview--list-patterns)
9. [Combobox Patterns](#combobox-patterns)
10. [Progress Log & Analysis Summary](#progress-log--analysis-summary)
11. [Error Messages](#error-messages)
12. [Card Patterns](#card-patterns)
13. [Dialog Patterns](#dialog-patterns)
14. [Layout Best Practices](#layout-best-practices)
15. [Icon Loading & Theme Updates](#icon-loading--theme-updates)
16. [Tooltips & Message Boxes](#tooltips--message-boxes)
17. [Claude Code Prompts](#claude-code-prompts)
18. [Checklist for New Tools](#checklist-for-new-tools)

---

## Quick Reference

### Essential Imports

```python
from core.constants import AppConstants
from core.theme_manager import get_theme_manager
from core.ui_base import (
    BaseToolTab,
    RoundedButton,
    SquareIconButton,
    SVGToggle,
    ThemedScrollbar,
    ModernScrolledText,
    ThemedMessageBox,
    ThemedInputDialog,
    ThemedContextMenu,
    Tooltip,
    # Template Widgets (composable building blocks)
    ActionButtonBar,
    FileInputSection,
    SplitLogSection
)
from core.widgets import FieldNavigator, DropTargetProtocol
```

### Style Names

| Style | Use For |
|-------|---------|
| `Section.TLabelframe` | Main sections (no border) |
| `Section.TFrame` | Content inside sections (uses `colors['background']`) |
| `Section.TLabelframe.Label` | Section titles |
| `Section.TEntry` | Entry fields inside Section.TFrame (blends with background) - AVOID `state='readonly'` |
| `AnalysisSummary.TFrame` | Analysis results (borderless) |
| `ProgressLog.TFrame` | Log sections (faint border) |
| `Card.TFrame` | Card containers |
| `Dialog.TFrame` | Help dialogs |

### Color Keys Quick Reference

| Color Key | Usage |
|-----------|-------|
| `background` | Main window/frame background (#0d0d1a dark / #ffffff light) |
| `section_bg` | Section frame backgrounds (#161627 dark / #f5f5f7 light) |
| `card_surface` | Card/panel backgrounds |
| `text_primary` | Primary text |
| `text_secondary` | Secondary/subtitle text |
| `text_muted` | Disabled/hint text |
| `border` | Borders and separators |
| `primary` | Button accent color (always teal #009999) |
| `title_color` | Section titles, selected text (blue #0084b7 in dark, teal #009999 in light) |
| `selection_bg` | Text widget selection background (#1a5a8a dark / #3b82f6 light) - ALWAYS blue |
| `risk_high` | High severity (#dc2626) |
| `risk_medium` | Medium severity (#d97706) |
| `risk_low` | Low severity (#059669) |

### Entry Widget Patterns

**For read-only file paths (Browse button pattern):**
```python
# DO NOT use state='readonly' - causes 1px corner artifacts on Windows
file_entry = ttk.Entry(parent, textvariable=self.file_path,
                       font=('Segoe UI', 10), style='Section.TEntry')
file_entry.grid(row=0, column=1, sticky=(tk.W, tk.E))

# Prevent keyboard input via event bindings instead
file_entry.bind('<Key>', lambda e: 'break')
file_entry.bind('<Button-1>', lambda e: file_entry.focus_set())
```

**For editable entries:**
```python
# Use Section.TEntry for entries inside content areas
entry = ttk.Entry(content_frame, textvariable=self.value,
                  font=('Segoe UI', 10), style='Section.TEntry')
```

---

## Critical Patterns & Common Mistakes

This section consolidates the most critical patterns and common pitfalls. These warnings are also documented in their relevant sections with full code examples.

### Theme & Colors

| Pattern | Common Mistake | Correct Approach | See Section |
|---------|---------------|------------------|-------------|
| Selected text color | Using `primary` for selected/highlighted text | Use `title_color` - it's blue in dark mode, teal in light mode | [Theme System](#theme-system) |
| RoundedButton background | Setting `canvas_bg` to a fixed color | `canvas_bg` must match exact parent container background color | [Button Patterns](#button-patterns) |
| Text widget background | Using `surface` or `card_surface` | Use `colors['section_bg']` for text widget backgrounds | [Theme System](#theme-system) |
| Popup/dialog backgrounds | Using `section_bg` in popups | ALL popup elements use `colors['background']` | [Dialog Patterns](#dialog-patterns) |
| Global ttk styles | Using `style.configure('TFrame', ...)` | Never override global styles - create custom named styles | [Icon Loading & Theme Updates](#icon-loading--theme-updates) |

### Treeview & Lists

| Pattern | Common Mistake | Correct Approach | See Section |
|---------|---------------|------------------|-------------|
| Border removal | Using `style.configure()` to remove border | Use `style.layout()` - border is in layout elements | [Treeview & List Patterns](#treeview--list-patterns) |
| Selection color mapping | Mapping both background AND foreground | Only map `background` - mapping foreground breaks tag colors | [Treeview & List Patterns](#treeview--list-patterns) |
| ThemedScrollbar update | Calling internal methods on theme change | Use public `on_theme_changed()` method | [Treeview & List Patterns](#treeview--list-patterns) |

### Layout & Frames

| Pattern | Common Mistake | Correct Approach | See Section |
|---------|---------------|------------------|-------------|
| Inner content frames | Using `tk.Frame` with `padding` property | `tk.Frame` has no `padding` - use `ttk.Frame` or pack's `padx/pady` | [Layout Best Practices](#layout-best-practices) |
| SectionPanelMixin wrapper | Placing content directly in LabelFrame | Create `Section.TFrame` wrapper with `fill=BOTH, expand=True` | [Layout Best Practices](#layout-best-practices) |
| Column width locking | Using only `minsize` to lock widths | Must use `weight=0` - minsize alone doesn't prevent drift | [Layout Best Practices](#layout-best-practices) |
| pack(before=widget) | Omitting `side=` parameter | Must specify `side=` matching target widget's side | [Layout Best Practices](#layout-best-practices) |
| Dialog geometry | Updating only one location | Geometry is set TWICE in dialogs - update both | [Dialog Patterns](#dialog-patterns) |

### Theme Change Handling

| Pattern | Common Mistake | Correct Approach | See Section |
|---------|---------------|------------------|-------------|
| Checkbox icons | Only reloading icons on theme change | Must also call `_update_checkboxes()` to refresh display | [Icon Loading & Theme Updates](#icon-loading--theme-updates) |
| Dynamic widgets | Not tracking dynamically created widgets | Store references and update in `on_theme_changed()` | [Icon Loading & Theme Updates](#icon-loading--theme-updates) |
| Column widths | Using weight=1 on fixed-width columns | Use `weight=0` to prevent width drift on theme toggle | [Icon Loading & Theme Updates](#icon-loading--theme-updates) |

### Mixins & Inheritance

| Pattern | Common Mistake | Correct Approach | See Section |
|---------|---------------|------------------|-------------|
| MRO with BaseToolTab | Putting BaseToolTab before mixins | Place mixins BEFORE BaseToolTab in inheritance list | [Claude Code Prompts](#claude-code-prompts) |

### Icons

| Pattern | Common Mistake | Correct Approach | See Section |
|---------|---------------|------------------|-------------|
| Icon filenames | Using incorrect icon name | Icon name must match exact SVG filename (without .svg) | [Icon Loading & Theme Updates](#icon-loading--theme-updates) |

---

## Theme System

### Theme Overview

The application supports **Dark Mode** and **Light Mode** with distinct color palettes.

| Element | Dark Mode | Light Mode |
|---------|-----------|------------|
| **Primary (Brand)** | `#009999` (Teal) | `#009999` (Teal) |
| **Secondary** | `#00587C` (Dark Blue) | `#00587C` (Dark Blue) |
| **Title Color** | `#0084b7` (Light Blue) | `#009999` (Teal) |
| **Background** | `#0d0d1a` | `#ffffff` |
| **Section Background** | `#161627` | `#f5f5f7` |
| **Card Surface** | `#1a1a2e` | `#e8e8f0` |

### title_color vs primary - CRITICAL Distinction

**Important:** Use the correct color for highlighted/selected text:

| Context | Color to Use | Why |
|---------|-------------|-----|
| Buttons, button hover | `primary` | Brand color for interactive elements |
| Section headers | `title_color` | Blue in dark mode for visual hierarchy |
| Selected items in lists | `title_color` | Must match section headers for consistency |
| Radio button selected text | `title_color` | Matches section header styling |

**Common Mistake:**
```python
# WRONG - Using primary for selected text (always teal)
fg_color = colors['primary'] if is_selected else colors['text_primary']

# CORRECT - Using title_color (blue in dark mode, teal in light mode)
fg_color = colors['title_color'] if is_selected else colors['text_primary']
```

**In Dark Mode:**
- `primary` = #009999 (teal) - for buttons only
- `title_color` = #0084b7 (blue) - for titles and selected text

**In Light Mode:**
- Both are #009999 (teal) - but still use the semantically correct key

### Button Colors

| State | Dark Mode | Light Mode |
|-------|-----------|------------|
| Primary Normal | `#00587C` (Blue) | `#009999` (Teal) |
| Primary Hover | `#004466` | `#007A7A` |
| Primary Pressed | `#003050` | `#005C5C` |
| Primary Disabled | `#3a3a4e` | `#c0c0cc` |
| Secondary Normal | `#2a2a40` | `#d8d8e0` |
| Secondary Hover | `#222236` | `#c8c8d0` |
| Button Text | `#ffffff` | `#ffffff` |
| Disabled Text | `#6a6a7a` | `#9a9aa8` |

### Status Colors

| Status | Color |
|--------|-------|
| Success | `#10b981` (dark) / `#059669` (light) |
| Warning | `#f5751f` |
| Error | `#ef4444` (dark) / `#dc2626` (light) |
| Info | `#3b82f6` (dark) / `#2563eb` (light) |

### Typography

All fonts are defined in `AppConstants.FONTS`:

| Key | Font | Usage |
|-----|------|-------|
| `section_header` | Segoe UI Semibold, 11 | Section titles |
| `dialog_title` | Segoe UI, 16, bold | Help dialog main titles |
| `dialog_section` | Segoe UI, 12, bold | Help dialog section headers |
| `button` | Segoe UI, 10, bold | Primary action buttons |
| `button_secondary` | Segoe UI, 10 | Secondary/reset buttons |
| `label` | Segoe UI, 10 | Standard labels |
| `label_bold` | Segoe UI, 10, bold | Bold labels/values |
| `body` | Segoe UI, 9 | Body text, table cells |
| `body_italic` | Segoe UI, 9, italic | Tips/hints |
| `table_header` | Segoe UI, 9, bold | Table column headers |
| `log` | Consolas, 9 | Monospace log text |

### Usage in Code

```python
from core.theme_manager import get_theme_manager
from core.constants import AppConstants

# Get current theme colors
colors = get_theme_manager().colors

# Use colors from theme
bg_color = colors['background']
title_color = colors['title_color']
button_bg = colors['button_primary']

# Use font constants
label = tk.Label(parent, text="Label", font=AppConstants.FONTS['label'])
```

### Background Color Rules

#### Two-Layer Section Structure

**Main window sections** have a two-layer structure:

```
+-- Section Frame (section_bg = gray) ---------------+
|  +-- Inner Content (background = white) ---------+ |
|  |  Labels, buttons, inputs use 'background'     | |
|  +-----------------------------------------------+ |
+----------------------------------------------------+
```

```python
colors = self._theme_manager.colors
section_bg = colors.get('section_bg', colors['background'])  # Gray border
content_bg = colors['background']  # White inner content

# Outer section with gray border/frame
section_frame = ttk.LabelFrame(parent, labelwidget=header_widget,
                              style='Section.TLabelframe', padding="12")

# Inner content frame with WHITE background
content_frame = tk.Frame(section_frame, bg=content_bg, padx=15, pady=15)
content_frame.pack(fill=tk.BOTH, expand=True)
```

#### Popup/Dialog vs Main Window

| Context | Background Color | Why |
|---------|-----------------|-----|
| Main window sections | Outer: `section_bg`, Inner: `background` | Creates visual hierarchy |
| Popup/dialog windows | `colors['background']` everywhere | Clean surface, no gray needed |

#### tk.Frame vs ttk.Frame

Use `tk.Frame` when you need reliable background color control:

```python
# RELIABLE - tk.Frame allows explicit bg control
content_frame = tk.Frame(parent, bg=section_bg, padx=15, pady=15)

# In on_theme_changed
content_frame.configure(bg=section_bg)  # Works!

# UNRELIABLE - ttk.Frame background comes from style
content_frame = ttk.Frame(parent, style='Section.TFrame')
content_frame.configure(bg=section_bg)  # TclError - ttk doesn't have 'bg' option
```

#### Inner Content Frame Pattern (CRITICAL)

**Problem:** Creating inner content frames inside `Section.TLabelframe` sections.

**WARNING - OUTER Container Frames:**

`Section.TFrame` should ONLY be used for frames INSIDE `Section.TLabelframe` sections. Do NOT use it for:
- The main content_frame that HOLDS your sections
- Wrapper frames between the tab and the sections

Using `Section.TFrame` on outer containers creates visible color gaps between sections.

```python
# WRONG - creates visible divider lines between sections
content_frame = ttk.Frame(self.frame, style='Section.TFrame')  # DON'T DO THIS

# CORRECT - plain ttk.Frame for outer containers
content_frame = ttk.Frame(self.frame)  # No style for outer containers
```

**Two Approaches:**

| Approach | Use Case | Properties |
|----------|----------|------------|
| `ttk.Frame(parent, style='Section.TFrame', padding="15")` | Inner content needing auto-themed bg + inner padding | Auto-themed via style system, has `padding` property |
| `tk.Frame(parent, bg=content_bg)` | Nested frames where you need explicit bg control | Requires manual bg, NO `padding` property |

**CRITICAL: Padding Difference**

`tk.Frame` does NOT have a `padding` property. Using `padx/pady` in `pack()` adds padding OUTSIDE the frame, not inside:

```python
# WRONG - padx/pady add padding OUTSIDE the frame (between frame and parent)
inner_frame = tk.Frame(parent, bg=content_bg)
inner_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=15)  # Padding is OUTSIDE

# CORRECT - ttk.Frame with style has inner padding
inner_frame = ttk.Frame(parent, style='Section.TFrame', padding="15")  # Padding is INSIDE
inner_frame.pack(fill=tk.BOTH, expand=True)
```

**When to Use Each:**

| Use `ttk.Frame` with `Section.TFrame` | Use `tk.Frame` with explicit `bg=` |
|---------------------------------------|-----------------------------------|
| Main inner content frame of a section | Nested frames for layout (e.g., button rows) |
| When you need inner padding | When you need to track frame for theme updates |
| When auto-theming is sufficient | When combining with child tk widgets |

**Standard Pattern for Section Content:**

```python
# Outer section (gray border)
section_frame = ttk.LabelFrame(parent, labelwidget=header_widget,
                              style='Section.TLabelframe', padding="12")
section_frame.pack(fill=tk.X, pady=(0, 15))

# Inner content (white background + inner padding)
inner_frame = ttk.Frame(section_frame, style='Section.TFrame', padding="15")
inner_frame.pack(fill=tk.BOTH, expand=True)

# Child widgets - use tk widgets with explicit bg for theme control
content_bg = colors['background']  # Match Section.TFrame's background
label = tk.Label(inner_frame, text="Label", bg=content_bg, fg=colors['text_primary'])
```

**Note:** Child widgets inside a ttk.Frame still need explicit `bg=colors['background']` and should be tracked for theme updates.

#### Background Color Summary

| Context | Use This Color |
|---------|---------------|
| Section frame/border | `section_bg` |
| Inner content area | `colors['background']` |
| Elements inside content | `colors['background']` |
| Section header labelwidget | `section_bg` |
| Popup/dialog everything | `colors['background']` |
| Canvas button corners | Match exact parent background |

#### Panel Class Pattern (SectionPanelMixin)

When creating panel classes that inherit from `ttk.LabelFrame` with `SectionPanelMixin`:

> **CRITICAL**: Always create a `Section.TFrame` wrapper inside LabelFrame panels that is packed with `fill=tk.BOTH, expand=True`. This wrapper fills the LabelFrame's padding area with the correct background color. Without it, the lighter `section_bg` color will show through around the edges.

```python
class MyPanel(SectionPanelMixin, ttk.LabelFrame):
    def __init__(self, parent, main_tab):
        ttk.LabelFrame.__init__(self, parent, style='Section.TLabelframe', padding="12")
        self._init_section_panel_mixin()

    def setup_ui(self):
        colors = self._theme_manager.colors
        content_bg = colors['background']

        # CRITICAL: Inner wrapper with Section.TFrame style fills LabelFrame
        self._content_wrapper = ttk.Frame(self, style='Section.TFrame', padding="10")
        self._content_wrapper.pack(fill=tk.BOTH, expand=True)

        # All inner frames use _content_wrapper as parent, NOT self!
        toolbar = tk.Frame(self._content_wrapper, bg=content_bg)
        toolbar.pack(fill=tk.X)

        # Track frames for theme updates
        self._inner_frames = [toolbar]

    def on_theme_changed(self):
        colors = self._theme_manager.colors
        content_bg = colors['background']
        for frame in self._inner_frames:
            frame.config(bg=content_bg)
```

**Why the wrapper is required:**
- `ttk.LabelFrame` with `padding="12"` creates a visible padding area
- This padding area uses the LabelFrame's background (`section_bg`)
- The `Section.TFrame` wrapper covers the entire interior with correct color

---

## Section Header Pattern

### Standard Section Structure

Every tool section follows this structure:

```
+--[ ICON  SECTION TITLE ]------------------[?]--+
|                                                 |
|   Content frame with padding (15px)             |
|                                                 |
+-------------------------------------------------+
```

### Implementation (Consistent Across ALL Tools)

```python
def _create_section_labelwidget(self, text: str, icon_name: str) -> tk.Frame:
    """Create a labelwidget for LabelFrame with icon + text"""
    colors = self._theme_manager.colors
    icon = self._button_icons.get(icon_name)
    bg_color = colors.get('section_bg', colors['background'])

    # Frame to hold icon + text
    header_frame = tk.Frame(self.frame, bg=bg_color)

    icon_label = None
    if icon:
        icon_label = tk.Label(header_frame, image=icon, bg=bg_color)
        icon_label.pack(side=tk.LEFT, padx=(0, 6))
        icon_label._icon_ref = icon  # Prevent garbage collection

    # Use title_color and Semibold font
    text_label = tk.Label(header_frame, text=text, bg=bg_color,
                         fg=colors['title_color'], font=('Segoe UI Semibold', 11))
    text_label.pack(side=tk.LEFT)

    # Store for theme updates
    self._section_header_widgets.append((header_frame, icon_label, text_label))

    return header_frame

# Usage
header_widget = self._create_section_labelwidget("PBIP File Source", "Power-BI")
section_frame = ttk.LabelFrame(self.frame, labelwidget=header_widget,
                             style='Section.TLabelframe', padding="12")
section_frame.pack(fill=tk.X, pady=(0, 15))
```

### Key Design Rules

1. **Icon + Title**: Every section has a 16px SVG icon paired with its title
2. **Title Case**: Section titles use Title Case (e.g., "Sensitivity Scanner Setup")
3. **Title Color**: Uses `title_color` (blue in dark mode, teal in light mode)
4. **No Border**: Sections use background color framing, NOT visible borders
5. **Help Icon**: Upper-right corner, positioned with `place(relx=1.0, y=-35, anchor=tk.NE)`

### Help Button Positioning

```python
from core.ui_base import SquareIconButton

# Load help icon and create button AFTER content_frame
help_icon = self._load_icon_for_button("question", size=14)
self._button_icons['question'] = help_icon

self._help_button = SquareIconButton(
    section_frame, icon=help_icon, command=self.show_help_dialog,
    tooltip_text="Help", size=26, radius=6,
    bg_normal_override={'dark': '#0d0d1a', 'light': '#ffffff'}
)
# Position in upper-right corner of the section title bar area
self._help_button.place(relx=1.0, y=-35, anchor=tk.NE, x=-0)
```

**Critical Notes:**
- Create `content_frame` BEFORE the help button for proper stacking order
- Do NOT call `lift()` or `tkraise()` - Canvas widgets override these
- Use `y=-35` to position in title bar region

### Theme Update for Section Headers

```python
def on_theme_changed(self, theme: str):
    colors = self._theme_manager.colors
    bg_color = colors.get('section_bg', colors['background'])

    for header_frame, icon_label, text_label in self._section_header_widgets:
        try:
            header_frame.configure(bg=bg_color)
            if icon_label:
                icon_label.configure(bg=bg_color)
            text_label.configure(bg=bg_color, fg=colors['title_color'])
        except Exception:
            pass
```

---

## Template Widgets

Template widgets are composable building blocks that encapsulate common UI patterns. They handle theme changes automatically and provide a consistent API for common operations.

### Available Template Widgets

| Widget | Purpose | Replaces |
|--------|---------|----------|
| `ActionButtonBar` | Bottom action buttons (Primary + Secondary) | Manual button creation |
| `FileInputSection` | File input with browse, label, optional action button | Manual file input sections |
| `SplitLogSection` | Split log with summary (left) and progress log (right) | `create_split_log_section()` |
| `HierarchicalFilterDropdown` | Hierarchical filter with checkbox groups | Manual filter UI |
| `FieldNavigator` | Hierarchical field browser with search, filter, drag-drop | Manual treeview field lists |

### ActionButtonBar

Standard action button bar with primary and optional secondary buttons.

```python
from core.ui_base import ActionButtonBar

# Create button bar
button_bar = ActionButtonBar(
    parent=self.frame,
    theme_manager=self._theme_manager,
    primary_text="EXECUTE MERGE",
    primary_command=self.start_merge,
    primary_icon=execute_icon,
    secondary_text="RESET ALL",
    secondary_command=self.reset_tab,
    secondary_icon=reset_icon,
    primary_starts_disabled=True  # Optional: start with primary disabled
)
button_bar.pack(side=tk.BOTTOM, pady=(15, 0))

# Control buttons
button_bar.set_primary_enabled(True)   # Enable primary button
button_bar.set_secondary_enabled(False) # Disable secondary button

# Access buttons directly
button_bar.primary_button.set_enabled(False)
button_bar.secondary_button.update_colors(...)
```

**Key Features:**
- Auto-registers for theme changes (unregisters on destroy)
- Handles disabled state colors automatically
- Secondary button is optional (pass `secondary_command=None` to hide)

### FileInputSection

Complete file input section with styled LabelFrame, entry, browse button, and optional action button.

```python
from core.ui_base import FileInputSection

# Create file input section
file_section = FileInputSection(
    parent=self.frame,
    theme_manager=self._theme_manager,
    section_title="PBIP File Source",
    section_icon="Power-BI",
    file_label="Project File (PBIP):",
    file_types=[("PBIP files", "*.pbip"), ("All files", "*.*")],
    action_button_text="ANALYZE REPORT",
    action_button_command=self.analyze,
    action_button_icon=analyze_icon,
    on_file_selected=self.validate_file,  # Callback when file is selected
    help_command=self.show_help  # Optional help button
)
file_section.pack(fill=tk.X, pady=(0, 15))

# Get/set file path
path = file_section.path
file_section.path = "C:/path/to/file.pbip"

# Control action button
file_section.set_action_enabled(True)

# Access underlying widgets
file_section.path_var  # tk.StringVar
file_section.browse_button  # RoundedButton
file_section.action_button  # RoundedButton (may be None)
file_section.content_frame  # For adding custom widgets
```

**Key Features:**
- Includes section header with icon
- Browse button with file dialog
- Optional action button (e.g., "ANALYZE", "SCAN")
- Callback when file is selected via browse
- Theme-aware with auto-registration

### SplitLogSection

Split log section with summary panel (left) and progress log (right).

```python
from core.ui_base import SplitLogSection

# Create split log section
log_section = SplitLogSection(
    parent=self.frame,
    theme_manager=self._theme_manager,
    section_title="Analysis & Progress",
    section_icon="analyze",
    summary_title="Analysis Summary",
    summary_icon="bar-chart",
    log_title="Progress Log",
    log_icon="log-file",
    summary_placeholder="Run analysis to see results",
    on_export=self.custom_export,  # Optional custom export handler
    on_clear=self.custom_clear     # Optional custom clear handler
)
log_section.pack(fill=tk.BOTH, expand=True, pady=(0, 15))

# Log messages
log_section.log("Processing file...")
log_section.log("Complete!", "success")

# Set summary content (hides placeholder)
log_section.set_summary("Found 5 issues\n- Issue 1\n- Issue 2")

# Clear operations
log_section.clear_log()      # Clear progress log
log_section.clear_summary()  # Clear summary and show placeholder

# Access underlying widgets
log_section.log_text       # ModernScrolledText for log
log_section.summary_text   # ModernScrolledText for summary
log_section.summary_frame  # For custom content
log_section.export_button  # SquareIconButton
log_section.clear_button   # SquareIconButton
log_section.placeholder_label  # tk.Label
```

**Key Features:**
- Left panel: Summary with placeholder text (no border)
- Right panel: Progress log with border, export/clear buttons
- `log()` method adds timestamped messages
- `set_summary()` hides placeholder and shows content
- Default export saves to file; override with `on_export` callback
- Theme-aware with auto-registration

**WARNING - Custom Content in summary_frame:**

If you destroy the placeholder_label to add custom content, you MUST understand the theme callback implications:

```python
# DANGEROUS - This pattern can break theme updates!
summary_frame = self.log_section.summary_frame
for widget in summary_frame.winfo_children():
    widget.destroy()  # Destroys placeholder_label!

# What happens: SplitLogSection._on_theme_changed() tries to configure
# the destroyed placeholder_label, raises TclError, and the callback
# exits early - leaving other widgets (like log header) un-updated.
```

**Safe patterns for custom content:**
1. **Hide instead of destroy:** Use `placeholder_label.grid_remove()` or `pack_forget()`
2. **The SplitLogSection handles this internally** with `winfo_exists()` checks, but be aware that destroying child widgets can cause cascading theme update failures if not handled carefully.

**Text Widget Selection Colors:**

When creating or modifying Text/ScrolledText widgets, ALWAYS use these selection color settings:

```python
# CORRECT - Use selection_bg for text selection highlight
text_widget = ModernScrolledText(
    parent,
    selectbackground=colors['selection_bg'],  # Blue: #1a5a8a (dark) / #3b82f6 (light)
    selectforeground='#ffffff',               # White text when selected
    ...
)

# WRONG - Never use primary for text selection
text_widget = ModernScrolledText(
    parent,
    selectbackground=colors['primary'],       # WRONG! This is teal, not for text selection
    ...
)
```

| Property | Value | Purpose |
|----------|-------|---------|
| `selectbackground` | `colors['selection_bg']` | Blue selection highlight (#1a5a8a dark / #3b82f6 light) |
| `selectforeground` | `'#ffffff'` | White text for readability when selected |

**Important:** The `primary` color (teal) is for buttons and accents, NOT for text selection. Text selection must always use `selection_bg` (blue) for visual consistency across all tools.

### HierarchicalFilterDropdown

Hierarchical filter dropdown with checkbox groups for filtering results. Shows parent items (e.g., severity levels, categories) with child items nested underneath.

```python
from core.filter_dropdown import HierarchicalFilterDropdown

# Create filter dropdown
self._filter_dropdown = HierarchicalFilterDropdown(
    parent=results_header_frame,
    theme_manager=self._theme_manager,
    on_filter_changed=self._apply_filters,
    group_names=["High", "Medium", "Low"],  # Custom groups (optional)
    group_colors={"High": "#ef4444", "Medium": "#f59e0b", "Low": "#22c55e"},  # Custom colors (optional)
    header_text="Filter Results",
    empty_message="No findings to filter.\nRun a scan first."
)
self._filter_dropdown.pack(side=tk.RIGHT)

# Populate with items (call after data is available)
self._filter_dropdown.set_items({
    "High": ["Email Address", "SSN", "Credit Card"],
    "Medium": ["IP Address", "Phone Number"],
    "Low": ["Date of Birth"]
})

# Check visibility in filter callback
def _apply_filters(self):
    for finding in self.all_findings:
        visible = self._filter_dropdown.is_item_visible(finding.rule_name)
        # Show/hide finding based on visibility

# Get selected state
selected_groups = self._filter_dropdown.get_selected_groups()  # Set of group names
selected_items = self._filter_dropdown.get_selected_items()    # Set of item names
```

**Constructor Parameters:**
| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `parent` | Widget | Required | Parent widget to attach to |
| `theme_manager` | ThemeManager | Required | Theme manager for colors |
| `on_filter_changed` | Callable | Required | Callback when filter selection changes |
| `group_names` | List[str] | `["High", "Medium", "Low"]` | List of group names in display order |
| `group_colors` | Dict[str, str] | None | Dict mapping group name to color (falls back to risk colors) |
| `header_text` | str | `"Filter Results"` | Text shown in dropdown header |
| `empty_message` | str | `"No items to filter..."` | Message shown when no items available |

**Key Methods:**
| Method | Description |
|--------|-------------|
| `set_items(items_by_group: Dict[str, List[str]])` | Populate the filter with items grouped by group name |
| `get_selected_groups()` | Get currently selected group names (Set[str]) |
| `get_selected_items()` | Get currently selected item names (Set[str]) |
| `is_item_visible(item_name: str)` | Check if an item should be visible based on filters |
| `on_theme_changed()` | Handle theme change (reload icons, update colors) |
| `pack(**kwargs)` / `grid(**kwargs)` | Geometry managers |

**Key Features:**
- Hierarchical checkbox structure with parent groups and nested child items
- Partial state checkbox icons (checked, unchecked, partial for mixed selection)
- Search functionality with 500ms debounce
- Expand/collapse group sections
- Select All / Clear All quick actions
- Scrollable content area (max height 350px)
- Theme-aware with automatic icon reload on theme change
- Two icon buttons: Filter (opens dropdown) + Clear (resets all filters)

**Customization Pattern:**

For tool-specific convenience methods, create a thin wrapper:

```python
from core.filter_dropdown import HierarchicalFilterDropdown as BaseFilterDropdown

class SensitivityFilterDropdown(BaseFilterDropdown):
    """Sensitivity Scanner specific filter dropdown."""

    def set_rules(self, rules_by_severity: Dict[str, List[str]]):
        """Convenience method with domain-specific naming."""
        self.set_items(rules_by_severity)

    def is_finding_visible(self, finding) -> bool:
        """Check if a finding should be visible based on filters."""
        return self.is_item_visible(finding.rule_name)
```

### FieldNavigator

Hierarchical field browser widget for displaying tables, folders, measures, and columns from Power BI models. Supports search, filtering by type, drag-and-drop, multi-select, and context menus with positional adds.

**Location:** `core.widgets.field_navigator`

```python
from core.widgets import FieldNavigator, DropTargetProtocol
from core.pbi_connector import FieldInfo, TableFieldsInfo

# Create field navigator (drop_target set later if target widget created after)
navigator = FieldNavigator(
    parent=self.frame,
    theme_manager=self._theme_manager,
    on_fields_selected=self._handle_fields_selected,
    drop_target=None,  # Set later via set_drop_target()
    section_title="Available Fields",
    section_icon="table",
    show_columns=True,
    show_add_button=True,
    duplicate_checker=self._is_duplicate,
    show_duplicate_dialogs=True,
    placeholder_text="Connect to view fields.",
    can_add_validator=self._validate_can_add,
)
navigator.pack(fill=tk.BOTH, expand=True)

# Create target panel, then connect drag-drop
target_panel = BuilderPanel(self.frame)
navigator.set_drop_target(target_panel)

# Populate with data
tables_data = {
    "Sales": TableFieldsInfo(name="Sales", fields=[...]),
    "Products": TableFieldsInfo(name="Products", fields=[...]),
}
navigator.set_fields(tables_data)

# Callback receives fields and optional position
def _handle_fields_selected(self, fields: List[FieldInfo], position: Optional[int]):
    for i, field in enumerate(fields):
        insert_pos = position + i if position is not None else None
        self.add_field(field.table_name, field.name, insert_pos)

# Duplicate checker returns True if field already exists
def _is_duplicate(self, table_name: str, field_name: str) -> bool:
    return any(f.table == table_name and f.name == field_name for f in self.fields)

# Validator returns True if adding is allowed
def _validate_can_add(self) -> bool:
    if not self.current_selection:
        ThemedMessageBox.showwarning(self, "No Selection", "Select a target first")
        return False
    return True
```

**Constructor Parameters:**

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `parent` | Widget | Required | Parent container |
| `theme_manager` | ThemeManager | Required | Theme manager instance |
| `on_fields_selected` | Callable | Required | Callback: `(fields: List[FieldInfo], position: Optional[int]) -> None` |
| `drop_target` | DropTargetProtocol | `None` | Widget receiving drag-drop (must implement protocol) |
| `section_title` | str | `"Available Fields"` | Header title text |
| `section_icon` | str | `"table"` | Header icon name (SVG without extension) |
| `show_columns` | bool | `True` | Show columns in tree (False = measures only) |
| `show_add_button` | bool | `True` | Show "Add" button in section header |
| `duplicate_checker` | Callable | `None` | Callback: `(table_name: str, field_name: str) -> bool` |
| `show_duplicate_dialogs` | bool | `True` | Show confirmation dialogs for duplicates |
| `placeholder_text` | str | `"No tables..."` | Text shown when no data loaded |
| `can_add_validator` | Callable | `None` | Callback: `() -> bool` - validates if adding is allowed |

**Key Methods:**

| Method | Description |
|--------|-------------|
| `set_fields(tables: Dict[str, TableFieldsInfo])` | Populate tree with table/field data |
| `clear()` | Clear all fields and reset state |
| `get_selected_fields() -> List[FieldInfo]` | Get currently selected field items |
| `set_enabled(enabled: bool)` | Enable/disable all controls |
| `set_drop_target(target: DropTargetProtocol)` | Set/update drop target (use when target created after navigator) |
| `on_theme_changed()` | Handle theme change (reload icons, update colors) |

**DropTargetProtocol:**

Any widget can receive drag-drop by implementing this protocol:

```python
from core.widgets import DropTargetProtocol

class BuilderPanel(ttk.Frame):
    # Implements DropTargetProtocol

    def show_drop_indicator(self, screen_y: int) -> None:
        """Show visual indicator at y position during drag."""
        # Calculate insertion point and show line/highlight
        ...

    def hide_drop_indicator(self) -> None:
        """Hide the drop indicator."""
        ...

    def get_drop_position(self) -> Optional[int]:
        """Get index where dropped item should be inserted."""
        return self._drop_index  # None = append to end
```

**Key Features:**
- Hierarchical display with tables, folders (via `\` separator), and fields
- Search filtering with 300ms debounce
- Type filter radio buttons (All/Measures/Columns)
- Multi-select with Shift+Click and Ctrl+Click
- Context menu with "Add", "Add to Top", "Add at Position" options
- Drag-and-drop with floating label and drop indicators
- Duplicate detection with bulk add dialogs
- Theme-aware SVG icons (measure/column types)
- Expand/collapse state preserved across searches

**Wrapper Pattern:**

For tool-specific convenience, create a thin wrapper:

```python
from core.widgets import FieldNavigator

class AvailableFieldsPanel(ttk.LabelFrame):
    """Thin wrapper for Field Parameters tool."""

    def __init__(self, parent, main_tab):
        ttk.LabelFrame.__init__(self, parent, style='Section.TLabelframe')
        self._navigator = FieldNavigator(
            parent=self,
            theme_manager=get_theme_manager(),
            on_fields_selected=self._handle_selected,
            drop_target=main_tab.builder_panel,
            duplicate_checker=self._is_duplicate,
        )
        self._navigator.pack(fill=tk.BOTH, expand=True)

    def _handle_selected(self, fields, position):
        for i, f in enumerate(fields):
            pos = position + i if position is not None else None
            self.main_tab.add_field(f.table_name, f.name, pos)

    def update_available_fields(self, tables):
        self._navigator.set_fields(tables)
```

### When to Use Template Widgets

| Scenario | Use Template Widget |
|----------|---------------------|
| New tool with standard bottom buttons | `ActionButtonBar` |
| New tool with file input section | `FileInputSection` |
| New tool with analysis + log section | `SplitLogSection` |
| New tool with filterable results | `HierarchicalFilterDropdown` |
| New tool with field/measure browser | `FieldNavigator` |
| Existing tool refactoring | Optional - existing code works |
| Custom button layout needed | Use `RoundedButton` directly |
| Custom file handling logic | Use manual file input pattern |

**Migration Note:** Template widgets are recommended for new tools. Existing tools can continue using manual patterns or be migrated as needed.

---

## Button Patterns

### Button Text Style

**IMPORTANT**: Action buttons use ALL CAPS text:

| Text Style | Example | Usage |
|------------|---------|-------|
| ALL CAPS | "ANALYZE REPORTS", "SCAN FOR ISSUES" | Primary action buttons |
| ALL CAPS | "RESET ALL", "EXPORT REPORT" | Secondary action buttons |
| Title Case | Section titles, labels | Headers and descriptive text |

### Button Types and Dimensions

| Type | Width | Height | Radius | Font | Usage |
|------|-------|--------|--------|------|-------|
| Primary | Auto | 38px | 6px | Bold | Main actions (Analyze, Execute, Scan) |
| Secondary | Auto | 38px | 6px | Normal | Supporting actions (Reset, Cancel) |
| Browse | 90px | 32px | 6px | Normal | Compact inline file selection |
| Dialog | Auto | 36px | 8px | Varies | Dialog buttons (Save, Close, Apply) |
| Compact | 58-68px | 26px | 5px | Normal | Quick actions (All, None, Select) |

**Auto-Sizing (Preferred)**: Omit the `width=` parameter to let buttons auto-size based on text content. This ensures consistent padding around text regardless of button label length.

**When to use hardcoded width:**
- Browse buttons (90px) - intentionally compact for inline use with file input
- Compact utility buttons (58-68px) - space-constrained UI areas
- **Avoid** hardcoding widths for main action buttons and dialog buttons

### RoundedButton Usage

```python
from core.ui_base import RoundedButton
from core.constants import AppConstants

colors = get_theme_manager().colors
is_dark = get_theme_manager().is_dark

# Primary action button (auto-sized, bold font)
action_button = RoundedButton(
    parent,
    text="ANALYZE REPORTS",
    command=self._analyze,
    bg=colors['button_primary'],
    hover_bg=colors['button_primary_hover'],
    pressed_bg=colors['button_primary_pressed'],
    fg='#ffffff',
    height=38,  # No width - auto-size
    radius=6,
    font=('Segoe UI', 10, 'bold'),  # Bold for primary actions
    icon=analyze_icon,
    disabled_bg=colors.get('button_primary_disabled', '#3a3a4e' if is_dark else '#c0c0cc'),
    disabled_fg=colors.get('button_text_disabled', '#6a6a7a' if is_dark else '#9a9aa8'),
    canvas_bg=colors['section_bg']  # CRITICAL: Must match parent background
)

# Secondary button (auto-sized, normal font)
reset_button = RoundedButton(
    parent,
    text="RESET ALL",
    command=self._reset,
    bg=colors['button_secondary'],
    hover_bg=colors['button_secondary_hover'],
    pressed_bg=colors['button_secondary_pressed'],
    fg=colors['text_primary'],
    height=38,  # No width - auto-size
    radius=6,
    font=('Segoe UI', 10),  # Non-bold for secondary actions
    icon=reset_icon,
    canvas_bg=colors['section_bg']
)

# Browse button (intentionally compact - 90px width)
browse_button = RoundedButton(
    input_row,
    text="Browse",
    command=self._browse_file,
    bg=colors['button_primary'],
    hover_bg=colors['button_primary_hover'],
    pressed_bg=colors['button_primary_pressed'],
    fg='#ffffff',
    width=90, height=32,  # Compact inline button
    radius=6,
    font=('Segoe UI', 10),
    icon=folder_icon,
    canvas_bg=input_row_bg
)
```

### CRITICAL: canvas_bg Parameter

The `canvas_bg` parameter must match the exact background color of the parent container, or the button corners will show a different color (corner rounding artifacts).

#### Background Locations and Their Colors

| Location | Style/Class | Background Color | Buttons Use |
|----------|-------------|------------------|-------------|
| Content area inside LabelFrame | `Section.TFrame` | `colors['background']` | `colors['background']` |
| Outer tab area (bottom buttons) | ttk.Frame | `colors['section_bg']` | `colors['section_bg']` |
| Help dialogs | `Dialog.TFrame` | `colors['background']` | `colors['background']` |

#### For Content Area Buttons (Browse, Scan, etc.)
Buttons inside `Section.TFrame` content areas use `colors['background']`:
```python
browse_btn = RoundedButton(
    content_frame,  # Section.TFrame uses colors['background']
    text="Browse",
    ...
    canvas_bg=colors['background']  # Match parent
)
```

#### For Bottom Action Buttons (Reset, Export, etc.)
Buttons on the outer tab area use `colors['section_bg']`:
```python
outer_canvas_bg = colors['section_bg']
reset_btn = RoundedButton(
    button_frame,  # Outer area uses section_bg
    text="RESET ALL",
    ...
    canvas_bg=outer_canvas_bg  # Match parent
)
```

**BaseToolTab Handles Primary Buttons Automatically:**
The `BaseToolTab.on_theme_changed()` updates `_primary_buttons` with `colors['background']` canvas_bg. Tools only need to explicitly handle bottom buttons that use `section_bg`:

```python
def on_theme_changed(self, theme: str):
    super().on_theme_changed(theme)  # Handles primary buttons
    colors = self._theme_manager.colors

    # Bottom buttons need explicit section_bg
    outer_canvas_bg = colors['section_bg']
    if hasattr(self, 'reset_btn') and self.reset_btn:
        self.reset_btn.update_canvas_bg(outer_canvas_bg)
```

### Disabled Button States

Buttons that can be disabled (e.g., action buttons before file selection) must use the standard disabled color keys from `constants.py`.

#### Standard Color Keys

| Key | Dark Mode | Light Mode | Usage |
|-----|-----------|------------|-------|
| `button_primary_disabled` | `#3a3a4e` | `#c0c0cc` | Disabled button background |
| `button_text_disabled` | `#6a6a7a` | `#9a9aa8` | Disabled button text |

#### Standard Pattern

```python
colors = self._theme_manager.colors
is_dark = self._theme_manager.is_dark

# Define disabled colors with proper dark/light fallbacks
disabled_bg = colors.get('button_primary_disabled', '#3a3a4e' if is_dark else '#c0c0cc')
disabled_fg = colors.get('button_text_disabled', '#6a6a7a' if is_dark else '#9a9aa8')

# Pass to RoundedButton
self.action_button = RoundedButton(
    parent, text="ACTION",
    command=self._on_action,
    bg=colors['button_primary'],
    hover_bg=colors['button_primary_hover'],
    pressed_bg=colors['button_primary_pressed'],
    fg='#ffffff',
    disabled_bg=disabled_bg,
    disabled_fg=disabled_fg,
    height=38, radius=6,
    font=('Segoe UI', 10, 'bold'),
    canvas_bg=outer_canvas_bg
)

# Control enabled state
self.action_button.set_enabled(False)  # Disable (grayed out, no hover)
self.action_button.set_enabled(True)   # Enable (normal interaction)
```

#### Common Mistakes

| Mistake | Correction |
|---------|------------|
| Wrong key: `'button_disabled_bg'` | Use `'button_primary_disabled'` |
| Missing light fallback: `'#3a3a4e'` only | Add `'#3a3a4e' if is_dark else '#c0c0cc'` |
| Inverted logic: `not is_dark` | Use `is_dark` for dark-first pattern |

#### RoundedButton Grid Layout Spacing

When placing multiple RoundedButtons in a row using grid with wrapper frames, use LEFT-side padding for spacing between buttons:

```python
# Wrapper frames for buttons in a grid row
add_wrapper.grid(row=0, column=0, sticky="ew")  # No padding for first
rename_wrapper.grid(row=0, column=1, sticky="ew", padx=(2, 0))  # Left padding
remove_wrapper.grid(row=0, column=2, padx=(2, 0))  # Left padding

# Pack buttons at natural size - maintains consistent text-to-edge padding
self.add_label_btn.pack()
self.rename_btn.pack()
```

**Important:** Do NOT use `pack(fill=tk.X, expand=True)` with RoundedButton. This causes buttons to expand and fill all available space, losing the standard internal padding between text and button edges. Always pack buttons at their natural calculated size.

### SquareIconButton Usage

For small icon-only buttons (help, export, clear):

```python
from core.ui_base import SquareIconButton

# Help button
help_button = SquareIconButton(
    section_frame, icon=help_icon,
    command=self.show_help_dialog,
    tooltip_text="Help", size=26, radius=6,
    bg_normal_override={'dark': '#0d0d1a', 'light': '#ffffff'}
)

# Export/Clear buttons in log header
export_button = SquareIconButton(
    icon_buttons_frame, icon=save_icon,
    command=lambda: self._export_log(self.log_text),
    tooltip_text="Export Log", size=26, radius=6
)
```

### SVGToggle Usage

For binary on/off states (used in Report Cleanup cards):

```python
from core.ui_base import SVGToggle

toggle = SVGToggle(
    parent,
    on_svg=str(self.base_path / "assets/Tool Icons/toggle-on.svg"),
    off_svg=str(self.base_path / "assets/Tool Icons/toggle-off.svg"),
    width=44, height=24,
    initial_state=True,
    command=self._on_toggle_changed,
    theme_manager=self._theme_manager
)
```

### Button Layout Standards

1. **Bottom Action Buttons**: Center-aligned, primary on left
   - Pack button frame with `side=tk.BOTTOM` FIRST (ensures visibility on resize)
   - Primary buttons: Auto-width, 38px height, bold font
   - Secondary buttons: Auto-width, 38px height, normal font
   - Both auto-size to text + consistent internal padding

2. **Inline Section Buttons**: Left-aligned within sections
   - Compact buttons: 58-68px width, 26px height (intentionally fixed for space-constrained areas)

3. **Dialog Buttons**: Right-aligned or center-aligned
   - Auto-width, 36px height, 8px radius
   - Primary action button on right (or left of Cancel)
   - Use bold for primary action, normal for others

---

## Checkbox & Radio Patterns

### SVG Checkbox Icons

**Theme-Aware Icons:**
- Light mode: `box.svg`, `box-checked.svg`, `box-partial.svg`
- Dark mode: `box-dark.svg`, `box-checked-dark.svg`, `box-partial-dark.svg`
- Size: 16px (rendered at 4x, downscaled with LANCZOS)

**Loading Pattern:**
```python
def _load_checkbox_icons(self):
    """Load themed checkbox SVG icons."""
    is_dark = self._theme_manager.is_dark

    box_name = 'box-dark' if is_dark else 'box'
    checked_name = 'box-checked-dark' if is_dark else 'box-checked'
    partial_name = 'box-partial-dark' if is_dark else 'box-partial'

    self._checkbox_off_icon = self._load_icon_for_button(box_name, size=16)
    self._checkbox_on_icon = self._load_icon_for_button(checked_name, size=16)
    self._checkbox_partial_icon = self._load_icon_for_button(partial_name, size=16)
```

**Usage in Selection Rows:**
```python
def _create_svg_checkbox(self, parent, text, var_key, bg_color):
    """Create an SVG checkbox with label."""
    colors = self._theme_manager.colors

    icon = self._checkbox_off_icon
    checkbox_label = tk.Label(parent, image=icon, cursor='hand2', bg=bg_color)
    checkbox_label.pack(side=tk.LEFT, padx=(0, 4))
    checkbox_label._icon_ref = icon  # Prevent garbage collection
    checkbox_label.bind('<Button-1>', lambda e: self._toggle_checkbox(var_key))

    text_label = tk.Label(parent, text=text, font=('Segoe UI', 9),
                         fg=colors['text_primary'], bg=bg_color, cursor='hand2')
    text_label.pack(side=tk.LEFT)
    text_label.bind('<Button-1>', lambda e: self._toggle_checkbox(var_key))

    return checkbox_label, text_label
```

### Partial Checkbox State

For hierarchical "Select All" parent items when some but not all children are selected:

```python
def _update_parent_checkbox(self, parent_key: str):
    """Update parent checkbox to show checked/partial/unchecked."""
    children = self._get_children_for_parent(parent_key)
    selected_count = len([c for c in children if c in self._selected_items])

    if selected_count == len(children) and len(children) > 0:
        icon = self._checkbox_on_icon      # All selected
    elif selected_count > 0:
        icon = self._checkbox_partial_icon  # Some selected
    else:
        icon = self._checkbox_off_icon      # None selected

    if icon:
        self._parent_labels[parent_key].configure(image=icon)
        self._parent_labels[parent_key]._icon_ref = icon
```

**Applied to:**
- Layout Optimizer: Diagram selection (Select All row)
- Sensitivity Scanner: Filter popup severity levels
- Table Column Widths: Visual selection treeview
- Advanced Copy: Page selection treeview

### SVG Radio Buttons

**Theme-Aware Icons:**
- `radio-on.svg` / `radio-on-dark.svg`
- `radio-off.svg` / `radio-off-dark.svg`

**Implementation with Hover Underline:**
```python
def _create_radio_row(self, parent, text: str, value: str, var: tk.StringVar, bg_color: str) -> dict:
    """Create a single SVG radio button row with hover underline."""
    colors = self._theme_manager.colors

    row_frame = tk.Frame(parent, bg=bg_color)
    row_frame.pack(side=tk.LEFT, padx=(0, 12))

    is_selected = var.get() == value
    icon = self._radio_on_icon if is_selected else self._radio_off_icon

    icon_label = tk.Label(row_frame, bg=bg_color, cursor='hand2')
    if icon:
        icon_label.configure(image=icon)
        icon_label._icon_ref = icon
    icon_label.pack(side=tk.LEFT, padx=(0, 4))

    # Selected uses title_color, unselected uses text_primary
    text_fg = colors['title_color'] if is_selected else colors['text_primary']

    text_label = tk.Label(row_frame, text=text, bg=bg_color, fg=text_fg,
                         font=('Segoe UI', 9), cursor='hand2')
    text_label.pack(side=tk.LEFT)

    # Click handler
    def on_click(event=None):
        var.set(value)
        self._update_radio_rows()

    icon_label.bind('<Button-1>', on_click)
    text_label.bind('<Button-1>', on_click)

    # Hover underline effect - applies to ALL items
    def on_enter(event=None):
        text_label.configure(font=('Segoe UI', 9, 'underline'))
    def on_leave(event=None):
        text_label.configure(font=('Segoe UI', 9))

    text_label.bind('<Enter>', on_enter)
    text_label.bind('<Leave>', on_leave)

    return {'frame': row_frame, 'icon_label': icon_label, 'text_label': text_label, 'value': value}
```

**Key Design Choices:**
- Hover underline shows on ALL items (including selected) for consistency
- Selected items use `title_color` (blue in dark mode)
- Unselected items use `text_primary`

### LabeledRadioGroup Reusable Widget

For radio button groups with SVG icons and disabled state support, use the `LabeledRadioGroup` class:

**Location:** `core/ui_base.py`

```python
from core.ui_base import LabeledRadioGroup

# Create radio group - format: (value, label) or (value, label, description)
self._conn_type_radio = LabeledRadioGroup(
    parent_frame,
    variable=self.connection_type_var,
    options=[
        ("pbiServiceLive", "Semantic Model"),
        ("analysisServicesDatabaseLive", "XMLA Endpoint"),
    ],
    command=self._on_connection_type_changed,
    orientation="horizontal",  # or "vertical"
    padding=12,
    bg=colors['section_bg']  # Optional custom background
)
self._conn_type_radio.pack(anchor=tk.W)

# With descriptions (e.g., in settings dialogs)
radio_group = LabeledRadioGroup(
    parent,
    variable=mode_var,
    options=[
        ("AA", "AA Standard (4.5:1)", "Recommended"),
        ("AAA", "AAA Enhanced (7:1)", "Stricter"),
    ]
)
```

**Disabling Options with Tooltips:**
```python
# Disable an option with explanatory tooltip (tooltip only shows when disabled)
self._conn_type_radio.set_option_enabled(
    "analysisServicesDatabaseLive",
    enabled=False,
    tooltip="XMLA endpoint not available for Pro workspaces.\nRequires Premium, PPU, or Fabric capacity."
)

# Re-enable an option (clears tooltip automatically)
self._conn_type_radio.set_option_enabled("analysisServicesDatabaseLive", enabled=True)
```

**Disabled State Visual:**
- Text color changes to `text_muted`
- Cursor changes from `hand2` to `arrow`
- Hover underline effect disabled
- Clicks ignored
- Tooltip shown on hover (only when disabled)

**Theme Support:**
```python
def on_theme_changed(self):
    self._conn_type_radio.on_theme_changed()  # Updates colors for new theme
    # Or update custom background:
    self._conn_type_radio.set_bg(colors['section_bg'])
```

### LabeledToggle Reusable Widget

For toggle switches with optional text labels, use the `LabeledToggle` class:

**Location:** `core/ui_base.py`

```python
from core.ui_base import LabeledToggle

# Create toggle with label
toggle = LabeledToggle(
    parent_frame,
    variable=my_bool_var,
    text="Enable feature",
    command=on_toggle_changed,
    icon_height=18
)
toggle.pack(anchor=tk.W)

# Toggle without label (icon only)
toggle = LabeledToggle(parent_frame, variable=my_bool_var)
```

**Features:**
- Uses toggle-on.svg and toggle-off.svg from assets/Tool Icons
- Class-level icon caching for efficiency
- Optional text label positioned to the right of the icon
- BooleanVar binding with command callback
- Theme-aware colors

**Setting Custom Background:**
```python
# For toggles inside section backgrounds
toggle.configure(bg=colors['section_bg'])
toggle._icon_label.configure(bg=colors['section_bg'])
if toggle._label:
    toggle._label.configure(bg=colors['section_bg'])
```

---

## Treeview & List Patterns

### Flat Treeview Style

Modern treeview styling uses flat design with groove headers:

```python
colors = self._theme_manager.colors
is_dark = self._theme_manager.is_dark

# Theme-aware heading colors
if is_dark:
    heading_bg = '#2a2a3c'
    heading_fg = '#e0e0e0'
    header_separator = '#0d0d1a'
else:
    heading_bg = '#f0f0f0'
    heading_fg = '#333333'
    header_separator = '#ffffff'

style = ttk.Style()
style.configure("Flat.Treeview",
    borderwidth=0,
    relief="flat",
    rowheight=25,
    background=colors['card_surface'],
    fieldbackground=colors['card_surface']
)
style.configure("Flat.Treeview.Heading",
    background=heading_bg,
    foreground=heading_fg,
    relief='groove',
    borderwidth=1,
    bordercolor=header_separator,
    font=('Segoe UI', 9, 'bold')
)
```

### Removing Treeview Inner Border

**Problem:** Treeviews have a default white/light inner border that doesn't match dark themes.

**Solution:** Use `style.layout()` to remove the border element from the Treeview layout:

```python
# Configure Treeview colors
style.configure("Treeview",
                background=tree_bg,
                fieldbackground=tree_bg,
                foreground=text_color)

# CRITICAL: Remove border via layout() - style.configure doesn't work for this
# The border is built into the layout elements, must be removed via layout()
style.layout("Treeview", [
    ('Treeview.treearea', {'sticky': 'nswe'})
])
```

**Why style.configure doesn't work:** The Treeview border is part of the widget's internal layout structure (e.g., `Treeview.border`), not a style property. Setting `borderwidth=0` or `relief="flat"` in configure() has no effect - you must redefine the layout to exclude the border element.

For container frames around treeviews, use `highlightthickness=0` to remove outer borders:

```python
# Container with NO border
tree_container = tk.Frame(parent, bg=tree_bg,
                          highlightbackground=border_color,
                          highlightcolor=border_color,
                          highlightthickness=0)  # No border
```

To ADD a border around a container (like Parameter Builder canvas):

```python
# Container WITH border
border_color = colors.get('border', '#3a3a4a')
canvas_frame = tk.Frame(parent, bg=bg_color,
                        highlightbackground=border_color,
                        highlightcolor=border_color,
                        highlightthickness=1)  # 1px border
```

### Treeview with Checkbox Column

For selection treeviews (Column Width, Advanced Copy):

```python
# Column #0 for checkbox icons (44px fixed width)
tree = ttk.Treeview(container, style="Flat.Treeview", show="tree headings")
tree.column("#0", width=44, minwidth=44, stretch=False, anchor='center')

# Insert item with checkbox icon
item_id = tree.insert("", tk.END, text="", image=self._checkbox_off_icon,
                     values=(name, type_val, page, fields))

# Toggle on click
def on_click(event):
    item = tree.identify_row(event.y)
    column = tree.identify_column(event.x)
    if column == '#0' and item:
        self._toggle_item_selection(item)

tree.bind('<Button-1>', on_click)
```

### ThemedScrollbar

Always use `ThemedScrollbar` instead of standard ttk.Scrollbar:

```python
from core.ui_base import ThemedScrollbar

scrollbar = ThemedScrollbar(
    parent_frame,
    command=self.tree.yview,
    theme_manager=self._theme_manager,
    width=12,
    auto_hide=True
)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)  # No padding
self.tree.configure(yscrollcommand=scrollbar.set)
```

**CRITICAL: Theme Updates for ThemedScrollbar**

When manually calling theme updates on a ThemedScrollbar, use the public `on_theme_changed()` method:

```python
# CORRECT - Use public method
def on_theme_changed(self):
    if hasattr(self, '_scrollbar'):
        self._scrollbar.on_theme_changed()

# WRONG - Private method requires 'theme' parameter
# This will cause TypeError and break theme updates!
self._scrollbar._on_theme_change()  # Missing 'theme' argument
```

Note: ThemedScrollbar auto-registers with theme_manager, so manual calls are only needed when you want to ensure immediate update or when the scrollbar was created without a theme_manager reference.

### Multi-Select Treeview

Enable multi-row selection with `selectmode=tk.EXTENDED`:

```python
tree = ttk.Treeview(
    container,
    columns=('id', 'name', 'risk'),
    show='headings',
    style="Flat.Treeview",
    selectmode=tk.EXTENDED  # Enable multi-select with Ctrl+Click and Shift+Click
)
```

### Selection Styling That Preserves Tag Colors

When using tags for color-coded rows (e.g., risk levels), configure selection to ONLY change background, not foreground. This preserves the tag colors when rows are selected:

```python
colors = self._theme_manager.colors
is_dark = self._theme_manager.is_dark

# Calculate selection background - use selection_highlight for visibility
selection_bg = colors.get('selection_highlight', colors.get('card_surface', '#1a3a5c' if is_dark else '#e6f3ff'))

# CRITICAL: Only map background for selected state, NOT foreground
# This preserves tag-based foreground colors (risk colors, etc.)
style.map("Flat.Treeview",
          background=[('selected', selection_bg)])
# Note: Intentionally NOT setting foreground mapping to preserve tag colors
```

**Why this matters:**
- Tags like `high_risk`, `medium_risk`, `low_risk` set foreground colors
- Default selection styling overrides ALL foreground colors
- By omitting foreground from the style.map, tag colors remain visible when rows are selected

**Selection Color Note:**
- Always use `selection_highlight` (not `card_surface`) for treeview selection backgrounds
- `card_surface` is too similar to tree/section backgrounds, making selections nearly invisible
- Color values in `constants.py`:
  - Dark: `selection_highlight` = `#1a3a5c` (blue-tinted, clearly visible)
  - Light: `selection_highlight` = `#e6f3ff` (light blue, clearly visible)

### Icon Button in Section Header

Add small icon buttons (like delete/trash) to section headers:

```python
# Create header frame with title, count, and action button
header_frame = tk.Frame(parent, bg=popup_bg)

# Title and count on left
header_title = tk.Label(header_frame, text="Current Patterns",
                        font=('Segoe UI', 10, 'bold'), bg=popup_bg, fg=colors['text_primary'])
header_title.pack(side=tk.LEFT)

count_label = tk.Label(header_frame, text=f"  ({count} patterns)",
                       font=('Segoe UI', 9), bg=popup_bg, fg=colors['text_muted'])
count_label.pack(side=tk.LEFT)

# Action button on right (e.g., delete icon)
delete_btn = tk.Label(
    header_frame,
    bg=popup_bg,
    cursor='hand2',
    padx=4, pady=2
)
if trash_icon:
    delete_btn.configure(image=trash_icon)
    delete_btn._icon_ref = trash_icon  # Prevent garbage collection
else:
    # Fallback to text
    delete_btn.configure(text="X", fg=colors['error'], font=('Segoe UI', 9, 'bold'))
delete_btn.pack(side=tk.RIGHT, padx=(10, 0))
delete_btn.bind('<Button-1>', lambda e: self._delete_selected())

# Hover effect
delete_btn.bind('<Enter>', lambda e: delete_btn.configure(
    bg=colors.get('card_surface_hover', colors.get('surface', colors['background']))))
delete_btn.bind('<Leave>', lambda e: delete_btn.configure(bg=popup_bg))

# Use as labelwidget
section_frame = ttk.LabelFrame(parent, labelwidget=header_frame, padding="10")
```

### Preserving Treeview Expanded State

When rebuilding tree content, preserve expand/collapse state:

```python
def _apply_filters(self):
    # 1. Save expanded state BEFORE clearing
    expanded_state = {}
    for item in self.tree.get_children():
        tags = self.tree.item(item, 'tags')
        if tags:
            tag = tags[0] if isinstance(tags, (list, tuple)) else tags
            expanded_state[tag] = self.tree.item(item, 'open')

    # 2. Clear and repopulate tree
    for item in self.tree.get_children():
        self.tree.delete(item)

    # 3. Insert with preserved state
    high_parent = self.tree.insert("", tk.END, text="HIGH RISK",
                                  tags=("high_risk",),
                                  open=expanded_state.get('high_risk', True))
```

### Virtual Scrolling Filter Refresh Pattern

**Problem:** In virtual scrolling lists with filtering, blank spaces appear during drag-drop or filter changes because widgets aren't properly hidden when the filtered item list changes.

**Root Cause:** When refreshing the view after a filter change:
1. Updating `_filtered_items` first changes which items are in the list
2. Trying to hide widgets based on the NEW list misses the old visible widgets
3. Old widgets remain visible at stale positions, causing visual artifacts

**Solution:** Capture old visible items BEFORE updating the filtered cache, then hide those specific widgets:

```python
def _refresh_virtual_view(self):
    """Refresh the virtual view after data changes"""
    if not self._virtual_mode:
        return

    # CRITICAL: Capture OLD visible items BEFORE updating filtered cache
    old_start, old_end = self._visible_range
    old_visible_items = list(self._filtered_items[old_start:old_end]) if self._filtered_items else []

    # Now update the filtered items cache
    self._update_filtered_items_cache()

    # Update scroll region based on NEW filtered count
    total_height = max(1, len(self._filtered_items) * self.ROW_HEIGHT)
    self.canvas.configure(scrollregion=(0, 0, self.canvas.winfo_width(), total_height))

    # Hide OLD visible widgets (not new ones!)
    for field_item in old_visible_items:
        field_id = id(field_item)
        if field_id in self.field_widgets:
            self.field_widgets[field_id]["frame"].place_forget()

    # Reset visible range and re-render with new items
    self._visible_range = (0, 0)
    self._update_virtual_view()
```

**Key Insight:** The order of operations matters. Save references to old state BEFORE modifying the underlying data structure, then use those saved references to clean up.

**Applied in:** Field Parameters tool (`field_parameters_builder.py`) - Parameter Builder panel virtual scrolling.

### Treeview with Capacity Icons

For displaying workspace capacity type icons (Premium, PPU) in treeviews:

**Icon Loading:**
```python
from PIL import Image, ImageTk
import cairosvg
import io
from pathlib import Path

def _load_capacity_icons(self):
    """Load Premium and PPU capacity icons for workspace/model display."""
    is_dark = self._theme_manager.is_dark
    icons_dir = Path(__file__).parent.parent.parent.parent.parent / "assets" / "Tool Icons"

    # Light icons for dark mode, dark icons for light mode
    icon_suffix = "" if is_dark else "-dark"
    icon_files = {
        'Premium': f"Premium{icon_suffix}.svg",
        'PPU': f"Premium-Per-User{icon_suffix}.svg"
    }

    for cap_type, filename in icon_files.items():
        icon_path = icons_dir / filename
        icon = self._load_svg_icon(icon_path, size=14)
        if icon:
            self._capacity_icons[cap_type] = icon
```

**Treeview Configuration for Icons:**
```python
# Enable tree column for icons with show="tree headings"
self.dataset_tree = ttk.Treeview(
    container,
    columns=("name", "workspace", "configured_by"),
    show="tree headings",  # Enable tree column for icons
    selectmode="browse",
    style=tree_style
)

# Configure tree column #0 for icons (narrow, icon only)
self.dataset_tree.column("#0", width=24, minwidth=24, stretch=False, anchor="center")
```

**Inserting Items with Icons:**
```python
# Get capacity icon for workspace
cap_icon = self._get_capacity_icon_for_workspace(ds.workspace_id)
if cap_icon:
    self.dataset_tree.insert(
        "", "end", iid=str(i), image=cap_icon,
        values=(ds.name, ds.workspace_name or "", ds.configured_by or "")
    )
else:
    self.dataset_tree.insert(
        "", "end", iid=str(i),
        values=(ds.name, ds.workspace_name or "", ds.configured_by or "")
    )
```

**Key Points:**
- Use `show="tree headings"` to enable tree column (#0) for icons
- Configure column #0 as narrow (24px) with `stretch=False`
- Pass icon via `image=` parameter in `insert()`
- Light SVG variants (white fill) for dark mode, dark variants for light mode

---

## Combobox Patterns

### ThemedCombobox (Recommended for Scrollable Dropdowns)

**Use `ThemedCombobox` instead of `ttk.Combobox` when you need modern scrollbar styling in dropdowns.**

The native `ttk.Combobox` uses OS-native scrollbars in its dropdown list, which cannot be styled with ThemedScrollbar. The `ThemedCombobox` widget provides a custom dropdown popup with `ThemedScrollbar` for consistent modern appearance.

```python
from core.widgets import ThemedCombobox

# Create themed combobox with modern scrollbar
combo = ThemedCombobox(
    parent_frame,
    textvariable=my_var,
    values=["Option 1", "Option 2", "Option 3"],
    state="readonly",
    width=25,
    font=("Segoe UI", 9),
    theme_manager=self._theme_manager
)
combo.pack(side=tk.LEFT, padx=5)

# Bind selection event (same as ttk.Combobox)
combo.bind("<<ComboboxSelected>>", self._on_selection_changed)

# Update values dynamically
combo['values'] = ["New Option 1", "New Option 2"]
combo.set("New Option 1")
```

**Key Features:**
- Uses `ThemedScrollbar` for dropdown list (auto-hides when not needed)
- Full dark/light theme support with automatic color updates
- API compatible with `ttk.Combobox` (`.get()`, `.set()`, `.current()`, bracket notation)
- Hover effects and proper focus handling
- Generates `<<ComboboxSelected>>` event for compatibility

**When to Use:**
- Dialogs with many dropdown options (e.g., Edit Field Categories)
- Any combobox where consistent modern scrollbar styling is important
- Bulk edit dropdowns and filter selectors

**Location:** `core/widgets/common.py`

### Standard Combobox Styling

Theme-aware combobox styling with proper border colors and cursor visibility:

```python
def _configure_combobox_styles(self, colors: dict):
    """Configure combobox styles - flat, clean, theme-consistent."""
    # Border color: blue in dark mode, teal in light mode
    border_color = colors['button_primary']

    self._style.configure('TCombobox',
        fieldbackground=colors['card_surface'],
        background=colors['card_surface'],
        foreground=colors['text_primary'],
        arrowcolor=colors['text_secondary'],
        bordercolor=border_color,
        lightcolor=border_color,
        darkcolor=border_color,
        insertbackground=colors['text_primary'],  # Cursor visibility
        relief='flat',
        padding=(6, 3)
    )

    self._style.map('TCombobox',
        fieldbackground=[
            ('readonly', colors['card_surface']),
            ('disabled', colors['background'])
        ],
        foreground=[
            ('readonly', colors['text_primary']),
            ('!disabled', colors['text_primary'])
        ],
        bordercolor=[
            ('focus', colors['primary']),
            ('!focus', border_color)
        ]
    )
```

### Dropdown Listbox Styling

Remove white corners and style the dropdown list:

```python
if self._root:
    self._root.option_add('*TCombobox*Listbox.background', colors['card_surface'])
    self._root.option_add('*TCombobox*Listbox.foreground', colors['text_primary'])
    self._root.option_add('*TCombobox*Listbox.selectBackground', colors['button_primary'])
    self._root.option_add('*TCombobox*Listbox.selectForeground', '#ffffff')
    self._root.option_add('*TCombobox*Listbox.font', ('Segoe UI', 9))
    # Remove white corner by eliminating highlight border
    self._root.option_add('*TCombobox*Listbox.relief', 'flat')
    self._root.option_add('*TCombobox*Listbox.borderWidth', 0)
    self._root.option_add('*TCombobox*Listbox.highlightThickness', 0)
    self._root.option_add('*TCombobox*Listbox.highlightBackground', colors['card_surface'])
    self._root.option_add('*TCombobox*Listbox.highlightColor', colors['card_surface'])
```

### Key Styling Properties

| Property | Purpose | Value |
|----------|---------|-------|
| `bordercolor` | Border color | `button_primary` (blue dark / teal light) |
| `insertbackground` | Cursor color | `text_primary` (prevents black cursor in dark mode) |
| `highlightThickness` | Corner highlight | `0` (removes white corner pixel) |
| `highlightBackground` | Corner color | `card_surface` (matches listbox bg) |

### Custom Combobox Style

For tool-specific combobox styling:

```python
style = ttk.Style()
style.configure('CloudDialog.TCombobox',
    fieldbackground=colors['card_surface'],
    background=colors['card_surface'],
    foreground=colors['text_primary'],
    arrowcolor=colors['text_secondary'],
    bordercolor=colors['button_primary'],
    lightcolor=colors['button_primary'],
    darkcolor=colors['button_primary'],
    insertbackground=colors['text_primary'],
    relief='flat',
    padding=(6, 3)
)
```

### Custom Dropdown Arrow Direction

For custom dropdown implementations (not standard ttk.Combobox), the dropdown arrow should indicate the current state:
- **Closed state**: Arrow points DOWN (`\u25BC`)
- **Open state**: Arrow points UP (`\u25B2`)

This provides clear visual feedback to users about whether clicking will open or close the dropdown.

```python
def _show_dropdown(self):
    """Show the dropdown popup."""
    # Change arrow to point up (indicates dropdown is open)
    self._dropdown_btn.configure(text="\u25B2")

    # ... create and show popup ...

def _close_dropdown(self):
    """Close the dropdown popup."""
    # Change arrow back to point down (indicates dropdown is closed)
    self._dropdown_btn.configure(text="\u25BC")

    # ... destroy popup ...
```

**Key Points:**
- Apply to ALL custom dropdown implementations (not just perspective dropdown)
- Arrow reversal provides intuitive open/close indication
- Use Unicode arrow characters: `\u25BC` (down) and `\u25B2` (up)

### IMPORTANT: Avoid `state='readonly'` on Windows

**Problem:** Using `state='readonly'` on ttk.Combobox causes 1px white corner artifacts on Windows due to theme border rendering issues.

**Solution:** Use `state='normal'` combined with key binding to block keyboard input:

```python
# WRONG - causes visual artifacts on Windows
self.combo.config(state="readonly")

# CORRECT - prevents typing while avoiding visual issues
self.combo.config(state="normal")
self.combo.bind('<Key>', lambda e: 'break')  # Block all keyboard input
```

**When to apply:**
- Any combobox that should be selectable but not editable
- Filter dropdowns, selection pickers, etc.

---

## Progress Log & Analysis Summary

### Split Log Section Pattern

Standard pattern with Analysis Summary on left, Progress Log on right:

```
+--[ ANALYSIS & PROGRESS ]--------------------+
|                         |                   |
|  Analysis Summary       |  Progress Log     |
|  (NO border)            |  (HAS border)     |
|                         |                   |
+-------------------------+-------------------+
```

### Critical Design Rules

| Panel | Border | highlightthickness |
|-------|--------|-------------------|
| LEFT (Summary/Details) | NO border | `0` |
| RIGHT (Log) | HAS border | `1` |

### Using Base Class Method

```python
# Use base class method for split log
log_header = self._create_section_labelwidget("Analysis & Progress", "bar-chart")
self.log_components = self.create_split_log_section(
    parent,
    title="Analysis & Progress",
    labelwidget=log_header
)
self.log_components['frame'].pack(fill=tk.BOTH, expand=True, pady=(0, 15))

# Access components
self.summary_frame = self.log_components['summary_frame']
self.log_text = self.log_components['log_text']
```

### Welcome Message Format

Tools with a Progress Log should display a consistent welcome message on load. Override `_show_welcome_message()` in your tool tab:

**Standard Format:**
```
{emoji} Welcome to {Tool Name}!
============================================================
This tool helps you {brief description}:
• {Feature 1}
• {Feature 2}
• {Feature 3-4 max bullets}

{emoji} {Start instruction}
{emoji} {Backup/safety tip or secondary instruction}
```

**Implementation:**
```python
def _show_welcome_message(self):
    """Show welcome message in log"""
    self.log_message("🔧 Welcome to My Tool!")
    self.log_message("=" * 60)
    self.log_message("This tool helps you do X and Y:")
    self.log_message("• Feature one description")
    self.log_message("• Feature two description")
    self.log_message("• Feature three description")
    self.log_message("")
    self.log_message("📂 Start by selecting your file and clicking 'ANALYZE'")
    self.log_message("💡 Consider backing up your file before making changes")
```

**Guidelines:**
- Welcome line: Use a relevant emoji + "Welcome to {Tool Name}!"
- Divider: Use `"=" * 60` immediately after welcome line
- Description: Brief intro followed by 3-4 bullet points maximum
- Keep it concise - avoid causing scroll on initial load
- End with 1-2 instruction lines using emojis (📂, 💡, ⚠️)
- Use bullet points (•) not dashes or asterisks

**CRITICAL: MRO Pitfall with Mixins**

If your tool uses multiple mixins AND inherits from `BaseToolTab`, be aware that `BaseToolTab` has a default `_show_welcome_message()` method. Due to Python's Method Resolution Order (MRO), `BaseToolTab`'s version will be found FIRST if it appears before your mixin in the inheritance list.

```python
# PROBLEM: BaseToolTab appears before HelpersMixin in MRO
class AdvancedCopyTab(BaseToolTab, FileInputMixin, HelpersMixin, ...):
    def setup_ui(self):
        self._show_welcome_message()  # WRONG: Calls BaseToolTab's 2-line version!

# SOLUTION: Explicitly call the mixin's method
class AdvancedCopyTab(BaseToolTab, FileInputMixin, HelpersMixin, ...):
    def setup_ui(self):
        HelpersMixin._show_welcome_message(self)  # CORRECT: Calls your 8-line version
```

**Symptom:** Welcome message only shows 2 lines (title + divider) instead of full message.

---

## Error Messages

### Centralized Error Constants

Use the `ErrorMessages` class from `core/constants.py` for consistent error dialogs:

```python
from core.constants import ErrorMessages

# Dialog titles
messagebox.showerror(ErrorMessages.INVALID_POSITION, ErrorMessages.INVALID_NUMBER)
messagebox.showinfo(ErrorMessages.SUCCESS, ErrorMessages.OPERATION_COMPLETE)

# Messages with dynamic paths
ThemedMessageBox.showerror(
    parent,
    ErrorMessages.FILE_NOT_FOUND,
    ErrorMessages.FILE_DOES_NOT_EXIST.format(path=file_path)
)
```

### Available Constants

**Dialog Titles:**
| Constant | Value |
|----------|-------|
| `NO_SELECTION` | "No Selection" |
| `INVALID_INPUT` | "Invalid Input" |
| `INVALID_POSITION` | "Invalid Position" |
| `FILE_NOT_FOUND` | "File Not Found" |
| `SUCCESS` | "Success" |

**Message Bodies:**
| Constant | Value |
|----------|-------|
| `NO_FILE_SELECTED` | "Please select a file first." |
| `NO_ITEM_SELECTED` | "Please select an item first." |
| `INVALID_NUMBER` | "Please enter a valid number." |
| `PBIP_NOT_FOUND` | "PBIP file not found: {path}" |
| `OPERATION_COMPLETE` | "Operation completed successfully." |

**Benefits:**
- Consistent wording across all tools
- Single source of truth for error text
- Easy to update messages globally
- Prevents typos in repeated strings

---

### Sub-header Styling

For panel titles like "Finding Details", "Scan Log":
- Font: `('Segoe UI Semibold', 11)` - NOT bold, uses Semibold weight
- Icon size: 16px
- Background: `colors['background']` (main bg), NOT `section_bg`
- Text color: `colors['title_color']`

### Text Widget Background Color

**CRITICAL:** Use `colors['section_bg']` for text widget backgrounds:

```python
# In widget creation
text_bg = colors['section_bg']  # Use section_bg, NOT surface
self.details_text = ModernScrolledText(
    parent, bg=text_bg, fg=colors['text_primary'],
    highlightthickness=0,  # NO border for left panel
    theme_manager=self._theme_manager
)

# In on_theme_changed
text_bg = colors['section_bg']
self.details_text.configure(bg=text_bg, fg=colors['text_primary'])
```

---

## Card Patterns

### CleanupOpportunityCard (Report Cleanup)

Visual card showing cleanup categories:

**States:**
- Initial: Shows "--" count, toggle hidden
- Has opportunities (count > 0): Count shown, toggle enabled, normal background
- No opportunities (count = 0): "0" in muted color, toggle hidden, greyed background
- Selected: 2px border (vs 1px normal)

**Structure:**
1. Outer frame: `card_surface` bg with `border` highlight (1px)
2. Inner padding: `padx=25, pady=15`
3. Icon area: SVG image or emoji fallback
4. Title/Subtitle labels
5. Toggle frame (hidden until analysis)
6. Count/Size labels

### AccessibilityCheckCard (Accessibility Checker)

Category cards with issue counts:

**Color Coding:**
- Green: 0 issues
- Yellow: warnings
- Red: errors

---

## Dialog Patterns

### Help Dialog Structure

```
+--[ TOOL ICON ]  TOOL NAME  Help  -----------[X]+
|                                                 |
|  +--[ WARNING ]--------------------------------+|
|  | Important disclaimers and requirements      ||
|  +---------------------------------------------+|
|                                                 |
|  +-- Section 1 -----+  +-- Section 2 ----------+|
|  | Content          |  | Content               ||
|  +------------------+  +------------------------+|
+-------------------------------------------------+
```

### Two-Column Vertical Alignment

Use a **SINGLE grid frame** with multiple rows:

```python
# CORRECT: Single grid with 2 rows ensures vertical alignment
sections_frame = tk.Frame(container, bg=help_bg)
sections_frame.pack(fill=tk.BOTH, expand=True)
sections_frame.columnconfigure(0, weight=1)
sections_frame.columnconfigure(1, weight=1)

# Row 0: First pair of sections
left_top.grid(row=0, column=0, sticky='nwe', padx=(0, 10), pady=(0, 15))
right_top.grid(row=0, column=1, sticky='nwe', padx=(10, 0), pady=(0, 15))

# Row 1: Second pair of sections
left_bottom.grid(row=1, column=0, sticky='nwe', padx=(0, 10))
right_bottom.grid(row=1, column=1, sticky='nwe', padx=(10, 0))
```

### Text Formatting Rules

1. **NO BULLETS**: Do not use bullet points in help dialog text
2. **Use `padx` for Indentation**: NOT string spaces
   - Wrong: `text=f"   {item}"` (inconsistent wrap)
   - Correct: `text=item, padx=(12, 0)` (consistent indentation)

### Warning Box

```python
warnings = [
    "This tool ONLY works with PBIP enhanced report format",
    "NOT officially supported by Microsoft - use at your own discretion",
    "Always keep backups before optimization"
]

for warning in warnings:
    tk.Label(warning_container, text=warning, font=('Segoe UI', 10),
             bg=colors['warning_bg'], fg=colors['warning_text']
    ).pack(anchor=tk.W, padx=(12, 0), pady=1)
```

### Dialog Sizing

| Dialog | Width | Height | Use For |
|--------|-------|--------|---------|
| Compact | 720 | 530-560 | Simple tools (Accessibility Checker) |
| Standard | 670-800 | 550-630 | Medium tools (Report Merger) |
| Medium | 1000 | 720-855 | Standard tools (Layout Optimizer, Sensitivity Scanner) |
| Large | 950 | 950 | Complex tools (Advanced Copy) |

### CRITICAL: Dialog Geometry

Dialog geometry is often set TWICE. Both must be updated:

```python
# Location 1: Initial creation
dialog.geometry("720x640")

# ... dialog content setup ...

# Location 2: Centering logic (THIS IS THE ONE ACTUALLY USED)
dialog.update_idletasks()
x = parent_window.winfo_rootx() + (parent_window.winfo_width() - 720) // 2
y = parent_window.winfo_rooty() + (parent_window.winfo_height() - 640) // 2
dialog.geometry(f"720x640+{x}+{y}")
```

### Popup Background Colors

**CRITICAL:** In popups/dialogs, ALL elements use `colors['background']`, NOT `section_bg`:

```python
# CORRECT - All popup elements use background
popup_bg = colors['background']
frame = tk.Frame(popup, bg=popup_bg)
label = tk.Label(popup, text="Label", bg=popup_bg, fg=colors['text_primary'])
```

### ttk.Label vs tk.Label

**Critical:** `ttk.Label` does NOT accept `bg=` parameter and foreground colors may be overridden by ttk styles.

**Problems with ttk.Label:**
1. Background: Inherits from ttk styles, can't be directly controlled
2. Foreground: Style system can override widget-specific settings during theme changes

**Solution:** Use `tk.Label` whenever you need reliable color control:

```python
# WRONG - ttk.Label can't have explicit background control
label = ttk.Label(popup, text="My Label", font=('Segoe UI', 9))

# CORRECT - tk.Label allows explicit control
label = tk.Label(popup, text="My Label", bg=popup_bg, fg=colors['text_primary'],
                font=('Segoe UI', 9))
```

### Custom Themed Confirmation Dialogs

For custom Yes/No dialogs that need RoundedButtons instead of ThemedMessageBox, use this pattern:

```python
def _show_custom_confirm_dialog(self, title: str, message: str) -> bool:
    """Show a themed Yes/No confirmation dialog with RoundedButtons."""
    from pathlib import Path

    colors = self._theme_manager.colors

    # Create modal dialog
    dialog = tk.Toplevel(self.frame)
    dialog.title(title)
    dialog.geometry("450x265")  # Width x Height
    dialog.resizable(False, False)
    dialog.transient(self.frame.winfo_toplevel())
    dialog.grab_set()
    dialog.configure(bg=colors['background'])

    # Center on parent
    dialog.update_idletasks()
    parent = self.frame.winfo_toplevel()
    x = parent.winfo_x() + (parent.winfo_width() - 450) // 2
    y = parent.winfo_y() + (parent.winfo_height() - 265) // 2
    dialog.geometry(f"+{x}+{y}")

    # Set AE favicon icon
    try:
        base_dir = Path(__file__).parent.parent.parent.parent
        icon_path = base_dir / 'assets' / 'favicon.ico'
        if icon_path.exists():
            dialog.iconbitmap(str(icon_path))
    except Exception:
        pass

    # Set dark/light title bar on Windows
    try:
        import ctypes
        dialog.update()
        hwnd = ctypes.windll.user32.GetParent(dialog.winfo_id())
        DWMWA_USE_IMMERSIVE_DARK_MODE = 20
        value = ctypes.c_int(1 if is_dark else 0)
        ctypes.windll.dwmapi.DwmSetWindowAttribute(
            hwnd, DWMWA_USE_IMMERSIVE_DARK_MODE,
            ctypes.byref(value), ctypes.sizeof(value)
        )
    except Exception:
        pass

    # Result variable
    result_var = tk.BooleanVar(value=False)

    # Content frame
    content = tk.Frame(dialog, bg=colors['background'], padx=20, pady=15)
    content.pack(fill=tk.BOTH, expand=True)

    # Message label
    tk.Label(
        content,
        text=message,
        font=("Segoe UI", 10),
        bg=colors['background'],
        fg=colors['text_primary'],
        justify='left',
        wraplength=400
    ).pack(anchor='w', pady=(0, 16))

    # Buttons - centered using inner frame
    button_frame = tk.Frame(content, bg=colors['background'])
    button_frame.pack(fill=tk.X)

    button_inner = tk.Frame(button_frame, bg=colors['background'])
    button_inner.pack(anchor='center')

    def on_yes():
        result_var.set(True)
        dialog.destroy()

    def on_no():
        result_var.set(False)
        dialog.destroy()

    RoundedButton(
        button_inner,
        text="Yes",
        command=on_yes,
        bg=colors['button_primary'],
        hover_bg=colors['button_primary_hover'],
        pressed_bg=colors.get('button_primary_pressed', colors['button_primary_hover']),
        fg='#ffffff',
        height=32, radius=5,
        font=('Segoe UI', 9, 'bold'),
        canvas_bg=colors['background'],
        width=90
    ).pack(side=tk.LEFT, padx=(0, 10))

    RoundedButton(
        button_inner,
        text="No",
        command=on_no,
        bg=colors['button_secondary'],
        hover_bg=colors['button_secondary_hover'],
        pressed_bg=colors['button_secondary_pressed'],
        fg=colors['text_primary'],
        height=32, radius=5,
        font=('Segoe UI', 9),
        canvas_bg=colors['background'],
        width=90
    ).pack(side=tk.LEFT)

    # Wait for dialog to close
    dialog.wait_window()
    return result_var.get()
```

**Key Pattern Elements:**
1. **Modal behavior**: `transient()` and `grab_set()` make dialog modal
2. **Centering**: Calculate position relative to parent window
3. **AE favicon**: Load from `assets/favicon.ico` for consistent branding
4. **Dark title bar**: Use `DwmSetWindowAttribute` with `DWMWA_USE_IMMERSIVE_DARK_MODE` (20) for Windows theme-aware title bar
5. **Result variable**: Use `tk.BooleanVar` to capture user's choice
6. **Centered buttons**: Use inner frame with `pack(anchor='center')`
7. **Themed colors**: All elements use colors from `_theme_manager.colors`
8. **canvas_bg**: Buttons use `colors['background']` to match dialog background

**When to Use:**
- Custom confirmation dialogs with specific messaging
- Dialogs needing RoundedButtons for visual consistency
- When ThemedMessageBox.askyesno() is too generic

---

## Layout Best Practices

### Button Positioning

```python
# Center action buttons - pack BOTTOM first for visibility
btn_frame = ttk.Frame(parent)
btn_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(15, 0))

inner_frame = ttk.Frame(btn_frame)
inner_frame.pack(anchor=tk.CENTER)

primary_btn.pack(side=tk.LEFT, padx=2)
secondary_btn.pack(side=tk.LEFT, padx=2)
```

### Section Spacing

- **Between sections:** 15-20px vertical padding
- **Inside sections:** 10-12px padding
- **Between related elements:** 5-8px
- **Bottom content to dialog buttons:** 15px minimum

### Avoiding Visible Container Backgrounds

**Problem Pattern:**
```python
# WRONG - Creates visible off-white rectangle
container = tk.Frame(parent, bg=option_bg, padx=10, pady=6)
container.pack(side=tk.LEFT)
```

**Correct Pattern:**
```python
# CORRECT - Elements blend seamlessly
section_bg = colors.get('section_bg', colors['background'])
mode_label = tk.Label(parent, text="Mode:", bg=section_bg, fg=colors['text_primary'])
mode_label.pack(side=tk.LEFT, padx=(15, 10))
```

### Avoiding Dark Divider Lines Between Sections

**Problem:** Using `padx` or `pady` on grid/pack creates gaps that show the parent frame's background color, causing visible dark divider lines in dark mode.

```python
# WRONG - padx creates gap showing parent's dark background
left_frame.grid(row=0, column=0, padx=(0, 8))   # 8px gap shows dark bg
right_frame.grid(row=0, column=1, padx=(8, 0))  # Creates dark divider line
```

**Solutions:**

1. **Remove gaps entirely** - LabelFrame borders provide sufficient visual separation:
```python
# CORRECT - No gaps, borders provide separation
left_frame.grid(row=0, column=0, sticky='nsew')
right_frame.grid(row=0, column=1, sticky='nsew')
```

2. **Use internal padding** - Add padding inside the child widgets instead:
```python
# CORRECT - Padding inside LabelFrame, not as gap
config_frame = ttk.LabelFrame(parent, padding="12")
config_frame.pack(fill=tk.BOTH, expand=True)  # No padx here
```

### Multi-Column Layouts with Stable Widths

For side-by-side columns that maintain ratio during theme changes:

```python
# Use tk.Frame (not ttk.Frame) for stable geometry during theme changes
middle_content = tk.Frame(self.frame, bg=colors['background'])
middle_content.pack(fill=tk.BOTH, expand=True)
middle_content.columnconfigure(0, weight=1)  # 1/3 width
middle_content.columnconfigure(1, weight=2)  # 2/3 width
middle_content.rowconfigure(0, weight=1)

# Store reference for theme updates
self._middle_content = middle_content

# Lock widths after initial layout to prevent theme toggle wiggle
self.frame.after_idle(self._lock_column_widths)
```

**Width Locking Pattern - CRITICAL: Use weight=0**

When locking column widths, you MUST set `weight=0` to disable proportional expansion. Using `minsize` alone does NOT prevent width drift because weights still control space distribution.

```python
def _lock_column_widths(self):
    """Lock column widths after initial layout."""
    if self._columns_width_locked:
        return
    try:
        left_width = self._left_column.winfo_width()
        right_width = self._right_column.winfo_width()
        if left_width > 10 and right_width > 10:
            # CRITICAL: weight=0 disables proportional expansion
            # minsize alone is NOT sufficient - weights override it!
            self._middle_content.columnconfigure(0, weight=0, minsize=left_width)
            self._middle_content.columnconfigure(1, weight=0, minsize=right_width)
            self._locked_left_width = left_width
            self._locked_right_width = right_width
            self._columns_width_locked = True
    except Exception:
        pass
```

**Common Mistake:**
```python
# WRONG - weights still control expansion, columns will drift
self._middle_content.columnconfigure(0, weight=1, minsize=left_width)
self._middle_content.columnconfigure(1, weight=2, minsize=right_width)

# CORRECT - weight=0 truly locks the widths
self._middle_content.columnconfigure(0, weight=0, minsize=left_width)
self._middle_content.columnconfigure(1, weight=0, minsize=right_width)
```

### Preventing Vertical Expansion

**Problem:** Treeviews and their containers grow vertically when data is loaded, causing section heights to change.

**Root Cause:** Multiple `expand=True` directives in the widget hierarchy allow vertical growth.

**Solution:** Use `expand=False` on tree-related pack() calls:

```python
# WRONG - allows vertical growth when data is added
tree_frame.pack(fill=tk.BOTH, expand=True)
tree_container.pack(fill=tk.BOTH, expand=True)

# CORRECT - prevents vertical expansion
tree_frame.pack(fill=tk.BOTH, expand=False)
tree_container.pack(fill=tk.BOTH, expand=False)
```

**Note:** Use `fill=tk.BOTH` to stretch horizontally, but `expand=False` to prevent vertical growth.

### Progress Bar Positioning with pack(before=)

**CRITICAL:** When using `pack(before=widget)`, you MUST specify `side=` to match the target widget's side:

```python
# WRONG - Defaults to side=TOP, progress bar appears at top of window
self.progress_frame.pack(before=self.button_frame, fill=tk.X)

# CORRECT - Matches button_frame's side=BOTTOM
self.progress_frame.pack(side=tk.BOTTOM, before=self.button_frame, fill=tk.X)
```

---

## Icon Loading & Theme Updates

### Icon Loading Methods Overview

There are several icon loading patterns in the codebase, each serving a specific purpose:

| Method | Location | Purpose | Use Case |
|--------|----------|---------|----------|
| `_load_icon_for_button()` | BaseToolTab | Button icons | Standard square icons from assets/Tool Icons |
| `_load_checkbox_icons()` | Per-tool | Checkbox states | Theme-aware checked/unchecked icons |
| `_load_radio_icons()` | Per-tool | Radio states | Theme-aware selected/unselected icons |
| `_load_svg_icon()` | LabeledToggle | Toggle states | On/off toggle switch icons |

### When to Use Each

- **Button icons**: Use `_load_icon_for_button()` from BaseToolTab for any icon that appears in a RoundedButton
- **Checkbox/Radio icons**: Implement `_load_checkbox_icons()` / `_load_radio_icons()` locally when you need theme-aware state icons
- **Custom widgets**: Widgets like LabeledToggle and LabeledRadioGroup have their own icon loading with class-level caching for efficiency

### Standard Pattern

```python
def _load_button_icons(self):
    if not hasattr(self, '_button_icons'):
        self._button_icons = {}

    # 16px icons for buttons and sections
    icon_names = ["Power-BI", "magnifying-glass", "reset", "folder", "save", "question"]
    for name in icon_names:
        icon = self._load_icon_for_button(name, size=16)
        if icon:
            self._button_icons[name] = icon

    # Load theme-aware checkbox/radio icons
    self._load_checkbox_icons()
    self._load_radio_icons()
```

### CRITICAL: File Naming

Icon names must match **exact SVG filename** (without extension):

```python
# Icons loaded by filename (minus .svg extension)
csv_icon = self._button_icons.get('csv-file')  # CORRECT - matches csv-file.svg
csv_icon = self._button_icons.get('csv')       # WRONG - file is csv-file.svg
```

### Icon Reference Storage

Prevent garbage collection during theme changes:

```python
# Store on widget
icon_label._icon_ref = icon

# Store on parent frame
row_data['frame']._icon_ref = icon

# Store in global list
if not hasattr(self, '_icon_refs'):
    self._icon_refs = []
self._icon_refs.append(icon)
```

### Consistent on_theme_changed() Structure

```python
def on_theme_changed(self, theme: str):
    colors = self._theme_manager.colors
    section_bg = colors.get('section_bg', colors['background'])
    content_bg = colors['background']

    # 1. Update section header widgets
    for header_frame, icon_label, text_label in self._section_header_widgets:
        try:
            header_frame.configure(bg=section_bg)
            if icon_label:
                icon_label.configure(bg=section_bg)
            text_label.configure(bg=section_bg, fg=colors['title_color'])
        except Exception:
            pass

    # 2. Reload themed icons
    self._load_checkbox_icons()
    self._load_radio_icons()

    # 3. Update treeview items with new icons
    # ... (tool-specific)

    # 4. Update button colors AND canvas_bg
    for btn in self._primary_buttons:
        try:
            btn.update_colors(
                bg=colors['button_primary'],
                hover_bg=colors['button_primary_hover'],
                pressed_bg=colors['button_primary_pressed'],
                fg='#ffffff'
            )
            btn.update_canvas_bg(section_bg)
        except Exception:
            pass

    # 5. Update text widget backgrounds
    if hasattr(self, 'log_text'):
        try:
            self.log_text.configure(bg=colors['section_bg'])
        except Exception:
            pass
```

### Register/Unregister Pattern

```python
def __init__(self, parent):
    self._theme_manager = get_theme_manager()
    self._theme_manager.register_callback(self.on_theme_changed)

def _on_close(self):
    """Cleanup on close - always unregister callbacks."""
    try:
        self._theme_manager.unregister_callback(self.on_theme_changed)
    except (ValueError, AttributeError):
        pass
    self.window.destroy()
```

### CRITICAL: Updating Checkbox Widgets After Theme Change

**Problem:** When theme changes, icons are reloaded into memory but existing checkbox widgets still display the OLD icons until explicitly updated.

**Symptom:** Partial checkbox icons (for "Select All" rows) show as empty/unchecked after theme toggle, but correct themselves when a child is toggled.

**Solution:** Always call `_update_checkboxes()` AFTER reloading icons in `on_theme_changed()`:

```python
def on_theme_changed(self):
    """Handle theme change - reload icons and update colors."""
    # 1. Reload checkbox icons for new theme
    self._load_checkbox_icons()

    # 2. CRITICAL: Update displayed checkboxes with new icons
    self._update_checkboxes()

    # 3. Rest of theme update...
```

**Also for Dropdown/Popup Content:** When building hierarchical checkbox content (like filter dropdowns), the initial icon selection often only handles binary checked/unchecked state. Call `_update_checkboxes()` after building to show partial state correctly:

```python
def _open_dropdown(self):
    # ... build dropdown content with checkboxes ...

    for severity in severity_order:
        # Initial icon only checks checked/unchecked
        is_checked = severity in self._selected_severities
        icon = self._checkbox_icons.get('checked' if is_checked else 'unchecked')
        # ... create checkbox widget ...

    # CRITICAL: Update to show partial icons where applicable
    self._update_checkboxes()

    # ... rest of dropdown setup ...
```

### Theme Update Methods for Complex Selection States

For UI with multiple selection modes or text states that need theme-aware colors, create dedicated update methods rather than inline logic:

```python
def _update_bookmark_selection_theme(self, colors):
    """Update bookmark selection UI for theme changes."""
    is_dark = self._theme_manager.current_theme == 'dark'

    # Update target page checkboxes - handle partial state for Select All
    if hasattr(self, '_target_page_widgets'):
        for row_data in self._target_page_widgets:
            is_select_all = row_data.get('is_select_all', False)
            is_all_selected = self._is_all_targets_selected()
            is_partial = self._is_partial_targets_selected()

            # CRITICAL: Handle all three states for Select All row
            if is_select_all and is_partial:
                checkbox_icon = self._button_icons.get('box-partial-dark' if is_dark else 'box-partial')
            elif is_select_all and is_all_selected:
                checkbox_icon = self._button_icons.get('box-checked-dark' if is_dark else 'box-checked')
            else:
                # ... normal checked/unchecked logic

    # Use dedicated method for mode text colors
    if hasattr(self, '_update_bookmark_mode_colors'):
        self._update_bookmark_mode_colors()
```

**Key Principle:** Consolidate related color/state logic into reusable methods that can be called from:
1. Initial widget creation
2. Selection change handlers
3. Theme change handlers

### CRITICAL: Dynamically Created Widget Containers

**Problem:** Widgets created dynamically (e.g., card panels built when a row is selected) are not tracked by the parent component's `on_theme_changed()` method. When the theme changes, these widgets retain their old colors.

**Symptom:** A section like "Connection Details" shows dark mode colors (dark background, light text) while the rest of the UI has switched to light mode.

**Solution:** Re-render dynamically created containers when theme changes by detecting if they exist and rebuilding them:

```python
def on_theme_changed(self, theme: str):
    # ... other theme updates ...

    # Re-render dynamically created card container if visible
    if hasattr(self, '_card_container') and self._card_container:
        try:
            # Get current selection from tree
            if hasattr(self, 'mapping_tree') and self.mapping_tree:
                selection = self.mapping_tree.selection()
                if selection:
                    item_id = selection[0]
                    idx = int(item_id)
                    if idx < len(self.mappings):
                        # Re-render with new theme colors
                        self._update_selected_connection_details(self.mappings[idx])
        except Exception:
            pass
```

**Key Points:**
1. Template widgets like `SplitLogSection` update their base frames automatically
2. But content injected INTO those frames (like card containers) must be manually re-rendered
3. The simplest approach is to re-call the same method that created the dynamic content
4. Always wrap in try/except to prevent theme toggle failures

### CRITICAL: Column Width Stability on Theme Change

**Problem:** Column widths in multi-panel layouts drift/change slightly each time the theme is toggled. This is caused by grid `weight` allowing columns to resize.

**Solution:** Lock column widths using `weight=0` and `minsize`:

```python
def setup_ui(self):
    # Configure column grid
    self.main_frame.columnconfigure(0, weight=1, minsize=250)  # Left panel - flexible
    self.main_frame.columnconfigure(1, weight=0, minsize=350)  # Middle panel - LOCKED
    self.main_frame.columnconfigure(2, weight=0, minsize=400)  # Right panel - LOCKED

def on_theme_changed(self):
    # Re-apply column configuration to prevent drift
    self.main_frame.columnconfigure(0, weight=1, minsize=250)
    self.main_frame.columnconfigure(1, weight=0, minsize=350)
    self.main_frame.columnconfigure(2, weight=0, minsize=400)
```

**Key Points:**
- `weight=0` prevents columns from growing when space is available
- `minsize` sets a fixed minimum (and effective) width
- Re-apply in `on_theme_changed()` to prevent accumulating drift
- Use `weight=1` only for columns that should flex with window resizing

### CRITICAL: Never Override Global ttk Styles from Tools

**Problem:** `style.configure('TFrame', ...)` modifies the style for ALL ttk.Frame widgets application-wide, not just within the calling tool. This causes style "pollution" where one tool's customization affects every other tool.

**Symptom:** Dark mode background colors appear wrong in all tools after visiting a specific tool. The issue persists until theme is toggled.

**Root Cause Example:**
```python
# BAD - This overwrites the global TFrame style for ALL tools
def on_theme_changed(self):
    style = ttk.Style()
    style.configure('TFrame', background=colors['surface'])  # Affects entire app!
```

**Solution:** Always use `theme_manager.py` as the single source of truth for global ttk styles. Individual tools should NEVER call `style.configure()` on base style names.

**If a tool needs a custom appearance:**
```python
# GOOD - Use a named style that only affects widgets explicitly using it
style = ttk.Style()
style.configure('MyTool.TFrame', background=colors['custom_color'])
my_frame.configure(style='MyTool.TFrame')  # Only affects this widget
```

**Global Styles Owned by theme_manager:**
- `TFrame` - base frame background
- `TLabel` - standard label styling
- `TButton` - button defaults
- `Section.TFrame` - section container styling
- `Treeview` - treeview styling
- (See `theme_manager.py` for full list)

**Best Practice:**
1. Tools should work with theme_manager's global styles out of the box
2. Re-apply named styles in `on_theme_changed()` if needed (e.g., `style='Section.TFrame'`)
3. Never create a new `ttk.Style()` instance to modify base style names
4. If custom styling is absolutely needed, prefix with tool name: `'FieldParam.TFrame'`

---

## Tooltips & Message Boxes

### Tooltip Usage

```python
from core.ui_base import Tooltip

# Simple tooltip
Tooltip(my_label, "This explains what this label does")

# Multi-line tooltip
Tooltip(my_button, "Line 1 explanation\nLine 2 with more details")
```

**Positioning:** Tooltips appear 15px right and 20px below mouse cursor.

### ThemedMessageBox

Use instead of native `messagebox` for consistent theming:

```python
from core.ui_base import ThemedMessageBox

parent = self.frame.winfo_toplevel()

ThemedMessageBox.showinfo(parent, "Success", "Operation completed")
ThemedMessageBox.showwarning(parent, "Warning", "This action cannot be undone")
ThemedMessageBox.showerror(parent, "Error", "Failed to save file")

if ThemedMessageBox.askyesno(parent, "Confirm", "Are you sure?"):
    # User clicked Yes
    pass
```

### ThemedInputDialog

Use instead of native `simpledialog.askstring` for consistent theming:

```python
from core.ui_base import ThemedInputDialog

parent = self.frame.winfo_toplevel()

# Basic usage - returns string or None if cancelled
result = ThemedInputDialog.askstring(
    parent,
    "Rename Item",
    "Enter new name:",
    initialvalue="current name"
)

if result:
    # User entered a value and clicked OK
    print(f"New name: {result}")
```

**Parameters:**
- `parent`: Parent window (required)
- `title`: Dialog window title
- `prompt`: Label text shown above the entry field
- `initialvalue`: Pre-populated text in entry field (default: "")
- `min_width`: Minimum dialog width in pixels (default: 300)

**Features:**
- Dark title bar support on Windows
- AE favicon
- Theme-aware colors (background, text, buttons)
- RoundedButton styling for OK/Cancel
- Entry pre-selects initialvalue for easy replacement
- Returns `None` if user cancels or closes dialog

### ThemedContextMenu

Use instead of `tk.Menu` for context menus - `tk.Menu` has intense white borders on Windows that cannot be styled away. ThemedContextMenu uses a custom Toplevel window for full border and styling control.

```python
from core.ui_base import ThemedContextMenu

def _on_right_click(self, event):
    menu = ThemedContextMenu(self, self._theme_manager)

    # Add clickable items
    menu.add_command("Edit Item", self._on_edit)
    menu.add_command("Duplicate", self._on_duplicate)

    # Add visual separator
    menu.add_separator()

    # Add section header (muted, non-clickable)
    menu.add_section_header("Danger Zone")
    menu.add_command("Delete", self._on_delete)

    # Show at cursor position
    menu.show(event.x_root, event.y_root)
```

**API:**
- `add_command(label, command, icon=None)` - Add clickable menu item
- `add_separator()` - Add visual divider line
- `add_section_header(text)` - Add muted, non-clickable section label
- `show(x, y)` - Display menu at screen coordinates
- `close()` - Hide menu (called automatically on item click or outside click)
- `destroy()` - Full cleanup

**Features:**
- 1px themed border (uses `border` color)
- Hover effects on items (uses `card_surface_hover` color)
- Escape key closes menu
- Clicking outside closes menu
- Section headers in muted text (`text_muted` color)

**Color Keys Used:**
- `surface` - Menu background
- `border` - 1px outer border
- `card_surface_hover` - Hover state background
- `text_primary` - Menu item text
- `text_muted` - Section header text

**When to Use:**
- Any visible right-click context menu
- Popup menus triggered by user action
- Use standard `tk.Menu` only for system-level menus (menu bars) where OS styling is expected

---

## Claude Code Prompts

When building a new tool with Claude Code, use these decision prompts:

### Section Type

> "I'm creating a new section. Should this be:
> 1. **Standard Section** - Icon + title, no border, content padding
> 2. **Card Section** - Selectable items with borders
> 3. **Split Section** - Two-column layout (e.g., Summary + Log)"

### Selection Pattern

> "This tool needs user selection. Should I use:
> 1. **Radio Buttons** - Single select with circular icons
> 2. **Checkboxes** - Multi-select with square icons
> 3. **TreeView** - Hierarchical or tabular selection
> 4. **Cards** - Visual selection with toggle switches"

### Analysis Display

> "Displaying analysis results. Should this be:
> 1. **Summary Table** - Grid of metrics (no border)
> 2. **Scrollable Cards** - Multiple result cards
> 3. **Text Log** - Streaming text output
> 4. **Split View** - Summary left, details right"

### Button Layout

> "Adding action buttons. Standard patterns:
> 1. **Single Primary** - Centered main action
> 2. **Primary + Secondary** - Main action + Reset
> 3. **Button Row** - Multiple equal actions
> 4. **Inline Buttons** - Small buttons within content"

---

## Checklist for New Tools

- [ ] Tool has icon in sidebar
- [ ] All sections have Icon + Title headers
- [ ] Section titles use `title_color` (blue dark / teal light)
- [ ] Sections have NO visible borders (use background framing)
- [ ] Help icon in upper-right of main section
- [ ] Primary buttons use `button_primary` colors
- [ ] Secondary/Reset buttons use `button_secondary` colors
- [ ] Buttons have hover and pressed states
- [ ] Disabled buttons have proper disabled styling
- [ ] Radio buttons for single-select with `title_color` for selected text
- [ ] Checkboxes for multi-select with theme-aware icons
- [ ] Analysis Summary has NO border
- [ ] Progress Log has faint border
- [ ] Help dialog has warning box with orange background
- [ ] All colors come from `colors` dict (not hardcoded)
- [ ] All fonts come from `AppConstants.FONTS`
- [ ] Theme change handler updates all custom widgets
- [ ] Icon references stored to prevent garbage collection

