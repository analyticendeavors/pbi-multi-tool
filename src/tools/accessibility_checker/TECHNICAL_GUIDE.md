# Accessibility Checker - Technical Guide

**Tool Version:** 1.0.0
**Last Updated:** January 2026
**Built by:** Reid Havens of Analytic Endeavors

---

## Table of Contents

1. [Overview](#overview)
2. [Architecture](#architecture)
3. [Core Components](#core-components)
4. [WCAG Checks](#wcag-checks)
5. [Color Contrast Analysis](#color-contrast-analysis)
6. [Data Models](#data-models)
7. [Configuration](#configuration)
8. [Usage Guide](#usage-guide)
9. [File Structure](#file-structure)
10. [Extending the Tool](#extending-the-tool)
11. [Troubleshooting](#troubleshooting)

---

## Overview

### What is the Accessibility Checker?

The **Accessibility Checker** is a static analysis tool that scans Power BI reports (PBIP/PBIR format) for accessibility compliance issues based on WCAG 2.1 guidelines. It provides automated detection of common accessibility problems.

### Key Features

- **Tab Order Analysis** - Validates keyboard navigation sequence
- **Alt Text Detection** - Checks for missing or generic alt text on data visuals
- **Color Contrast Checking** - WCAG AA/AAA contrast ratio validation
- **Page Title Validation** - Ensures pages have meaningful names
- **Visual Title Checking** - Verifies data visuals have descriptive titles
- **Data Labels Detection** - Identifies charts missing data labels
- **Bookmark Name Review** - Flags generic bookmark names
- **Hidden Page Warnings** - Alerts about hidden pages with potential content

### What It Does NOT Do

- Does not modify report files (read-only analysis)
- Does not guarantee 100% WCAG compliance (automated checks are a starting point)
- Does not test runtime behavior or dynamic content
- Does not replace manual accessibility audits
- Is NOT certified for compliance validation

---

## Architecture

### Component Overview

```
accessibility_checker/
├── tool.py                    # BaseTool implementation
├── accessibility_analyzer.py  # Core analysis engine
├── accessibility_types.py     # Data models and enums
├── accessibility_config.py    # Configuration management
├── accessibility_ui.py        # UI tab implementation
└── TECHNICAL_GUIDE.md         # This document
```

### Data Flow

```
User selects PBIP/PBIX file
    |
    v
PBIPReader validates PBIR format
    |
    v
AccessibilityAnalyzer scans report
    |
    +-> Scan pages (metadata, visibility)
    +-> Scan visuals (properties, alt text, titles)
    +-> Scan bookmarks (names, groups)
    |
    v
Run accessibility checks (configurable)
    |
    +-> Tab order analysis
    +-> Alt text validation
    +-> Color contrast checking
    +-> Page title validation
    +-> Visual title checking
    +-> Data label detection
    +-> Bookmark name review
    +-> Hidden page detection
    |
    v
AccessibilityAnalysisResult returned
    |
    v
UI displays issues by category with severity
```

---

## Core Components

### 1. AccessibilityAnalyzer (accessibility_analyzer.py)

**Role**: Main analysis engine that orchestrates all accessibility checks.

**Key Methods**:
```python
analyze_pbip_report(pbip_path: str) -> AccessibilityAnalysisResult
    # Main entry point - runs all enabled checks

_scan_pages(report_dir: Path) -> List[PageInfo]
    # Extract page metadata (names, visibility, visual counts)

_scan_all_visuals(report_dir: Path) -> List[VisualInfo]
    # Extract visual properties (alt text, titles, tab order, etc.)

_scan_bookmarks(report_dir: Path) -> List[BookmarkInfo]
    # Extract bookmark information

_get_bookmark_toggled_visuals(report_dir: Path) -> Set[str]
    # Find visuals that can be toggled visible by bookmarks
```

**Progress Tracking**:
- Accepts optional `progress_callback` for UI progress bar updates
- Reports progress at each analysis phase (5%, 10%, 20%, etc.)

### 2. AccessibilityCheckConfig (accessibility_config.py)

**Role**: Manages check configuration and user preferences.

**Key Features**:
- Enable/disable individual checks
- Configure contrast level (AA, AA_large, AAA)
- Flag AAA failures when checking at AA level
- Persistent configuration storage

### 3. Data Models (accessibility_types.py)

**Key Data Classes**:

| Class | Purpose |
|-------|---------|
| `AccessibilityIssue` | Single accessibility issue with severity, description, recommendation |
| `TabOrderInfo` | Tab order data for a visual element |
| `ColorContrastResult` | Contrast analysis result with pass/fail status |
| `VisualInfo` | Visual element metadata (alt text, title, type) |
| `PageInfo` | Page metadata (name, visibility, visual count) |
| `BookmarkInfo` | Bookmark metadata (name, is generic) |
| `AccessibilityAnalysisResult` | Complete analysis result container |

**Severity Levels**:
```python
class AccessibilitySeverity(Enum):
    ERROR = "error"       # Critical - fails WCAG compliance
    WARNING = "warning"   # Should fix for better accessibility
    INFO = "info"         # Suggestion/best practice
```

---

## WCAG Checks

### 1. Tab Order (WCAG 2.4.3 Focus Order)

**What It Checks**:
- Missing tab order on visuals (-1 value)
- Duplicate tab order values on same page
- Tab order vs visual reading order alignment

**Detection Logic**:
```python
# Missing tab order
if visual.tab_order == -1:
    # Flag as WARNING

# Duplicate tab orders
if tab_order_values.count(value) > 1:
    # Flag as WARNING

# Reading order mismatch (top-to-bottom, left-to-right)
sorted_by_position = sorted(visuals, key=(y_row, x))
if mismatches > total // 2:
    # Flag as INFO
```

### 2. Alt Text (WCAG 1.1.1 Non-text Content)

**What It Checks**:
- Data visuals missing alt text
- Generic/non-descriptive alt text
- Alt text length (>125 chars flagged as INFO)

**Data Visual Types** requiring alt text:
- Charts (bar, line, pie, etc.)
- Tables and matrices
- Cards and KPIs
- Maps
- AI visuals (Q&A, Key Influencers)

**Hidden Visual Handling**:
- Skips visuals marked as hidden
- Includes visuals that can be toggled visible by bookmarks

### 3. Color Contrast (WCAG 1.4.3/1.4.6)

See [Color Contrast Analysis](#color-contrast-analysis) section.

### 4. Page Titles (WCAG 2.4.2 Page Titled)

**What It Checks**:
- Missing page display names
- Generic titles (e.g., "Page 1", "Sheet 2", "Untitled")

**Generic Patterns Detected**:
```python
GENERIC_PAGE_TITLE_PATTERNS = [
    r"^page\s*\d*$",
    r"^sheet\s*\d*$",
    r"^untitled$",
    r"^new\s*page$",
]
```

### 5. Visual Titles (WCAG 2.4.6 Headings and Labels)

**What It Checks**:
- Data visuals without displayed titles
- Skips hidden visuals (unless bookmark-toggled)

### 6. Data Labels (WCAG 1.4.1 Use of Color)

**What It Checks**:
- Charts without data labels enabled
- Higher severity for charts where labels are critical (pie, donut, treemap)

**Label-Critical Chart Types**:
- Pie and donut charts
- Treemaps
- Bar and column charts (stacked variants)

### 7. Bookmark Names (WCAG 2.4.6 Headings and Labels)

**What It Checks**:
- Generic bookmark names (e.g., "Bookmark 1", "BM 2")

### 8. Hidden Pages (WCAG 2.4.1 Bypass Blocks)

**What It Checks**:
- Pages with visibility=1 (hidden)
- INFO level - alerts user to verify no essential content is hidden

---

## Color Contrast Analysis

### WCAG Contrast Requirements

| Level | Text Type | Minimum Ratio |
|-------|-----------|---------------|
| AA | Normal text | 4.5:1 |
| AA | Large text (18pt+ or 14pt+ bold) | 3:1 |
| AAA | Normal text | 7:1 |
| AAA | Large text | 4.5:1 |
| All | UI Components | 3:1 |

### Elements Checked

The analyzer checks contrast for multiple text elements per visual:
- Subtitle text
- Title text
- Data labels
- Category labels
- Data point colors
- Legend text
- Axis labels (X, Y, Category, Value)
- Table headers (column, row, values)

### Color Extraction

**Priority Order**:
1. Literal hex values (`#RRGGBB` or `#RRGGBBAA`)
2. ThemeDataColor references (resolved via theme palette)
3. SolidColor theme references
4. FillRule gradients (uses min color)

**Theme Color Resolution**:
```python
# Checks in order:
# 1. RegisteredResources (custom themes)
# 2. SharedResources/BaseThemes
# 3. Fallback to white background
```

### Transparency Handling

**Alpha Blending**:
- Colors with alpha channels are blended with background
- Uses standard alpha compositing formula
- Flags visuals with 3+ transparent layers for manual review

### Contrast Calculation

```python
# WCAG relative luminance formula
def _calculate_luminance(r, g, b):
    # sRGB to linear RGB conversion
    # Returns 0.2126*R + 0.7152*G + 0.0722*B

# Contrast ratio
ratio = (lighter_luminance + 0.05) / (darker_luminance + 0.05)
```

---

## Data Models

### AccessibilityAnalysisResult

```python
@dataclass
class AccessibilityAnalysisResult:
    report_path: str
    report_name: str

    # Collected data
    pages: List[PageInfo]
    visuals: List[VisualInfo]
    bookmarks: List[BookmarkInfo]

    # Analysis results
    issues: List[AccessibilityIssue]
    tab_orders: List[TabOrderInfo]
    color_contrasts: List[ColorContrastResult]

    # Summary counts
    total_issues: int
    errors: int
    warnings: int
    info_count: int
    issues_by_type: Dict[AccessibilityCheckType, int]

    # Metadata
    analysis_timestamp: str
    analysis_duration_ms: int
```

### AccessibilityIssue

```python
@dataclass
class AccessibilityIssue:
    check_type: AccessibilityCheckType
    severity: AccessibilitySeverity
    page_name: str
    issue_description: str
    recommendation: str
    visual_name: Optional[str]
    visual_type: Optional[str]
    current_value: Optional[str]
    wcag_reference: Optional[str]
```

---

## Configuration

### Check Configuration

Users can enable/disable individual checks:
```python
config = get_config()
config.is_check_enabled("tab_order")      # True/False
config.is_check_enabled("color_contrast") # True/False
```

### Contrast Level Configuration

```python
config.contrast_level = "AA"       # Standard (4.5:1)
config.contrast_level = "AA_large" # Large text only (3:1)
config.contrast_level = "AAA"      # Enhanced (7:1)

config.flag_aaa_failures = True    # Flag AAA failures as INFO
config.flag_aa_failures = True     # Flag AA failures when using AA_large
```

---

## Usage Guide

### Basic Workflow

1. **Select Report**: Choose a .pbip or .pbix file (PBIR format required)
2. **Configure Checks**: Optionally adjust which checks to run
3. **Analyze Report**: Click analyze to scan for issues
4. **Review Results**: Issues displayed by category cards
5. **Drill Into Details**: Click category to see specific issues
6. **Export Report**: Generate documentation for remediation

### Interpreting Results

**ERROR (Red)**: Critical issues failing WCAG compliance
- Must fix for accessibility compliance
- Example: Missing alt text on data visual

**WARNING (Yellow)**: Should fix for better accessibility
- Important for good accessibility
- Example: Generic page title

**INFO (Blue)**: Suggestions and best practices
- Consider addressing
- Example: Hidden page notification

---

## File Structure

```
accessibility_checker/
├── __init__.py               # Package exports
├── tool.py                   # BaseTool implementation
├── accessibility_analyzer.py # Core analysis engine (~1800 lines)
├── accessibility_types.py    # Data models (~490 lines)
├── accessibility_config.py   # Configuration management
├── accessibility_ui.py       # UI tab with category cards
└── TECHNICAL_GUIDE.md        # This document
```

---

## Extending the Tool

### Adding a New Check Type

1. **Add to AccessibilityCheckType enum**:
```python
class AccessibilityCheckType(Enum):
    # ... existing
    MY_NEW_CHECK = "my_new_check"
```

2. **Add display name and icon**:
```python
CHECK_TYPE_DISPLAY_NAMES[AccessibilityCheckType.MY_NEW_CHECK] = "My New Check"
CHECK_TYPE_ICONS[AccessibilityCheckType.MY_NEW_CHECK] = "New"
```

3. **Implement analysis method in AccessibilityAnalyzer**:
```python
def _analyze_my_new_check(self, visuals: List[VisualInfo]) -> List[AccessibilityIssue]:
    issues = []
    # ... analysis logic
    return issues
```

4. **Call from analyze_pbip_report**:
```python
if config.is_check_enabled("my_new_check"):
    self._update_progress(XX, "Checking my new check...")
    issues = self._analyze_my_new_check(result.visuals)
    result.issues.extend(issues)
```

### Adding Generic Pattern Detection

Add patterns to the appropriate list in `accessibility_types.py`:
```python
GENERIC_PAGE_TITLE_PATTERNS = [
    # ... existing
    r"^my_new_pattern$",
]
```

---

## Troubleshooting

### Common Issues

**"File is not in PBIR format"**
- The PBIX file must be saved with "Store reports using enhanced metadata format (PBIR)" enabled
- PBIR is the default format since January 2026 Power BI Desktop

**"No .Report folder found"**
- Ensure the PBIP structure includes the .Report directory
- For PBIX files, save in PBIR format first

**Color contrast shows "Unable to automatically detect"**
- Visual uses complex color expressions (conditional formatting)
- Colors defined in unsupported format
- Solution: Manual verification recommended

**Analysis takes a long time**
- Large reports with many pages/visuals
- Many color contrast checks per visual
- Consider disabling checks not needed

### Performance Tips

- Disable checks you don't need via configuration
- Reports with 50+ pages may take 30+ seconds
- Color contrast is the most intensive check

---

## Appendix

### WCAG References

| Criterion | Name | Level |
|-----------|------|-------|
| 1.1.1 | Non-text Content | A |
| 1.4.1 | Use of Color | A |
| 1.4.3 | Contrast (Minimum) | AA |
| 1.4.6 | Contrast (Enhanced) | AAA |
| 2.4.1 | Bypass Blocks | A |
| 2.4.2 | Page Titled | A |
| 2.4.3 | Focus Order | A |
| 2.4.6 | Headings and Labels | AA |

### Dependencies

- Python 3.8+
- tkinter (standard library)
- pathlib (standard library)
- json (standard library)
- re (standard library)
- Core PBI File Reader (shared)
- Core UI Base classes

---

**Built by Reid Havens of Analytic Endeavors**
**Website**: https://www.analyticendeavors.com
**Email**: reid@analyticendeavors.com
