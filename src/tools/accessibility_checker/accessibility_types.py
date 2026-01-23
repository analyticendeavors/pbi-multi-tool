"""
Accessibility Checker - Data Types and Models
Built by Reid Havens of Analytic Endeavors

Data classes for representing accessibility check results, issues, and analysis outputs.
"""

from dataclasses import dataclass, field
from typing import List, Optional, Dict, Any
from enum import Enum

# Ownership fingerprint
_AE_FP = "QWNjZXNzaWJpbGl0eUNoZWNrZXI6QUUtMjAyNA=="


class AccessibilitySeverity(Enum):
    """Severity levels for accessibility issues"""
    ERROR = "error"       # Critical - fails WCAG compliance
    WARNING = "warning"   # Should fix for better accessibility
    INFO = "info"         # Suggestion/best practice


class AccessibilityCheckType(Enum):
    """Types of accessibility checks performed"""
    TAB_ORDER = "tab_order"
    ALT_TEXT = "alt_text"
    COLOR_CONTRAST = "color_contrast"
    PAGE_TITLE = "page_title"
    VISUAL_TITLE = "visual_title"
    DATA_LABELS = "data_labels"
    BOOKMARK_NAME = "bookmark_name"
    HIDDEN_PAGE = "hidden_page"


# Display names for check types
CHECK_TYPE_DISPLAY_NAMES = {
    AccessibilityCheckType.TAB_ORDER: "Tab Order",
    AccessibilityCheckType.ALT_TEXT: "Alt Text",
    AccessibilityCheckType.COLOR_CONTRAST: "Color Contrast",
    AccessibilityCheckType.PAGE_TITLE: "Page Titles",
    AccessibilityCheckType.VISUAL_TITLE: "Visual Titles",
    AccessibilityCheckType.DATA_LABELS: "Data Labels",
    AccessibilityCheckType.BOOKMARK_NAME: "Bookmark Names",
    AccessibilityCheckType.HIDDEN_PAGE: "Hidden Pages",
}

# Icons for check types (emoji fallbacks)
CHECK_TYPE_ICONS = {
    AccessibilityCheckType.TAB_ORDER: "Tab",
    AccessibilityCheckType.ALT_TEXT: "Alt",
    AccessibilityCheckType.COLOR_CONTRAST: "Clr",
    AccessibilityCheckType.PAGE_TITLE: "Pg",
    AccessibilityCheckType.VISUAL_TITLE: "Vis",
    AccessibilityCheckType.DATA_LABELS: "Lbl",
    AccessibilityCheckType.BOOKMARK_NAME: "Bkm",
    AccessibilityCheckType.HIDDEN_PAGE: "Hid",
}


@dataclass
class AccessibilityIssue:
    """Represents a single accessibility issue found in the report"""
    check_type: AccessibilityCheckType
    severity: AccessibilitySeverity
    page_name: str
    issue_description: str
    recommendation: str
    visual_name: Optional[str] = None
    visual_type: Optional[str] = None
    location: str = ""  # File path or location within report
    current_value: Optional[str] = None  # e.g., current tab order, color hex
    wcag_reference: Optional[str] = None  # e.g., "WCAG 2.1 AA 1.4.3"

    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary for serialization"""
        return {
            "check_type": self.check_type.value,
            "severity": self.severity.value,
            "page_name": self.page_name,
            "visual_name": self.visual_name,
            "visual_type": self.visual_type,
            "issue_description": self.issue_description,
            "recommendation": self.recommendation,
            "location": self.location,
            "current_value": self.current_value,
            "wcag_reference": self.wcag_reference,
        }


@dataclass
class TabOrderInfo:
    """Tab order information for a visual element"""
    page_name: str
    visual_name: str
    visual_type: str
    tab_order: int  # -1 if not set
    position_x: float
    position_y: float
    position_z: int
    width: float
    height: float
    parent_group: Optional[str] = None
    nesting_level: int = 0
    visual_id: Optional[str] = None

    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary for serialization"""
        return {
            "page_name": self.page_name,
            "visual_name": self.visual_name,
            "visual_type": self.visual_type,
            "tab_order": self.tab_order,
            "position_x": self.position_x,
            "position_y": self.position_y,
            "position_z": self.position_z,
            "width": self.width,
            "height": self.height,
            "parent_group": self.parent_group,
            "nesting_level": self.nesting_level,
        }


@dataclass
class ColorContrastResult:
    """Result of a color contrast check for a visual element"""
    page_name: str
    visual_name: str
    visual_type: str
    foreground_color: str
    background_color: str
    contrast_ratio: float
    passes_aa_normal: bool  # 4.5:1 for normal text
    passes_aa_large: bool   # 3:1 for large text
    passes_aaa_normal: bool # 7:1 for normal text
    passes_aaa_large: bool  # 4.5:1 for large text
    luminance_fg: float
    luminance_bg: float
    element_type: str = "text"  # text, ui_component, etc.
    # Transparency handling fields
    fg_has_transparency: bool = False
    bg_has_transparency: bool = False
    transparency_layer_count: int = 0
    requires_review: bool = False  # True if too many layers for certain calculation
    blended_fg_color: Optional[str] = None  # Calculated blended foreground color
    blended_bg_color: Optional[str] = None  # Calculated blended background color

    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary for serialization"""
        return {
            "page_name": self.page_name,
            "visual_name": self.visual_name,
            "visual_type": self.visual_type,
            "foreground_color": self.foreground_color,
            "background_color": self.background_color,
            "contrast_ratio": round(self.contrast_ratio, 2),
            "passes_aa_normal": self.passes_aa_normal,
            "passes_aa_large": self.passes_aa_large,
            "passes_aaa_normal": self.passes_aaa_normal,
            "passes_aaa_large": self.passes_aaa_large,
            "element_type": self.element_type,
            "fg_has_transparency": self.fg_has_transparency,
            "bg_has_transparency": self.bg_has_transparency,
            "transparency_layer_count": self.transparency_layer_count,
            "requires_review": self.requires_review,
            "blended_fg_color": self.blended_fg_color,
            "blended_bg_color": self.blended_bg_color,
        }


@dataclass
class VisualContrastSummary:
    """Summary of all contrast results for a single visual.

    Groups multiple ColorContrastResult objects by visual to enable
    hierarchical display where each visual is shown with all its
    failing elements underneath.
    """
    page_name: str
    visual_name: str
    visual_type: str
    visual_id: str
    element_results: List['ColorContrastResult']  # All elements checked
    worst_ratio: float  # Lowest contrast ratio found
    worst_element: str  # Element type with worst ratio
    total_elements: int  # Total elements checked
    failing_count: int  # Count of elements failing current threshold

    @property
    def has_issues(self) -> bool:
        """True if any element fails the current threshold"""
        return self.failing_count > 0

    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary for serialization"""
        return {
            "page_name": self.page_name,
            "visual_name": self.visual_name,
            "visual_type": self.visual_type,
            "visual_id": self.visual_id,
            "element_results": [r.to_dict() for r in self.element_results],
            "worst_ratio": round(self.worst_ratio, 2),
            "worst_element": self.worst_element,
            "total_elements": self.total_elements,
            "failing_count": self.failing_count
        }


@dataclass
class VisualInfo:
    """Information about a visual element"""
    page_name: str
    visual_name: str
    visual_type: str
    visual_id: str
    has_alt_text: bool = False
    alt_text: Optional[str] = None
    has_title: bool = False
    title_text: Optional[str] = None
    has_data_labels: Optional[bool] = None  # None if not applicable
    tab_order: int = -1
    position_x: float = 0
    position_y: float = 0
    width: float = 0
    height: float = 0
    is_data_visual: bool = False  # Charts, tables, KPIs, etc.
    is_hidden: bool = False  # True if visual is hidden (not visible)
    is_group: bool = False  # True if this is a visual group container
    parent_group_name: Optional[str] = None  # Parent group ID if visual is in a group

    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary for serialization"""
        return {
            "page_name": self.page_name,
            "visual_name": self.visual_name,
            "visual_type": self.visual_type,
            "visual_id": self.visual_id,
            "has_alt_text": self.has_alt_text,
            "alt_text": self.alt_text,
            "has_title": self.has_title,
            "title_text": self.title_text,
            "has_data_labels": self.has_data_labels,
            "tab_order": self.tab_order,
            "is_data_visual": self.is_data_visual,
            "is_hidden": self.is_hidden,
            "is_group": self.is_group,
            "parent_group_name": self.parent_group_name,
        }


@dataclass
class PageInfo:
    """Information about a report page"""
    page_name: str
    display_name: str
    page_id: str
    is_hidden: bool = False
    has_title: bool = False
    title_text: Optional[str] = None
    visual_count: int = 0
    ordinal: int = 0

    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary for serialization"""
        return {
            "page_name": self.page_name,
            "display_name": self.display_name,
            "page_id": self.page_id,
            "is_hidden": self.is_hidden,
            "has_title": self.has_title,
            "title_text": self.title_text,
            "visual_count": self.visual_count,
            "ordinal": self.ordinal,
        }


@dataclass
class BookmarkInfo:
    """Information about a bookmark"""
    bookmark_name: str
    display_name: str
    bookmark_id: str
    is_generic_name: bool = False

    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary for serialization"""
        return {
            "bookmark_name": self.bookmark_name,
            "display_name": self.display_name,
            "bookmark_id": self.bookmark_id,
            "is_generic_name": self.is_generic_name,
        }


@dataclass
class AccessibilityAnalysisResult:
    """Complete result of an accessibility analysis"""
    report_path: str
    report_name: str = ""

    # Collected data
    pages: List[PageInfo] = field(default_factory=list)
    visuals: List[VisualInfo] = field(default_factory=list)
    bookmarks: List[BookmarkInfo] = field(default_factory=list)

    # Analysis results
    issues: List[AccessibilityIssue] = field(default_factory=list)
    tab_orders: List[TabOrderInfo] = field(default_factory=list)
    color_contrasts: List[ColorContrastResult] = field(default_factory=list)

    # Summary counts
    total_issues: int = 0
    errors: int = 0
    warnings: int = 0
    info_count: int = 0

    # Per-check issue counts
    issues_by_type: Dict[AccessibilityCheckType, int] = field(default_factory=dict)

    # Metadata
    analysis_timestamp: str = ""
    analysis_duration_ms: int = 0

    def __post_init__(self):
        """Initialize issues_by_type with all check types"""
        if not self.issues_by_type:
            self.issues_by_type = {check_type: 0 for check_type in AccessibilityCheckType}

    def update_counts(self):
        """Recalculate summary counts from issues list"""
        self.total_issues = len(self.issues)
        self.errors = sum(1 for i in self.issues if i.severity == AccessibilitySeverity.ERROR)
        self.warnings = sum(1 for i in self.issues if i.severity == AccessibilitySeverity.WARNING)
        self.info_count = sum(1 for i in self.issues if i.severity == AccessibilitySeverity.INFO)

        # Reset and recalculate per-type counts
        self.issues_by_type = {check_type: 0 for check_type in AccessibilityCheckType}
        for issue in self.issues:
            self.issues_by_type[issue.check_type] += 1

    def get_issues_by_type(self, check_type: AccessibilityCheckType) -> List[AccessibilityIssue]:
        """Get all issues of a specific type"""
        return [i for i in self.issues if i.check_type == check_type]

    def get_issues_by_severity(self, severity: AccessibilitySeverity) -> List[AccessibilityIssue]:
        """Get all issues of a specific severity"""
        return [i for i in self.issues if i.severity == severity]

    def get_issues_by_page(self, page_name: str) -> List[AccessibilityIssue]:
        """Get all issues for a specific page"""
        return [i for i in self.issues if i.page_name == page_name]

    def has_errors(self) -> bool:
        """Check if there are any error-level issues"""
        return self.errors > 0

    def has_warnings(self) -> bool:
        """Check if there are any warning-level issues"""
        return self.warnings > 0

    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary for serialization"""
        return {
            "report_path": self.report_path,
            "report_name": self.report_name,
            "total_issues": self.total_issues,
            "errors": self.errors,
            "warnings": self.warnings,
            "info_count": self.info_count,
            "issues_by_type": {k.value: v for k, v in self.issues_by_type.items()},
            "pages": [p.to_dict() for p in self.pages],
            "issues": [i.to_dict() for i in self.issues],
            "analysis_timestamp": self.analysis_timestamp,
            "analysis_duration_ms": self.analysis_duration_ms,
        }


# Visual types that are considered "data visuals" requiring more accessibility attention
DATA_VISUAL_TYPES = {
    "barChart",
    "clusteredBarChart",
    "stackedBarChart",
    "hundredPercentStackedBarChart",
    "columnChart",
    "clusteredColumnChart",
    "stackedColumnChart",
    "hundredPercentStackedColumnChart",
    "lineChart",
    "areaChart",
    "stackedAreaChart",
    "lineStackedColumnComboChart",
    "lineClusteredColumnComboChart",
    "ribbonChart",
    "waterfallChart",
    "funnelChart",
    "pieChart",
    "donutChart",
    "treemap",
    "map",
    "filledMap",
    "shapeMap",
    "azureMap",
    "scatterChart",
    "table",
    "matrix",
    "card",
    "multiRowCard",
    "kpi",
    "gauge",
    "slicer",
    "decompositionTreeVisual",
    "keyInfluencersVisual",
    "qnaVisual",
    "smartNarrativeVisual",
}

# Visual types that are decorative/UI elements
DECORATIVE_VISUAL_TYPES = {
    "shape",
    "textbox",
    "image",
    "basicShape",
    "actionButton",
    "bookmarkNavigator",
    "pageNavigator",
}

# Generic alt text patterns that should be flagged
GENERIC_ALT_TEXT_PATTERNS = [
    "image",
    "chart",
    "visual",
    "picture",
    "graph",
    "table",
    "data",
    "figure",
    "diagram",
    "illustration",
]

# Generic page title patterns that should be flagged
GENERIC_PAGE_TITLE_PATTERNS = [
    r"^page\s*\d*$",
    r"^sheet\s*\d*$",
    r"^section\s*\d*$",
    r"^tab\s*\d*$",
    r"^report\s*page\s*\d*$",
    r"^untitled$",
    r"^new\s*page$",
]

# Generic bookmark name patterns that should be flagged
GENERIC_BOOKMARK_PATTERNS = [
    r"^bookmark\s*\d*$",
    r"^bm\s*\d*$",
    r"^new\s*bookmark$",
    r"^untitled$",
]


# WCAG contrast ratio requirements
WCAG_CONTRAST_AA_NORMAL = 4.5  # Normal text
WCAG_CONTRAST_AA_LARGE = 3.0   # Large text (18pt+ or 14pt+ bold)
WCAG_CONTRAST_AAA_NORMAL = 7.0  # Enhanced: Normal text
WCAG_CONTRAST_AAA_LARGE = 4.5   # Enhanced: Large text
WCAG_CONTRAST_UI_COMPONENT = 3.0  # UI components and graphical objects


__all__ = [
    'AccessibilitySeverity',
    'AccessibilityCheckType',
    'CHECK_TYPE_DISPLAY_NAMES',
    'CHECK_TYPE_ICONS',
    'AccessibilityIssue',
    'TabOrderInfo',
    'ColorContrastResult',
    'VisualInfo',
    'PageInfo',
    'BookmarkInfo',
    'AccessibilityAnalysisResult',
    'DATA_VISUAL_TYPES',
    'DECORATIVE_VISUAL_TYPES',
    'GENERIC_ALT_TEXT_PATTERNS',
    'GENERIC_PAGE_TITLE_PATTERNS',
    'GENERIC_BOOKMARK_PATTERNS',
    'WCAG_CONTRAST_AA_NORMAL',
    'WCAG_CONTRAST_AA_LARGE',
    'WCAG_CONTRAST_AAA_NORMAL',
    'WCAG_CONTRAST_AAA_LARGE',
    'WCAG_CONTRAST_UI_COMPONENT',
]
