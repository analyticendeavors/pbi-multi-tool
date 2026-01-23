"""
Accessibility Checker UI - User interface for accessibility analysis
Built by Reid Havens of Analytic Endeavors

This module provides the user interface for the Accessibility Checker tool,
following the established patterns from other tools in the suite.
"""

import tkinter as tk
from tkinter import ttk, filedialog
from pathlib import Path
from typing import Dict, List, Any, Optional, Callable
import threading
import csv
import io
from datetime import datetime
from collections import defaultdict
import re

# Icon loading support
try:
    from PIL import Image, ImageTk
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False

try:
    import cairosvg
    CAIROSVG_AVAILABLE = True
except ImportError:
    CAIROSVG_AVAILABLE = False

from core.ui_base import BaseToolTab, FileInputMixin, ValidationMixin, RoundedButton, SquareIconButton, Tooltip, ThemedMessageBox, ActionButtonBar, SplitLogSection, FileInputSection, LabeledToggle, LabeledRadioGroup, ThemedScrollbar
from core.constants import AppConstants
from core.theme_manager import get_theme_manager
from tools.accessibility_checker.accessibility_types import (
    AccessibilityCheckType,
    AccessibilitySeverity,
    AccessibilityIssue,
    AccessibilityAnalysisResult,
    CHECK_TYPE_DISPLAY_NAMES,
)
from tools.accessibility_checker.accessibility_analyzer import AccessibilityAnalyzer
from tools.accessibility_checker.accessibility_config import (
    get_config,
    save_config,
    reset_config,
    AccessibilityCheckConfig,
    CONTRAST_THRESHOLDS,
)


class AccessibilityCheckCard(tk.Frame):
    """
    Visual card showing an accessibility check category.
    Shows "--" before analysis, then updates with "X issues / Y total" after analysis.
    Color-coded: green (0 issues), yellow (warnings), red (errors).
    """

    def __init__(self, parent, title: str, subtitle: str,
                 check_type: AccessibilityCheckType,
                 on_card_click: Callable = None,
                 icon: 'ImageTk.PhotoImage' = None, **kwargs):
        self._theme_manager = get_theme_manager()
        colors = self._theme_manager.colors

        super().__init__(parent, bg=colors['card_surface'],
                        highlightbackground=colors['border'],
                        highlightthickness=1, takefocus=False, **kwargs)

        self.check_type = check_type
        self.on_card_click = on_card_click
        self._issue_count = None
        self._total_count = None
        self._has_errors = False
        self._has_warnings = False
        self._enabled = False
        self._title = title
        self._subtitle = subtitle
        self._selected = False
        self._icon = icon

        # Build card UI
        self._build_card()
        self._bind_card_click()

        # Register for theme changes
        self._theme_manager.register_theme_callback(self._handle_theme_change)

    def _build_card(self):
        """Build the card UI elements"""
        colors = self._theme_manager.colors

        # Inner padding frame - compact padding
        inner = tk.Frame(self, bg=colors['card_surface'], padx=15, pady=8, takefocus=False)
        inner.pack(fill=tk.BOTH, expand=True)

        # Icon (if provided)
        self.icon_label = None
        if self._icon:
            self.icon_label = tk.Label(inner, image=self._icon, bg=colors['card_surface'])
            self.icon_label.pack(pady=(0, 2))
            self.icon_label._icon_ref = self._icon  # Prevent garbage collection

        # Title (e.g., "Tab Order")
        self.title_label = tk.Label(inner, text=self._title, font=('Segoe UI', 9, 'bold'),
                                    bg=colors['card_surface'], fg=colors['text_primary'])
        self.title_label.pack(pady=(2 if not self._icon else 0, 0))

        # Subtitle (e.g., "Navigation sequence")
        self.subtitle_label = tk.Label(inner, text=self._subtitle, font=('Segoe UI', 8),
                                       bg=colors['card_surface'], fg=colors['text_secondary'])
        self.subtitle_label.pack()

        # Count label (shows "--" initially, then issue count)
        self.count_label = tk.Label(inner, text="--", font=('Segoe UI', 14, 'bold'),
                                    bg=colors['card_surface'], fg=colors['text_muted'])
        self.count_label.pack(pady=(4, 0))

        # Status label (shows "X issues / Y total" or "Pass" after analysis)
        self.status_label = tk.Label(inner, text="", font=('Segoe UI', 8),
                                     bg=colors['card_surface'], fg=colors['text_muted'])
        self.status_label.pack()

        # Store inner frame for theme updates
        self._inner = inner

    def _bind_card_click(self):
        """Bind click events to show details"""
        def on_click(event):
            if self._enabled and self.on_card_click:
                self.on_card_click(self.check_type)

        self.bind('<Button-1>', on_click)
        self._inner.bind('<Button-1>', on_click)
        if self.icon_label:
            self.icon_label.bind('<Button-1>', on_click)
        self.title_label.bind('<Button-1>', on_click)
        self.subtitle_label.bind('<Button-1>', on_click)
        self.count_label.bind('<Button-1>', on_click)
        self.status_label.bind('<Button-1>', on_click)

        # Change cursor to pointer on hover when enabled
        def on_enter(event):
            if self._enabled:
                self.config(cursor='hand2')

        def on_leave(event):
            self.config(cursor='')

        self.bind('<Enter>', on_enter)
        self.bind('<Leave>', on_leave)

    def set_selected(self, selected: bool):
        """Set the card as selected (showing details)"""
        self._selected = selected
        colors = self._theme_manager.colors

        if selected and self._enabled:
            self.config(highlightthickness=2, highlightbackground=colors['button_primary'])
        else:
            self.config(highlightthickness=1, highlightbackground=colors['border'])

    def update_results(self, issue_count: int, total_count: int,
                       has_errors: bool = False, has_warnings: bool = False):
        """Update card with analysis results showing issues / total format"""
        colors = self._theme_manager.colors
        self._issue_count = issue_count
        self._total_count = total_count
        self._has_errors = has_errors
        self._has_warnings = has_warnings
        self._enabled = issue_count > 0

        if issue_count > 0:
            self.count_label.config(text=str(issue_count))

            # Color-code based on severity and show status with new terminology
            if has_errors:
                self.count_label.config(fg=colors['error'])
                self.status_label.config(text=f"Critical / {total_count} checked", fg=colors['error'])
            elif has_warnings:
                self.count_label.config(fg=colors['warning'])
                self.status_label.config(text=f"Should Fix / {total_count} checked", fg=colors['warning'])
            else:
                self.count_label.config(fg=colors['info'])
                self.status_label.config(text=f"Review / {total_count} checked", fg=colors['info'])

            # Normal card appearance
            self.config(bg=colors['card_surface'])
            self._inner.config(bg=colors['card_surface'])
            self._update_label_backgrounds(colors['card_surface'])
        else:
            # No issues - green/success
            self.count_label.config(text="0", fg=colors['success'])
            if total_count > 0:
                self.status_label.config(text=f"Pass / {total_count} checked", fg=colors['success'])
            else:
                self.status_label.config(text="Pass", fg=colors['success'])

            # Grey out the card when no issues (matches Report Cleanup pattern)
            grey_bg = colors['card_surface_hover']
            self.config(bg=grey_bg)
            self._inner.config(bg=grey_bg)
            self._update_label_backgrounds(grey_bg)

    def _update_label_backgrounds(self, bg_color: str):
        """Update all label backgrounds"""
        if self.icon_label:
            self.icon_label.config(bg=bg_color)
        self.title_label.config(bg=bg_color)
        self.subtitle_label.config(bg=bg_color)
        self.count_label.config(bg=bg_color)
        self.status_label.config(bg=bg_color)

    def reset(self):
        """Reset card to initial state"""
        colors = self._theme_manager.colors

        self.count_label.config(text="--", fg=colors['text_muted'])
        self.status_label.config(text="", fg=colors['text_muted'])
        self._issue_count = None
        self._total_count = None
        self._has_errors = False
        self._has_warnings = False
        self._enabled = False
        self._selected = False

        # Restore normal appearance
        self.config(bg=colors['card_surface'], highlightthickness=1, highlightbackground=colors['border'])
        self._inner.config(bg=colors['card_surface'])
        self._update_label_backgrounds(colors['card_surface'])

    def _handle_theme_change(self, theme: str):
        """Handle theme change"""
        colors = self._theme_manager.colors

        # Determine background based on state
        if self._enabled:
            bg = colors['card_surface']
        elif self._issue_count == 0:
            bg = colors['card_surface_hover']  # Grey out when no issues
        else:
            bg = colors['card_surface']

        if self._selected and self._enabled:
            self.config(bg=bg, highlightthickness=2, highlightbackground=colors['button_primary'])
        else:
            self.config(bg=bg, highlightthickness=1, highlightbackground=colors['border'])

        self._inner.config(bg=bg)
        self._update_label_backgrounds(bg)

        self.title_label.config(fg=colors['text_primary'])
        self.subtitle_label.config(fg=colors['text_secondary'])

        # Update count/status colors based on state
        if self._issue_count is None:
            self.count_label.config(fg=colors['text_muted'])
            self.status_label.config(fg=colors['text_muted'])
        elif self._issue_count == 0:
            self.count_label.config(fg=colors['success'])
            self.status_label.config(fg=colors['success'])
        elif self._has_errors:
            self.count_label.config(fg=colors['error'])
            self.status_label.config(fg=colors['error'])
        elif self._has_warnings:
            self.count_label.config(fg=colors['warning'])
            self.status_label.config(fg=colors['warning'])
        else:
            self.count_label.config(fg=colors['info'])
            self.status_label.config(fg=colors['info'])


class AccessibilityCheckerTab(BaseToolTab, FileInputMixin, ValidationMixin):
    """
    Accessibility Checker UI Tab - provides interface for analyzing Power BI report accessibility
    """

    def __init__(self, parent, main_app):
        super().__init__(parent, main_app, "accessibility_checker", "Accessibility Checker")

        # Tool instance
        self.analyzer = AccessibilityAnalyzer(
            logger_callback=self.log_message,
            progress_callback=self._update_progress_safe
        )

        # UI state
        self.current_results: Optional[AccessibilityAnalysisResult] = None
        self.selected_check_type: Optional[AccessibilityCheckType] = None

        # Store totals for each check type (for X/Y display)
        self._check_totals: Dict[AccessibilityCheckType, int] = {}

        # UI components
        self.file_section = None  # FileInputSection instance
        self.pbip_path_var = None  # Set from file_section.path_var
        self.analyze_button = None  # Set from file_section.action_button
        self.check_cards: Dict[AccessibilityCheckType, AccessibilityCheckCard] = {}
        self.export_button = None
        self.reset_button = None
        self.progress_components = None

        self._cards_frame = None

        # Split log section components
        self.log_section = None
        self.details_text = None
        self._summary_frame = None
        self._placeholder_label = None
        self._selected_card_key = None

        # Page filter state for Issue Details
        self._page_filter_selection = "All Pages"  # Can be "All Pages" or a page display name
        self._page_filter_popup = None
        self._page_filter_icon = None
        self._page_filter_clear_btn = None
        self._page_filter_search_var = None
        self._page_filter_search_after_id = None

        # Setup UI and show welcome message
        self.setup_ui()
        self._show_welcome_message()

    def _show_welcome_message(self):
        """Show welcome message in log"""
        self.log_message("✨ Welcome to Accessibility Checker!")
        self.log_message("=" * 60)
        self.log_message("This tool checks your Power BI reports for WCAG compliance:")
        self.log_message("• Tab order, alt text, and keyboard navigation issues")
        self.log_message("• Color contrast and visual accessibility problems")
        self.log_message("• Missing titles, labels, and generic naming issues")
        self.log_message("• Detailed remediation guidance for all findings")
        self.log_message("")
        self.log_message("📂 Start by selecting your Power BI file and clicking 'ANALYZE REPORT'")
        self.log_message("💡 Consider backing up your report before making changes")

    def on_theme_changed(self, theme: str):
        """Update theme-dependent widgets when theme changes"""
        super().on_theme_changed(theme)
        colors = self._theme_manager.colors
        is_dark = self._theme_manager.is_dark

        # FileInputSection handles its own theme updates via _on_theme_changed

        # Update cards frame background
        if self._cards_frame:
            try:
                self._cards_frame.configure(bg=colors['background'])
            except Exception:
                pass

        # Update filter elements (positioned above section frame)
        outer_bg = colors.get('section_bg', colors['background'])
        if hasattr(self, '_filter_frame') and self._filter_frame:
            try:
                self._filter_frame.configure(bg=outer_bg)
            except Exception:
                pass
        if hasattr(self, '_filter_label') and self._filter_label:
            try:
                self._filter_label.configure(bg=outer_bg, fg=colors['text_secondary'])
            except Exception:
                pass

        # Update severity dropdown button
        if hasattr(self, '_severity_dropdown_btn') and self._severity_dropdown_btn:
            try:
                self._severity_dropdown_btn.configure(
                    bg=colors['card_surface'],
                    fg=colors['text_primary'],
                    highlightbackground=colors['border']
                )
            except Exception:
                pass

        # Update bottom buttons with correct outer background
        outer_canvas_bg = colors.get('outer_bg', '#161627' if is_dark else '#f5f5f7')
        try:
            if hasattr(self, '_button_frame') and self._button_frame:
                self._button_frame.configure(bg=outer_canvas_bg)
            if hasattr(self, '_button_container') and self._button_container:
                self._button_container.configure(bg=outer_canvas_bg)
            if self.export_button:
                self.export_button.update_colors(
                    bg=colors['button_primary'],
                    hover_bg=colors['button_primary_hover'],
                    pressed_bg=colors['button_primary_pressed'],
                    fg='#ffffff',
                    disabled_bg=colors.get('button_primary_disabled', '#3a3a4e' if is_dark else '#c0c0cc'),
                    disabled_fg=colors.get('button_text_disabled', '#6a6a7a' if is_dark else '#9a9aa8'),
                    canvas_bg=outer_canvas_bg
                )
                self.export_button._draw_button()
            if hasattr(self, 'export_pdf_button') and self.export_pdf_button:
                self.export_pdf_button.update_colors(
                    bg=colors['button_primary'],
                    hover_bg=colors['button_primary_hover'],
                    pressed_bg=colors['button_primary_pressed'],
                    fg='#ffffff',
                    disabled_bg=colors.get('button_primary_disabled', '#3a3a4e' if is_dark else '#c0c0cc'),
                    disabled_fg=colors.get('button_text_disabled', '#6a6a7a' if is_dark else '#9a9aa8'),
                    canvas_bg=outer_canvas_bg
                )
                self.export_pdf_button._draw_button()
            if self.reset_button:
                self.reset_button.update_colors(
                    bg=colors['button_secondary'],
                    hover_bg=colors['button_secondary_hover'],
                    pressed_bg=colors['button_secondary_pressed'],
                    fg=colors['text_primary'],
                    canvas_bg=outer_canvas_bg
                )
                self.reset_button._draw_button()
        except Exception:
            pass

        # SplitLogSection handles its own theme updates internally

        # Update details text widget text tags for custom issue display
        if hasattr(self, 'details_text') and self.details_text:
            try:
                self.details_text.configure(
                    bg=colors['section_bg'],
                    fg=colors['text_primary']
                )
                # Update text tags for theme-aware colors
                self.details_text.tag_config('header', foreground=colors['text_primary'])
                self.details_text.tag_config('subheader', foreground=colors['text_primary'])
                self.details_text.tag_config('separator', foreground=colors['text_muted'])
                self.details_text.tag_config('item', foreground=colors['text_primary'])
                self.details_text.tag_config('error', foreground=colors['error'])
                self.details_text.tag_config('warning', foreground=colors['warning'])
                self.details_text.tag_config('info', foreground=colors['info'])
                self.details_text.tag_config('success', foreground=colors['success'])
                self.details_text.tag_config('tip', foreground=colors['text_secondary'])
            except Exception:
                pass

        # Refresh issue details display to update dynamic content with new theme colors
        # This ensures expand/collapse icons and other themed elements update on theme switch
        if self.selected_check_type and self.current_results:
            self._show_card_details(self.selected_check_type)

    def setup_ui(self) -> None:
        """Setup the Accessibility Checker UI"""
        # Load UI icons
        if not hasattr(self, '_button_icons'):
            self._button_icons = {}

        icon_names_16 = ["Power-BI", "magnifying-glass", "reset", "analyze", "folder", "question", "checker", "cogwheel", "csv-file", "pdf", "save"]
        for name in icon_names_16:
            icon = self._load_icon_for_button(name, size=16)
            if icon:
                self._button_icons[name] = icon

        # Load card icons (compact size for smaller cards)
        card_icon_names = ["target", "alt text", "paint", "file", "bar-chart", "tooltip", "bookmark", "hidden"]
        for name in card_icon_names:
            icon = self._load_icon_for_button(name, size=22)
            if icon:
                self._button_icons[f"card_{name}"] = icon

        # Create action buttons FIRST with side=BOTTOM (ensures visibility)
        self._setup_action_buttons()

        # Progress bar above buttons
        self._setup_progress_section()

        # Now create other sections from top
        self._setup_file_input_section()
        self._setup_check_cards()
        self._setup_log_section()

        # Setup path cleaning
        self.setup_path_cleaning(self.pbip_path_var)

    def _setup_file_input_section(self):
        """Setup the file input section using FileInputSection template"""
        # Get analyze icon
        analyze_icon = self._button_icons.get('magnifying-glass')

        # Create FileInputSection with all components
        # Accepts both PBIP files and PBIX files saved in PBIR format
        self.file_section = FileInputSection(
            parent=self.frame,
            theme_manager=self._theme_manager,
            section_title="Report Source",
            section_icon="Power-BI",
            file_label="Power BI File:",
            file_types=[("Power BI", "*.pbip *.pbix"), ("Power BI Project", "*.pbip"), ("Power BI File", "*.pbix")],
            action_button_text="ANALYZE REPORT",
            action_button_command=self._analyze_report,
            action_button_icon=analyze_icon,
            help_command=self.show_help_dialog
        )
        self.file_section.pack(fill=tk.X, pady=(0, 15))

        # Add settings button - positioned left of help button
        cogwheel_icon = self._button_icons.get('cogwheel')
        self._settings_button = SquareIconButton(
            self.file_section.section_frame, icon=cogwheel_icon, command=self._show_settings_dialog,
            tooltip_text="Settings", size=26, radius=6,
            bg_normal_override=AppConstants.CORNER_ICON_BG
        )
        self._settings_button.place(relx=1.0, y=-35, anchor=tk.NE, x=-30)

        # Store references for backward compatibility
        self.pbip_path_var = self.file_section.path_var
        self.analyze_button = self.file_section.action_button

    def _setup_check_cards(self):
        """Setup the accessibility check category cards"""
        colors = self._theme_manager.colors
        is_dark = self._theme_manager.is_dark
        # Outer background for filter (matches page background)
        outer_bg = colors.get('section_bg', colors['background'])

        # Create section frame FIRST
        header_widget = self.create_section_header(self.frame, "Accessibility Checks", "checker")[0]
        section_frame = ttk.LabelFrame(self.frame, labelwidget=header_widget,
                                       style='Section.TLabelframe', padding="12")
        section_frame.pack(fill=tk.X, pady=(0, 15))
        self._checks_section_frame = section_frame

        # Severity filter options
        config = get_config()
        self._severity_filter_var = tk.StringVar(value=config.min_severity)

        severity_options = [
            ("info", "All"),
            ("warning", "Important"),  # Critical + Should Fix (excludes Review)
            ("error", "Critical"),
        ]
        self._severity_display_map = {v: l for v, l in severity_options}
        self._severity_value_map = {l: v for v, l in severity_options}

        # Container for Show: label + dropdown - positioned ABOVE section using place()
        # Same positioning pattern as settings button above file input section
        filter_container = tk.Frame(self.frame, bg=outer_bg)
        self._filter_frame = filter_container

        # "Show:" label
        self._filter_label = tk.Label(
            filter_container, text="Show:",
            bg=outer_bg, fg=colors['text_secondary'],
            font=('Segoe UI', 9)
        )
        self._filter_label.pack(side=tk.LEFT, padx=(0, 4))

        # Custom themed dropdown button
        self._severity_dropdown_popup = None
        self._severity_dropdown_options = severity_options

        # Dropdown button (shows current selection with arrow)
        initial_display = self._severity_display_map.get(config.min_severity, "All")
        self._severity_dropdown_btn = tk.Label(
            filter_container,
            text=f" {initial_display}  \u25BC",
            bg=colors['card_surface'],
            fg=colors['text_primary'],
            font=('Segoe UI', 9),
            cursor='hand2',
            padx=6, pady=3,
            highlightbackground=colors['border'],
            highlightthickness=1
        )
        self._severity_dropdown_btn.pack(side=tk.LEFT)
        self._severity_dropdown_btn.bind('<Button-1>', self._toggle_severity_dropdown)

        # Position filter above section frame (like settings button above file input)
        filter_container.place(in_=section_frame, relx=1.0, y=-35, anchor=tk.NE, x=0)

        # Inner content frame with reduced padding for compactness
        content_frame = ttk.Frame(section_frame, style='Section.TFrame', padding="10")
        content_frame.pack(fill=tk.BOTH, expand=True)

        # Cards container - use background color not card_surface
        cards_frame = tk.Frame(content_frame, bg=colors['background'])
        cards_frame.pack(fill=tk.BOTH, expand=True)
        self._cards_frame = cards_frame

        # Configure columns for equal sizing (uniform ensures same width)
        for col in range(4):
            cards_frame.columnconfigure(col, weight=1, uniform='card')

        # Card definitions: (check_type, title, subtitle, icon_name)
        card_defs = [
            (AccessibilityCheckType.TAB_ORDER, "Tab Order", "Navigation sequence", "target"),
            (AccessibilityCheckType.ALT_TEXT, "Alt Text", "Screen reader text", "alt text"),
            (AccessibilityCheckType.COLOR_CONTRAST, "Contrast", "Color readability", "paint"),
            (AccessibilityCheckType.PAGE_TITLE, "Page Titles", "Descriptive names", "file"),
            (AccessibilityCheckType.VISUAL_TITLE, "Visual Titles", "Chart headings", "bar-chart"),
            (AccessibilityCheckType.DATA_LABELS, "Data Labels", "Chart labels", "tooltip"),
            (AccessibilityCheckType.BOOKMARK_NAME, "Bookmarks", "Bookmark names", "bookmark"),
            (AccessibilityCheckType.HIDDEN_PAGE, "Hidden Pages", "Visibility", "hidden"),
        ]

        # Create 2 rows of 4 cards
        for i, (check_type, title, subtitle, icon_name) in enumerate(card_defs):
            row = i // 4
            col = i % 4

            card_icon = self._button_icons.get(f"card_{icon_name}")
            card = AccessibilityCheckCard(
                cards_frame,
                title=title,
                subtitle=subtitle,
                check_type=check_type,
                on_card_click=self._on_card_clicked,
                icon=card_icon
            )
            card.grid(row=row, column=col, padx=10, pady=6, sticky='nsew')
            self.check_cards[check_type] = card

        # Configure row weights (uniform ensures same height for both rows)
        for row in range(2):
            cards_frame.rowconfigure(row, weight=1, uniform='cardrow')

    def _setup_log_section(self):
        """Setup the split log section with Details (left) and Progress Log (right)"""
        # Use SplitLogSection template widget
        self.log_section = SplitLogSection(
            parent=self.frame,
            theme_manager=self._theme_manager,
            section_title="Analysis & Progress",
            section_icon="analyze",
            summary_title="Issue Details",
            summary_icon="bar-chart",
            log_title="Progress Log",
            log_icon="log-file",
            summary_placeholder="Select an issue to see details"
        )
        self.log_section.pack(fill=tk.BOTH, expand=True, pady=(0, 0))

        # Store references for compatibility
        self.log_section_frame = self.log_section
        self._summary_frame = self.log_section.summary_frame
        self._placeholder_label = self.log_section.placeholder_label

        # Use the provided summary_text (scrolledtext) for details
        self.details_text = self.log_section.summary_text

        # Store reference to log text (right side)
        self.log_text = self.log_section.log_text

        # Show initial placeholder in details panel
        self._show_details_placeholder()

        # Add page filter controls to Issue Details header
        self._add_page_filter_to_header()

    def _add_page_filter_to_header(self):
        """Add page filter dropdown and icons to Issue Details header (right side)"""
        colors = self._theme_manager.colors

        # Get the summary header frame from the log section
        header_frame = self.log_section.summary_header_frame

        # Clear filter button (eraser) - pack FIRST with side=RIGHT so it's rightmost
        eraser_icon = self._load_icon_for_button("eraser", size=14)
        self._page_filter_clear_btn = SquareIconButton(
            header_frame, icon=eraser_icon,
            command=self._clear_page_filter,
            tooltip_text="Clear Page Filter", size=26, radius=6
        )
        self._page_filter_clear_btn.pack(side=tk.RIGHT, padx=(4, 0))
        self._button_icons['page_filter_clear'] = eraser_icon

        # Filter icon button - opens dropdown on click
        filter_icon = self._load_icon_for_button("filter", size=14)
        self._page_filter_icon = SquareIconButton(
            header_frame, icon=filter_icon,
            command=self._toggle_page_filter_popup,
            tooltip_text="Filter by Page", size=26, radius=6
        )
        self._page_filter_icon.pack(side=tk.RIGHT)
        self._button_icons['page_filter'] = filter_icon

    def _get_all_pages(self) -> List[str]:
        """Get all page display names for filtering.

        Returns list of display_name values (matching what AccessibilityIssue.page_name uses).
        Issues store display names in page_name field, so we use display names for filtering.
        """
        if not self.current_results or not self.current_results.pages:
            return []

        # Collect display names (what issues use for page_name)
        page_names = []
        for page in self.current_results.pages:
            display_name = page.display_name or page.page_name
            if display_name:
                page_names.append(display_name)

        # Sort alphabetically
        page_names.sort(key=lambda p: p.lower())
        return page_names

    def _get_page_display_name(self, page_name: str) -> str:
        """Get the display name for a page. Since issues use display names, just return as-is."""
        return page_name

    def _get_pages_with_issues_for_current_card(self) -> set:
        """Get set of page names that have issues for the currently selected check type"""
        if not self.current_results or not self.selected_check_type:
            return set()

        # Get all issues for current check type
        all_issues = self.current_results.get_issues_by_type(self.selected_check_type)

        # Apply severity filter (same as display logic)
        config = get_config()
        filtered_issues = [i for i in all_issues if config.should_show_severity(i.severity.value)]

        # Collect unique page names
        return {issue.page_name for issue in filtered_issues if issue.page_name}

    def _toggle_page_filter_popup(self, event=None):
        """Toggle the page filter dropdown popup visibility."""
        if self._page_filter_popup and self._page_filter_popup.winfo_exists():
            self._close_page_filter_popup()
        else:
            self._show_page_filter_popup()

    def _show_page_filter_popup(self, event=None):
        """Show the page filter dropdown popup with search and scrolling."""
        if self._page_filter_popup:
            self._close_page_filter_popup()

        colors = self._theme_manager.colors
        popup_bg = colors.get('surface', colors['card_surface'])

        # Get pages data
        all_pages = self._get_all_pages()
        pages_with_issues = self._get_pages_with_issues_for_current_card()

        # Create popup window
        self._page_filter_popup = tk.Toplevel(self.frame)
        self._page_filter_popup.withdraw()
        self._page_filter_popup.overrideredirect(True)

        # Border frame (1px border)
        border_frame = tk.Frame(
            self._page_filter_popup,
            bg=colors['border'],
            padx=1, pady=1
        )
        border_frame.pack(fill=tk.BOTH, expand=True)

        # Main content frame
        main_frame = tk.Frame(border_frame, bg=popup_bg)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Header row with label and search on same line
        header_frame = tk.Frame(main_frame, bg=popup_bg)
        header_frame.pack(fill=tk.X, padx=8, pady=(8, 0))

        header = tk.Label(
            header_frame,
            text="Filter by Page",
            font=('Segoe UI', 10, 'bold'),
            bg=popup_bg,
            fg=colors['text_primary']
        )
        header.pack(side=tk.LEFT)

        # Search container on the right
        search_container = tk.Frame(header_frame, bg=popup_bg)
        search_container.pack(side=tk.RIGHT, padx=(16, 0))

        # Magnifying glass icon
        search_icon = self._load_icon_for_button("magnifying-glass", size=14)
        search_icon_label = tk.Label(search_container, bg=popup_bg)
        if search_icon:
            search_icon_label.configure(image=search_icon)
            search_icon_label._icon_ref = search_icon
        search_icon_label.pack(side=tk.LEFT, padx=(0, 4))

        # Search entry with border
        entry_border = tk.Frame(search_container, bg=colors['border'])
        entry_border.pack(side=tk.LEFT)

        entry_bg = colors['background']
        entry_inner = tk.Frame(entry_border, bg=entry_bg)
        entry_inner.pack(padx=1, pady=1)

        self._page_filter_search_var = tk.StringVar()
        search_entry = tk.Entry(
            entry_inner,
            textvariable=self._page_filter_search_var,
            font=('Segoe UI', 9),
            width=20,
            bg=entry_bg,
            fg=colors['text_primary'],
            insertbackground=colors['text_primary'],
            relief=tk.FLAT,
            highlightthickness=0
        )
        search_entry.pack(padx=4, pady=4)

        # Bind search with debounce
        self._page_filter_search_var.trace_add('write', self._on_page_filter_search_changed)

        # Separator
        sep = tk.Frame(main_frame, bg=colors['border'], height=1)
        sep.pack(fill=tk.X, padx=8, pady=(8, 4))

        # Store page item widgets for search filtering
        self._page_filter_item_widgets = {}

        # Check if there are any pages
        if not all_pages:
            empty_frame = tk.Frame(main_frame, bg=popup_bg)
            empty_frame.pack(fill=tk.X, padx=8, pady=16)
            tk.Label(
                empty_frame,
                text="No pages available.\nRun analysis first.",
                font=('Segoe UI', 9, 'italic'),
                bg=popup_bg,
                fg=colors['text_muted'],
                justify=tk.CENTER
            ).pack(anchor=tk.CENTER)
            self._position_page_filter_popup()
            return

        # Scrollable content area with max height
        MAX_HEIGHT = 300
        MIN_WIDTH = 280

        # Create canvas for scrolling
        canvas_frame = tk.Frame(main_frame, bg=popup_bg)
        canvas_frame.pack(fill=tk.BOTH, expand=True, padx=8)

        canvas = tk.Canvas(canvas_frame, bg=popup_bg, highlightthickness=0, width=MIN_WIDTH)
        scrollbar = ThemedScrollbar(canvas_frame, command=canvas.yview,
                                    theme_manager=self._theme_manager, width=12)

        # Inner frame for content
        inner_frame = tk.Frame(canvas, bg=popup_bg)
        canvas_window = canvas.create_window((0, 0), window=inner_frame, anchor=tk.NW)

        # Store references for height recalculation
        self._page_filter_canvas = canvas
        self._page_filter_inner_frame = inner_frame
        self._page_filter_scrollbar = scrollbar
        self._page_filter_max_height = MAX_HEIGHT

        # "All Pages" option (always at top, always selectable)
        is_selected = self._page_filter_selection == "All Pages"
        all_frame = tk.Frame(inner_frame, bg=popup_bg)
        all_frame.pack(fill=tk.X, pady=2)

        all_label = tk.Label(
            all_frame,
            text="  All Pages",
            bg=popup_bg,
            fg=colors['text_primary'],
            font=('Segoe UI', 9, 'bold' if is_selected else 'normal'),
            anchor='w',
            padx=8, pady=3,
            cursor='hand2'
        )
        all_label.pack(fill=tk.X)
        all_label.bind('<Enter>', lambda e, l=all_label: l.configure(
            bg=colors['button_primary'], fg='#ffffff'))
        all_label.bind('<Leave>', lambda e, l=all_label: l.configure(
            bg=popup_bg, fg=colors['text_primary']))
        all_label.bind('<Button-1>', lambda e: self._select_page_filter("All Pages"))

        # Store in widgets dict for search (always visible)
        self._page_filter_item_widgets["All Pages"] = {
            'frame': all_frame,
            'label': all_label,
            'display_name': "All Pages",
            'always_visible': True
        }

        # Separator after All Pages
        sep2 = tk.Frame(inner_frame, height=1, bg=colors['border'])
        sep2.pack(fill=tk.X, pady=4)

        # Individual page options
        for page_name in all_pages:
            has_issues = page_name in pages_with_issues
            is_selected = self._page_filter_selection == page_name
            display_name = self._get_page_display_name(page_name)

            # Truncate long display names
            display_text = display_name if len(display_name) <= 30 else display_name[:27] + "..."

            page_frame = tk.Frame(inner_frame, bg=popup_bg)
            page_frame.pack(fill=tk.X, pady=1)

            opt_label = tk.Label(
                page_frame,
                text=f"  {display_text}",
                bg=popup_bg,
                fg=colors['text_primary'] if has_issues else colors['text_muted'],
                font=('Segoe UI', 9, 'bold' if is_selected else 'normal'),
                anchor='w',
                padx=8, pady=3,
                cursor='hand2' if has_issues else ''
            )
            opt_label.pack(fill=tk.X)

            if has_issues:
                # Selectable - add hover effects and click binding
                opt_label.bind('<Enter>', lambda e, l=opt_label: l.configure(
                    bg=colors['button_primary'], fg='#ffffff'))
                opt_label.bind('<Leave>', lambda e, l=opt_label, hi=has_issues: l.configure(
                    bg=popup_bg, fg=colors['text_primary'] if hi else colors['text_muted']))
                opt_label.bind('<Button-1>', lambda e, p=page_name: self._select_page_filter(p))

            # Store in widgets dict for search filtering
            self._page_filter_item_widgets[page_name] = {
                'frame': page_frame,
                'label': opt_label,
                'display_name': display_name,
                'has_issues': has_issues,
                'always_visible': False
            }

        # Update canvas scroll region after content is added
        inner_frame.update_idletasks()
        content_height = inner_frame.winfo_reqheight()

        # Only show scrollbar if content exceeds max height
        if content_height > MAX_HEIGHT:
            canvas.configure(height=MAX_HEIGHT, yscrollcommand=scrollbar.set)
            scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        else:
            canvas.configure(height=content_height)
            canvas.pack(fill=tk.BOTH, expand=True)

        canvas.configure(scrollregion=canvas.bbox("all"))

        # Configure canvas width to match frame
        def on_canvas_configure(event):
            canvas.itemconfig(canvas_window, width=event.width)
        canvas.bind('<Configure>', on_canvas_configure)

        # Mouse wheel scrolling
        def on_mousewheel(event):
            if content_height > MAX_HEIGHT:
                canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        canvas.bind_all('<MouseWheel>', on_mousewheel)
        self._page_filter_popup._on_mousewheel = on_mousewheel

        # Bottom padding
        tk.Frame(main_frame, bg=popup_bg, height=8).pack(fill=tk.X)

        # Position and show popup
        self._position_page_filter_popup()

    def _position_page_filter_popup(self):
        """Position popup below and right-anchored to filter button, then show it."""
        self._page_filter_popup.update_idletasks()
        popup_width = self._page_filter_popup.winfo_reqwidth()
        btn_right_x = self._page_filter_icon.winfo_rootx() + self._page_filter_icon.winfo_width()
        btn_y = self._page_filter_icon.winfo_rooty() + self._page_filter_icon.winfo_height()

        # Anchor right edge of popup to right edge of button
        popup_x = btn_right_x - popup_width
        self._page_filter_popup.geometry(f"+{popup_x}+{btn_y}")
        self._page_filter_popup.deiconify()
        self._page_filter_popup.lift()
        self._page_filter_popup.focus_set()

        # Bind click outside to close
        self.frame.winfo_toplevel().bind('<Button-1>', self._on_page_filter_click_outside, add='+')
        self.frame.winfo_toplevel().bind('<Configure>', self._on_page_filter_window_configure, add='+')

    def _on_page_filter_window_configure(self, event):
        """Handle parent window move/resize to keep dropdown anchored."""
        if not self._page_filter_popup or not self._page_filter_popup.winfo_exists():
            return

        try:
            popup_width = self._page_filter_popup.winfo_width()
            btn_right_x = self._page_filter_icon.winfo_rootx() + self._page_filter_icon.winfo_width()
            btn_y = self._page_filter_icon.winfo_rooty() + self._page_filter_icon.winfo_height()
            popup_x = btn_right_x - popup_width
            self._page_filter_popup.geometry(f"+{popup_x}+{btn_y}")
        except Exception:
            pass

    def _on_page_filter_search_changed(self, *args):
        """Handle search text change with debounce (300ms delay)."""
        if self._page_filter_search_after_id:
            try:
                self.frame.after_cancel(self._page_filter_search_after_id)
            except Exception:
                pass
        self._page_filter_search_after_id = self.frame.after(300, self._apply_page_filter_search)

    def _apply_page_filter_search(self):
        """Apply search filter to show/hide matching pages."""
        self._page_filter_search_after_id = None

        if not self._page_filter_search_var:
            return

        search_text = self._page_filter_search_var.get().strip().lower()

        if not search_text:
            # Show all items
            for widgets in self._page_filter_item_widgets.values():
                widgets['frame'].pack(fill=tk.X, pady=1 if not widgets.get('always_visible') else 2)
            return

        # Filter items by display name
        for page_key, widgets in self._page_filter_item_widgets.items():
            display_name = widgets['display_name'].lower()
            if widgets.get('always_visible') or search_text in display_name:
                widgets['frame'].pack(fill=tk.X, pady=1 if not widgets.get('always_visible') else 2)
            else:
                widgets['frame'].pack_forget()

    def _on_page_filter_click_outside(self, event):
        """Handle click outside the page filter dropdown"""
        if not self._page_filter_popup or not self._page_filter_popup.winfo_exists():
            return

        x, y = event.x_root, event.y_root

        # Check if click is inside dropdown popup
        dx = self._page_filter_popup.winfo_rootx()
        dy = self._page_filter_popup.winfo_rooty()
        dw = self._page_filter_popup.winfo_width()
        dh = self._page_filter_popup.winfo_height()

        # Also check if click is on the filter icon button (to allow toggle)
        bx = self._page_filter_icon.winfo_rootx()
        by = self._page_filter_icon.winfo_rooty()
        bw = self._page_filter_icon.winfo_width()
        bh = self._page_filter_icon.winfo_height()

        click_in_dropdown = dx <= x <= dx + dw and dy <= y <= dy + dh
        click_in_button = bx <= x <= bx + bw and by <= y <= by + bh

        if not click_in_dropdown and not click_in_button:
            self._close_page_filter_popup()

    def _close_page_filter_popup(self):
        """Close the page filter popup"""
        if self._page_filter_popup and self._page_filter_popup.winfo_exists():
            try:
                self._page_filter_popup.unbind_all('<MouseWheel>')
            except Exception:
                pass
            self._page_filter_popup.destroy()
        self._page_filter_popup = None

        # Unbind handlers
        try:
            self.frame.winfo_toplevel().unbind('<Button-1>')
            self.frame.winfo_toplevel().unbind('<Configure>')
        except Exception:
            pass

        # Cancel any pending search debounce
        if self._page_filter_search_after_id:
            try:
                self.frame.after_cancel(self._page_filter_search_after_id)
            except Exception:
                pass
            self._page_filter_search_after_id = None

    def _select_page_filter(self, page_name: str):
        """Apply page filter selection"""
        self._page_filter_selection = page_name

        # Close the popup
        self._close_page_filter_popup()

        # Refresh the issue details display with the filter applied
        if self.selected_check_type:
            self._show_card_details(self.selected_check_type)

    def _clear_page_filter(self):
        """Clear the page filter and show all pages"""
        self._page_filter_selection = "All Pages"

        # Close popup if open
        self._close_page_filter_popup()

        # Refresh the issue details display
        if self.selected_check_type:
            self._show_card_details(self.selected_check_type)

    def _show_details_placeholder(self):
        """Show placeholder text in details panel - hide text widget, show placeholder label"""
        # Hide the details text widget
        if hasattr(self, 'details_text') and self.details_text:
            self.details_text.grid_remove()

        # Show the placeholder label (centered)
        if hasattr(self, '_placeholder_label') and self._placeholder_label:
            self._placeholder_label.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.E, tk.W))

    def _setup_progress_section(self):
        """Setup the progress bar section - positioned above bottom buttons"""
        self.progress_components = self.create_progress_bar(self.frame)
        # Initially hidden
        self.progress_components['frame'].pack_forget()

    def _position_progress_frame(self):
        """Position the progress frame appropriately for this layout"""
        if self.progress_components and self.progress_components['frame']:
            # Position above the action buttons
            self.progress_components['frame'].pack(side=tk.BOTTOM, fill=tk.X, pady=(10, 5))

    def _setup_action_buttons(self):
        """Setup the action buttons at the bottom"""
        colors = self._theme_manager.colors
        is_dark = self._theme_manager.is_dark

        # Use outer background color for button frame
        outer_canvas_bg = colors.get('outer_bg', '#161627' if is_dark else '#f5f5f7')

        # Button frame - pack to bottom first
        button_frame = tk.Frame(self.frame, bg=outer_canvas_bg)
        button_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(10, 0))
        self._button_frame = button_frame

        # Center the buttons
        button_container = tk.Frame(button_frame, bg=outer_canvas_bg)
        button_container.pack(anchor=tk.CENTER, pady=5)
        self._button_container = button_container

        # Export CSV button - primary style with icon
        csv_icon = self._button_icons.get('csv-file')
        self.export_button = RoundedButton(
            button_container, text="EXPORT CSV", icon=csv_icon,
            height=38, radius=6,
            bg=colors['button_primary'], fg='#ffffff',
            hover_bg=colors['button_primary_hover'],
            pressed_bg=colors['button_primary_pressed'],
            disabled_bg=colors.get('button_primary_disabled', '#3a3a4e' if is_dark else '#c0c0cc'),
            disabled_fg=colors.get('button_text_disabled', '#6a6a7a' if is_dark else '#9a9aa8'),
            font=('Segoe UI', 10, 'bold'),
            command=self._export_report,
            canvas_bg=outer_canvas_bg
        )
        self.export_button.pack(side=tk.LEFT, padx=(0, 15))
        self.export_button.set_enabled(False)
        self._primary_buttons.append(self.export_button)

        # Export PDF/HTML button - primary style with icon
        pdf_icon = self._button_icons.get('pdf')
        self.export_pdf_button = RoundedButton(
            button_container, text="EXPORT PDF", icon=pdf_icon,
            height=38, radius=6,
            bg=colors['button_primary'], fg='#ffffff',
            hover_bg=colors['button_primary_hover'],
            pressed_bg=colors['button_primary_pressed'],
            disabled_bg=colors.get('button_primary_disabled', '#3a3a4e' if is_dark else '#c0c0cc'),
            disabled_fg=colors.get('button_text_disabled', '#6a6a7a' if is_dark else '#9a9aa8'),
            font=('Segoe UI', 10, 'bold'),
            command=self._export_html_report,
            canvas_bg=outer_canvas_bg
        )
        self.export_pdf_button.pack(side=tk.LEFT, padx=(0, 15))
        self.export_pdf_button.set_enabled(False)
        self._primary_buttons.append(self.export_pdf_button)

        # Reset button - secondary style
        reset_icon = self._button_icons.get('reset')
        self.reset_button = RoundedButton(
            button_container, text="RESET ALL", icon=reset_icon,
            height=38, radius=6,
            bg=colors['button_secondary'], fg=colors['text_primary'],
            hover_bg=colors['button_secondary_hover'],
            pressed_bg=colors['button_secondary_pressed'],
            font=('Segoe UI', 10),
            command=self.reset_tab,
            canvas_bg=outer_canvas_bg
        )
        self.reset_button.pack(side=tk.LEFT)
        self._secondary_buttons.append(self.reset_button)

    def _toggle_severity_dropdown(self, event=None):
        """Toggle the severity filter dropdown visibility"""
        if self._severity_dropdown_popup and self._severity_dropdown_popup.winfo_exists():
            self._close_severity_dropdown()
        else:
            self._open_severity_dropdown()

    def _open_severity_dropdown(self):
        """Open the custom severity filter dropdown"""
        if self._severity_dropdown_popup:
            self._close_severity_dropdown()

        colors = self._theme_manager.colors

        # Create popup window
        self._severity_dropdown_popup = tk.Toplevel(self.frame)
        self._severity_dropdown_popup.withdraw()
        self._severity_dropdown_popup.overrideredirect(True)

        # Border frame (1px border)
        border_frame = tk.Frame(
            self._severity_dropdown_popup,
            bg=colors['border'],
            padx=1, pady=1
        )
        border_frame.pack(fill=tk.BOTH, expand=True)

        # Content frame
        content_frame = tk.Frame(border_frame, bg=colors['card_surface'])
        content_frame.pack(fill=tk.BOTH, expand=True)

        # Store option labels for theme updates
        self._severity_option_labels = []

        # Create option items
        for value, display in self._severity_dropdown_options:
            opt_label = tk.Label(
                content_frame,
                text=f"  {display}",
                bg=colors['card_surface'],
                fg=colors['text_primary'],
                font=('Segoe UI', 9),
                anchor='w',
                padx=8, pady=4,
                cursor='hand2'
            )
            opt_label.pack(fill=tk.X)

            # Hover effect
            opt_label.bind('<Enter>', lambda e, l=opt_label: l.configure(
                bg=colors['button_primary'], fg='#ffffff'))
            opt_label.bind('<Leave>', lambda e, l=opt_label: l.configure(
                bg=colors['card_surface'], fg=colors['text_primary']))
            opt_label.bind('<Button-1>', lambda e, v=value, d=display: self._on_severity_option_selected(v, d))

            self._severity_option_labels.append(opt_label)

        # Position popup below the button
        self._severity_dropdown_popup.update_idletasks()
        btn = self._severity_dropdown_btn
        btn_x = btn.winfo_rootx()
        btn_y = btn.winfo_rooty() + btn.winfo_height()

        self._severity_dropdown_popup.geometry(f"+{btn_x}+{btn_y}")
        self._severity_dropdown_popup.deiconify()
        self._severity_dropdown_popup.lift()

        # Bind click outside to close
        self.frame.winfo_toplevel().bind('<Button-1>', self._on_dropdown_click_outside, add='+')

    def _on_dropdown_click_outside(self, event):
        """Handle click outside the dropdown"""
        if not self._severity_dropdown_popup or not self._severity_dropdown_popup.winfo_exists():
            return

        # Check if click is inside dropdown
        dx = self._severity_dropdown_popup.winfo_rootx()
        dy = self._severity_dropdown_popup.winfo_rooty()
        dw = self._severity_dropdown_popup.winfo_width()
        dh = self._severity_dropdown_popup.winfo_height()

        # Also check if click is on the button (to allow toggle)
        bx = self._severity_dropdown_btn.winfo_rootx()
        by = self._severity_dropdown_btn.winfo_rooty()
        bw = self._severity_dropdown_btn.winfo_width()
        bh = self._severity_dropdown_btn.winfo_height()

        click_in_dropdown = dx <= event.x_root <= dx + dw and dy <= event.y_root <= dy + dh
        click_in_button = bx <= event.x_root <= bx + bw and by <= event.y_root <= by + bh

        if not click_in_dropdown and not click_in_button:
            self._close_severity_dropdown()

    def _close_severity_dropdown(self):
        """Close the severity dropdown popup"""
        if self._severity_dropdown_popup and self._severity_dropdown_popup.winfo_exists():
            self._severity_dropdown_popup.destroy()
        self._severity_dropdown_popup = None

        # Unbind click outside handler
        try:
            self.frame.winfo_toplevel().unbind('<Button-1>')
        except Exception:
            pass

    def _on_severity_option_selected(self, value, display):
        """Handle selection of a severity filter option"""
        # Update button text
        self._severity_dropdown_btn.configure(text=f" {display}  \u25BC")

        # Close dropdown
        self._close_severity_dropdown()

        # Update config
        config = get_config()
        config.min_severity = value
        save_config()

        # Update the internal variable
        self._severity_filter_var.set(value)

        # Re-display current card details with new filter if showing details
        if self.selected_check_type and self.current_results:
            self._show_card_details(self.selected_check_type)

        # Update analysis summary if showing
        if self.current_results and not self.selected_check_type:
            self._show_analysis_summary(self.current_results)

    def _update_progress_safe(self, percent: int, message: str):
        """Thread-safe progress update"""
        self.frame.after(0, lambda: self.update_progress(percent, message))

    def _analyze_report(self):
        """Run accessibility analysis in background thread"""
        pbip_path = self.pbip_path_var.get().strip()

        if not pbip_path:
            ThemedMessageBox.showwarning(self.frame.winfo_toplevel(), "Missing File", "Please select a Power BI file first.")
            return

        # Validate file exists
        if not Path(pbip_path).exists():
            ThemedMessageBox.showerror(self.frame.winfo_toplevel(), "File Not Found", f"The file does not exist:\n{pbip_path}")
            return

        # Disable UI during analysis
        self.analyze_button.set_enabled(False)

        # Clear previous results
        for card in self.check_cards.values():
            card.reset()

        # Show progress
        self.update_progress(5, "Starting analysis...", show=True)

        self.log_message("")
        self.log_message("🚀 Starting accessibility analysis...")
        self.log_message(f"📁 Analyzing: {pbip_path}")

        # Run analysis in background
        def run_analysis():
            try:
                result = self.analyzer.analyze_pbip_report(pbip_path)
                self.frame.after(0, lambda: self._on_analysis_complete(result))
            except Exception as e:
                error_msg = str(e)
                self.frame.after(0, lambda msg=error_msg: self._on_analysis_error(msg))

        thread = threading.Thread(target=run_analysis, daemon=True)
        thread.start()

    def _on_analysis_complete(self, result: AccessibilityAnalysisResult):
        """Handle completed analysis"""
        self.current_results = result

        # Calculate totals for each check type
        self._check_totals = {
            AccessibilityCheckType.TAB_ORDER: len(result.visuals),
            AccessibilityCheckType.ALT_TEXT: len([v for v in result.visuals if v.is_data_visual]),
            AccessibilityCheckType.COLOR_CONTRAST: len(result.color_contrasts) if result.color_contrasts else len(result.visuals),
            AccessibilityCheckType.PAGE_TITLE: len(result.pages),
            AccessibilityCheckType.VISUAL_TITLE: len([v for v in result.visuals if v.is_data_visual]),
            AccessibilityCheckType.DATA_LABELS: len([v for v in result.visuals if v.is_data_visual]),
            AccessibilityCheckType.BOOKMARK_NAME: len(result.bookmarks),
            AccessibilityCheckType.HIDDEN_PAGE: len(result.pages),
        }

        # Update cards with results
        for check_type, card in self.check_cards.items():
            issues = result.get_issues_by_type(check_type)
            total_count = self._check_totals.get(check_type, 0)
            has_errors = any(i.severity == AccessibilitySeverity.ERROR for i in issues)
            has_warnings = any(i.severity == AccessibilitySeverity.WARNING for i in issues)

            # For contrast, count unique visuals with issues (not total element failures)
            if check_type == AccessibilityCheckType.COLOR_CONTRAST:
                visuals_with_issues = set()
                for issue in issues:
                    if issue.visual_name:
                        visuals_with_issues.add((issue.page_name, issue.visual_name))
                issue_count = len(visuals_with_issues)
            else:
                issue_count = len(issues)

            card.update_results(issue_count, total_count, has_errors, has_warnings)

        # Update details panel with summary
        self._show_analysis_summary(result)

        # Log summary
        self.log_message("")
        self.log_message("✅ Analysis completed successfully!")
        self.log_message("")
        self.log_message(f"📊 Found {result.total_issues} accessibility issues:")
        self.log_message(f"   🔴 {result.errors} Critical (must fix for WCAG compliance)")
        self.log_message(f"   🟡 {result.warnings} Should Fix (recommended improvements)")
        self.log_message(f"   🔵 {result.info_count} Review (manual verification needed)")
        self.log_message("")
        self.log_message("💡 Click on a card above to view detailed issues for that category")

        # Re-enable UI
        self.analyze_button.set_enabled(True)
        self.export_button.set_enabled(True)  # Enable export even if no issues (to export pass report)
        self.export_pdf_button.set_enabled(True)
        self.update_progress(100, "Analysis complete!", show=True)

    def _show_analysis_summary(self, result: AccessibilityAnalysisResult):
        """Show analysis summary in the details panel"""
        if not hasattr(self, 'details_text') or not self.details_text:
            return

        colors = self._theme_manager.colors

        # Hide placeholder and show details text widget
        if hasattr(self, '_placeholder_label') and self._placeholder_label:
            self._placeholder_label.grid_remove()
        self.details_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Build summary text
        self.details_text.config(state=tk.NORMAL)
        self.details_text.delete(1.0, tk.END)

        self.details_text.insert(tk.END, f"Accessibility Analysis Summary\n", 'header')
        self.details_text.insert(tk.END, "─" * 35 + "\n\n", 'separator')

        if result.total_issues > 0:
            self.details_text.insert(tk.END, f"Found {result.total_issues} total issues:\n\n", 'item')
            self.details_text.insert(tk.END, f"  🔴 Critical: {result.errors}\n", 'error')
            self.details_text.insert(tk.END, f"  🟡 Should Fix: {result.warnings}\n", 'warning')
            self.details_text.insert(tk.END, f"  🔵 Review: {result.info_count}\n\n", 'info')

            self.details_text.insert(tk.END, "Issues by Category:\n", 'subheader')
            for check_type in AccessibilityCheckType:
                issues = result.get_issues_by_type(check_type)
                count = len(issues)
                total = self._check_totals.get(check_type, 0)
                name = CHECK_TYPE_DISPLAY_NAMES.get(check_type, check_type.value)
                status = "✅" if count == 0 else "⚠️" if count < 5 else "❌"
                self.details_text.insert(tk.END, f"  {status} {name}: {count}/{total}\n", 'item')

            self.details_text.insert(tk.END, "\n💡 Click a card above to see details\n", 'tip')
        else:
            self.details_text.insert(tk.END, "🎉 No accessibility issues found!\n\n", 'success')
            self.details_text.insert(tk.END, "Your report passes the basic accessibility checks.\n", 'item')

        # Configure text tags
        # Only non-colored text (black/white) gets selectforeground so it's readable when selected
        # Colored text (error, warning, info, success) keeps its color even when selected
        select_fg = '#ffffff'  # White text when selected (for non-colored text only)
        self.details_text.tag_config('header', font=('Segoe UI', 11, 'bold'), foreground=colors['text_primary'], selectforeground=select_fg)
        self.details_text.tag_config('subheader', font=('Segoe UI', 10, 'bold'), foreground=colors['text_primary'], selectforeground=select_fg)
        self.details_text.tag_config('separator', foreground=colors['text_muted'], selectforeground=select_fg)
        self.details_text.tag_config('item', font=('Segoe UI', 9), foreground=colors['text_primary'], selectforeground=select_fg)
        self.details_text.tag_config('tip', font=('Segoe UI', 9, 'italic'), foreground=colors['text_secondary'], selectforeground=select_fg)
        # Colored tags - no selectforeground so they keep their distinctive colors when selected
        self.details_text.tag_config('error', font=('Segoe UI', 9), foreground=colors['error'])
        self.details_text.tag_config('warning', font=('Segoe UI', 9), foreground=colors['warning'])
        self.details_text.tag_config('info', font=('Segoe UI', 9), foreground=colors['info'])
        self.details_text.tag_config('success', font=('Segoe UI', 10, 'bold'), foreground=colors['success'])

        # Scroll to top and finalize
        self.details_text.see("1.0")
        self.details_text.config(state=tk.DISABLED)
        self.details_text.update_idletasks()

    def _on_analysis_error(self, error_message: str):
        """Handle analysis error"""
        # Check for PBIR format error (special case with Learn More button)
        if error_message.startswith("PBIR_FORMAT_ERROR:"):
            # Strip the marker prefix for display
            display_message = error_message.replace("PBIR_FORMAT_ERROR:", "")
            self.log_message(f"❌ Analysis failed: {display_message}")

            # Show dialog with Learn More button
            result = ThemedMessageBox.show(
                self.frame.winfo_toplevel(),
                "Analysis Error",
                f"An error occurred during analysis:\n{display_message}",
                msg_type="warning",
                buttons=["OK", "Learn More"]
            )

            # If user clicked Learn More, open the Microsoft documentation
            if result == "Learn More":
                import webbrowser
                webbrowser.open("https://learn.microsoft.com/en-us/power-bi/developer/projects/projects-report?tabs=v2%2Cdesktop#pbir-format")
        else:
            self.log_message(f"❌ Analysis failed: {error_message}")
            ThemedMessageBox.showerror(self.frame.winfo_toplevel(), "Analysis Error", f"An error occurred during analysis:\n{error_message}")

        self.analyze_button.set_enabled(True)
        self.update_progress(0, "", show=False)

    def _on_card_clicked(self, check_type: AccessibilityCheckType):
        """Show details for clicked card"""
        if not self.current_results:
            return

        # Update card selection state
        for ct, card in self.check_cards.items():
            card.set_selected(ct == check_type)

        self.selected_check_type = check_type

        # Reset page filter when switching cards
        self._page_filter_selection = "All Pages"

        # Show details for the clicked card
        self._show_card_details(check_type)

    def _show_card_details(self, check_type: AccessibilityCheckType):
        """Show details for the selected card category in the details panel"""
        if not hasattr(self, 'details_text') or not self.details_text:
            return

        colors = self._theme_manager.colors

        # Hide placeholder and show details text widget
        if hasattr(self, '_placeholder_label') and self._placeholder_label:
            self._placeholder_label.grid_remove()
        self.details_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Get issues for this check type
        all_issues = self.current_results.get_issues_by_type(check_type)

        # Filter by severity
        config = get_config()
        severity_filtered = [i for i in all_issues if config.should_show_severity(i.severity.value)]

        # Filter by page (if page filter is active)
        if self._page_filter_selection != "All Pages":
            issues = [i for i in severity_filtered if i.page_name == self._page_filter_selection]
        else:
            issues = severity_filtered

        # Build details text
        self.details_text.config(state=tk.NORMAL)
        self.details_text.delete(1.0, tk.END)

        title = CHECK_TYPE_DISPLAY_NAMES.get(check_type, check_type.value)
        total = self._check_totals.get(check_type, 0)

        # Show filter info if any filters are active
        filters_active = []
        if len(severity_filtered) != len(all_issues):
            filters_active.append(self._severity_display_map.get(config.min_severity, "Severity"))
        if self._page_filter_selection != "All Pages":
            page_display_name = self._get_page_display_name(self._page_filter_selection)
            page_display = page_display_name if len(page_display_name) <= 20 else page_display_name[:17] + "..."
            filters_active.append(f"Page: {page_display}")

        # For contrast, count unique visuals instead of individual element issues
        if check_type == AccessibilityCheckType.COLOR_CONTRAST:
            shown_visuals = {(i.page_name, i.visual_name) for i in issues if i.visual_name}
            all_visuals = {(i.page_name, i.visual_name) for i in all_issues if i.visual_name}
            shown_count = len(shown_visuals)
            all_count = len(all_visuals)
            unit = "visual" if shown_count == 1 else "visuals"
        else:
            shown_count = len(issues)
            all_count = len(all_issues)
            unit = "issue" if shown_count == 1 else "issues"

        if filters_active:
            self.details_text.insert(tk.END, f"{title} ({shown_count} of {all_count} {unit} shown / {total} checked)\n", 'header')
            self.details_text.insert(tk.END, f"Filter: {', '.join(filters_active)}\n", 'tip')
        else:
            self.details_text.insert(tk.END, f"{title} ({shown_count} {unit} / {total} checked)\n", 'header')
        self.details_text.insert(tk.END, "─" * 35 + "\n\n", 'separator')

        if not issues:
            self.details_text.insert(tk.END, "✅ No issues found for this category.\n", 'success')
        elif check_type == AccessibilityCheckType.COLOR_CONTRAST:
            # Special hierarchical display for contrast issues - group by visual
            self._display_contrast_issues_hierarchical(issues, colors)
        else:
            for i, issue in enumerate(issues, 1):
                # Severity tag - map to new terminology
                severity_tag = issue.severity.value
                severity_display, _ = self._get_severity_display(issue.severity)

                # Issue header
                self.details_text.insert(tk.END, f"{i}. ", 'header')
                self.details_text.insert(tk.END, f"[{severity_display.upper()}] ", severity_tag)
                self.details_text.insert(tk.END, f"{issue.issue_description}\n")

                # Details
                if issue.page_name:
                    self.details_text.insert(tk.END, f"   Page: {issue.page_name}\n", 'item')
                if issue.visual_name:
                    self.details_text.insert(tk.END, f"   Visual: {issue.visual_name}", 'item')
                    if issue.visual_type:
                        self.details_text.insert(tk.END, f" ({issue.visual_type})", 'item')
                    self.details_text.insert(tk.END, "\n")
                if issue.current_value:
                    # Special handling for color contrast - render color swatches
                    if issue.check_type == AccessibilityCheckType.COLOR_CONTRAST and '|' in issue.current_value:
                        parts = issue.current_value.split('|')
                        if len(parts) == 3:
                            element_type, fg_color, bg_color = parts
                            self.details_text.insert(tk.END, f"   Elements: {element_type}\n", 'item')
                            self.details_text.insert(tk.END, f"   Colors: ", 'item')
                            # Create color swatches as embedded canvas widgets with clean borders
                            border_color = colors.get('border', '#3d3d5c')
                            swatch_size = 16
                            border_width = 1
                            # Get internal text widget for embedding (required for window_create)
                            text_parent = self.details_text.text_widget

                            # Foreground color swatch
                            fg_canvas = tk.Canvas(
                                text_parent,
                                width=swatch_size,
                                height=swatch_size,
                                highlightthickness=border_width,
                                highlightbackground=border_color,
                                bg=fg_color,
                                bd=0
                            )
                            self.details_text.window_create(tk.END, window=fg_canvas)
                            self.details_text.insert(tk.END, f" {fg_color}", 'item')
                            self.details_text.insert(tk.END, "  →  ", 'item')

                            # Background color swatch
                            bg_canvas = tk.Canvas(
                                text_parent,
                                width=swatch_size,
                                height=swatch_size,
                                highlightthickness=border_width,
                                highlightbackground=border_color,
                                bg=bg_color,
                                bd=0
                            )
                            self.details_text.window_create(tk.END, window=bg_canvas)
                            self.details_text.insert(tk.END, f" {bg_color}\n", 'item')
                        else:
                            self.details_text.insert(tk.END, f"   Current: {issue.current_value}\n", 'item')
                    else:
                        self.details_text.insert(tk.END, f"   Current: {issue.current_value}\n", 'item')

                # Recommendation
                self.details_text.insert(tk.END, f"   💡 ", 'tip')
                self.details_text.insert(tk.END, f"{issue.recommendation}\n", 'tip')

                # WCAG reference
                if issue.wcag_reference:
                    self.details_text.insert(tk.END, f"   📖 {issue.wcag_reference}\n", 'info')

                self.details_text.insert(tk.END, "\n")

        # Configure text tags
        self.details_text.tag_config('header', font=('Segoe UI', 11, 'bold'), foreground=colors['text_primary'])
        self.details_text.tag_config('subheader', font=('Segoe UI', 10, 'bold'), foreground=colors['text_primary'])
        self.details_text.tag_config('separator', foreground=colors['text_muted'])
        self.details_text.tag_config('item', font=('Segoe UI', 9), foreground=colors['text_primary'])
        self.details_text.tag_config('error', font=('Segoe UI', 9, 'bold'), foreground=colors['error'])
        self.details_text.tag_config('warning', font=('Segoe UI', 9, 'bold'), foreground=colors['warning'])
        self.details_text.tag_config('info', font=('Segoe UI', 9), foreground=colors['info'])
        self.details_text.tag_config('success', font=('Segoe UI', 10, 'bold'), foreground=colors['success'])
        self.details_text.tag_config('tip', font=('Segoe UI', 9), foreground=colors['success'])

        # Scroll to top and finalize
        self.details_text.see("1.0")
        self.details_text.config(state=tk.DISABLED)
        self.details_text.update_idletasks()

    def _display_contrast_issues_hierarchical(self, issues: List[AccessibilityIssue], colors: Dict[str, Any]):
        """Display contrast issues grouped by visual with collapsible sections.

        Uses text widget elide tags to hide/show children while preserving rich formatting
        including color swatches, severity badges, and tips.

        Groups issues by (page_name, visual_name) and displays:
        - Visual as collapsible parent (collapsed by default) with worst ratio and element count
        - Element details as children with colors and ratios
        """
        from collections import defaultdict
        import re

        # Always fetch fresh colors from theme manager to avoid stale color issues
        colors = self._theme_manager.colors

        # Group issues by (page_name, visual_name)
        issues_by_visual = defaultdict(list)
        for issue in issues:
            key = (issue.page_name or "Unknown Page", issue.visual_name or "Unknown Visual")
            issues_by_visual[key].append(issue)

        # Get internal text widget for embedding canvases
        text_parent = self.details_text.text_widget
        border_color = colors.get('border', '#3d3d5c')
        swatch_size = 16
        border_width = 1

        # Initialize collapsed state tracking if needed
        if not hasattr(self, '_collapsed_visuals'):
            self._collapsed_visuals = {}

        # Store canvas references to prevent garbage collection
        if not hasattr(self, '_contrast_canvases'):
            self._contrast_canvases = []
        else:
            self._contrast_canvases.clear()

        # Add hint text for expand/collapse functionality
        self.details_text.insert(tk.END, "Click ", 'hint_text')
        self.details_text.insert(tk.END, "▶", 'hint_icon')
        self.details_text.insert(tk.END, " to expand visual details, ", 'hint_text')
        self.details_text.insert(tk.END, "▼", 'hint_icon')
        self.details_text.insert(tk.END, " to collapse\n\n", 'hint_text')

        # Configure hint text styles
        hint_icon_color = colors['button_primary']  # Teal in light mode, blue in dark mode
        self.details_text.tag_configure('hint_text', font=('Segoe UI', 9), foreground=colors['text_muted'])
        self.details_text.tag_configure('hint_icon', font=('Segoe UI', 11, 'bold'), foreground=hint_icon_color)

        visual_index = 0
        for (page_name, visual_name), visual_issues in sorted(issues_by_visual.items()):
            children_tag = f"children_{visual_index}"
            header_tag = f"header_{visual_index}"

            # Calculate visual summary (worst ratio, element count)
            worst_ratio = float('inf')
            worst_severity = AccessibilitySeverity.INFO
            for issue in visual_issues:
                ratio_match = re.search(r'\((\d+\.?\d*):1\)', issue.issue_description)
                if ratio_match:
                    ratio = float(ratio_match.group(1))
                    if ratio < worst_ratio:
                        worst_ratio = ratio
                if issue.severity.value == 'error':
                    worst_severity = issue.severity
                elif issue.severity.value == 'warning' and worst_severity.value != 'error':
                    worst_severity = issue.severity

            # Default collapsed state (start collapsed)
            if children_tag not in self._collapsed_visuals:
                self._collapsed_visuals[children_tag] = True

            is_collapsed = self._collapsed_visuals[children_tag]

            # Visual header with expand/collapse icon
            icon = "▶" if is_collapsed else "▼"
            visual_type = visual_issues[0].visual_type if visual_issues else ""
            severity_display, _ = self._get_severity_display(worst_severity)

            # Insert visual header
            self.details_text.insert(tk.END, f"{visual_index + 1}. ", 'header')
            self.details_text.insert(tk.END, f"[{severity_display.upper()}] ", worst_severity.value)

            # Expand/collapse icon (clickable)
            self.details_text.insert(tk.END, f"{icon} ", header_tag)

            # Visual name and type
            type_suffix = f" ({visual_type})" if visual_type else ""
            self.details_text.insert(tk.END, f"{visual_name}{type_suffix}\n", 'subheader')

            # Summary line
            self.details_text.insert(tk.END, f"   Page: {page_name}\n", 'item')
            ratio_text = f"{worst_ratio:.2f}:1" if worst_ratio != float('inf') else "?"
            self.details_text.insert(tk.END, f"   Worst Ratio: {ratio_text} | {len(visual_issues)} element(s) failing\n", 'item')

            # Mark start of children content
            children_start = self.details_text.index(tk.END)

            # Element children (will be tagged for elide)
            for i, issue in enumerate(visual_issues):
                # Element sub-header
                self.details_text.insert(tk.END, f"\n   {visual_index + 1}.{i + 1} ", 'item')

                # Parse element info from current_value
                if issue.current_value and '|' in issue.current_value:
                    parts = issue.current_value.split('|')
                    if len(parts) == 3:
                        element_type, fg_color, bg_color = parts
                        ratio_match = re.search(r'\((\d+\.?\d*):1\)', issue.issue_description)
                        elem_ratio = ratio_match.group(1) if ratio_match else "?"

                        # Element type header
                        self.details_text.insert(tk.END, f"{element_type}\n", 'subheader')
                        self.details_text.insert(tk.END, f"      Ratio: {elem_ratio}:1\n", 'item')

                        # Color swatches (embedded canvases)
                        self.details_text.insert(tk.END, "      Colors: ", 'item')

                        # Foreground color swatch
                        fg_canvas = tk.Canvas(
                            text_parent,
                            width=swatch_size,
                            height=swatch_size,
                            highlightthickness=border_width,
                            highlightbackground=border_color,
                            bg=fg_color,
                            bd=0
                        )
                        self._contrast_canvases.append(fg_canvas)
                        self.details_text.window_create(tk.END, window=fg_canvas)
                        self.details_text.insert(tk.END, f" {fg_color}", 'item')
                        self.details_text.insert(tk.END, "  →  ", 'item')

                        # Background color swatch
                        bg_canvas = tk.Canvas(
                            text_parent,
                            width=swatch_size,
                            height=swatch_size,
                            highlightthickness=border_width,
                            highlightbackground=border_color,
                            bg=bg_color,
                            bd=0
                        )
                        self._contrast_canvases.append(bg_canvas)
                        self.details_text.window_create(tk.END, window=bg_canvas)
                        self.details_text.insert(tk.END, f" {bg_color}\n", 'item')
                    else:
                        self.details_text.insert(tk.END, f"{issue.issue_description[:60]}\n", 'item')
                else:
                    self.details_text.insert(tk.END, f"{issue.issue_description[:60]}\n", 'item')

                # Recommendation
                self.details_text.insert(tk.END, f"      💡 ", 'tip')
                self.details_text.insert(tk.END, f"{issue.recommendation}\n", 'tip')

            # Mark end of children content
            children_end = self.details_text.index(tk.END)

            # Apply children tag for elide functionality
            self.details_text.tag_add(children_tag, children_start, children_end)
            if is_collapsed:
                self.details_text.tag_configure(children_tag, elide=True)

            # Configure clickable header icon
            self.details_text.tag_configure(header_tag, foreground=colors['button_primary'], font=('Segoe UI', 10, 'bold'))
            self.details_text.tag_bind(header_tag, '<Button-1>',
                lambda e, tag=children_tag, idx=visual_index: self._toggle_visual_collapse(tag, idx))
            self.details_text.tag_bind(header_tag, '<Enter>',
                lambda e: self.details_text.config(cursor='hand2'))
            self.details_text.tag_bind(header_tag, '<Leave>',
                lambda e: self.details_text.config(cursor=''))

            # Add spacing between visuals
            self.details_text.insert(tk.END, "\n")
            visual_index += 1

    def _toggle_visual_collapse(self, children_tag: str, visual_index: int):
        """Toggle expand/collapse for a visual's children in contrast display."""
        is_collapsed = self._collapsed_visuals.get(children_tag, True)
        new_state = not is_collapsed
        self._collapsed_visuals[children_tag] = new_state

        # Toggle elide on children
        self.details_text.config(state=tk.NORMAL)
        self.details_text.tag_configure(children_tag, elide=new_state)

        # Update header icon text
        header_tag = f"header_{visual_index}"
        new_icon = "▶ " if new_state else "▼ "

        # Find and replace the icon text
        try:
            ranges = self.details_text.tag_ranges(header_tag)
            if ranges and len(ranges) >= 2:
                start, end = str(ranges[0]), str(ranges[1])
                self.details_text.delete(start, end)
                self.details_text.insert(start, new_icon, header_tag)
        except Exception:
            pass  # Ignore if tag manipulation fails

        self.details_text.config(state=tk.DISABLED)

    def _export_report(self):
        """Export accessibility report to CSV - includes both issues and passing items"""
        if not self.current_results:
            ThemedMessageBox.showwarning(self.frame.winfo_toplevel(), "No Results", "Please run an analysis first.")
            return

        # Ask for save location
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        default_name = f"accessibility_report_{timestamp}.csv"

        file_path = filedialog.asksaveasfilename(
            title="Save Accessibility Report",
            defaultextension=".csv",
            initialfile=default_name,
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )

        if not file_path:
            return

        try:
            with open(file_path, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)

                # Header
                writer.writerow([
                    'Check Type', 'Severity', 'Page', 'Visual Name', 'Visual Type',
                    'Issue Description', 'Recommendation', 'Current Value', 'WCAG Reference'
                ])

                # Group issues by check type, then by page
                for check_type in AccessibilityCheckType:
                    check_issues = self.current_results.get_issues_by_type(check_type)
                    if not check_issues:
                        continue

                    # Group by page within this check type
                    issues_by_page = defaultdict(list)
                    for issue in check_issues:
                        page_name = issue.page_name or "Unknown Page"
                        issues_by_page[page_name].append(issue)

                    # Write check type header
                    writer.writerow([])
                    writer.writerow([f"=== {CHECK_TYPE_DISPLAY_NAMES.get(check_type, check_type.value).upper()} ==="])

                    # Write each page's issues
                    for page_name in sorted(issues_by_page.keys()):
                        writer.writerow([f"--- {page_name} ---"])

                        for issue in issues_by_page[page_name]:
                            severity_display, _ = self._get_severity_display(issue.severity)
                            # Format current_value for readability in CSV
                            current_val = issue.current_value or ''
                            if issue.check_type == AccessibilityCheckType.COLOR_CONTRAST and '|' in current_val:
                                parts = current_val.split('|')
                                if len(parts) == 3:
                                    element_type, fg_color, bg_color = parts
                                    current_val = f"{element_type}: {fg_color} on {bg_color}"
                            writer.writerow([
                                CHECK_TYPE_DISPLAY_NAMES.get(issue.check_type, issue.check_type.value),
                                severity_display.upper(),
                                issue.page_name,
                                issue.visual_name or '',
                                issue.visual_type or '',
                                issue.issue_description,
                                issue.recommendation,
                                current_val,
                                issue.wcag_reference or ''
                            ])

                # Add summary section
                writer.writerow([])  # Empty row
                writer.writerow(['--- SUMMARY ---'])
                writer.writerow(['Total Issues', self.current_results.total_issues])
                writer.writerow(['Critical', self.current_results.errors])
                writer.writerow(['Should Fix', self.current_results.warnings])
                writer.writerow(['Review', self.current_results.info_count])
                writer.writerow([])

                # Add per-category summary
                writer.writerow(['--- BY CATEGORY ---'])
                writer.writerow(['Check Type', 'Issues', 'Total Checked', 'Pass Rate'])
                for check_type in AccessibilityCheckType:
                    issues = len(self.current_results.get_issues_by_type(check_type))
                    total = self._check_totals.get(check_type, 0)
                    pass_rate = f"{((total - issues) / total * 100):.1f}%" if total > 0 else "N/A"
                    writer.writerow([
                        CHECK_TYPE_DISPLAY_NAMES.get(check_type, check_type.value),
                        issues,
                        total,
                        pass_rate
                    ])

            self.log_message(f"📄 Report exported to: {file_path}")
            ThemedMessageBox.showsuccess(self.frame.winfo_toplevel(), "Export Complete", f"Report saved to:\n{file_path}")

        except Exception as e:
            self.log_message(f"❌ Failed to export report: {e}")
            ThemedMessageBox.showerror(self.frame.winfo_toplevel(), "Export Error", f"Failed to export report:\n{e}")

    def _export_html_report(self):
        """Export accessibility report as formatted HTML (printable to PDF)"""
        if not self.current_results:
            ThemedMessageBox.showwarning(self.frame.winfo_toplevel(), "No Results", "Please run an analysis first.")
            return

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        default_name = f"accessibility_report_{timestamp}.html"

        file_path = filedialog.asksaveasfilename(
            title="Save Accessibility Report (HTML)",
            defaultextension=".html",
            initialfile=default_name,
            filetypes=[("HTML files", "*.html"), ("All files", "*.*")]
        )

        if not file_path:
            return

        try:
            html_content = self._generate_html_report()

            with open(file_path, 'w', encoding='utf-8') as f:
                f.write(html_content)

            # Open in browser for printing
            import webbrowser
            webbrowser.open(f'file:///{file_path}')

            self.log_message(f"📄 HTML report saved: {file_path}")
            ThemedMessageBox.showsuccess(self.frame.winfo_toplevel(), "Export Complete",
                "Report saved as HTML and opened in browser.\n\n"
                "Use your browser's Print function (Ctrl+P) to save as PDF.")

        except Exception as e:
            self.log_message(f"❌ Failed to export HTML report: {e}")
            ThemedMessageBox.showerror(self.frame.winfo_toplevel(), "Export Error", f"Failed to export report:\n{e}")

    def _get_severity_display(self, severity: AccessibilitySeverity) -> tuple:
        """Get display name and CSS class for severity level.

        Returns (display_name, css_class) tuple.
        Maps internal severity to user-friendly terminology:
        - ERROR -> Critical (must fix for WCAG compliance)
        - WARNING -> Should Fix (recommended improvement)
        - INFO -> Review (manual review suggested)
        """
        severity_map = {
            AccessibilitySeverity.ERROR: ("Critical", "critical"),
            AccessibilitySeverity.WARNING: ("Should Fix", "should-fix"),
            AccessibilitySeverity.INFO: ("Review", "review"),
        }
        return severity_map.get(severity, (severity.value.title(), severity.value))

    def _generate_html_report(self) -> str:
        """Generate professional HTML report with TOC and Executive Summary"""
        result = self.current_results
        from datetime import datetime

        # Format analysis timestamp
        try:
            analysis_time = datetime.fromisoformat(result.analysis_timestamp).strftime("%B %d, %Y at %I:%M %p")
        except:
            analysis_time = result.analysis_timestamp

        # Calculate category statistics for executive summary
        category_stats = []
        for check_type in AccessibilityCheckType:
            issues = result.get_issues_by_type(check_type)
            total = self._check_totals.get(check_type, 0)
            name = CHECK_TYPE_DISPLAY_NAMES.get(check_type, check_type.value)
            anchor = check_type.value.replace("_", "-")

            # Count by severity using new terminology
            critical_count = sum(1 for i in issues if i.severity == AccessibilitySeverity.ERROR)
            should_fix_count = sum(1 for i in issues if i.severity == AccessibilitySeverity.WARNING)
            review_count = sum(1 for i in issues if i.severity == AccessibilitySeverity.INFO)

            # For contrast, count unique visuals with issues (not total element failures)
            if check_type == AccessibilityCheckType.COLOR_CONTRAST:
                visuals_with_issues = set()
                for issue in issues:
                    if issue.visual_name:
                        visuals_with_issues.add((issue.page_name, issue.visual_name))
                display_issue_count = len(visuals_with_issues)
            else:
                display_issue_count = len(issues)

            category_stats.append({
                'name': name,
                'anchor': anchor,
                'check_type': check_type,
                'issue_count': display_issue_count,
                'total_checked': total,
                'pass_rate': round(((total - display_issue_count) / total * 100) if total > 0 else 100, 1),
                'critical': critical_count,
                'should_fix': should_fix_count,
                'review': review_count,
                'issues': issues
            })

        # Calculate overall stats
        total_checked = sum(s['total_checked'] for s in category_stats)
        total_issues = result.total_issues
        overall_pass_rate = round(((total_checked - total_issues) / total_checked * 100) if total_checked > 0 else 100, 1)

        # Build table of contents
        toc_items = ""
        for stat in category_stats:
            issue_badge = f'<span class="toc-count">{stat["issue_count"]}</span>' if stat['issue_count'] > 0 else '<span class="toc-pass">✓</span>'
            toc_items += f'<li><a href="#{stat["anchor"]}">{stat["name"]}</a> {issue_badge}</li>\n'

        # Build executive summary cards
        summary_cards = ""
        for stat in category_stats:
            status_class = "pass" if stat['issue_count'] == 0 else "critical" if stat['critical'] > 0 else "should-fix" if stat['should_fix'] > 0 else "review"
            summary_cards += f'''
            <div class="summary-card {status_class}">
                <div class="card-title">{stat['name']}</div>
                <div class="card-count">{stat['issue_count']}</div>
                <div class="card-detail">{stat['pass_rate']}% pass rate</div>
            </div>'''

        # Build detailed category sections
        categories_html = ""
        for stat in category_stats:
            issues = stat['issues']

            if issues:
                # Group issues by page within this category
                issues_by_page = defaultdict(list)
                for issue in issues:
                    page_name = issue.page_name or "Unknown Page"
                    issues_by_page[page_name].append(issue)

                issues_html = ""
                # Build HTML for each page's issues
                for page_name in sorted(issues_by_page.keys()):
                    page_issues = issues_by_page[page_name]

                    # Page sub-header
                    issues_html += f'''
                    <div class="page-section">
                        <h3 class="page-header">{page_name}</h3>
                        <div class="page-issues">'''

                    # Special hierarchical handling for contrast issues
                    if stat['check_type'] == AccessibilityCheckType.COLOR_CONTRAST:
                        # Group issues by visual within this page
                        issues_by_visual = defaultdict(list)
                        for issue in page_issues:
                            visual_key = issue.visual_name or "Unknown Visual"
                            issues_by_visual[visual_key].append(issue)

                        for visual_name in sorted(issues_by_visual.keys()):
                            visual_issues = issues_by_visual[visual_name]

                            # Get visual type from first issue (all issues for same visual should have same type)
                            visual_type = visual_issues[0].visual_type if visual_issues else None

                            # Check if visual_name is just a hex ID (no meaningful name set)
                            # Power BI visual IDs are typically 20-char hex strings like "06c8c18fb4151a739c3d"
                            is_unnamed_visual = bool(re.match(r'^[a-f0-9]{16,24}$', visual_name.lower()))

                            # Build display name: include visual type if name is just an ID
                            if is_unnamed_visual and visual_type:
                                display_name = f"{visual_name} ({visual_type})"
                            else:
                                display_name = visual_name

                            # Calculate summary for this visual
                            worst_ratio = float('inf')
                            worst_severity = None
                            for issue in visual_issues:
                                if issue.current_value and '|' in issue.current_value:
                                    # Extract ratio from issue description
                                    desc = issue.issue_description
                                    if 'Ratio:' in desc:
                                        try:
                                            ratio_str = desc.split('Ratio:')[1].split(':1')[0].strip()
                                            ratio = float(ratio_str)
                                            if ratio < worst_ratio:
                                                worst_ratio = ratio
                                                worst_severity = issue.severity
                                        except:
                                            pass

                            # Determine element class based on worst severity
                            if worst_severity == AccessibilitySeverity.ERROR:
                                ratio_class = "critical"
                            elif worst_severity == AccessibilitySeverity.WARNING:
                                ratio_class = "should-fix"
                            else:
                                ratio_class = "review"

                            worst_ratio_display = f"{worst_ratio:.1f}:1" if worst_ratio != float('inf') else "N/A"

                            # Add naming tip if visual is unnamed
                            naming_tip = ""
                            if is_unnamed_visual:
                                naming_tip = '<span class="naming-tip">Tip: Give this visual a meaningful name in Power BI for better accessibility reports</span>'

                            # Visual group header (expanded in PDF for readability)
                            issues_html += f'''
                            <details class="visual-contrast-group" open>
                                <summary>
                                    <span class="visual-name">{display_name}</span>
                                    <span class="visual-worst-ratio">Worst: {worst_ratio_display}</span>
                                    <span class="visual-element-count">{len(visual_issues)} element(s)</span>
                                    {naming_tip}
                                </summary>
                                <div class="visual-elements-list">'''

                            # Individual elements within this visual
                            for issue in visual_issues:
                                severity_display, severity_class = self._get_severity_display(issue.severity)
                                element_class = "critical-element" if severity_class == "critical" else "review-element" if severity_class == "review" else ""

                                # Parse current_value for element details
                                element_name = "Element"
                                fg_color = ""
                                bg_color = ""
                                if issue.current_value and '|' in issue.current_value:
                                    parts = issue.current_value.split('|')
                                    if len(parts) == 3:
                                        element_name, fg_color, bg_color = parts

                                # Extract ratio from description
                                ratio_display = ""
                                if 'Ratio:' in issue.issue_description:
                                    try:
                                        ratio_display = issue.issue_description.split('Ratio:')[1].split(' ')[0].strip()
                                    except:
                                        ratio_display = ""

                                color_swatches = ""
                                if fg_color and bg_color:
                                    color_swatches = f'''
                                    <span class="element-colors">
                                        <span class="color-swatch" style="background:{fg_color};" title="{fg_color}"></span>
                                        <span class="color-arrow">→</span>
                                        <span class="color-swatch" style="background:{bg_color};" title="{bg_color}"></span>
                                    </span>'''

                                issues_html += f'''
                                <div class="contrast-element {element_class}">
                                    <span class="element-name">{element_name}</span>
                                    <span class="element-ratio">{ratio_display}</span>
                                    {color_swatches}
                                    <span class="severity-badge {severity_class}" style="font-size:10px;padding:2px 6px;">{severity_display}</span>
                                </div>'''

                            issues_html += '''
                                </div>
                            </details>'''
                    else:
                        # Standard flat display for non-contrast issues
                        for issue in page_issues:
                            severity_display, severity_class = self._get_severity_display(issue.severity)
                            # Build visual info with type if name is just an ID
                            visual_info = ""
                            if issue.visual_name:
                                is_unnamed = bool(re.match(r'^[a-f0-9]{16,24}$', issue.visual_name.lower()))
                                if is_unnamed and issue.visual_type:
                                    visual_display = f"{issue.visual_name} ({issue.visual_type})"
                                    visual_info = f'<br><strong>Visual:</strong> {visual_display} <span class="naming-tip" style="display:inline;border:none;margin:0;padding:0;">- Consider naming this visual</span>'
                                else:
                                    visual_info = f"<br><strong>Visual:</strong> {issue.visual_name}"
                            # Format current_value with color swatches for contrast issues
                            current_info = ""
                            if issue.current_value:
                                if issue.check_type == AccessibilityCheckType.COLOR_CONTRAST and '|' in issue.current_value:
                                    parts = issue.current_value.split('|')
                                    if len(parts) == 3:
                                        element_type, fg_color, bg_color = parts
                                        current_info = f'''<br><strong>Elements:</strong> {element_type}
                                        <br><strong>Colors:</strong>
                                        <span style="display:inline-block;width:20px;height:14px;background:{fg_color};border:1px solid #333;vertical-align:middle;"></span>
                                        <code>{fg_color}</code> →
                                        <span style="display:inline-block;width:20px;height:14px;background:{bg_color};border:1px solid #333;vertical-align:middle;"></span>
                                        <code>{bg_color}</code>'''
                                    else:
                                        current_info = f"<br><strong>Current Value:</strong> <code>{issue.current_value}</code>"
                                else:
                                    current_info = f"<br><strong>Current Value:</strong> <code>{issue.current_value}</code>"
                            wcag_info = f'<div class="wcag-ref">📖 {issue.wcag_reference}</div>' if issue.wcag_reference else ""

                            issues_html += f'''
                            <div class="issue {severity_class}">
                                <div class="issue-header">
                                    <span class="severity-badge {severity_class}">{severity_display}</span>
                                    <span class="issue-title">{issue.issue_description}</span>
                                </div>
                                <div class="issue-details">
                                    {visual_info.replace("<br><strong>Visual:", "<strong>Visual:") if visual_info else ""}{current_info}
                                </div>
                                <div class="recommendation">
                                    <strong>Recommendation:</strong> {issue.recommendation}
                                </div>
                                {wcag_info}
                            </div>'''

                    issues_html += '''
                        </div>
                    </div>'''
            else:
                issues_html = '''
                <div class="all-pass">
                    <span class="pass-icon">✅</span>
                    <span>All items in this category pass accessibility checks</span>
                </div>'''

            # Category status indicator
            if stat['issue_count'] == 0:
                status_indicator = '<span class="section-status pass">PASS</span>'
            elif stat['critical'] > 0:
                status_indicator = f'<span class="section-status critical">{stat["critical"]} Critical</span>'
            elif stat['should_fix'] > 0:
                status_indicator = f'<span class="section-status should-fix">{stat["should_fix"]} Should Fix</span>'
            else:
                status_indicator = f'<span class="section-status review">{stat["review"]} Review</span>'

            categories_html += f'''
            <section id="{stat['anchor']}" class="category-section">
                <div class="section-header">
                    <h2>{stat['name']}</h2>
                    {status_indicator}
                </div>
                <div class="section-meta">
                    Checked {stat['total_checked']} items • {stat['issue_count']} issues found • {stat['pass_rate']}% pass rate
                </div>
                <div class="issues-container">
                    {issues_html}
                </div>
            </section>'''

        # Determine priority areas for executive summary
        priority_areas = sorted([s for s in category_stats if s['issue_count'] > 0],
                               key=lambda x: (x['critical'] * 3 + x['should_fix'] * 2 + x['review']),
                               reverse=True)[:3]

        priority_html = ""
        if priority_areas:
            for i, area in enumerate(priority_areas, 1):
                priority_html += f'<li><strong>{area["name"]}</strong>: {area["issue_count"]} issues ({area["critical"]} critical, {area["should_fix"]} should fix, {area["review"]} review)</li>'
        else:
            priority_html = '<li>No issues found - excellent accessibility compliance!</li>'

        return f'''<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Accessibility Report - {result.report_name}</title>
    <style>
        :root {{
            --critical-color: #c62828;
            --critical-bg: #ffebee;
            --should-fix-color: #ef6c00;
            --should-fix-bg: #fff3e0;
            --review-color: #1565c0;
            --review-bg: #e3f2fd;
            --pass-color: #2e7d32;
            --pass-bg: #e8f5e9;
            --primary-color: #1a73e8;
            --text-color: #333;
            --border-color: #e0e0e0;
        }}

        * {{ box-sizing: border-box; }}

        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            max-width: 1000px;
            margin: 0 auto;
            padding: 30px 40px;
            line-height: 1.6;
            color: var(--text-color);
            background: #fafafa;
        }}

        /* Header */
        .report-header {{
            background: linear-gradient(135deg, var(--primary-color), #1557b0);
            color: white;
            padding: 30px 40px;
            border-radius: 12px;
            margin-bottom: 30px;
        }}
        .report-header h1 {{
            margin: 0 0 10px 0;
            font-size: 28px;
        }}
        .report-meta {{
            opacity: 0.9;
            font-size: 14px;
        }}
        .report-meta strong {{
            opacity: 1;
        }}

        /* Table of Contents */
        .toc {{
            background: white;
            padding: 25px 30px;
            border-radius: 10px;
            border: 1px solid var(--border-color);
            margin-bottom: 30px;
        }}
        .toc h2 {{
            margin: 0 0 15px 0;
            font-size: 18px;
            color: var(--text-color);
        }}
        .toc ul {{
            list-style: none;
            padding: 0;
            margin: 0;
            display: grid;
            grid-template-columns: repeat(2, 1fr);
            gap: 8px 30px;
        }}
        .toc li {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            padding: 6px 0;
            border-bottom: 1px dashed #eee;
        }}
        .toc a {{
            color: var(--primary-color);
            text-decoration: none;
        }}
        .toc a:hover {{
            text-decoration: underline;
        }}
        .toc-count {{
            background: var(--should-fix-color);
            color: white;
            padding: 2px 8px;
            border-radius: 10px;
            font-size: 12px;
            font-weight: bold;
        }}
        .toc-pass {{
            color: var(--pass-color);
            font-weight: bold;
        }}

        /* Executive Summary */
        .executive-summary {{
            background: white;
            padding: 30px;
            border-radius: 10px;
            border: 1px solid var(--border-color);
            margin-bottom: 30px;
        }}
        .executive-summary h2 {{
            margin: 0 0 20px 0;
            color: var(--text-color);
            font-size: 20px;
        }}
        .overall-score {{
            text-align: center;
            padding: 20px;
            background: linear-gradient(135deg, #f8f9fa, #e9ecef);
            border-radius: 10px;
            margin-bottom: 25px;
        }}
        .score-number {{
            font-size: 48px;
            font-weight: bold;
            color: {('var(--pass-color)' if overall_pass_rate >= 90 else 'var(--should-fix-color)' if overall_pass_rate >= 70 else 'var(--critical-color)')};
        }}
        .score-label {{
            font-size: 14px;
            color: #666;
            margin-top: 5px;
        }}
        .summary-grid {{
            display: grid;
            grid-template-columns: repeat(4, 1fr);
            gap: 15px;
            margin-bottom: 25px;
        }}
        .summary-card {{
            padding: 15px;
            border-radius: 8px;
            text-align: center;
            border: 2px solid transparent;
        }}
        .summary-card.pass {{ background: var(--pass-bg); border-color: var(--pass-color); }}
        .summary-card.critical {{ background: var(--critical-bg); border-color: var(--critical-color); }}
        .summary-card.should-fix {{ background: var(--should-fix-bg); border-color: var(--should-fix-color); }}
        .summary-card.review {{ background: var(--review-bg); border-color: var(--review-color); }}
        .card-title {{
            font-size: 11px;
            font-weight: 600;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            margin-bottom: 5px;
        }}
        .card-count {{
            font-size: 24px;
            font-weight: bold;
        }}
        .summary-card.pass .card-count {{ color: var(--pass-color); }}
        .summary-card.critical .card-count {{ color: var(--critical-color); }}
        .summary-card.should-fix .card-count {{ color: var(--should-fix-color); }}
        .summary-card.review .card-count {{ color: var(--review-color); }}
        .card-detail {{
            font-size: 11px;
            color: #666;
        }}
        .priority-areas {{
            background: #f8f9fa;
            padding: 20px;
            border-radius: 8px;
        }}
        .priority-areas h3 {{
            margin: 0 0 12px 0;
            font-size: 14px;
            color: #444;
        }}
        .priority-areas ol {{
            margin: 0;
            padding-left: 20px;
        }}
        .priority-areas li {{
            margin-bottom: 8px;
        }}

        /* Category Sections */
        .category-section {{
            background: white;
            padding: 25px 30px;
            border-radius: 10px;
            border: 1px solid var(--border-color);
            margin-bottom: 20px;
        }}
        .section-header {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 10px;
        }}
        .section-header h2 {{
            margin: 0;
            font-size: 18px;
            color: var(--text-color);
        }}
        .section-status {{
            padding: 4px 12px;
            border-radius: 15px;
            font-size: 12px;
            font-weight: 600;
        }}
        .section-status.pass {{ background: var(--pass-bg); color: var(--pass-color); }}
        .section-status.critical {{ background: var(--critical-bg); color: var(--critical-color); }}
        .section-status.should-fix {{ background: var(--should-fix-bg); color: var(--should-fix-color); }}
        .section-status.review {{ background: var(--review-bg); color: var(--review-color); }}
        .section-meta {{
            font-size: 13px;
            color: #666;
            margin-bottom: 20px;
            padding-bottom: 15px;
            border-bottom: 1px solid #eee;
        }}

        /* Issues */
        .issues-container {{
            display: flex;
            flex-direction: column;
            gap: 15px;
        }}

        /* Page Sections (within category) */
        .page-section {{
            margin-bottom: 20px;
        }}
        .page-header {{
            font-size: 14px;
            color: #555;
            font-weight: 600;
            margin: 0 0 12px 0;
            padding: 8px 0 8px 12px;
            border-left: 3px solid #009999;
            background: #f8f9fa;
        }}
        .page-issues {{
            margin-left: 15px;
            display: flex;
            flex-direction: column;
            gap: 12px;
        }}

        .issue {{
            padding: 15px 20px;
            border-radius: 8px;
            border-left: 4px solid;
        }}
        .issue.critical {{
            background: var(--critical-bg);
            border-left-color: var(--critical-color);
        }}
        .issue.should-fix {{
            background: var(--should-fix-bg);
            border-left-color: var(--should-fix-color);
        }}
        .issue.review {{
            background: var(--review-bg);
            border-left-color: var(--review-color);
        }}
        .issue-header {{
            display: flex;
            align-items: flex-start;
            gap: 10px;
            margin-bottom: 10px;
        }}
        .severity-badge {{
            padding: 3px 10px;
            border-radius: 4px;
            font-size: 11px;
            font-weight: 600;
            text-transform: uppercase;
            white-space: nowrap;
        }}
        .severity-badge.critical {{ background: var(--critical-color); color: white; }}
        .severity-badge.should-fix {{ background: var(--should-fix-color); color: white; }}
        .severity-badge.review {{ background: var(--review-color); color: white; }}
        .issue-title {{
            font-weight: 500;
            color: #333;
        }}
        .issue-details {{
            font-size: 13px;
            color: #555;
            margin-bottom: 10px;
            padding: 10px;
            background: rgba(255,255,255,0.7);
            border-radius: 4px;
        }}
        .issue-details code {{
            background: #f5f5f5;
            padding: 2px 6px;
            border-radius: 3px;
            font-family: 'Consolas', monospace;
            font-size: 12px;
        }}
        .recommendation {{
            font-size: 13px;
            color: var(--pass-color);
            background: rgba(255,255,255,0.7);
            padding: 10px;
            border-radius: 4px;
        }}
        .wcag-ref {{
            font-size: 12px;
            color: #666;
            margin-top: 8px;
            font-style: italic;
        }}
        .all-pass {{
            display: flex;
            align-items: center;
            gap: 10px;
            padding: 20px;
            background: var(--pass-bg);
            border-radius: 8px;
            color: var(--pass-color);
            font-weight: 500;
        }}
        .pass-icon {{
            font-size: 20px;
        }}

        /* Hierarchical Contrast Visual Grouping */
        .visual-contrast-group {{
            margin-bottom: 15px;
            border: 1px solid #ddd;
            border-radius: 8px;
            overflow: hidden;
        }}
        .visual-contrast-group summary {{
            padding: 12px 16px;
            background: #f5f5f5;
            cursor: pointer;
            font-weight: 600;
            display: flex;
            align-items: center;
            gap: 12px;
            list-style: none;
        }}
        .visual-contrast-group summary::-webkit-details-marker {{
            display: none;
        }}
        .visual-contrast-group summary::before {{
            content: "▶";
            font-size: 10px;
            transition: transform 0.2s;
        }}
        .visual-contrast-group[open] summary::before {{
            transform: rotate(90deg);
        }}
        .visual-contrast-group summary:hover {{
            background: #eee;
        }}
        .visual-name {{
            flex: 1;
            color: #333;
        }}
        .visual-type-badge {{
            font-size: 11px;
            color: #666;
            font-weight: normal;
        }}
        .visual-worst-ratio {{
            font-size: 12px;
            padding: 2px 8px;
            border-radius: 4px;
            background: var(--critical-bg);
            color: var(--critical-color);
        }}
        .visual-element-count {{
            font-size: 11px;
            color: #888;
        }}
        .naming-tip {{
            display: block;
            width: 100%;
            font-size: 11px;
            font-style: italic;
            color: #009999;
            margin-top: 6px;
            padding-top: 6px;
            border-top: 1px dashed #ddd;
        }}
        .visual-elements-list {{
            padding: 12px 16px 12px 32px;
            background: white;
        }}
        .contrast-element {{
            display: flex;
            align-items: center;
            gap: 12px;
            padding: 8px 12px;
            margin-bottom: 8px;
            background: #fafafa;
            border-radius: 6px;
            border-left: 3px solid var(--should-fix-color);
        }}
        .contrast-element.critical-element {{
            border-left-color: var(--critical-color);
            background: var(--critical-bg);
        }}
        .contrast-element.review-element {{
            border-left-color: var(--review-color);
            background: var(--review-bg);
        }}
        .element-name {{
            flex: 1;
            font-size: 13px;
            color: #444;
        }}
        .element-ratio {{
            font-weight: 600;
            font-size: 13px;
        }}
        .element-colors {{
            display: flex;
            align-items: center;
            gap: 6px;
            font-size: 12px;
        }}
        .color-swatch {{
            width: 18px;
            height: 18px;
            border: 1px solid #333;
            border-radius: 3px;
            display: inline-block;
        }}
        .color-arrow {{
            color: #888;
        }}

        /* Footer */
        footer {{
            margin-top: 40px;
            padding: 25px;
            background: #f5f5f5;
            border-radius: 10px;
            text-align: center;
            color: #666;
            font-size: 13px;
        }}
        footer strong {{
            color: var(--primary-color);
        }}

        /* Print styles */
        @media print {{
            body {{
                max-width: none;
                margin: 0;
                padding: 20px;
                background: white;
            }}
            .report-header {{
                background: var(--primary-color) !important;
                -webkit-print-color-adjust: exact;
                print-color-adjust: exact;
            }}
            .category-section {{
                page-break-inside: avoid;
                break-inside: avoid;
            }}
            .issue {{
                page-break-inside: avoid;
                break-inside: avoid;
            }}
            .toc {{
                page-break-after: always;
            }}
        }}
    </style>
</head>
<body>
    <!-- Header -->
    <header class="report-header">
        <h1>🔍 Power BI Accessibility Report</h1>
        <div class="report-meta">
            <strong>Report:</strong> {result.report_name}<br>
            <strong>Analyzed:</strong> {analysis_time}<br>
            <strong>Analysis Duration:</strong> {result.analysis_duration_ms}ms
        </div>
    </header>

    <!-- Table of Contents -->
    <nav class="toc">
        <h2>📑 Contents</h2>
        <ul>
            <li><a href="#executive-summary">Executive Summary</a></li>
            {toc_items}
        </ul>
    </nav>

    <!-- Executive Summary -->
    <section id="executive-summary" class="executive-summary">
        <h2>📊 Executive Summary</h2>

        <div class="overall-score">
            <div class="score-number">{overall_pass_rate}%</div>
            <div class="score-label">Overall Accessibility Score ({total_checked - total_issues} of {total_checked} items pass)</div>
        </div>

        <div class="summary-grid">
            {summary_cards}
        </div>

        <div class="priority-areas">
            <h3>🎯 Priority Areas</h3>
            <ol>
                {priority_html}
            </ol>
        </div>
    </section>

    <!-- Category Sections -->
    {categories_html}

    <!-- Footer -->
    <footer>
        <p>Generated by <strong>AE Multi-Tool Accessibility Checker</strong></p>
        <p>Built by Reid Havens of Analytic Endeavors</p>
        <p style="font-size: 11px; margin-top: 15px;">
            This report analyzes WCAG 2.1 Level AA/AAA compliance for Power BI reports.
            Automated checks are a starting point — manual review is recommended for complete accessibility verification.
        </p>
    </footer>
</body>
</html>'''

    def reset_tab(self) -> None:
        """Reset tab to initial state"""
        # Clear path
        self.pbip_path_var.set("")

        # Reset cards
        for card in self.check_cards.values():
            card.reset()

        # Clear results
        self.current_results = None
        self.selected_check_type = None
        self._check_totals.clear()

        # Show placeholder in details panel
        self._show_details_placeholder()

        # Clear log
        if self.log_text:
            self.log_text.delete(1.0, tk.END)

        # Reset progress
        self.update_progress(0, "", show=False)

        # Disable buttons
        self.analyze_button.set_enabled(False)
        self.export_button.set_enabled(False)
        self.export_pdf_button.set_enabled(False)

        # Show welcome message
        self._show_welcome_message()

    def _show_settings_dialog(self) -> None:
        """Show settings dialog for configuring accessibility checks"""
        from core.ui_base import RoundedButton

        colors = self._theme_manager.colors
        is_dark = self._theme_manager.is_dark
        config = get_config()

        # Get correct parent window
        parent_window = None
        if hasattr(self, 'main_app') and hasattr(self.main_app, 'root'):
            parent_window = self.main_app.root
        elif hasattr(self, 'master'):
            parent_window = self.master
        else:
            parent_window = self.frame.winfo_toplevel()

        # Create settings dialog - sized to fit all content
        dialog = tk.Toplevel(parent_window)
        dialog.withdraw()
        dialog.title("Accessibility Check Settings")
        dialog.geometry("520x650")
        dialog.resizable(False, False)
        dialog.transient(parent_window)
        dialog.grab_set()
        dialog.configure(bg=colors['background'])

        # Set icon
        try:
            if hasattr(self, 'main_app') and hasattr(self.main_app, 'config'):
                dialog.iconbitmap(self.main_app.config.icon_path)
        except Exception:
            pass

        # Set dark title bar if dark mode
        if hasattr(self, '_set_dialog_title_bar_color'):
            self._set_dialog_title_bar_color(dialog, is_dark)

        # Section styling - use section_bg for better radio button visibility in light mode
        section_bg = colors['section_bg']
        section_border = colors['border']

        # Main container - reduced padding
        main_frame = tk.Frame(dialog, bg=colors['background'], padx=20, pady=15)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Header - smaller padding
        tk.Label(main_frame, text="Settings",
                font=('Segoe UI', 14, 'bold'),
                bg=colors['background'], fg=colors['title_color']).pack(anchor=tk.W, pady=(0, 12))

        # Helper to create styled section frames (like Categorical/Measure containers)
        def create_section(parent, title, show_link=False, link_text=None, link_command=None):
            """Create a section with clean styling matching Table Column Widths

            Args:
                show_link: If True, adds a link label to the right of the title
                link_text: Text for the link
                link_command: Callback when link is clicked
            """
            section = tk.Frame(parent, bg=section_bg, highlightbackground=section_border,
                              highlightthickness=1, padx=12, pady=10)
            section.pack(fill=tk.X, pady=(0, 10))

            # Title row (frame to hold title and optional link)
            title_row = tk.Frame(section, bg=section_bg)
            title_row.pack(fill=tk.X, pady=(0, 8))

            # Title label - use title_color (blue in dark, teal in light)
            title_label = tk.Label(title_row, text=title,
                                  bg=section_bg, fg=colors['title_color'],
                                  font=('Segoe UI', 10, 'bold'))
            title_label.pack(side=tk.LEFT)

            # Optional link on the right
            link_label = None
            if show_link and link_text:
                link_label = tk.Label(title_row, text=link_text,
                                     bg=section_bg, fg=colors['info'],
                                     font=('Segoe UI', 8, 'underline'), cursor='hand2')
                link_label.pack(side=tk.RIGHT)
                if link_command:
                    link_label.bind('<Button-1>', lambda e: link_command())

            # Content frame
            content = tk.Frame(section, bg=section_bg)
            content.pack(fill=tk.X)
            return content, section, title_row, title_label, link_label

        # === Enabled Checks Section ===
        checks_content, checks_section, checks_title_row, checks_title_label, _ = create_section(main_frame, "Enabled Checks")

        # Toggle variables and widgets
        check_vars = {}
        toggle_widgets = []
        check_names = {
            "tab_order": "Tab Order",
            "alt_text": "Alt Text",
            "color_contrast": "Color Contrast",
            "page_title": "Page Titles",
            "visual_title": "Visual Titles",
            "data_labels": "Data Labels",
            "bookmark_name": "Bookmarks",
            "hidden_page": "Hidden Pages",
        }

        # Create 2 columns of toggles - more compact grid
        checks_content.columnconfigure(0, weight=1)
        checks_content.columnconfigure(1, weight=1)

        check_items = list(check_names.items())
        for i, (key, label) in enumerate(check_items):
            row = i // 2
            col = i % 2

            var = tk.BooleanVar(value=config.enabled_checks.get(key, True))
            check_vars[key] = var

            # Create toggle with section background
            toggle_frame = tk.Frame(checks_content, bg=section_bg)
            toggle_frame.grid(row=row, column=col, sticky=tk.W, padx=(0, 10), pady=2)

            toggle = LabeledToggle(toggle_frame, variable=var, text=label)
            toggle.configure(bg=section_bg)
            toggle._icon_label.configure(bg=section_bg)
            if toggle._label:
                toggle._label.configure(bg=section_bg)
            toggle.pack(anchor=tk.W)
            toggle_widgets.append(toggle)

        # === WCAG Contrast Level Section ===
        import webbrowser
        def open_wcag_link():
            webbrowser.open("https://webaim.org/articles/contrast/")

        contrast_content, contrast_section, contrast_title_row, contrast_title_label, contrast_link = create_section(
            main_frame, "WCAG Contrast Level",
            show_link=True, link_text="Learn more", link_command=open_wcag_link
        )

        contrast_var = tk.StringVar(value=config.contrast_level)

        contrast_options = [
            ("AA", "AA Standard (4.5:1)", "Recommended"),
            ("AAA", "AAA Enhanced (7:1)", "Stricter"),
            ("AA_large", "Large Text Only (3:1)", "Relaxed"),
        ]

        # Tooltip explanations for each contrast level
        contrast_tooltips = {
            "AA": "WCAG Level AA: Requires 4.5:1 contrast for normal text\nand 3:1 for large text (18pt+ or 14pt bold).\nThis is the standard for most accessibility compliance.",
            "AAA": "WCAG Level AAA: Requires 7:1 contrast for normal text\nand 4.5:1 for large text. This is the highest standard\nand may be difficult to achieve with brand colors.",
            "AA_large": "Only checks large text threshold (3:1 ratio).\nUse this for reports with primarily large text\nor when stricter checks aren't required.",
        }

        contrast_radio = LabeledRadioGroup(
            contrast_content, options=contrast_options, variable=contrast_var, bg=section_bg
        )
        # Add tooltips to each radio option
        for item in contrast_radio._radio_items:
            tooltip_text = contrast_tooltips.get(item['value'], "")
            if tooltip_text:
                Tooltip(item['frame'], tooltip_text)
        contrast_radio.pack(fill=tk.X)

        # === Flag Options Row (AAA and AA toggles side by side) ===
        flag_toggle_frame = tk.Frame(contrast_content, bg=section_bg)
        flag_toggle_frame.pack(fill=tk.X, pady=(6, 0))

        # AAA Flag Toggle
        aaa_var = tk.BooleanVar(value=config.flag_aaa_failures)
        aaa_toggle = LabeledToggle(flag_toggle_frame, variable=aaa_var,
                                   text="Flag AAA failures as 'Review'")
        aaa_toggle.configure(bg=section_bg)
        aaa_toggle._icon_label.configure(bg=section_bg)
        if aaa_toggle._label:
            aaa_toggle._label.configure(bg=section_bg)
        aaa_toggle.pack(side=tk.LEFT, anchor=tk.W)

        # AA Flag Toggle (to the right of AAA)
        aa_var = tk.BooleanVar(value=config.flag_aa_failures)
        aa_toggle = LabeledToggle(flag_toggle_frame, variable=aa_var,
                                  text="Flag AA failures as 'Review'")
        aa_toggle.configure(bg=section_bg)
        aa_toggle._icon_label.configure(bg=section_bg)
        if aa_toggle._label:
            aa_toggle._label.configure(bg=section_bg)
        aa_toggle.pack(side=tk.LEFT, anchor=tk.W, padx=(16, 0))

        # === Severity Filter Section ===
        severity_content, severity_section, severity_title_row, severity_title_label, _ = create_section(main_frame, "Default Severity Filter")

        severity_var = tk.StringVar(value=config.min_severity)

        severity_options = [
            ("info", "All Issues", "Critical + Should Fix + Review"),
            ("warning", "Critical & Should Fix", "Hide Review items"),
            ("error", "Critical Only", "Most severe only"),
        ]

        severity_radio = LabeledRadioGroup(
            severity_content, options=severity_options, variable=severity_var, bg=section_bg
        )
        severity_radio.pack(fill=tk.X)

        # === Info Text ===
        info_text = tk.Label(main_frame,
                            text="Settings persist between sessions",
                            bg=colors['background'], fg=colors['text_secondary'],
                            font=('Segoe UI', 8, 'italic'))
        info_text.pack(pady=(6, 10))

        # === Buttons ===
        # Load icons for buttons
        save_icon = self._load_icon_for_button('save', size=16)
        reset_icon = self._load_icon_for_button('reset', size=16)

        button_frame = tk.Frame(main_frame, bg=colors['background'])
        button_frame.pack(pady=(10, 0))  # Centered, no fill=tk.X

        def save_settings():
            # Update config with new values
            for key, var in check_vars.items():
                config.enabled_checks[key] = var.get()
            config.contrast_level = contrast_var.get()
            config.flag_aaa_failures = aaa_var.get()
            config.flag_aa_failures = aa_var.get()
            config.min_severity = severity_var.get()

            # Save to file
            save_config()

            # Update severity filter dropdown if it exists
            if hasattr(self, '_severity_filter_var'):
                self._severity_filter_var.set(config.min_severity)
                # Also update the dropdown button display
                if hasattr(self, '_severity_dropdown_btn'):
                    display = self._severity_display_map.get(config.min_severity, "All")
                    self._severity_dropdown_btn.configure(text=f" {display}  \u25BC")

            self.log_message("Settings saved successfully")
            # Use cleanup_and_close if available (set after theme callback is registered)
            if hasattr(save_settings, '_cleanup'):
                save_settings._cleanup()
            else:
                dialog.destroy()

        def reset_defaults():
            # Reset to defaults
            new_config = reset_config()

            # Update UI
            for key, var in check_vars.items():
                var.set(new_config.enabled_checks.get(key, True))
            contrast_var.set(new_config.contrast_level)
            contrast_radio.set(new_config.contrast_level)
            aaa_var.set(new_config.flag_aaa_failures)
            aa_var.set(new_config.flag_aa_failures)
            severity_var.set(new_config.min_severity)
            severity_radio.set(new_config.min_severity)

            self.log_message("Settings reset to defaults")

        # Save button with icon (auto-sized)
        save_btn = RoundedButton(
            button_frame, text="SAVE SETTINGS", icon=save_icon,
            height=36, radius=8,  # Auto-width, dialog standard
            bg=colors['button_primary'], fg='#ffffff',
            hover_bg=colors['button_primary_hover'],
            pressed_bg=colors['button_primary_pressed'],
            font=('Segoe UI', 10, 'bold'),
            command=save_settings,
            canvas_bg=colors['background']
        )
        save_btn.pack(side=tk.LEFT, padx=(0, 15))

        # Reset to Defaults button with icon (auto-sized)
        reset_btn = RoundedButton(
            button_frame, text="RESET DEFAULTS", icon=reset_icon,
            height=36, radius=8,  # Auto-width, dialog standard
            bg=colors['button_secondary'], fg=colors['text_primary'],
            hover_bg=colors['button_secondary_hover'],
            pressed_bg=colors['button_secondary_pressed'],
            font=('Segoe UI', 10),
            command=reset_defaults,
            canvas_bg=colors['background']
        )
        reset_btn.pack(side=tk.LEFT)

        # Theme change handler for the dialog
        def on_theme_change(new_theme):
            """Update all dialog colors when theme changes"""
            new_colors = self._theme_manager.colors
            new_is_dark = self._theme_manager.is_dark
            new_section_bg = new_colors['card_surface']
            new_section_border = new_colors['border']
            new_title_color = new_colors['title_color']

            # Update dialog and main containers
            dialog.configure(bg=new_colors['background'])
            main_frame.configure(bg=new_colors['background'])
            button_frame.configure(bg=new_colors['background'])
            info_text.configure(bg=new_colors['background'], fg=new_colors['text_secondary'])

            # Update main header
            for child in main_frame.winfo_children():
                if isinstance(child, tk.Label) and child.cget('text') == "Settings":
                    child.configure(bg=new_colors['background'], fg=new_title_color)
                    break

            # Update title bar color
            if hasattr(self, '_set_dialog_title_bar_color'):
                self._set_dialog_title_bar_color(dialog, new_is_dark)

            # Update all sections and their frames
            for section in [checks_section, contrast_section, severity_section]:
                section.configure(bg=new_section_bg, highlightbackground=new_section_border)
                for child in section.winfo_children():
                    if isinstance(child, tk.Frame):
                        child.configure(bg=new_section_bg)

            # Update title rows and labels
            for title_row, title_label in [(checks_title_row, checks_title_label),
                                            (contrast_title_row, contrast_title_label),
                                            (severity_title_row, severity_title_label)]:
                title_row.configure(bg=new_section_bg)
                title_label.configure(bg=new_section_bg, fg=new_title_color)

            # Update contrast link
            if contrast_link:
                contrast_link.configure(bg=new_section_bg, fg=new_colors['info'])

            # Update content frames
            for content in [checks_content, contrast_content, severity_content]:
                content.configure(bg=new_section_bg)

            # Update toggle frames and widgets
            for toggle in toggle_widgets:
                parent_frame = toggle.master
                if parent_frame:
                    parent_frame.configure(bg=new_section_bg)
                toggle.configure(bg=new_section_bg)
                toggle._icon_label.configure(bg=new_section_bg)
                if toggle._label:
                    toggle._label.configure(bg=new_section_bg, fg=new_colors['text_primary'])

            # Update AAA toggle
            aaa_toggle_frame.configure(bg=new_section_bg)
            aaa_toggle.configure(bg=new_section_bg)
            aaa_toggle._icon_label.configure(bg=new_section_bg)
            if aaa_toggle._label:
                aaa_toggle._label.configure(bg=new_section_bg, fg=new_colors['text_primary'])

            # Update radio groups
            for radio in [contrast_radio, severity_radio]:
                radio.configure(bg=new_section_bg)
                for item in radio._radio_items:
                    item['frame'].configure(bg=new_section_bg)
                    item['icon_label'].configure(bg=new_section_bg)
                    item['labels_frame'].configure(bg=new_section_bg)
                    item['main_label'].configure(bg=new_section_bg, fg=new_colors['text_primary'])
                    if item['desc_label']:
                        item['desc_label'].configure(bg=new_section_bg, fg=new_colors['text_secondary'])

            # Update buttons
            save_btn.update_colors(
                bg=new_colors['button_primary'],
                hover_bg=new_colors['button_primary_hover'],
                pressed_bg=new_colors['button_primary_pressed'],
                canvas_bg=new_colors['background']
            )
            reset_btn.update_colors(
                bg=new_colors['button_secondary'],
                fg=new_colors['text_primary'],
                hover_bg=new_colors['button_secondary_hover'],
                pressed_bg=new_colors['button_secondary_pressed'],
                canvas_bg=new_colors['background']
            )

        # Register the theme callback
        self._theme_manager.register_theme_callback(on_theme_change)

        # Unregister callback when dialog closes
        def on_close():
            self._theme_manager.unregister_theme_callback(on_theme_change)
            dialog.destroy()

        dialog.protocol("WM_DELETE_WINDOW", on_close)
        dialog.bind('<Escape>', lambda event: on_close())

        # Set save_settings cleanup to use on_close
        save_settings._cleanup = on_close

        # Center and show dialog
        dialog.update_idletasks()
        x = parent_window.winfo_rootx() + (parent_window.winfo_width() - 520) // 2
        y = parent_window.winfo_rooty() + (parent_window.winfo_height() - 650) // 2
        dialog.geometry(f"520x650+{x}+{y}")
        dialog.deiconify()

    def show_help_dialog(self) -> None:
        """Show help dialog for accessibility checker - modern styling with orange warning and two-column layout"""
        from core.ui_base import RoundedButton
        from tools.accessibility_checker.tool import AccessibilityCheckerTool
        tool = AccessibilityCheckerTool()
        help_content = tool.get_help_content()

        colors = self._theme_manager.colors

        # Get correct parent window
        parent_window = None
        if hasattr(self, 'main_app') and hasattr(self.main_app, 'root'):
            parent_window = self.main_app.root
        elif hasattr(self, 'master'):
            parent_window = self.master
        else:
            parent_window = self.frame.winfo_toplevel()

        # Create help dialog
        dialog = tk.Toplevel(parent_window)
        dialog.withdraw()  # Hide until styled
        dialog.title(help_content['title'])
        dialog.geometry("900x850")
        dialog.resizable(False, False)
        dialog.transient(parent_window)
        dialog.grab_set()
        dialog.configure(bg=colors['background'])

        # Set AE favicon icon
        try:
            if hasattr(self, 'main_app') and hasattr(self.main_app, 'config'):
                dialog.iconbitmap(self.main_app.config.icon_path)
        except Exception:
            pass

        # Set dark title bar if dark mode
        if hasattr(self, '_set_dialog_title_bar_color'):
            self._set_dialog_title_bar_color(dialog, self._theme_manager.is_dark)

        # Main container
        main_frame = tk.Frame(dialog, bg=colors['background'], padx=20, pady=15)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Header - centered, no emoji
        tk.Label(main_frame, text="Accessibility Checker - Help",
                font=('Segoe UI', 16, 'bold'),
                bg=colors['background'], fg=colors['title_color']).pack(pady=(0, 15))

        # Orange warning section - two-column layout to save vertical space
        warning_bg = colors.get('warning_bg', '#d97706')
        warning_text = colors.get('warning_text', '#ffffff')
        warning_container = tk.Frame(main_frame, bg=warning_bg,
                                   padx=15, pady=10, relief='flat', borderwidth=0)
        warning_container.pack(fill=tk.X, pady=(0, 15))

        tk.Label(warning_container, text="IMPORTANT DISCLAIMERS & REQUIREMENTS",
                font=('Segoe UI', 11, 'bold'),
                bg=warning_bg, fg=warning_text).pack(anchor=tk.W)

        warnings = help_content.get('warnings', [
            "This tool analyzes PBIP enhanced report format (PBIR) files only",
            "Results are guidelines - always verify in Power BI Desktop",
            "NOT officially supported by Microsoft"
        ])

        # Two-column grid for warnings
        warnings_grid = tk.Frame(warning_container, bg=warning_bg)
        warnings_grid.pack(fill=tk.X, pady=(5, 0))
        warnings_grid.columnconfigure(0, weight=1)
        warnings_grid.columnconfigure(1, weight=1)

        # Split warnings into two columns
        mid = (len(warnings) + 1) // 2
        for i, warning in enumerate(warnings):
            col = 0 if i < mid else 1
            row = i if i < mid else i - mid
            tk.Label(warnings_grid, text=f"- {warning}",
                    font=('Segoe UI', 9),
                    bg=warning_bg, fg=warning_text,
                    anchor=tk.W, justify=tk.LEFT).grid(row=row, column=col, sticky='w', padx=(8, 15), pady=1)

        # Two-column layout for sections
        columns_frame = tk.Frame(main_frame, bg=colors['background'])
        columns_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        columns_frame.columnconfigure(0, weight=1)
        columns_frame.columnconfigure(1, weight=1)

        # Split sections by content length for better balance
        # "Check Categories & WCAG Criteria" (index 2) is very long, so balance accordingly
        sections = help_content.get('sections', [])
        # Left: Quick Start, Important Notes, WCAG Levels, Issue Severity, File Requirements
        # Right: Check Categories (long), Color Contrast Requirements
        left_indices = [0, 6, 1, 4, 5]  # Shorter sections (Important Notes after Quick Start)
        right_indices = [2, 3]  # Longer sections (Check Categories, Color Contrast)
        left_sections = [sections[i] for i in left_indices if i < len(sections)]
        right_sections = [sections[i] for i in right_indices if i < len(sections)]

        # Left column
        left_col = tk.Frame(columns_frame, bg=colors['background'])
        left_col.grid(row=0, column=0, sticky='nsew', padx=(0, 10))

        for section in left_sections:
            tk.Label(left_col, text=section['title'],
                    font=('Segoe UI', 11, 'bold'),
                    bg=colors['background'], fg=colors['title_color']).pack(anchor=tk.W, pady=(10, 5))
            for item in section['items']:
                tk.Label(left_col, text=item,
                        font=('Segoe UI', 10),
                        bg=colors['background'], fg=colors['text_primary'],
                        wraplength=400, justify=tk.LEFT).pack(anchor=tk.W, padx=(12, 0))

        # Right column
        right_col = tk.Frame(columns_frame, bg=colors['background'])
        right_col.grid(row=0, column=1, sticky='nsew', padx=(10, 0))

        for section in right_sections:
            tk.Label(right_col, text=section['title'],
                    font=('Segoe UI', 11, 'bold'),
                    bg=colors['background'], fg=colors['title_color']).pack(anchor=tk.W, pady=(10, 5))
            for item in section['items']:
                tk.Label(right_col, text=item,
                        font=('Segoe UI', 10),
                        bg=colors['background'], fg=colors['text_primary'],
                        wraplength=400, justify=tk.LEFT).pack(anchor=tk.W, padx=(12, 0))

        dialog.bind('<Escape>', lambda event: dialog.destroy())

        # Center and show dialog
        dialog.update_idletasks()
        x = parent_window.winfo_rootx() + (parent_window.winfo_width() - 900) // 2
        y = parent_window.winfo_rooty() + (parent_window.winfo_height() - 850) // 2
        dialog.geometry(f"900x850+{x}+{y}")
        dialog.deiconify()

    def log_message(self, message: str) -> None:
        """Add message to log (thread-safe)"""
        def _append():
            if self.log_text:
                self.log_text.config(state=tk.NORMAL)
                self.log_text.insert(tk.END, f"{message}\n")
                self.log_text.config(state=tk.DISABLED)
                self.log_text.see(tk.END)

        # Use after() for thread-safe UI updates
        if threading.current_thread() is threading.main_thread():
            _append()
        else:
            self.frame.after(0, _append)
