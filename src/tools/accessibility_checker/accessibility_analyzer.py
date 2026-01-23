"""
Accessibility Analyzer - Core analysis engine for Power BI accessibility checks
Built by Reid Havens of Analytic Endeavors

This module analyzes Power BI reports for accessibility compliance:
- Tab order navigation sequence
- Alt text for visual elements
- Color contrast (WCAG 2.1 AA/AAA)
- Page titles
- Visual titles
- Data labels
- Bookmark names
- Hidden page warnings
"""

import json
import re
import time
from pathlib import Path
from typing import Dict, List, Any, Optional, Callable, Set, Tuple
from datetime import datetime

from tools.accessibility_checker.accessibility_types import (
    AccessibilitySeverity,
    AccessibilityCheckType,
    AccessibilityIssue,
    TabOrderInfo,
    ColorContrastResult,
    VisualInfo,
    PageInfo,
    BookmarkInfo,
    AccessibilityAnalysisResult,
    DATA_VISUAL_TYPES,
    DECORATIVE_VISUAL_TYPES,
    GENERIC_ALT_TEXT_PATTERNS,
    GENERIC_PAGE_TITLE_PATTERNS,
    GENERIC_BOOKMARK_PATTERNS,
    WCAG_CONTRAST_AA_NORMAL,
    WCAG_CONTRAST_AA_LARGE,
    WCAG_CONTRAST_AAA_NORMAL,
    WCAG_CONTRAST_AAA_LARGE,
)
from tools.accessibility_checker.accessibility_config import (
    get_config,
    AccessibilityCheckConfig,
    CONTRAST_THRESHOLDS,
)
from core.pbi_file_reader import PBIPReader

# Ownership fingerprint
_AE_FP = "QWNjZXNzQW5hbHl6ZXI6QUUtMjAyNA=="


class AccessibilityAnalyzer:
    """
    Analyzes Power BI reports for accessibility compliance and WCAG issues.
    """

    def __init__(self, logger_callback: Optional[Callable[[str], None]] = None,
                 progress_callback: Optional[Callable[[int, str], None]] = None):
        self.log_callback = logger_callback or self._default_log
        self.progress_callback = progress_callback

    def _default_log(self, message: str) -> None:
        """Default logging function"""
        print(message)

    def _update_progress(self, percent: int, message: str) -> None:
        """Update progress if callback is set"""
        if self.progress_callback:
            self.progress_callback(percent, message)

    def analyze_pbip_report(self, pbip_path: str) -> AccessibilityAnalysisResult:
        """
        Main entry point - analyzes a PBIP or PBIR-format PBIX report for accessibility issues.

        Args:
            pbip_path: Path to the .pbip or .pbix file (PBIX must be in PBIR format)

        Returns:
            AccessibilityAnalysisResult with all findings

        Raises:
            FileNotFoundError: If file or .Report folder not found
            ValueError: If PBIX file is not in PBIR format
        """
        start_time = time.time()
        self.log_callback("Starting accessibility analysis...")
        self._update_progress(5, "Loading report files...")

        # Validate report file using shared validation
        reader = PBIPReader()
        validation = reader.validate_report_file(pbip_path)
        if not validation['valid']:
            if validation['error_type'] == 'not_pbir_format':
                raise ValueError(validation['error'])
            else:
                raise FileNotFoundError(validation['error'])

        try:
            report_dir = validation['report_dir']

            # For backward compatibility, keep using pbip_file variable name
            pbip_file = Path(pbip_path)

            self.log_callback(f"Analyzing report: {pbip_file.stem}")

            # Initialize result
            result = AccessibilityAnalysisResult(
                report_path=str(pbip_path),
                report_name=pbip_file.stem
            )

            # Phase 1: Scan pages and extract metadata
            self._update_progress(10, "Scanning pages...")
            result.pages = self._scan_pages(report_dir)
            self.log_callback(f"  Found {len(result.pages)} pages")

            # Phase 2: Scan visuals and extract properties
            self._update_progress(20, "Scanning visuals...")
            result.visuals = self._scan_all_visuals(report_dir)
            self.log_callback(f"  Found {len(result.visuals)} visuals")

            # Phase 3: Scan bookmarks
            self._update_progress(30, "Scanning bookmarks...")
            result.bookmarks = self._scan_bookmarks(report_dir)
            self.log_callback(f"  Found {len(result.bookmarks)} bookmarks")

            # Get config for check filtering
            config = get_config()

            # Phase 4: Run accessibility checks (skip disabled ones)
            if config.is_check_enabled("tab_order"):
                self._update_progress(40, "Checking tab order...")
                result.tab_orders = self._analyze_tab_orders(result.visuals)
                tab_order_issues = self._find_tab_order_issues(result.tab_orders, result.visuals)
                result.issues.extend(tab_order_issues)
                self.log_callback(f"  Tab order: {len(tab_order_issues)} issues")
            else:
                self._update_progress(40, "Skipping tab order (disabled)...")
                result.tab_orders = self._analyze_tab_orders(result.visuals)  # Still get data for totals
                self.log_callback(f"  Tab order: skipped (disabled in settings)")

            # Get visuals/groups that can be toggled visible by bookmarks (needed for alt text AND contrast checks)
            bookmark_toggled_ids = self._get_bookmark_toggled_visuals(report_dir)
            # Build set of hidden group IDs (to check if child visuals inherit hidden state)
            hidden_group_ids = {v.visual_id for v in result.visuals if v.is_hidden and v.is_group}

            if config.is_check_enabled("alt_text"):
                self._update_progress(50, "Checking alt text...")
                alt_text_issues = self._analyze_alt_text(result.visuals, bookmark_toggled_ids, hidden_group_ids)
                result.issues.extend(alt_text_issues)
                self.log_callback(f"  Alt text: {len(alt_text_issues)} issues")
            else:
                self._update_progress(50, "Skipping alt text (disabled)...")
                self.log_callback(f"  Alt text: skipped (disabled in settings)")

            if config.is_check_enabled("color_contrast"):
                self._update_progress(60, "Checking color contrast...")
                result.color_contrasts, contrast_review_issues = self._analyze_color_contrast(
                    report_dir, result.visuals, config, bookmark_toggled_ids, hidden_group_ids
                )
                contrast_issues = self._find_contrast_issues(result.color_contrasts, config)
                result.issues.extend(contrast_issues)
                result.issues.extend(contrast_review_issues)
                total_contrast_issues = len(contrast_issues) + len(contrast_review_issues)
                self.log_callback(f"  Color contrast: {total_contrast_issues} issues ({len(contrast_review_issues)} need manual review)")
            else:
                self._update_progress(60, "Skipping color contrast (disabled)...")
                self.log_callback(f"  Color contrast: skipped (disabled in settings)")

            if config.is_check_enabled("page_title"):
                self._update_progress(70, "Checking page titles...")
                page_title_issues = self._analyze_page_titles(result.pages)
                result.issues.extend(page_title_issues)
                self.log_callback(f"  Page titles: {len(page_title_issues)} issues")
            else:
                self._update_progress(70, "Skipping page titles (disabled)...")
                self.log_callback(f"  Page titles: skipped (disabled in settings)")

            if config.is_check_enabled("visual_title"):
                self._update_progress(75, "Checking visual titles...")
                visual_title_issues = self._analyze_visual_titles(result.visuals, bookmark_toggled_ids, hidden_group_ids)
                result.issues.extend(visual_title_issues)
                self.log_callback(f"  Visual titles: {len(visual_title_issues)} issues")
            else:
                self._update_progress(75, "Skipping visual titles (disabled)...")
                self.log_callback(f"  Visual titles: skipped (disabled in settings)")

            if config.is_check_enabled("data_labels"):
                self._update_progress(80, "Checking data labels...")
                data_label_issues = self._analyze_data_labels(result.visuals, bookmark_toggled_ids, hidden_group_ids)
                result.issues.extend(data_label_issues)
                self.log_callback(f"  Data labels: {len(data_label_issues)} issues")
            else:
                self._update_progress(80, "Skipping data labels (disabled)...")
                self.log_callback(f"  Data labels: skipped (disabled in settings)")

            if config.is_check_enabled("bookmark_name"):
                self._update_progress(85, "Checking bookmark names...")
                bookmark_issues = self._analyze_bookmark_names(result.bookmarks)
                result.issues.extend(bookmark_issues)
                self.log_callback(f"  Bookmark names: {len(bookmark_issues)} issues")
            else:
                self._update_progress(85, "Skipping bookmark names (disabled)...")
                self.log_callback(f"  Bookmark names: skipped (disabled in settings)")

            if config.is_check_enabled("hidden_page"):
                self._update_progress(90, "Checking hidden pages...")
                hidden_page_issues = self._check_hidden_pages(result.pages)
                result.issues.extend(hidden_page_issues)
                self.log_callback(f"  Hidden pages: {len(hidden_page_issues)} issues")
            else:
                self._update_progress(90, "Skipping hidden pages (disabled)...")
                self.log_callback(f"  Hidden pages: skipped (disabled in settings)")

            # Update summary counts
            result.update_counts()
            result.analysis_timestamp = datetime.now().isoformat()
            result.analysis_duration_ms = int((time.time() - start_time) * 1000)

            self._update_progress(100, "Analysis complete")
            self.log_callback(f"Analysis complete. Found {result.total_issues} total issues "
                             f"({result.errors} errors, {result.warnings} warnings, {result.info_count} info)")

            return result

        finally:
            # Clean up extracted temp directory if we extracted from embedded PBIR
            reader.cleanup_extracted_pbir(validation)

    # =========================================================================
    # PAGE SCANNING
    # =========================================================================

    def _scan_pages(self, report_dir: Path) -> List[PageInfo]:
        """Scan all pages and extract page metadata"""
        pages = []

        pages_dir = report_dir / "definition" / "pages"
        if not pages_dir.exists():
            return pages

        page_dirs = [d for d in pages_dir.iterdir() if d.is_dir()]

        for ordinal, page_dir in enumerate(sorted(page_dirs, key=lambda d: d.name)):
            page_json = page_dir / "page.json"
            if not page_json.exists():
                continue

            try:
                with open(page_json, 'r', encoding='utf-8') as f:
                    page_data = json.load(f)

                page_name = page_dir.name
                display_name = page_data.get('displayName', page_name)

                # Check for hidden status
                visibility = page_data.get('visibility')
                is_hidden = visibility == 1 or visibility == "1"

                # Count visuals on page
                visuals_dir = page_dir / "visuals"
                visual_count = len([d for d in visuals_dir.iterdir() if d.is_dir()]) if visuals_dir.exists() else 0

                page_info = PageInfo(
                    page_name=page_name,
                    display_name=display_name,
                    page_id=page_name,
                    is_hidden=is_hidden,
                    has_title=bool(display_name and display_name != page_name),
                    title_text=display_name,
                    visual_count=visual_count,
                    ordinal=ordinal
                )
                pages.append(page_info)

            except Exception as e:
                self.log_callback(f"  Warning: Could not read page {page_dir.name}: {e}")

        return pages

    def _scan_all_visuals(self, report_dir: Path) -> List[VisualInfo]:
        """Scan all visuals across all pages"""
        visuals = []

        pages_dir = report_dir / "definition" / "pages"
        if not pages_dir.exists():
            return visuals

        for page_dir in pages_dir.iterdir():
            if not page_dir.is_dir():
                continue

            # Get page display name
            page_name = page_dir.name
            page_json = page_dir / "page.json"
            if page_json.exists():
                try:
                    with open(page_json, 'r', encoding='utf-8') as f:
                        page_data = json.load(f)
                    page_name = page_data.get('displayName', page_dir.name)
                except:
                    pass

            visuals_dir = page_dir / "visuals"
            if not visuals_dir.exists():
                continue

            for visual_dir in visuals_dir.iterdir():
                if not visual_dir.is_dir():
                    continue

                visual_json = visual_dir / "visual.json"
                if not visual_json.exists():
                    continue

                try:
                    with open(visual_json, 'r', encoding='utf-8') as f:
                        visual_data = json.load(f)

                    visual_info = self._extract_visual_info(visual_data, page_name, visual_dir.name)
                    if visual_info:
                        visuals.append(visual_info)

                except Exception as e:
                    pass  # Silently skip problematic visuals

        return visuals

    def _extract_visual_info(self, visual_data: Dict, page_name: str, visual_dir_name: str) -> Optional[VisualInfo]:
        """Extract visual information from visual.json data"""
        visual_config = visual_data.get('visual', {})
        visual_group = visual_data.get('visualGroup', {})
        is_group = bool(visual_group)

        # Handle visual groups (containers) - they need tracking for hidden state
        if is_group and not visual_config:
            visual_id = visual_data.get('name', visual_dir_name)
            position = visual_data.get('position', {})

            # Check if group is hidden
            is_hidden = visual_data.get('isHidden', False) == True

            return VisualInfo(
                page_name=page_name,
                visual_name=visual_group.get('displayName', visual_id),
                visual_type='visualGroup',
                visual_id=visual_dir_name,
                tab_order=position.get('tabOrder', -1),
                position_x=position.get('x', 0),
                position_y=position.get('y', 0),
                width=position.get('width', 0),
                height=position.get('height', 0),
                is_hidden=is_hidden,
                is_group=True,
                parent_group_name=visual_data.get('parentGroupName')
            )

        if not visual_config:
            return None

        visual_type = visual_config.get('visualType', 'unknown')
        visual_id = visual_data.get('name', visual_dir_name)

        # Get position and size
        position = visual_data.get('position', {})
        x = position.get('x', 0)
        y = position.get('y', 0)
        width = position.get('width', 0)
        height = position.get('height', 0)

        # Get tab order
        tab_order = position.get('tabOrder', -1)

        # Get parent group (if visual is in a group)
        parent_group_name = visual_data.get('parentGroupName')

        # Check if visual is hidden (multiple possible locations)
        is_hidden = False

        # Method 1: Root-level isHidden property
        if visual_data.get('isHidden') == True:
            is_hidden = True

        # Method 2: position.visible property (explicit false)
        if position.get('visible') == False:
            is_hidden = True

        # Get alt text - check both objects and visualContainerObjects
        alt_text = None
        has_alt_text = False

        # Check visualContainerObjects first (more common location)
        container_objects = visual_config.get('visualContainerObjects', {})
        objects = visual_config.get('objects', {})

        # Method 3: visualContainerObjects.general.show property
        general_config = container_objects.get('general', [])
        for obj in general_config:
            props = obj.get('properties', {})
            show_prop = props.get('show', {})
            if show_prop:
                expr = show_prop.get('expr', {})
                literal = expr.get('Literal', {})
                if literal.get('Value', 'true').lower() == 'false':
                    is_hidden = True
                    break

        for source in [container_objects, objects]:
            alt_text_config = source.get('general', [])
            if alt_text_config:
                for obj in alt_text_config:
                    props = obj.get('properties', {})
                    alt_text_prop = props.get('altText', {})
                    if alt_text_prop:
                        expr = alt_text_prop.get('expr', {})
                        literal = expr.get('Literal', {})
                        value = literal.get('Value', '')
                        if value:
                            alt_text = value.strip("'\"")
                            has_alt_text = bool(alt_text)
                            break
                if alt_text:
                    break

        # Get visual title from visualContainerObjects
        title_text = None
        has_title = False
        title_explicitly_disabled = False
        title_config = container_objects.get('title', [])
        if title_config:
            for obj in title_config:
                props = obj.get('properties', {})

                # Check if title is explicitly disabled
                show_prop = props.get('show', {})
                if show_prop:
                    expr = show_prop.get('expr', {})
                    literal = expr.get('Literal', {})
                    value = literal.get('Value', 'true')
                    if value.lower() == 'false':
                        title_explicitly_disabled = True
                        continue

                # Title section exists and is not disabled - title is enabled
                # (Power BI auto-generates title from field names if no explicit text)
                has_title = True

                text_prop = props.get('text', {}) or props.get('titleText', {})
                if text_prop:
                    expr = text_prop.get('expr', {})
                    literal = expr.get('Literal', {})
                    value = literal.get('Value', '')
                    if value:
                        title_text = value.strip("'\"")
                        break

        # If no title config at all, or only disabled entries, check if disabled
        if not title_config:
            has_title = False
        elif title_explicitly_disabled and not has_title:
            has_title = False

        # Try to extract auto-generated title from query displayName
        auto_title = None
        query = visual_config.get('query', {})
        query_state = query.get('queryState', {})
        # Check common value wells: Y, Values, Value
        for well_name in ['Y', 'Values', 'Value']:
            well = query_state.get(well_name, {})
            projections = well.get('projections', [])
            if projections and isinstance(projections, list):
                first_proj = projections[0]
                if isinstance(first_proj, dict):
                    display_name = first_proj.get('displayName')
                    if display_name:
                        auto_title = display_name
                        break

        # Build visual_name: prefer explicit title, then auto-title, then ID
        # Always append visual type in parentheses
        display_title = title_text if title_text else (auto_title if auto_title else visual_id)
        visual_name = f"{display_title} ({visual_type})"

        # Check for data labels (in objects, not visualContainerObjects)
        has_data_labels = None
        if visual_type.lower() in [v.lower() for v in DATA_VISUAL_TYPES]:
            has_data_labels = False
            labels_config = objects.get('labels', [])
            if labels_config:
                for obj in labels_config:
                    props = obj.get('properties', {})
                    show_prop = props.get('show', {})
                    if show_prop:
                        expr = show_prop.get('expr', {})
                        literal = expr.get('Literal', {})
                        value = literal.get('Value', 'false')
                        if value.lower() == 'true':
                            has_data_labels = True
                            break

        # Determine if this is a data visual
        is_data_visual = visual_type.lower() in [v.lower() for v in DATA_VISUAL_TYPES]

        return VisualInfo(
            page_name=page_name,
            visual_name=visual_name,
            visual_type=visual_type,
            visual_id=visual_dir_name,
            has_alt_text=has_alt_text,
            alt_text=alt_text,
            has_title=has_title,
            title_text=title_text,
            has_data_labels=has_data_labels,
            tab_order=tab_order,
            position_x=x,
            position_y=y,
            width=width,
            height=height,
            is_data_visual=is_data_visual,
            is_hidden=is_hidden,
            is_group=False,
            parent_group_name=parent_group_name
        )

    def _get_bookmark_toggled_visuals(self, report_dir: Path) -> Set[str]:
        """Scan all bookmarks to find visual IDs that can be toggled visible.

        Returns set of visual IDs (and group IDs) referenced by any bookmark.
        """
        toggled_ids = set()
        bookmarks_dir = report_dir / "definition" / "bookmarks"

        if not bookmarks_dir.exists():
            return toggled_ids

        for bookmark_dir in bookmarks_dir.iterdir():
            if not bookmark_dir.is_dir():
                continue
            bookmark_file = bookmark_dir / "bookmark.json"
            if not bookmark_file.exists():
                continue

            try:
                bookmark_data = json.loads(bookmark_file.read_text(encoding='utf-8'))

                # Method 1: targetVisualNames in options
                options = bookmark_data.get('options', {})
                target_visuals = options.get('targetVisualNames', [])
                toggled_ids.update(target_visuals)

                # Method 2: visualContainerGroups in explorationState (group toggles)
                exploration = bookmark_data.get('explorationState', {})
                sections = exploration.get('sections', {})
                for page_id, page_data in sections.items():
                    if isinstance(page_data, dict):
                        groups = page_data.get('visualContainerGroups', {})
                        toggled_ids.update(groups.keys())
            except Exception:
                pass

        return toggled_ids

    def _scan_bookmarks(self, report_dir: Path) -> List[BookmarkInfo]:
        """Scan bookmarks from the report"""
        bookmarks = []

        bookmarks_dir = report_dir / "definition" / "bookmarks"
        if not bookmarks_dir.exists():
            return bookmarks

        bookmarks_json = bookmarks_dir / "bookmarks.json"
        if not bookmarks_json.exists():
            return bookmarks

        try:
            with open(bookmarks_json, 'r', encoding='utf-8') as f:
                bookmarks_data = json.load(f)

            bookmark_items = bookmarks_data.get('items', [])

            for bookmark_meta in bookmark_items:
                bookmark_id = bookmark_meta.get('name', '')
                display_name = bookmark_meta.get('displayName', bookmark_id)

                # Check for generic naming
                is_generic = self._is_generic_bookmark_name(display_name)

                bookmarks.append(BookmarkInfo(
                    bookmark_name=bookmark_id,
                    display_name=display_name,
                    bookmark_id=bookmark_id,
                    is_generic_name=is_generic
                ))

                # Process children if present
                children = bookmark_meta.get('children', [])
                for child_id in children:
                    child_file = bookmarks_dir / f"{child_id}.bookmark.json"
                    if child_file.exists():
                        try:
                            with open(child_file, 'r', encoding='utf-8') as f:
                                child_data = json.load(f)

                            child_display = child_data.get('displayName', child_id)
                            is_child_generic = self._is_generic_bookmark_name(child_display)

                            bookmarks.append(BookmarkInfo(
                                bookmark_name=child_id,
                                display_name=child_display,
                                bookmark_id=child_id,
                                is_generic_name=is_child_generic
                            ))
                        except:
                            pass

        except Exception as e:
            self.log_callback(f"  Warning: Could not read bookmarks: {e}")

        return bookmarks

    # =========================================================================
    # TAB ORDER ANALYSIS
    # =========================================================================

    def _analyze_tab_orders(self, visuals: List[VisualInfo]) -> List[TabOrderInfo]:
        """Extract tab order information from visuals"""
        tab_orders = []

        for visual in visuals:
            tab_order_info = TabOrderInfo(
                page_name=visual.page_name,
                visual_name=visual.visual_name,
                visual_type=visual.visual_type,
                tab_order=visual.tab_order,
                position_x=visual.position_x,
                position_y=visual.position_y,
                position_z=0,
                width=visual.width,
                height=visual.height,
                visual_id=visual.visual_id
            )
            tab_orders.append(tab_order_info)

        return tab_orders

    def _find_tab_order_issues(self, tab_orders: List[TabOrderInfo], visuals: List[VisualInfo]) -> List[AccessibilityIssue]:
        """Find tab order issues"""
        issues = []

        # Group by page
        pages = {}
        for tab_info in tab_orders:
            if tab_info.page_name not in pages:
                pages[tab_info.page_name] = []
            pages[tab_info.page_name].append(tab_info)

        for page_name, page_tab_orders in pages.items():
            # Check for missing tab orders
            missing_tab_orders = [t for t in page_tab_orders if t.tab_order == -1]
            if missing_tab_orders:
                for tab_info in missing_tab_orders:
                    issues.append(AccessibilityIssue(
                        check_type=AccessibilityCheckType.TAB_ORDER,
                        severity=AccessibilitySeverity.WARNING,
                        page_name=page_name,
                        visual_name=tab_info.visual_name,
                        visual_type=tab_info.visual_type,
                        issue_description="Visual does not have a tab order defined",
                        recommendation="Set a tab order for this visual to ensure keyboard navigation works correctly",
                        current_value="Not set",
                        wcag_reference="WCAG 2.1 - 2.4.3 Focus Order"
                    ))

            # Check for duplicate tab orders
            valid_tab_orders = [t for t in page_tab_orders if t.tab_order >= 0]
            tab_order_values = [t.tab_order for t in valid_tab_orders]
            duplicates = set([x for x in tab_order_values if tab_order_values.count(x) > 1])

            if duplicates:
                for dup_order in duplicates:
                    dup_visuals = [t for t in valid_tab_orders if t.tab_order == dup_order]
                    visual_names = ", ".join([t.visual_name for t in dup_visuals])
                    issues.append(AccessibilityIssue(
                        check_type=AccessibilityCheckType.TAB_ORDER,
                        severity=AccessibilitySeverity.WARNING,
                        page_name=page_name,
                        issue_description=f"Multiple visuals share the same tab order ({dup_order}): {visual_names}",
                        recommendation="Assign unique tab order values to each visual for predictable keyboard navigation",
                        current_value=str(dup_order),
                        wcag_reference="WCAG 2.1 - 2.4.3 Focus Order"
                    ))

            # Check if tab order matches logical reading order (top-to-bottom, left-to-right)
            if valid_tab_orders:
                sorted_by_tab = sorted(valid_tab_orders, key=lambda t: t.tab_order)
                sorted_by_position = sorted(valid_tab_orders, key=lambda t: (round(t.position_y / 50) * 50, t.position_x))

                # Compare order - allow some tolerance
                mismatches = 0
                for i, (by_tab, by_pos) in enumerate(zip(sorted_by_tab, sorted_by_position)):
                    if by_tab.visual_id != by_pos.visual_id:
                        mismatches += 1

                if mismatches > len(valid_tab_orders) // 2:  # More than half mismatched
                    issues.append(AccessibilityIssue(
                        check_type=AccessibilityCheckType.TAB_ORDER,
                        severity=AccessibilitySeverity.INFO,
                        page_name=page_name,
                        issue_description="Tab order may not match visual reading order (top-to-bottom, left-to-right)",
                        recommendation="Consider reordering tab sequence to match the visual layout for intuitive navigation",
                        wcag_reference="WCAG 2.1 - 2.4.3 Focus Order"
                    ))

        return issues

    # =========================================================================
    # ALT TEXT ANALYSIS
    # =========================================================================

    def _analyze_alt_text(
        self,
        visuals: List[VisualInfo],
        bookmark_toggled_ids: Set[str] = None,
        hidden_group_ids: Set[str] = None
    ) -> List[AccessibilityIssue]:
        """Analyze alt text for visuals

        Args:
            visuals: List of visual info objects to analyze
            bookmark_toggled_ids: Set of visual/group IDs that can be toggled visible by bookmarks
            hidden_group_ids: Set of group IDs that are hidden
        """
        issues = []
        bookmark_toggled_ids = bookmark_toggled_ids or set()
        hidden_group_ids = hidden_group_ids or set()

        for visual in visuals:
            # Skip visual groups (they're containers, not content)
            if visual.is_group:
                continue

            # Check if visual is effectively hidden (directly or via parent group)
            is_effectively_hidden = visual.is_hidden or (
                visual.parent_group_name and visual.parent_group_name in hidden_group_ids
            )

            if is_effectively_hidden:
                # Skip ONLY if no bookmark can toggle this visual/group visible
                can_be_toggled = (
                    visual.visual_id in bookmark_toggled_ids or
                    (visual.parent_group_name and visual.parent_group_name in bookmark_toggled_ids)
                )
                if not can_be_toggled:
                    continue  # Skip truly hidden visuals

            # Data visuals need alt text
            if visual.is_data_visual:
                if not visual.has_alt_text:
                    issues.append(AccessibilityIssue(
                        check_type=AccessibilityCheckType.ALT_TEXT,
                        severity=AccessibilitySeverity.WARNING,
                        page_name=visual.page_name,
                        visual_name=visual.visual_name,
                        visual_type=visual.visual_type,
                        issue_description="Data visual is missing alt text for screen readers",
                        recommendation="Add descriptive alt text that conveys the meaning of the data",
                        wcag_reference="WCAG 2.1 - 1.1.1 Non-text Content"
                    ))
                elif visual.alt_text:
                    # Check if alt text is meaningful
                    if not self._is_alt_text_meaningful(visual.alt_text, visual.visual_type):
                        issues.append(AccessibilityIssue(
                            check_type=AccessibilityCheckType.ALT_TEXT,
                            severity=AccessibilitySeverity.WARNING,
                            page_name=visual.page_name,
                            visual_name=visual.visual_name,
                            visual_type=visual.visual_type,
                            issue_description="Alt text appears to be generic or not descriptive",
                            recommendation="Replace with meaningful alt text that describes what the visual shows",
                            current_value=visual.alt_text,
                            wcag_reference="WCAG 2.1 - 1.1.1 Non-text Content"
                        ))

                    # Check alt text length
                    if len(visual.alt_text) > 125:
                        issues.append(AccessibilityIssue(
                            check_type=AccessibilityCheckType.ALT_TEXT,
                            severity=AccessibilitySeverity.INFO,
                            page_name=visual.page_name,
                            visual_name=visual.visual_name,
                            visual_type=visual.visual_type,
                            issue_description=f"Alt text is longer than recommended (125 chars): {len(visual.alt_text)} chars",
                            recommendation="Consider shortening alt text to be more concise",
                            current_value=visual.alt_text[:50] + "...",
                            wcag_reference="WCAG 2.1 - 1.1.1 Non-text Content"
                        ))

            # Decorative visuals - just info if missing alt text
            elif visual.visual_type.lower() in [v.lower() for v in DECORATIVE_VISUAL_TYPES]:
                if visual.has_alt_text and visual.alt_text:
                    # Has alt text - check if it should be marked decorative instead
                    pass  # This is fine

        return issues

    def _is_alt_text_meaningful(self, alt_text: str, visual_type: str) -> bool:
        """Check if alt text is meaningful (not generic)"""
        if not alt_text:
            return False

        normalized = alt_text.lower().strip()

        # Check against generic patterns
        for pattern in GENERIC_ALT_TEXT_PATTERNS:
            if normalized == pattern or normalized == pattern + "s":
                return False

        # Check if it's just the visual type
        if normalized == visual_type.lower():
            return False

        # Check if too short
        if len(normalized) < 5:
            return False

        return True

    # =========================================================================
    # COLOR CONTRAST ANALYSIS
    # =========================================================================

    def _extract_page_settings(self, page_data: Dict, theme_colors: Dict[str, Any]) -> Dict[str, Any]:
        """Extract page background and wallpaper settings from page.json.

        Args:
            page_data: Parsed page.json data
            theme_colors: Theme colors for resolving ThemeDataColor references

        Returns:
            Dict with 'background_color', 'background_transparency', 'wallpaper_color'
        """
        settings = {
            'background_color': None,
            'background_transparency': 0,
            'wallpaper_color': None
        }

        objects = page_data.get('objects', {})

        # Extract canvas background (objects.background)
        background_config = objects.get('background', [])
        for obj in background_config:
            props = obj.get('properties', {})

            # Extract background color
            color_prop = props.get('color', {})
            settings['background_color'] = self._extract_color_from_prop(color_prop, theme_colors)

            # Extract transparency (format: "0D", "50D" where D suffix = decimal)
            trans_prop = props.get('transparency', {})
            trans_expr = trans_prop.get('expr', {})
            trans_literal = trans_expr.get('Literal', {})
            trans_value = trans_literal.get('Value', '0D')
            if trans_value:
                # Parse "50D" → 50, handle potential string or int
                if isinstance(trans_value, str):
                    trans_value = trans_value.rstrip('D')
                    try:
                        settings['background_transparency'] = int(float(trans_value))
                    except (ValueError, TypeError):
                        settings['background_transparency'] = 0
                elif isinstance(trans_value, (int, float)):
                    settings['background_transparency'] = int(trans_value)
            break

        # Extract wallpaper (objects.outspace)
        outspace_config = objects.get('outspace', [])
        for obj in outspace_config:
            props = obj.get('properties', {})
            color_prop = props.get('color', {})
            settings['wallpaper_color'] = self._extract_color_from_prop(color_prop, theme_colors)
            break

        return settings

    def _analyze_color_contrast(self, report_dir: Path, visuals: List[VisualInfo],
                                  config: AccessibilityCheckConfig = None,
                                  bookmark_toggled_ids: Set[str] = None,
                                  hidden_group_ids: Set[str] = None) -> Tuple[List[ColorContrastResult], List[AccessibilityIssue]]:
        """Analyze color contrast in the report.

        Args:
            report_dir: Path to report directory
            visuals: List of visual info objects
            config: Optional config for contrast thresholds
            bookmark_toggled_ids: Set of visual/group IDs that can be toggled visible by bookmarks
            hidden_group_ids: Set of group IDs that are hidden

        Returns:
            Tuple of (ColorContrastResult list, AccessibilityIssue list for unknown colors)
        """
        if config is None:
            config = get_config()
        results = []
        review_issues = []

        # Load theme colors if available
        theme_colors = self._load_theme_colors(report_dir)

        pages_dir = report_dir / "definition" / "pages"

        # Initialize sets if not provided
        bookmark_toggled_ids = bookmark_toggled_ids or set()
        hidden_group_ids = hidden_group_ids or set()

        # For each visual, load its config and check contrast
        for visual in visuals:
            # Check if visual is effectively hidden (directly or via parent group)
            is_effectively_hidden = visual.is_hidden or (
                visual.parent_group_name and visual.parent_group_name in hidden_group_ids
            )

            if is_effectively_hidden:
                # Skip ONLY if no bookmark can toggle this visual/group visible
                can_be_toggled = (
                    visual.visual_id in bookmark_toggled_ids or
                    (visual.parent_group_name and visual.parent_group_name in bookmark_toggled_ids)
                )
                if not can_be_toggled:
                    continue  # Skip truly hidden visuals

            # We need to find the page directory - it might be the page_name or we need to search
            visual_json = None
            page_settings = {}  # Will hold page background/wallpaper settings

            # Try different page directory naming patterns
            for page_dir in pages_dir.iterdir():
                if not page_dir.is_dir():
                    continue

                # Check if this page matches by looking at page.json display name
                page_json_file = page_dir / "page.json"
                if page_json_file.exists():
                    try:
                        with open(page_json_file, 'r', encoding='utf-8') as f:
                            page_data = json.load(f)
                        page_display_name = page_data.get('displayName', page_dir.name)
                        if page_display_name == visual.page_name or page_dir.name == visual.page_name:
                            # Found the page, extract page settings for background compositing
                            page_settings = self._extract_page_settings(page_data, theme_colors)
                            # Now look for the visual
                            potential_visual_json = page_dir / "visuals" / visual.visual_id / "visual.json"
                            if potential_visual_json.exists():
                                visual_json = potential_visual_json
                                break
                    except:
                        pass

            if not visual_json:
                continue

            try:
                with open(visual_json, 'r', encoding='utf-8') as f:
                    visual_config = json.load(f)

                # _check_visual_contrast now returns a list of all element contrasts
                contrast_results = self._check_visual_contrast(visual, theme_colors, visual_config, page_settings)
                if contrast_results:
                    results.extend(contrast_results)
                else:
                    # Could not determine colors - flag for manual review
                    # Only flag data visuals (charts, tables, etc.) not decorative elements
                    if visual.is_data_visual:
                        review_issues.append(AccessibilityIssue(
                            check_type=AccessibilityCheckType.COLOR_CONTRAST,
                            severity=AccessibilitySeverity.INFO,
                            page_name=visual.page_name,
                            visual_name=visual.visual_name,
                            visual_type=visual.visual_type,
                            issue_description=f"Unable to automatically detect text/data colors for '{visual.visual_name}' ({visual.visual_type})",
                            recommendation="Manually verify color contrast meets WCAG AA (4.5:1) or AAA (7:1) requirements using a contrast checker tool",
                            wcag_reference="WCAG 2.1 - 1.4.3 Contrast (Minimum)"
                        ))
            except Exception:
                pass

        return results, review_issues

    def _load_theme_colors(self, report_dir: Path) -> Dict[str, Any]:
        """Load theme colors from the report.

        Returns dict with:
        - 'dataColors': list of hex colors indexed by ColorId
        - 'background': default background color
        - 'foreground': default foreground color
        """
        colors = {
            'dataColors': [],
            'background': '#FFFFFF',
            'foreground': '#000000'
        }

        # Priority 1: Load from RegisteredResources (custom themes) - these override base themes
        registered_dir = report_dir / "StaticResources" / "RegisteredResources"
        if registered_dir.exists():
            for theme_file in registered_dir.glob("*.json"):
                try:
                    with open(theme_file, 'r', encoding='utf-8') as f:
                        theme_data = json.load(f)

                    # Only process if it has dataColors (it's a theme file)
                    if 'dataColors' in theme_data:
                        colors['dataColors'] = theme_data.get('dataColors', [])
                        colors['background'] = theme_data.get('background', '#FFFFFF')
                        colors['foreground'] = theme_data.get('foreground', '#000000')
                        return colors  # Use first custom theme found
                except:
                    pass

        # Priority 2: Load from base themes
        themes_dir = report_dir / "StaticResources" / "SharedResources" / "BaseThemes"
        if themes_dir.exists():
            for theme_file in themes_dir.glob("*.json"):
                try:
                    with open(theme_file, 'r', encoding='utf-8') as f:
                        theme_data = json.load(f)

                    if 'dataColors' in theme_data:
                        colors['dataColors'] = theme_data.get('dataColors', [])
                        colors['background'] = theme_data.get('background', '#FFFFFF')
                        colors['foreground'] = theme_data.get('foreground', '#000000')
                        return colors
                except:
                    pass

        return colors

    def _extract_color_from_prop(self, prop: Dict, theme_colors: Dict[str, Any] = None) -> Optional[str]:
        """Extract hex color from vcObjects property structure.

        Handles:
        - Literal hex values ('#FF0000')
        - ThemeDataColor references (ColorId + Percent)
        - SolidColor theme references

        Args:
            prop: The property dict containing color info
            theme_colors: Dict with 'dataColors' list for resolving ThemeDataColor

        Returns:
            Hex color string or None if not extractable
        """
        if not prop:
            return None

        # Handle nested 'solid' > 'color' structure
        if 'solid' in prop:
            prop = prop['solid'].get('color', prop)

        expr = prop.get('expr', {})

        # 1. Check for Literal value (explicit hex color)
        literal = expr.get('Literal', {})
        if literal:
            value = literal.get('Value', '')
            if value:
                color = value.strip("'\"")
                if color.startswith('#') and len(color) in (4, 7, 9):
                    return color

        # 2. Check for ThemeDataColor (theme color by index)
        theme_data_color = expr.get('ThemeDataColor', {})
        if theme_data_color and theme_colors:
            color_id = theme_data_color.get('ColorId', 0)
            percent = theme_data_color.get('Percent', 0)

            # ColorId 0-1 map to background/foreground, 2+ map to dataColors
            # This matches Power BI's color picker: columns 1-2 are white/black gradients,
            # columns 3-10 are theme dataColors
            if color_id == 0:
                base_color = theme_colors.get('background', '#FFFFFF')
            elif color_id == 1:
                base_color = theme_colors.get('foreground', '#000000')
            else:
                data_colors = theme_colors.get('dataColors', [])
                adjusted_id = color_id - 2
                if data_colors and 0 <= adjusted_id < len(data_colors):
                    base_color = data_colors[adjusted_id]
                else:
                    return None

            if percent != 0:
                return self._adjust_color_brightness(base_color, percent)
            return base_color

        # 3. Check for SolidColor (direct theme reference)
        solid = expr.get('SolidColor', {})
        if solid:
            color = solid.get('color', None)
            if color and color.startswith('#'):
                return color

        # 4. Check for FillRule (conditional formatting/gradients)
        # Extract the min or max color from gradient rules
        fill_rule = expr.get('FillRule', {})
        if fill_rule:
            inner_rule = fill_rule.get('FillRule', {})
            # Try linearGradient2 (two-color gradient)
            gradient2 = inner_rule.get('linearGradient2', {})
            if gradient2:
                # Use min color as the "base" color for contrast check
                min_color = gradient2.get('min', {}).get('color', {})
                if min_color:
                    literal = min_color.get('Literal', {})
                    if literal:
                        value = literal.get('Value', '')
                        if value:
                            color = value.strip("'\"")
                            if color.startswith('#'):
                                return color
            # Try linearGradient3 (three-color gradient)
            gradient3 = inner_rule.get('linearGradient3', {})
            if gradient3:
                min_color = gradient3.get('min', {}).get('color', {})
                if min_color:
                    literal = min_color.get('Literal', {})
                    if literal:
                        value = literal.get('Value', '')
                        if value:
                            color = value.strip("'\"")
                            if color.startswith('#'):
                                return color

        return None

    def _adjust_color_brightness(self, hex_color: str, percent: float) -> str:
        """Adjust color brightness by percentage.

        Args:
            hex_color: Hex color string (e.g., '#FF0000')
            percent: Adjustment percentage (-1 to 1, negative=darker, positive=lighter)

        Returns:
            Adjusted hex color string
        """
        try:
            hex_color = hex_color.lstrip('#')
            r = int(hex_color[0:2], 16)
            g = int(hex_color[2:4], 16)
            b = int(hex_color[4:6], 16)

            if percent > 0:
                # Lighten: move toward white
                r = int(r + (255 - r) * percent)
                g = int(g + (255 - g) * percent)
                b = int(b + (255 - b) * percent)
            else:
                # Darken: move toward black
                r = int(r * (1 + percent))
                g = int(g * (1 + percent))
                b = int(b * (1 + percent))

            # Clamp to valid range
            r = max(0, min(255, r))
            g = max(0, min(255, g))
            b = max(0, min(255, b))

            return f'#{r:02x}{g:02x}{b:02x}'
        except:
            return hex_color

    def _calculate_contrast_result(
        self, visual: VisualInfo,
        fg_color: str, bg_color: str,
        fg_element: str, bg_element: str,
        theme_colors: Dict[str, Any]
    ) -> Optional[ColorContrastResult]:
        """Calculate contrast between fg and bg colors, return result.

        Args:
            visual: Visual info for page/visual metadata
            fg_color: Foreground hex color (may include alpha)
            bg_color: Background hex color (may include alpha)
            fg_element: Description of foreground element (e.g., "Title text")
            bg_element: Description of background element (e.g., "Visual background")
            theme_colors: Theme colors dict for fallback colors

        Returns:
            ColorContrastResult with contrast ratio and pass/fail status, or None on error.
        """
        try:
            # Check for transparency
            fg_has_transparency = self._has_transparency(fg_color)
            bg_has_transparency = self._has_transparency(bg_color)
            transparency_layer_count = 0
            blended_fg = None
            blended_bg = None

            # Handle transparency by blending colors
            if fg_has_transparency or bg_has_transparency:
                blended_fg, layer_count = self._blend_colors(fg_color, bg_color)
                transparency_layer_count = layer_count
                fg_rgb = self._hex_to_rgb(blended_fg)
            else:
                fg_rgb = self._hex_to_rgb(fg_color)

            bg_rgb = self._hex_to_rgb(bg_color[:7] if len(bg_color) > 7 else bg_color)

            fg_lum = self._calculate_luminance(*fg_rgb)
            bg_lum = self._calculate_luminance(*bg_rgb)
            ratio = self._calculate_contrast_ratio(fg_lum, bg_lum)

            # Create descriptive element type showing what was checked
            element_type = f"{fg_element} vs {bg_element}"

            # Mark as requiring review if too many transparent layers
            requires_review = transparency_layer_count >= self.MAX_TRANSPARENT_LAYERS

            return ColorContrastResult(
                page_name=visual.page_name,
                visual_name=visual.visual_name,
                visual_type=visual.visual_type,
                foreground_color=fg_color,
                background_color=bg_color,
                contrast_ratio=ratio,
                passes_aa_normal=ratio >= WCAG_CONTRAST_AA_NORMAL,
                passes_aa_large=ratio >= WCAG_CONTRAST_AA_LARGE,
                passes_aaa_normal=ratio >= WCAG_CONTRAST_AAA_NORMAL,
                passes_aaa_large=ratio >= WCAG_CONTRAST_AAA_LARGE,
                luminance_fg=fg_lum,
                luminance_bg=bg_lum,
                element_type=element_type,
                fg_has_transparency=fg_has_transparency,
                bg_has_transparency=bg_has_transparency,
                transparency_layer_count=transparency_layer_count,
                requires_review=requires_review,
                blended_fg_color=blended_fg,
                blended_bg_color=blended_bg
            )
        except Exception:
            return None

    def _check_visual_contrast(self, visual: VisualInfo, theme_colors: Dict[str, Any],
                               visual_config: Dict, page_settings: Dict[str, Any] = None) -> List[ColorContrastResult]:
        """Check contrast for ALL text elements in a visual.

        Returns list of ColorContrastResult for each element with explicit colors.
        Empty list if no explicit colors found (caller creates "Review Required" issue).

        Checks elements in order: Subtitle, Title, Data labels, Category labels,
        Data points, Legend, Axis labels, Table headers. Each element with an
        explicit color gets its own contrast result.

        Args:
            visual: VisualInfo object
            theme_colors: Theme colors dict
            visual_config: Parsed visual.json
            page_settings: Page background/wallpaper settings from page.json
        """
        if page_settings is None:
            page_settings = {}

        visual_node = visual_config.get('visual', {})
        # Power BI uses two different property containers:
        # - 'objects': data-related (labels, dataPoint, legend, axes)
        # - 'visualContainerObjects': container-related (title, background, border)
        objects = visual_node.get('objects', {})
        container_objects = visual_node.get('visualContainerObjects', {})

        # Collect all contrast results for this visual
        results: List[ColorContrastResult] = []

        # =========================================================================
        # FALLBACK BACKGROUND COLOR EXTRACTION WITH LAYER COMPOSITING
        # Power BI layer stack (bottom to top):
        # 1. Wallpaper (page outspace)
        # 2. Page canvas background (with optional transparency)
        # 3. Visual background (if enabled, with optional transparency)
        # 4. Visual content (text, labels, etc.)
        # =========================================================================

        # First, compute the effective page background (wallpaper + canvas)
        effective_page_bg = self._compute_effective_page_background(page_settings, theme_colors)

        fallback_bg_color = None
        fallback_bg_source = "Page background"

        # 1. Check if visual background is ENABLED (show property)
        bg_config = container_objects.get('background', [])
        visual_bg_enabled = False
        visual_bg_color = None
        visual_bg_transparency = 0

        for obj in bg_config:
            props = obj.get('properties', {})

            # Check show property - default to true if not specified
            show_prop = props.get('show', {})
            if show_prop:
                show_expr = show_prop.get('expr', {})
                show_literal = show_expr.get('Literal', {})
                show_value = show_literal.get('Value', 'true')
                # Handle both string and boolean values
                if isinstance(show_value, bool):
                    visual_bg_enabled = show_value
                else:
                    visual_bg_enabled = str(show_value).lower() != 'false'
            else:
                # No show property = enabled by default
                visual_bg_enabled = True

            if visual_bg_enabled:
                visual_bg_color = self._extract_color_from_prop(props.get('color', {}), theme_colors)

                # Extract transparency (parse "50D" → 50)
                trans_prop = props.get('transparency', {})
                trans_expr = trans_prop.get('expr', {})
                trans_literal = trans_expr.get('Literal', {})
                trans_value = trans_literal.get('Value', '0D')
                if trans_value:
                    if isinstance(trans_value, str):
                        trans_value = trans_value.rstrip('D')
                        try:
                            visual_bg_transparency = int(float(trans_value))
                        except (ValueError, TypeError):
                            visual_bg_transparency = 0
                    elif isinstance(trans_value, (int, float)):
                        visual_bg_transparency = int(trans_value)
            break

        # 2. Compute effective background based on visual background state
        if visual_bg_enabled and visual_bg_color:
            if visual_bg_transparency == 0:
                # Fully opaque visual background
                fallback_bg_color = visual_bg_color
                fallback_bg_source = "Visual background"
            else:
                # Visual background with transparency - blend over page background
                alpha = 1.0 - (visual_bg_transparency / 100.0)
                fallback_bg_color = self._blend_with_alpha(visual_bg_color, effective_page_bg, alpha)
                fallback_bg_source = f"Visual background ({visual_bg_transparency}% transparent)"
        else:
            # Visual background is OFF or no color - use effective page background
            fallback_bg_color = effective_page_bg
            fallback_bg_source = "Page background"

        # 3. Check general background as secondary source (could be in either)
        if not fallback_bg_color or fallback_bg_source == "Page background":
            for source in [container_objects, objects]:
                general_config = source.get('general', [])
                for obj in general_config:
                    props = obj.get('properties', {})
                    general_bg = self._extract_color_from_prop(props.get('background', {}), theme_colors)
                    if not general_bg:
                        general_bg = self._extract_color_from_prop(props.get('backgroundColor', {}), theme_colors)
                    if general_bg:
                        fallback_bg_color = general_bg
                        fallback_bg_source = "Visual background"
                        break
                if fallback_bg_source == "Visual background":
                    break

        # 4. Check stylePreset or visualContainerStyle
        if not fallback_bg_color or fallback_bg_source == "Page background":
            for style_name in ['visualContainerStyle', 'stylePreset']:
                style_config = container_objects.get(style_name, [])
                for obj in style_config:
                    props = obj.get('properties', {})
                    style_bg = self._extract_color_from_prop(props.get('background', {}), theme_colors)
                    if style_bg:
                        fallback_bg_color = style_bg
                        fallback_bg_source = "Visual style background"
                        break
                if fallback_bg_source == "Visual style background":
                    break

        # 5. Check slicer items background (objects.items for slicers)
        # Slicers have a separate items background that sits on top of the visual container
        slicer_items_bg = None
        if not fallback_bg_color or fallback_bg_source == "Page background":
            items_config = objects.get('items', [])
            for obj in items_config:
                props = obj.get('properties', {})
                items_bg = self._extract_color_from_prop(props.get('background', {}), theme_colors)
                if items_bg:
                    slicer_items_bg = items_bg
                    fallback_bg_color = items_bg
                    fallback_bg_source = "Slicer items background"
                    break
        else:
            # Even if we have a visual container BG, check for slicer items BG which sits on top
            items_config = objects.get('items', [])
            for obj in items_config:
                props = obj.get('properties', {})
                items_bg = self._extract_color_from_prop(props.get('background', {}), theme_colors)
                if items_bg:
                    slicer_items_bg = items_bg
                    # Slicer items background takes precedence for slicer text
                    fallback_bg_color = items_bg
                    fallback_bg_source = "Slicer items background"
                    break

        # 6. Final fallback to effective page background if still nothing
        if not fallback_bg_color:
            fallback_bg_color = effective_page_bg
            fallback_bg_source = "Page background"

        # =========================================================================
        # CHECK ALL TEXT ELEMENTS FOR CONTRAST
        # Each element with explicit colors gets its own contrast result
        # =========================================================================

        # 1. Check subtitle text (in visualContainerObjects)
        subtitle_config = container_objects.get('subTitle', [])
        if not subtitle_config:
            subtitle_config = container_objects.get('subtitle', [])
        for obj in subtitle_config:
            props = obj.get('properties', {})
            # Skip if subtitle is disabled (show=false)
            show_prop = props.get('show', {})
            if show_prop:
                expr = show_prop.get('expr', {})
                literal = expr.get('Literal', {})
                if literal.get('Value', 'true').lower() == 'false':
                    continue
            fg_color = self._extract_color_from_prop(props.get('fontColor', {}), theme_colors)
            if not fg_color:
                fg_color = self._extract_color_from_prop(props.get('color', {}), theme_colors)
            if fg_color:
                # Check for subtitle-specific background
                sub_bg = self._extract_color_from_prop(props.get('background', {}), theme_colors)
                if not sub_bg:
                    sub_bg = self._extract_color_from_prop(props.get('backgroundColor', {}), theme_colors)
                bg_color = sub_bg if sub_bg else fallback_bg_color
                bg_source = "Subtitle background" if sub_bg else fallback_bg_source
                result = self._calculate_contrast_result(
                    visual, fg_color, bg_color, "Subtitle text", bg_source, theme_colors
                )
                if result:
                    results.append(result)

        # 2. Check title text (in visualContainerObjects)
        title_config = container_objects.get('title', [])
        for obj in title_config:
            props = obj.get('properties', {})
            # Skip if title is disabled (show=false)
            show_prop = props.get('show', {})
            if show_prop:
                expr = show_prop.get('expr', {})
                literal = expr.get('Literal', {})
                if literal.get('Value', 'true').lower() == 'false':
                    continue
            fg_color = self._extract_color_from_prop(props.get('fontColor', {}), theme_colors)
            if not fg_color:
                fg_color = self._extract_color_from_prop(props.get('color', {}), theme_colors)
            if fg_color:
                # Check for title-specific background
                title_bg = self._extract_color_from_prop(props.get('background', {}), theme_colors)
                if not title_bg:
                    title_bg = self._extract_color_from_prop(props.get('backgroundColor', {}), theme_colors)
                bg_color = title_bg if title_bg else fallback_bg_color
                bg_source = "Title background" if title_bg else fallback_bg_source
                result = self._calculate_contrast_result(
                    visual, fg_color, bg_color, "Title text", bg_source, theme_colors
                )
                if result:
                    results.append(result)

        # 3. Check data labels (in objects)
        labels_config = objects.get('labels', [])
        for obj in labels_config:
            props = obj.get('properties', {})
            fg_color = self._extract_color_from_prop(props.get('color', {}), theme_colors)
            if not fg_color:
                fg_color = self._extract_color_from_prop(props.get('fontColor', {}), theme_colors)
            if fg_color:
                result = self._calculate_contrast_result(
                    visual, fg_color, fallback_bg_color, "Data labels", fallback_bg_source, theme_colors
                )
                if result:
                    results.append(result)

        # 4. Check category labels (in objects)
        cat_labels = objects.get('categoryLabels', [])
        for obj in cat_labels:
            props = obj.get('properties', {})
            fg_color = self._extract_color_from_prop(props.get('color', {}), theme_colors)
            if fg_color:
                result = self._calculate_contrast_result(
                    visual, fg_color, fallback_bg_color, "Category labels", fallback_bg_source, theme_colors
                )
                if result:
                    results.append(result)

        # 5. Check data point fill colors (in objects - for chart data series)
        data_point = objects.get('dataPoint', [])
        for obj in data_point:
            props = obj.get('properties', {})
            fg_color = self._extract_color_from_prop(props.get('fill', {}), theme_colors)
            if not fg_color:
                fg_color = self._extract_color_from_prop(props.get('color', {}), theme_colors)
            if fg_color:
                result = self._calculate_contrast_result(
                    visual, fg_color, fallback_bg_color, "Data point colors", fallback_bg_source, theme_colors
                )
                if result:
                    results.append(result)

        # 6. Check legend text color (in objects)
        legend = objects.get('legend', [])
        for obj in legend:
            props = obj.get('properties', {})
            fg_color = self._extract_color_from_prop(props.get('labelColor', {}), theme_colors)
            if not fg_color:
                fg_color = self._extract_color_from_prop(props.get('fontColor', {}), theme_colors)
            if fg_color:
                result = self._calculate_contrast_result(
                    visual, fg_color, fallback_bg_color, "Legend text", fallback_bg_source, theme_colors
                )
                if result:
                    results.append(result)

        # 7. Check axis label colors (in objects) - check all axis types
        axis_display = {
            'xAxis': 'X-axis', 'yAxis': 'Y-axis',
            'categoryAxis': 'Category axis', 'valueAxis': 'Value axis',
            'x': 'X-axis', 'y': 'Y-axis'
        }
        for axis_name in ['xAxis', 'yAxis', 'categoryAxis', 'valueAxis', 'x', 'y']:
            axis_config = objects.get(axis_name, [])
            for obj in axis_config:
                props = obj.get('properties', {})
                fg_color = self._extract_color_from_prop(props.get('labelColor', {}), theme_colors)
                if not fg_color:
                    fg_color = self._extract_color_from_prop(props.get('fontColor', {}), theme_colors)
                if not fg_color:
                    fg_color = self._extract_color_from_prop(props.get('color', {}), theme_colors)
                if fg_color:
                    fg_element = f"{axis_display.get(axis_name, 'Axis')} labels"
                    result = self._calculate_contrast_result(
                        visual, fg_color, fallback_bg_color, fg_element, fallback_bg_source, theme_colors
                    )
                    if result:
                        results.append(result)

        # 8. Check table column/row headers (in objects)
        header_display = {
            'columnHeaders': 'Column headers', 'rowHeaders': 'Row headers',
            'values': 'Table values', 'header': 'Table header'
        }
        for header_name in ['columnHeaders', 'rowHeaders', 'values', 'header']:
            header_config = objects.get(header_name, [])
            for obj in header_config:
                props = obj.get('properties', {})
                fg_color = self._extract_color_from_prop(props.get('fontColor', {}), theme_colors)
                if not fg_color:
                    fg_color = self._extract_color_from_prop(props.get('color', {}), theme_colors)
                if fg_color:
                    fg_element = header_display.get(header_name, 'Table text')
                    result = self._calculate_contrast_result(
                        visual, fg_color, fallback_bg_color, fg_element, fallback_bg_source, theme_colors
                    )
                    if result:
                        results.append(result)

        # 9. Check slicer items text color (in objects.items for slicers)
        items_config = objects.get('items', [])
        for obj in items_config:
            props = obj.get('properties', {})
            # Slicer items can have fontColor or color property
            fg_color = self._extract_color_from_prop(props.get('fontColor', {}), theme_colors)
            if not fg_color:
                fg_color = self._extract_color_from_prop(props.get('color', {}), theme_colors)
            # If no explicit font color, try to use theme foreground as default
            if not fg_color:
                fg_color = theme_colors.get('foreground')
            if fg_color:
                # Use slicer items background if available, otherwise fallback
                items_bg = self._extract_color_from_prop(props.get('background', {}), theme_colors)
                bg_color = items_bg if items_bg else fallback_bg_color
                bg_source = "Slicer items background" if items_bg else fallback_bg_source
                result = self._calculate_contrast_result(
                    visual, fg_color, bg_color, "Slicer values text", bg_source, theme_colors
                )
                if result:
                    results.append(result)

        return results

    def _find_contrast_issues(self, contrasts: List[ColorContrastResult],
                               config: AccessibilityCheckConfig = None) -> List[AccessibilityIssue]:
        """Find contrast issues from analyzed colors based on config settings.

        Args:
            contrasts: List of color contrast results
            config: Optional config for contrast level settings

        Returns:
            List of accessibility issues for contrast failures
        """
        if config is None:
            config = get_config()

        issues = []

        # Get thresholds based on config
        thresholds = CONTRAST_THRESHOLDS.get(config.contrast_level, CONTRAST_THRESHOLDS["AA"])
        normal_threshold = thresholds["normal"]
        large_threshold = thresholds["large"]

        # Determine WCAG reference based on level
        if config.contrast_level == "AAA":
            wcag_ref = "WCAG 2.1 - 1.4.6 Contrast (Enhanced)"
            level_name = "AAA"
        elif config.contrast_level == "AA_large":
            wcag_ref = "WCAG 2.1 - 1.4.3 Contrast (Large Text)"
            level_name = "AA Large"
        else:
            wcag_ref = "WCAG 2.1 - 1.4.3 Contrast (Minimum)"
            level_name = "AA"

        for contrast in contrasts:
            # Format: element_type|fg_color|bg_color for UI parsing with color swatches
            element_info = contrast.element_type if contrast.element_type else "Text vs Background"
            current_val = f"{element_info}|{contrast.foreground_color}|{contrast.background_color}"

            # Check against configured threshold
            passes_configured = contrast.contrast_ratio >= normal_threshold

            if not passes_configured:
                # Fails the configured level - this is an error
                issues.append(AccessibilityIssue(
                    check_type=AccessibilityCheckType.COLOR_CONTRAST,
                    severity=AccessibilitySeverity.ERROR,
                    page_name=contrast.page_name,
                    visual_name=contrast.visual_name,
                    visual_type=contrast.visual_type,
                    issue_description=f"Color contrast ratio ({contrast.contrast_ratio:.2f}:1) fails WCAG {level_name} requirement ({normal_threshold}:1)",
                    recommendation="Increase contrast by using darker text or lighter background",
                    current_value=current_val,
                    wcag_reference=wcag_ref
                ))
            elif config.contrast_level != "AAA" and config.flag_aaa_failures:
                # Determine which AAA threshold to use based on current level
                # AA_large mode (3:1) -> AAA large threshold is 4.5:1
                # AA mode (4.5:1) -> AAA normal threshold is 7:1
                use_large_aaa = config.contrast_level == "AA_large"
                passes_aaa = contrast.passes_aaa_large if use_large_aaa else contrast.passes_aaa_normal
                aaa_threshold = 4.5 if use_large_aaa else 7.0

                if not passes_aaa:
                    # Passes configured level but fails AAA - flag as review
                    issues.append(AccessibilityIssue(
                        check_type=AccessibilityCheckType.COLOR_CONTRAST,
                        severity=AccessibilitySeverity.INFO,
                        page_name=contrast.page_name,
                        visual_name=contrast.visual_name,
                        visual_type=contrast.visual_type,
                        issue_description=f"Color contrast ratio ({contrast.contrast_ratio:.2f}:1) passes {level_name} but fails AAA ({aaa_threshold}:1)",
                        recommendation="Consider increasing contrast for enhanced accessibility",
                        current_value=current_val,
                        wcag_reference="WCAG 2.1 - 1.4.6 Contrast (Enhanced)"
                    ))

            # Check for AA failures when using AA_large (3:1) threshold and flag_aa_failures is enabled
            if config.contrast_level == "AA_large" and config.flag_aa_failures:
                # AA requires 4.5:1 for normal text
                if not contrast.passes_aa_normal:
                    # Passes AA_large (3:1) but fails AA (4.5:1) - flag as review
                    issues.append(AccessibilityIssue(
                        check_type=AccessibilityCheckType.COLOR_CONTRAST,
                        severity=AccessibilitySeverity.INFO,
                        page_name=contrast.page_name,
                        visual_name=contrast.visual_name,
                        visual_type=contrast.visual_type,
                        issue_description=f"Color contrast ratio ({contrast.contrast_ratio:.2f}:1) passes Large Text Only (3:1) but fails AA (4.5:1)",
                        recommendation="Consider increasing contrast to meet standard AA requirements",
                        current_value=current_val,
                        wcag_reference="WCAG 2.1 - 1.4.3 Contrast (Minimum)"
                    ))

            # Check for transparency that requires manual review
            if contrast.requires_review:
                issues.append(AccessibilityIssue(
                    check_type=AccessibilityCheckType.COLOR_CONTRAST,
                    severity=AccessibilitySeverity.INFO,
                    page_name=contrast.page_name,
                    visual_name=contrast.visual_name,
                    visual_type=contrast.visual_type,
                    issue_description=f"Color contrast includes {contrast.transparency_layer_count} transparent layers - calculated ratio ({contrast.contrast_ratio:.2f}:1) may not reflect true visible contrast",
                    recommendation="Manually verify contrast with overlapping transparent elements",
                    current_value=current_val,
                    wcag_reference="WCAG 2.1 - 1.4.3 Contrast (Minimum)"
                ))

        return issues

    def _hex_to_rgb(self, hex_color: str) -> Tuple[int, int, int]:
        """Convert hex color to RGB tuple"""
        hex_color = hex_color.lstrip('#')
        if len(hex_color) == 3:
            hex_color = ''.join([c*2 for c in hex_color])
        return tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))

    def _calculate_luminance(self, r: int, g: int, b: int) -> float:
        """Calculate relative luminance using WCAG formula"""
        def adjust(c):
            c = c / 255.0
            return c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4

        return 0.2126 * adjust(r) + 0.7152 * adjust(g) + 0.0722 * adjust(b)

    def _calculate_contrast_ratio(self, l1: float, l2: float) -> float:
        """Calculate contrast ratio between two luminance values"""
        lighter = max(l1, l2)
        darker = min(l1, l2)
        return (lighter + 0.05) / (darker + 0.05)

    # Maximum transparent layers before marking as "requires review"
    MAX_TRANSPARENT_LAYERS = 3

    def _hex_to_rgba(self, hex_color: str) -> Tuple[int, int, int, float]:
        """Parse hex color to RGBA, alpha as 0.0-1.0.

        Supports formats: #RGB, #RRGGBB, #RRGGBBAA
        """
        hex_color = hex_color.lstrip('#')
        if len(hex_color) == 3:
            hex_color = ''.join([c*2 for c in hex_color])

        r = int(hex_color[0:2], 16)
        g = int(hex_color[2:4], 16)
        b = int(hex_color[4:6], 16)

        # Parse alpha if present (8-digit hex)
        if len(hex_color) >= 8:
            alpha = int(hex_color[6:8], 16) / 255.0
        else:
            alpha = 1.0

        return (r, g, b, alpha)

    def _has_transparency(self, hex_color: str) -> bool:
        """Check if hex color has alpha < 1.0 (not fully opaque)"""
        hex_color = hex_color.lstrip('#')
        if len(hex_color) >= 8:
            alpha = int(hex_color[6:8], 16)
            return alpha < 255
        return False

    def _blend_colors(self, fg_hex: str, bg_hex: str) -> Tuple[str, int]:
        """Blend foreground over background using alpha compositing.

        Args:
            fg_hex: Foreground color in hex (may have alpha)
            bg_hex: Background color in hex (may have alpha)

        Returns:
            Tuple of (blended_hex_color, layer_count)
            layer_count indicates how many transparent layers were blended
        """
        layer_count = 0

        fg_r, fg_g, fg_b, fg_a = self._hex_to_rgba(fg_hex)
        bg_r, bg_g, bg_b, bg_a = self._hex_to_rgba(bg_hex)

        if fg_a < 1.0:
            layer_count += 1
        if bg_a < 1.0:
            layer_count += 1

        # If foreground is fully opaque, no blending needed
        if fg_a >= 1.0:
            return fg_hex[:7] if len(fg_hex) > 7 else fg_hex, layer_count

        # Alpha blending formula: result = fg * alpha + bg * (1 - alpha)
        result_r = int(fg_r * fg_a + bg_r * (1 - fg_a))
        result_g = int(fg_g * fg_a + bg_g * (1 - fg_a))
        result_b = int(fg_b * fg_a + bg_b * (1 - fg_a))

        # Clamp to valid range
        result_r = max(0, min(255, result_r))
        result_g = max(0, min(255, result_g))
        result_b = max(0, min(255, result_b))

        return f"#{result_r:02X}{result_g:02X}{result_b:02X}", layer_count

    def _rgb_to_hex(self, r: int, g: int, b: int) -> str:
        """Convert RGB values to hex color string"""
        return f"#{r:02X}{g:02X}{b:02X}"

    def _blend_with_alpha(self, fg_hex: str, bg_hex: str, alpha: float) -> str:
        """Blend foreground over background with explicit alpha (0.0-1.0).

        Used for transparency percentage blending where alpha is provided
        separately rather than embedded in the hex color.

        Args:
            fg_hex: Foreground hex color
            bg_hex: Background hex color
            alpha: Opacity of foreground (0.0 = transparent, 1.0 = opaque)

        Returns:
            Blended hex color string
        """
        fg_r, fg_g, fg_b, _ = self._hex_to_rgba(fg_hex)
        bg_r, bg_g, bg_b, _ = self._hex_to_rgba(bg_hex)

        r = int(fg_r * alpha + bg_r * (1 - alpha))
        g = int(fg_g * alpha + bg_g * (1 - alpha))
        b = int(fg_b * alpha + bg_b * (1 - alpha))

        # Clamp to valid range
        r = max(0, min(255, r))
        g = max(0, min(255, g))
        b = max(0, min(255, b))

        return f"#{r:02X}{g:02X}{b:02X}"

    def _compute_effective_page_background(self, page_settings: Dict[str, Any],
                                           theme_colors: Dict[str, Any]) -> str:
        """Compute effective page background by compositing wallpaper + canvas.

        Power BI layer stack (bottom to top):
        1. Wallpaper (outspace)
        2. Page canvas background (with optional transparency)

        Args:
            page_settings: Dict with 'background_color', 'background_transparency', 'wallpaper_color'
            theme_colors: Theme colors for fallbacks

        Returns:
            Effective page background hex color
        """
        # Get wallpaper (outspace) color - defaults to theme background
        wallpaper = page_settings.get('wallpaper_color')
        if not wallpaper:
            wallpaper = theme_colors.get('background', '#FFFFFF')

        # Get canvas background color - defaults to theme background
        canvas_color = page_settings.get('background_color')
        if not canvas_color:
            canvas_color = theme_colors.get('background', '#FFFFFF')

        # Get canvas transparency (0 = opaque, 100 = fully transparent)
        canvas_transparency = page_settings.get('background_transparency', 0)

        # If canvas is fully opaque, just return the canvas color
        if canvas_transparency == 0:
            return canvas_color

        # Convert transparency % to alpha (0% transparency = 1.0 alpha, 100% = 0.0)
        alpha = 1.0 - (canvas_transparency / 100.0)

        # Blend canvas over wallpaper
        return self._blend_with_alpha(canvas_color, wallpaper, alpha)

    # =========================================================================
    # PAGE TITLE ANALYSIS
    # =========================================================================

    def _analyze_page_titles(self, pages: List[PageInfo]) -> List[AccessibilityIssue]:
        """Check that all pages have meaningful titles"""
        issues = []

        for page in pages:
            # Check for missing or generic title
            if not page.display_name:
                issues.append(AccessibilityIssue(
                    check_type=AccessibilityCheckType.PAGE_TITLE,
                    severity=AccessibilitySeverity.ERROR,
                    page_name=page.page_name,
                    issue_description="Page is missing a display name",
                    recommendation="Add a descriptive display name that helps users understand the page content",
                    wcag_reference="WCAG 2.1 - 2.4.2 Page Titled"
                ))
            elif self._is_generic_page_title(page.display_name):
                issues.append(AccessibilityIssue(
                    check_type=AccessibilityCheckType.PAGE_TITLE,
                    severity=AccessibilitySeverity.WARNING,
                    page_name=page.page_name,
                    issue_description="Page title appears to be generic",
                    recommendation="Replace with a descriptive title that describes the page content",
                    current_value=page.display_name,
                    wcag_reference="WCAG 2.1 - 2.4.2 Page Titled"
                ))

        return issues

    def _is_generic_page_title(self, title: str) -> bool:
        """Check if page title is generic"""
        for pattern in GENERIC_PAGE_TITLE_PATTERNS:
            if re.match(pattern, title.lower().strip()):
                return True
        return False

    # =========================================================================
    # VISUAL TITLE ANALYSIS
    # =========================================================================

    def _analyze_visual_titles(
        self,
        visuals: List[VisualInfo],
        bookmark_toggled_ids: Set[str] = None,
        hidden_group_ids: Set[str] = None
    ) -> List[AccessibilityIssue]:
        """Check that data visuals have title text

        Args:
            visuals: List of visual info objects to analyze
            bookmark_toggled_ids: Set of visual/group IDs that can be toggled visible by bookmarks
            hidden_group_ids: Set of group IDs that are hidden
        """
        issues = []
        bookmark_toggled_ids = bookmark_toggled_ids or set()
        hidden_group_ids = hidden_group_ids or set()

        for visual in visuals:
            # Skip visual groups (they're containers, not content)
            if visual.is_group:
                continue

            # Check if visual is effectively hidden (directly or via parent group)
            is_effectively_hidden = visual.is_hidden or (
                visual.parent_group_name and visual.parent_group_name in hidden_group_ids
            )

            if is_effectively_hidden:
                # Skip ONLY if no bookmark can toggle this visual/group visible
                can_be_toggled = (
                    visual.visual_id in bookmark_toggled_ids or
                    (visual.parent_group_name and visual.parent_group_name in bookmark_toggled_ids)
                )
                if not can_be_toggled:
                    continue  # Skip truly hidden visuals

            if visual.is_data_visual and not visual.has_title:
                issues.append(AccessibilityIssue(
                    check_type=AccessibilityCheckType.VISUAL_TITLE,
                    severity=AccessibilitySeverity.WARNING,
                    page_name=visual.page_name,
                    visual_name=visual.visual_name,
                    visual_type=visual.visual_type,
                    issue_description="Data visual does not have a title displayed",
                    recommendation="Add a descriptive title to help users understand the visual's purpose",
                    wcag_reference="WCAG 2.1 - 2.4.6 Headings and Labels"
                ))

        return issues

    # =========================================================================
    # DATA LABELS ANALYSIS
    # =========================================================================

    def _analyze_data_labels(
        self,
        visuals: List[VisualInfo],
        bookmark_toggled_ids: Set[str] = None,
        hidden_group_ids: Set[str] = None
    ) -> List[AccessibilityIssue]:
        """Check if data visualizations have data labels enabled

        Args:
            visuals: List of visual info objects to analyze
            bookmark_toggled_ids: Set of visual/group IDs that can be toggled visible by bookmarks
            hidden_group_ids: Set of group IDs that are hidden
        """
        issues = []
        bookmark_toggled_ids = bookmark_toggled_ids or set()
        hidden_group_ids = hidden_group_ids or set()

        # Charts that benefit most from data labels
        label_important_types = {
            'pieChart', 'donutChart', 'treemap',
            'clusteredBarChart', 'clusteredColumnChart',
            'stackedBarChart', 'stackedColumnChart'
        }

        for visual in visuals:
            # Skip visual groups (they're containers, not content)
            if visual.is_group:
                continue

            # Check if visual is effectively hidden (directly or via parent group)
            is_effectively_hidden = visual.is_hidden or (
                visual.parent_group_name and visual.parent_group_name in hidden_group_ids
            )

            if is_effectively_hidden:
                # Skip ONLY if no bookmark can toggle this visual/group visible
                can_be_toggled = (
                    visual.visual_id in bookmark_toggled_ids or
                    (visual.parent_group_name and visual.parent_group_name in bookmark_toggled_ids)
                )
                if not can_be_toggled:
                    continue  # Skip truly hidden visuals

            if visual.has_data_labels is False:  # Explicitly False, not None
                visual_type_lower = visual.visual_type.lower()

                # Check if it's a chart type where labels are important
                is_important = any(t.lower() in visual_type_lower for t in label_important_types)

                severity = AccessibilitySeverity.WARNING if is_important else AccessibilitySeverity.INFO

                issues.append(AccessibilityIssue(
                    check_type=AccessibilityCheckType.DATA_LABELS,
                    severity=severity,
                    page_name=visual.page_name,
                    visual_name=visual.visual_name,
                    visual_type=visual.visual_type,
                    issue_description="Chart does not have data labels enabled",
                    recommendation="Consider enabling data labels to help users who cannot distinguish colors",
                    wcag_reference="WCAG 2.1 - 1.4.1 Use of Color"
                ))

        return issues

    # =========================================================================
    # BOOKMARK NAME ANALYSIS
    # =========================================================================

    def _analyze_bookmark_names(self, bookmarks: List[BookmarkInfo]) -> List[AccessibilityIssue]:
        """Check that bookmarks have descriptive names"""
        issues = []

        for bookmark in bookmarks:
            if bookmark.is_generic_name:
                issues.append(AccessibilityIssue(
                    check_type=AccessibilityCheckType.BOOKMARK_NAME,
                    severity=AccessibilitySeverity.WARNING,
                    page_name="Report",
                    issue_description=f"Bookmark has a generic name: {bookmark.display_name}",
                    recommendation="Use a descriptive name that helps users understand the bookmark's purpose",
                    current_value=bookmark.display_name,
                    wcag_reference="WCAG 2.1 - 2.4.6 Headings and Labels"
                ))

        return issues

    def _is_generic_bookmark_name(self, name: str) -> bool:
        """Check if bookmark name is generic"""
        for pattern in GENERIC_BOOKMARK_PATTERNS:
            if re.match(pattern, name.lower().strip()):
                return True
        return False

    # =========================================================================
    # HIDDEN PAGES WARNING
    # =========================================================================

    def _check_hidden_pages(self, pages: List[PageInfo]) -> List[AccessibilityIssue]:
        """Info-level notification about hidden pages"""
        issues = []

        hidden_pages = [p for p in pages if p.is_hidden]

        for page in hidden_pages:
            issues.append(AccessibilityIssue(
                check_type=AccessibilityCheckType.HIDDEN_PAGE,
                severity=AccessibilitySeverity.INFO,
                page_name=page.display_name,
                issue_description="Page is hidden from report navigation",
                recommendation="Ensure hidden pages don't contain content that users need to access",
                wcag_reference="WCAG 2.1 - 2.4.1 Bypass Blocks"
            ))

        return issues


__all__ = ['AccessibilityAnalyzer']
