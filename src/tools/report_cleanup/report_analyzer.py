"""
Report Cleanup Analyzer - Analysis engine for unused themes and custom visuals
Built by Reid Havens of Analytic Endeavors

This module analyzes Power BI reports to identify cleanup opportunities:
- Unused themes in BaseThemes and RegisteredResources  
- Unused custom visuals from AppSource (in publicCustomVisuals)
- Hidden custom visuals in CustomVisuals directory
"""

import json
from pathlib import Path
from typing import Dict, List, Any, Optional, Callable, Set, Tuple
import uuid

# Import shared data classes
from tools.report_cleanup.shared_types import CleanupOpportunity


class ReportAnalyzer:
    """
    Analyzes Power BI reports for cleanup opportunities related to themes and custom visuals
    """
    
    def __init__(self, logger_callback: Optional[Callable[[str], None]] = None):
        self.log_callback = logger_callback or self._default_log
    
    def _default_log(self, message: str) -> None:
        """Default logging function"""
        print(message)
    
    def analyze_pbip_report(self, pbip_path: str) -> Tuple[Dict[str, Any], List[CleanupOpportunity]]:
        """
        Analyze a PBIP report for cleanup opportunities
        
        Args:
            pbip_path: Path to the .pbip file
            
        Returns:
            Tuple of (analysis_data, cleanup_opportunities)
        """
        self.log_callback("🔍 Starting report cleanup analysis...")
        
        pbip_file = Path(pbip_path)
        if not pbip_file.exists():
            raise FileNotFoundError(f"PBIP file not found: {pbip_path}")
        
        # Get report directory
        report_dir = pbip_file.parent / f"{pbip_file.stem}.Report"
        if not report_dir.exists():
            raise FileNotFoundError(f"Report directory not found: {report_dir}")
        
        self.log_callback(f"📁 Analyzing report: {report_dir}")
        
        # Load report.json for theme and resource analysis
        report_json_path = report_dir / "definition" / "report.json"
        if not report_json_path.exists():
            raise FileNotFoundError(f"report.json not found: {report_json_path}")
        
        with open(report_json_path, 'r', encoding='utf-8') as f:
            report_data = json.load(f)
        
        # First: Comprehensively scan pages to find ALL visual usage
        used_visual_info = self._scan_pages_for_all_visuals(report_dir)
        
        # Analyze themes, custom visuals, bookmarks, and visual filters
        analysis_data = {
            'report_path': str(report_dir),
            'used_visual_info': used_visual_info,
            'themes': self._analyze_themes(report_data, report_dir),
            'custom_visuals': self._analyze_custom_visuals(report_data, report_dir, used_visual_info),
            'bookmarks': self._analyze_bookmarks(report_data, report_dir),
            'visual_filters': self._analyze_visual_filters(report_dir),
        }
        
        # Find cleanup opportunities
        opportunities = []
        theme_opportunities = self._find_theme_opportunities(analysis_data['themes'])
        visual_opportunities = self._find_custom_visual_opportunities(analysis_data['custom_visuals'])
        bookmark_opportunities = self._find_bookmark_opportunities(analysis_data['bookmarks'])
        filter_opportunities = self._find_visual_filter_opportunities(analysis_data['visual_filters'])
        
        opportunities.extend(theme_opportunities)
        opportunities.extend(visual_opportunities)
        opportunities.extend(bookmark_opportunities)
        opportunities.extend(filter_opportunities)
        
        # Create detailed summary
        theme_count = len(theme_opportunities)
        visual_count = len(visual_opportunities)
        bookmark_count = len(bookmark_opportunities)
        filter_count = len(filter_opportunities)
        total_count = len(opportunities)
        
        if total_count > 0:
            summary_parts = []
            if theme_count > 0:
                summary_parts.append(f"{theme_count} theme{'s' if theme_count != 1 else ''}")
            if visual_count > 0:
                summary_parts.append(f"{visual_count} custom visual{'s' if visual_count != 1 else ''}")
            if bookmark_count > 0:
                summary_parts.append(f"{bookmark_count} bookmark{'s' if bookmark_count != 1 else ''}")
            if filter_count > 0:
                summary_parts.append(f"{filter_count} visual filter group{'s' if filter_count != 1 else ''}")
            
            detailed_summary = f"{total_count} cleanup opportunities ({', '.join(summary_parts)})"
        else:
            detailed_summary = "0 cleanup opportunities"
        
        self.log_callback(f"✅ Analysis complete. Found {detailed_summary}.")
        
        return analysis_data, opportunities
    
    def _scan_pages_for_all_visuals(self, report_dir: Path) -> Dict[str, Set[str]]:
        """Comprehensive scan to find ALL visual types used in pages"""
        self.log_callback("🔍 Comprehensively scanning pages for visual usage...")
        
        visual_usage = {
            'visual_types': set(),      # All visualType values found
            'custom_guids': set(),      # All customVisualGuid values found
            'visual_names': set(),      # All visual names/identifiers found
        }
        
        pages_dir = report_dir / "definition" / "pages"
        if not pages_dir.exists():
            self.log_callback("  ⚠️ No pages directory found")
            return visual_usage
        
        page_dirs = [d for d in pages_dir.iterdir() if d.is_dir()]
        self.log_callback(f"  📄 Found {len(page_dirs)} pages to scan")
        
        total_visuals = 0
        custom_visuals_found = {}  # Dictionary to track custom visuals by page
        
        for page_dir in page_dirs:
            page_json = page_dir / "page.json"
            if page_json.exists():
                try:
                    with open(page_json, 'r', encoding='utf-8') as f:
                        page_data = json.load(f)
                    
                    page_name = page_data.get('displayName', page_dir.name)
                    
                    # Check for visuals directory
                    visuals_dir = page_dir / "visuals"
                    if visuals_dir.exists():
                        visual_dirs = [d for d in visuals_dir.iterdir() if d.is_dir()]
                        if visual_dirs:
                            total_visuals += len(visual_dirs)
                            
                            for visual_dir in visual_dirs:
                                visual_json = visual_dir / "visual.json"
                                if visual_json.exists():
                                    try:
                                        with open(visual_json, 'r', encoding='utf-8') as f:
                                            visual_data = json.load(f)
                                        
                                        # Get visual configuration
                                        visual_config = visual_data.get('visual', {})
                                        
                                        if visual_config:
                                            # Get visual type
                                            visual_type = visual_config.get('visualType', '')
                                            if visual_type:
                                                visual_usage['visual_types'].add(visual_type)
                                                
                                                # Only track custom visuals
                                                if self._is_custom_visual_type(visual_type):
                                                    if page_name not in custom_visuals_found:
                                                        custom_visuals_found[page_name] = {}
                                                    if visual_type not in custom_visuals_found[page_name]:
                                                        custom_visuals_found[page_name][visual_type] = 0
                                                    custom_visuals_found[page_name][visual_type] += 1
                                            
                                            # Get custom visual GUID
                                            custom_guid = visual_config.get('customVisualGuid', '')
                                            if custom_guid:
                                                visual_usage['custom_guids'].add(custom_guid)
                                            
                                            # Get visual name/identifier from the visual data
                                            visual_name = visual_data.get('name', '')
                                            if visual_name:
                                                visual_usage['visual_names'].add(visual_name)
                                            
                                            # Also check for other identifying fields
                                            for field in ['id', 'visualId', 'guid']:
                                                value = visual_config.get(field, '') or visual_data.get(field, '')
                                                if value:
                                                    visual_usage['visual_names'].add(value)
                                    
                                    except Exception as e:
                                        pass  # Silently skip problematic visuals
                        
                except Exception as e:
                    pass  # Silently skip problematic pages
        
        # Combine all found visual identifiers
        all_used_visuals = visual_usage['visual_types'] | visual_usage['custom_guids'] | visual_usage['visual_names']
        
        # Summary logging
        self.log_callback(f"  🎯 Summary: {total_visuals} total visuals scanned across {len(page_dirs)} pages")
        
        if custom_visuals_found:
            # Count total custom visual instances
            total_custom_instances = sum(sum(visuals.values()) for visuals in custom_visuals_found.values())
            self.log_callback(f"  🔮 Custom visuals found in use ({total_custom_instances} total instances):")
            for page_name, visuals in custom_visuals_found.items():
                for visual_type, count in visuals.items():
                    if count == 1:
                        self.log_callback(f"    ✅ {visual_type} (on page '{page_name}')")
                    else:
                        self.log_callback(f"    ✅ {visual_type} (on page '{page_name}' - {count} instances)")
        else:
            self.log_callback(f"  🔮 No custom visuals found in use on any pages")
        
        return visual_usage
    
    def _is_custom_visual_type(self, visual_type: str) -> bool:
        """Determine if a visual type is a custom visual (not built-in)"""
        if not visual_type:
            return False
        
        # Custom visual patterns
        custom_patterns = [
            '_CV_',  # Common custom visual pattern
            'PBI_CV_',  # Power BI custom visual pattern  
            'ChordChart',  # Specific visual types
            'sparklineChart',
            'waterCup',
            'searchVisual', 
            'calendarVisual',
            'hierarchySlicer',
            'timelineStoryteller',
            'wordCloud'
        ]
        
        # Check if it matches custom patterns
        for pattern in custom_patterns:
            if pattern.lower() in visual_type.lower():
                return True
        
        # If it's not in the built-in visual types, it's likely custom
        builtin_types = self._get_builtin_visual_types()
        return visual_type not in builtin_types
    
    def _get_builtin_visual_types(self) -> Set[str]:
        """Get set of built-in Power BI visual types"""
        return {
            # Core built-in visuals
            'clusteredColumnChart', 'table', 'slicer', 'card', 'lineChart', 'pieChart', 'map',
            'clusteredBarChart', 'scatterChart', 'gauge', 'multiRowCard', 'kpi', 'donutChart', 
            'matrix', 'waterfallChart', 'funnelChart', 'treemap', 'ribbonChart', 'histogram', 
            'filledMap', 'stackedColumnChart', 'stackedBarChart', 'lineStackedColumnChart',
            'lineClusteredColumnChart', 'hundredPercentStackedBarChart', 'hundredPercentStackedColumnChart',
            'shape', 'textbox', 'image', 'actionButton', 'decompositionTreeVisual', 'smartNarrativeVisual',
            'keyInfluencersVisual', 'qnaVisual', 'paginator',
            # Additional built-in visuals that were showing up as "custom"
            'advancedSlicerVisual', 'barChart', 'pivotTable', 'basicShape', 'areaChart',
            'columnChart', 'pieDonutChart', 'lineAreaChart', 'comboChart', 'multiRowCard',
            'tableEx', 'matrixEx', 'card', 'multiCardVisual'
        }
    
    def _analyze_themes(self, report_data: Dict, report_dir: Path) -> Dict[str, Any]:
        """Analyze themes in the report - only one theme can be active at a time"""
        self.log_callback("🎨 Analyzing themes...")
        
        theme_analysis = {
            'active_theme': None,  # Only one theme can be active
            'available_themes': {},
            'theme_files': {},
        }
        
        # Find THE SINGLE active theme from themeCollection
        theme_collection = report_data.get('themeCollection', {})
        if theme_collection:
            # Check for custom theme first (takes precedence)
            custom_theme = theme_collection.get('customTheme', {})
            if custom_theme and custom_theme.get('name'):
                theme_name = custom_theme.get('name', '')
                theme_type = custom_theme.get('type', 'SharedResources')
                theme_analysis['active_theme'] = (theme_name, theme_type)
                self.log_callback(f"  ✅ Active custom theme: {theme_name} ({theme_type})")
            else:
                # Fall back to base theme if no custom theme
                base_theme = theme_collection.get('baseTheme', {})
                if base_theme and base_theme.get('name'):
                    theme_name = base_theme.get('name', '')
                    theme_type = base_theme.get('type', 'SharedResources')
                    theme_analysis['active_theme'] = (theme_name, theme_type)
                    self.log_callback(f"  ✅ Active base theme: {theme_name} ({theme_type})")
        
        # If no active theme found, assume default
        if not theme_analysis['active_theme']:
            self.log_callback("  💬 No explicit active theme found - using default")
            theme_analysis['active_theme'] = ('Default', 'Built-in')
        
        # Scan for available theme files in SharedResources/BaseThemes
        base_themes_dir = report_dir / "StaticResources" / "SharedResources" / "BaseThemes"
        if base_themes_dir.exists():
            for theme_file in base_themes_dir.glob("*.json"):
                theme_name = theme_file.stem
                theme_analysis['available_themes'][theme_name] = {
                    'type': 'BaseTheme',
                    'location': 'SharedResources',
                    'path': theme_file,
                    'size': theme_file.stat().st_size if theme_file.exists() else 0
                }
                theme_analysis['theme_files'][theme_name] = theme_file
        
        # Scan for custom themes in RegisteredResources
        reg_resources_dir = report_dir / "StaticResources" / "RegisteredResources"
        if reg_resources_dir.exists():
            for theme_file in reg_resources_dir.glob("*.json"):
                # Check if this is actually a theme file
                try:
                    with open(theme_file, 'r', encoding='utf-8') as f:
                        data = json.load(f)
                    # Look for theme-like structure
                    if any(key in data for key in ['name', 'dataColors', 'background', 'foreground']):
                        theme_name = theme_file.name  # Keep full filename for registered resources
                        theme_analysis['available_themes'][theme_name] = {
                            'type': 'CustomTheme',
                            'location': 'RegisteredResources',
                            'path': theme_file,
                            'size': theme_file.stat().st_size
                        }
                        theme_analysis['theme_files'][theme_name] = theme_file
                except:
                    pass  # Skip files that can't be parsed as JSON
        
        active_theme_name = theme_analysis['active_theme'][0] if theme_analysis['active_theme'] else 'None'
        self.log_callback(f"  🎨 Current active theme: {active_theme_name}")
        self.log_callback(f"  📊 Found {len(theme_analysis['available_themes'])} available theme files")
        
        return theme_analysis
    
    def _analyze_custom_visuals(self, report_data: Dict, report_dir: Path, used_visual_info: Dict) -> Dict[str, Any]:
        """Analyze custom visuals in the report with enhanced detection"""
        self.log_callback("🔮 Analyzing custom visuals...")
        
        # Combine all used visual identifiers
        all_used_visuals = (used_visual_info['visual_types'] | 
                           used_visual_info['custom_guids'] | 
                           used_visual_info['visual_names'])
        
        visual_analysis = {
            'used_visuals': all_used_visuals,
            'build_pane_visuals': {},    # Visuals that appear in build pane
            'hidden_visuals': {},        # Visuals in CustomVisuals folder but not in build pane
            'appsource_visuals': {},     # Visuals from AppSource (publicCustomVisuals)
        }
        
        # Step 1: Analyze AppSource visuals (publicCustomVisuals)
        public_visuals = report_data.get('publicCustomVisuals', [])
        self.log_callback(f"  📱 Found {len(public_visuals)} AppSource/public custom visuals")
        
        for visual_id in public_visuals:
            # Check if this visual is actually used
            is_used = self._is_visual_used(visual_id, all_used_visuals)
            
            visual_analysis['appsource_visuals'][visual_id] = {
                'visual_id': visual_id,
                'display_name': self._get_visual_display_name(visual_id),
                'location': 'AppSource/publicCustomVisuals',
                'used': is_used,
                'size': 0  # AppSource visuals don't take up local space
            }
            
            status = "USED" if is_used else "UNUSED"
            self.log_callback(f"    📱 {status}: {visual_id}")
        
        # Step 2: Analyze resourcePackages CustomVisual entries (build pane visuals)
        resource_packages = report_data.get('resourcePackages', [])
        
        for package in resource_packages:
            if package.get('type') == 'CustomVisual':
                visual_id = package.get('name', '')
                if visual_id:
                    # Check if this visual is actually used
                    is_used = self._is_visual_used(visual_id, all_used_visuals)
                    
                    # Get metadata info
                    items = package.get('items', [])
                    metadata_file = None
                    for item in items:
                        if item.get('type') == 'CustomVisualMetadata':
                            metadata_file = item.get('path', '')
                            break
                    
                    visual_analysis['build_pane_visuals'][visual_id] = {
                        'visual_id': visual_id,
                        'display_name': self._get_visual_display_name(visual_id),
                        'metadata_file': metadata_file,
                        'location': 'resourcePackages',
                        'used': is_used,
                        'size': self._estimate_visual_size(report_dir, visual_id)
                    }
                    
                    status = "USED" if is_used else "UNUSED"
                    self.log_callback(f"    🔧 Build Pane {status}: {visual_id}")
        
        # Step 3: Analyze CustomVisuals directory for hidden/orphaned visuals
        custom_visuals_dir = report_dir / "CustomVisuals"
        if custom_visuals_dir.exists():
            self.log_callback(f"  📁 Scanning CustomVisuals directory...")
            visual_dirs = [d for d in custom_visuals_dir.iterdir() if d.is_dir()]
            self.log_callback(f"  📁 Found {len(visual_dirs)} visual directories")
            
            # Get list of visuals that are already registered in build pane
            registered_visuals = set(visual_analysis['build_pane_visuals'].keys())
            
            for visual_dir in visual_dirs:
                visual_id = visual_dir.name
                
                # Check if this is already registered (appears in build pane)
                if visual_id not in registered_visuals:
                    # This is a hidden visual
                    display_name = self._get_visual_display_name_from_folder(visual_dir)
                    size = sum(f.stat().st_size for f in visual_dir.rglob('*') if f.is_file())
                    
                    visual_analysis['hidden_visuals'][visual_id] = {
                        'visual_id': visual_id,
                        'display_name': display_name,
                        'location': 'CustomVisuals (hidden)',
                        'path': visual_dir,
                        'used': False,  # Hidden visuals can't be used
                        'size': size
                    }
                    
                    self.log_callback(f"    🚫 Hidden visual: {display_name} ({visual_id}) - {self._format_bytes(size)}")
        
        # Log summary
        total_build_pane = len(visual_analysis['build_pane_visuals'])
        total_hidden = len(visual_analysis['hidden_visuals'])
        total_appsource = len(visual_analysis['appsource_visuals'])
        total_used = len(all_used_visuals)
        
        self.log_callback(f"  📊 Custom Visual Summary:")
        self.log_callback(f"    AppSource visuals: {total_appsource}")
        self.log_callback(f"    Build pane visuals: {total_build_pane}")
        self.log_callback(f"    Hidden visuals: {total_hidden}")
        self.log_callback(f"    Total used in pages: {total_used}")
        
        return visual_analysis
    
    def _is_visual_used(self, visual_id: str, all_used_visuals: Set[str]) -> bool:
        """Check if a visual is used by matching against all found visual identifiers"""
        # Direct match
        if visual_id in all_used_visuals:
            return True
        
        # Partial matches for different naming patterns
        for used_visual in all_used_visuals:
            # Check if visual_id is contained in used_visual or vice versa
            if (visual_id.lower() in used_visual.lower() or 
                used_visual.lower() in visual_id.lower()):
                return True
            
            # Check for GUID pattern matches
            if len(visual_id) > 10 and len(used_visual) > 10:
                # Extract potential GUID parts and compare
                visual_parts = visual_id.replace('_', '').replace('-', '').lower()
                used_parts = used_visual.replace('_', '').replace('-', '').lower()
                if visual_parts in used_parts or used_parts in visual_parts:
                    return True
        
        return False
    
    def _get_visual_display_name(self, visual_id: str) -> str:
        """Get display name for a visual ID"""
        # Map known visual IDs to display names
        known_mappings = {
            'ChordChart1444757060245': 'Chord 2.4.1.0',
            'PBI_CV_25997FEB_F466_44FA_B562_AC4063283C4C': 'Sparkline by OKViz',
            'waterCupVisual821A487582524ED6878C0A8F277EF02F': 'Water Cup'
        }
        
        return known_mappings.get(visual_id, visual_id)
    
    def _get_visual_display_name_from_folder(self, visual_dir: Path) -> str:
        """Get display name from visual folder's package.json"""
        package_json = visual_dir / "package.json"
        if package_json.exists():
            try:
                with open(package_json, 'r', encoding='utf-8') as f:
                    package_data = json.load(f)
                
                visual_info = package_data.get('visual', {})
                return visual_info.get('displayName', visual_info.get('name', visual_dir.name))
            except:
                pass
        
        return visual_dir.name
    
    def _estimate_visual_size(self, report_dir: Path, visual_id: str) -> int:
        """Estimate the size of a custom visual"""
        # Check CustomVisuals directory
        custom_visuals_dir = report_dir / "CustomVisuals" / visual_id
        if custom_visuals_dir.exists():
            return sum(f.stat().st_size for f in custom_visuals_dir.rglob('*') if f.is_file())
        
        return 0
    
    def _find_theme_opportunities(self, theme_analysis: Dict[str, Any]) -> List[CleanupOpportunity]:
        """Find theme cleanup opportunities - only one theme can be active"""
        opportunities = []
        
        active_theme = theme_analysis['active_theme']  # Tuple of (name, type) or None
        available_themes = theme_analysis['available_themes']
        
        # Find unused themes
        for theme_name, theme_info in available_themes.items():
            # Check if this theme is the currently active one
            is_active = False
            if active_theme:
                active_name, active_type = active_theme
                # Match by name (handle both filename and theme name)
                is_active = (theme_name == active_name or 
                           theme_name.replace('.json', '') == active_name or
                           active_name in theme_name)
            
            if not is_active and theme_info.get('path'):
                reason = "Theme is not currently active"
                size_bytes = theme_info.get('size', 0)
                
                opportunity = CleanupOpportunity(
                    item_type='theme',
                    item_name=theme_name,
                    location=f"{theme_info['location']}/{theme_info['type']}",
                    reason=reason,
                    safety_level='safe',
                    size_bytes=size_bytes
                )
                opportunities.append(opportunity)
        
        return opportunities
    
    def _find_custom_visual_opportunities(self, visual_analysis: Dict[str, Any]) -> List[CleanupOpportunity]:
        """Find custom visual cleanup opportunities with proper categorization"""
        opportunities = []
        
        # Category 1: Unused AppSource/Build Pane visuals (publicCustomVisuals)
        appsource_visuals = visual_analysis['appsource_visuals']
        for visual_id, visual_info in appsource_visuals.items():
            if not visual_info['used']:
                opportunity = CleanupOpportunity(
                    item_type='custom_visual_build_pane',
                    item_name=visual_info['display_name'],
                    location='AppSource (publicCustomVisuals)',
                    reason="AppSource visual is not used in any pages",
                    safety_level='safe',
                    size_bytes=0,  # AppSource visuals don't take local space
                    visual_id=visual_id
                )
                opportunities.append(opportunity)
        
        # Category 2: Unused Build Pane visuals (resourcePackages)
        build_pane_visuals = visual_analysis['build_pane_visuals']
        for visual_id, visual_info in build_pane_visuals.items():
            if not visual_info['used']:
                opportunity = CleanupOpportunity(
                    item_type='custom_visual_build_pane',
                    item_name=visual_info['display_name'],
                    location='Build Pane (resourcePackages)',
                    reason="Visual appears in build pane but is not used in any pages",
                    safety_level='safe',
                    size_bytes=visual_info['size'],
                    visual_id=visual_id
                )
                opportunities.append(opportunity)
        
        # Category 3: Hidden visuals (CustomVisuals folder but not in build pane)
        hidden_visuals = visual_analysis['hidden_visuals']
        for visual_id, visual_info in hidden_visuals.items():
            opportunity = CleanupOpportunity(
                item_type='custom_visual_hidden',
                item_name=visual_info['display_name'],
                location='Hidden in CustomVisuals folder',
                reason="Visual is installed but hidden (not accessible in build pane)",
                safety_level='safe',
                size_bytes=visual_info['size'],
                visual_id=visual_id
            )
            opportunities.append(opportunity)
        
        return opportunities
    
    def _analyze_bookmarks(self, report_data: Dict, report_dir: Path) -> Dict[str, Any]:
        """Analyze bookmarks in the report"""
        self.log_callback("📖 Analyzing bookmarks...")
        
        bookmark_analysis = {
            'bookmarks': {},
            'existing_pages': set(),
            'bookmark_usage': {},
        }
        
        # Get existing pages
        pages_dir = report_dir / "definition" / "pages"
        if pages_dir.exists():
            for page_dir in pages_dir.iterdir():
                if page_dir.is_dir():
                    # Try to get the display name from page.json
                    page_json = page_dir / "page.json"
                    if page_json.exists():
                        try:
                            with open(page_json, 'r', encoding='utf-8') as f:
                                page_data = json.load(f)
                            page_name = page_data.get('displayName', page_dir.name)
                            bookmark_analysis['existing_pages'].add(page_name)
                            # Also add the internal page ID
                            bookmark_analysis['existing_pages'].add(page_dir.name)
                        except:
                            bookmark_analysis['existing_pages'].add(page_dir.name)
                    else:
                        bookmark_analysis['existing_pages'].add(page_dir.name)
        
        # Get bookmarks from bookmarks directory (PBIP structure)
        bookmarks_dir = report_dir / "definition" / "bookmarks"
        if not bookmarks_dir.exists():
            self.log_callback("  📖 No bookmarks directory found")
            return bookmark_analysis
        
        # Read bookmarks.json to get bookmark metadata
        bookmarks_json = bookmarks_dir / "bookmarks.json"
        if not bookmarks_json.exists():
            self.log_callback("  📖 No bookmarks.json found")
            return bookmark_analysis
        
        try:
            with open(bookmarks_json, 'r', encoding='utf-8') as f:
                bookmarks_meta = json.load(f)
            
            bookmark_items = bookmarks_meta.get('items', [])
            if bookmark_items:
                # Count total bookmark items (including parent groups)
                total_bookmark_items = len(bookmark_items)
                
                # Count individual bookmarks vs groups
                individual_bookmarks = 0
                bookmark_groups = 0
                total_child_bookmarks = 0
                
                for bookmark_meta in bookmark_items:
                    children = bookmark_meta.get('children', [])
                    if children:
                        bookmark_groups += 1
                        total_child_bookmarks += len(children)
                    else:
                        individual_bookmarks += 1
                
                # Count actual bookmark files in the directory
                actual_bookmark_files = [f for f in bookmarks_dir.glob("*.bookmark.json")]
                total_actual_bookmarks = len(actual_bookmark_files)
                
                # Create detailed count message
                if bookmark_groups > 0:
                    self.log_callback(f"  📖 Found {total_actual_bookmarks} bookmarks ({individual_bookmarks} standalone, {total_child_bookmarks} in {bookmark_groups} groups)")
                else:
                    self.log_callback(f"  📖 Found {total_actual_bookmarks} bookmarks")
                
                for bookmark_meta in bookmark_items:
                    bookmark_id = bookmark_meta.get('name', '')
                    if not bookmark_id:
                        continue
                    
                    # Handle parent bookmarks with children
                    children = bookmark_meta.get('children', [])
                    bookmark_display_name = bookmark_meta.get('displayName', bookmark_id)
                    
                    # If this is a parent bookmark (has children), process both parent and children
                    if children:
                        # Process parent bookmark
                        bookmark_analysis['bookmarks'][bookmark_id] = {
                            'name': bookmark_id,
                            'display_name': bookmark_display_name,
                            'page_id': '',  # Parent groups don't have page references
                            'page_exists': True,  # Parent groups are always "valid"
                            'bookmark_data': bookmark_meta,
                            'is_parent': True
                        }
                        
                        # Log parent group first
                        self.log_callback(f"    📁 {bookmark_display_name} (group with {len(children)} children)")
                        
                        # Process each child bookmark
                        for child_id in children:
                            child_file = bookmarks_dir / f"{child_id}.bookmark.json"
                            if child_file.exists():
                                try:
                                    with open(child_file, 'r', encoding='utf-8') as f:
                                        child_data = json.load(f)
                                    
                                    child_name = child_data.get('displayName', child_id)
                                    exploration_state = child_data.get('explorationState', {})
                                    active_section = exploration_state.get('activeSection', '')
                                    page_exists = active_section in bookmark_analysis['existing_pages'] if active_section else True
                                    
                                    bookmark_analysis['bookmarks'][child_id] = {
                                        'name': child_id,
                                        'display_name': child_name,
                                        'page_id': active_section,
                                        'page_exists': page_exists,
                                        'bookmark_data': child_data,
                                        'parent_id': bookmark_id
                                    }
                                    
                                    exists_text = "exists" if page_exists else "MISSING"
                                    self.log_callback(f"      📖 {child_name} → page '{active_section[:12]}...' ({exists_text})")
                                    
                                except Exception as e:
                                    self.log_callback(f"      ⚠️ Could not read child bookmark {child_id}: {e}")
                    else:
                        # Regular bookmark without children
                        # Read individual bookmark file
                        bookmark_file = bookmarks_dir / f"{bookmark_id}.bookmark.json"
                        if bookmark_file.exists():
                            try:
                                with open(bookmark_file, 'r', encoding='utf-8') as f:
                                    bookmark_data = json.load(f)
                                
                                bookmark_name = bookmark_data.get('displayName', bookmark_id)
                                
                                # Get page reference from explorationState
                                exploration_state = bookmark_data.get('explorationState', {})
                                active_section = exploration_state.get('activeSection', '')
                                page_exists = active_section in bookmark_analysis['existing_pages'] if active_section else True
                                
                                bookmark_analysis['bookmarks'][bookmark_id] = {
                                    'name': bookmark_id,
                                    'display_name': bookmark_name,
                                    'page_id': active_section,
                                    'page_exists': page_exists,
                                    'bookmark_data': bookmark_data
                                }
                                
                                exists_text = "exists" if page_exists else "MISSING"
                                self.log_callback(f"    📖 {bookmark_name} → page '{active_section[:12]}...' ({exists_text})")
                                
                            except Exception as e:
                                self.log_callback(f"    ⚠️ Could not read bookmark file {bookmark_id}: {e}")
                                continue
            else:
                self.log_callback("  📖 No bookmark items found in bookmarks.json")
                
        except Exception as e:
            self.log_callback(f"  ⚠️ Could not read bookmarks.json: {e}")
            return bookmark_analysis
        
        # Scan for bookmark usage in visuals
        bookmark_usage = self._scan_pages_for_bookmark_usage(report_dir)
        bookmark_analysis['bookmark_usage'] = bookmark_usage
        
        # Create simplified summary for logging
        self._log_bookmark_summary(bookmark_analysis)
        
        return bookmark_analysis
    
    def _log_bookmark_summary(self, bookmark_analysis: Dict[str, Any]) -> None:
        """Log a simplified bookmark summary with hierarchical display"""
        bookmarks = bookmark_analysis.get('bookmarks', {})
        existing_pages = bookmark_analysis.get('existing_pages', set())
        bookmark_usage = bookmark_analysis.get('bookmark_usage', {})
        used_bookmarks = bookmark_usage.get('used_bookmarks', set())
        
        if not bookmarks:
            return
        
        # Organize bookmarks by category and hierarchy
        page_missing = []
        page_exists_unused = []
        page_exists_used = []
        parent_groups = {}  # Track parent groups and their children
        
        # First pass: identify parent groups and their children
        for bookmark_id, bookmark_info in bookmarks.items():
            if bookmark_info.get('is_parent', False):
                parent_groups[bookmark_id] = {
                    'info': bookmark_info,
                    'children': [],
                    'child_usage': {'used': 0, 'total': 0}
                }
        
        # Second pass: categorize all bookmarks and track parent-child relationships
        for bookmark_id, bookmark_info in bookmarks.items():
            bookmark_name = bookmark_info['display_name']
            page_id = bookmark_info['page_id']
            page_exists = bookmark_info['page_exists']
            parent_id = bookmark_info.get('parent_id')
            
            # If this is a child bookmark, add to parent's children list
            if parent_id and parent_id in parent_groups:
                is_child_used = self._is_bookmark_used_including_children(bookmark_id, bookmark_name, bookmark_info, bookmark_analysis)
                parent_groups[parent_id]['children'].append({
                    'id': bookmark_id,
                    'info': bookmark_info,
                    'is_used': is_child_used
                })
                parent_groups[parent_id]['child_usage']['total'] += 1
                if is_child_used:
                    parent_groups[parent_id]['child_usage']['used'] += 1
                continue  # Skip individual processing for child bookmarks
            
            # Process non-child bookmarks
            is_bookmark_used = self._is_bookmark_used_including_children(bookmark_id, bookmark_name, bookmark_info, bookmark_analysis)
            
            if page_id and not page_exists:
                page_missing.append(bookmark_name)
            elif is_bookmark_used:
                if bookmark_info.get('is_parent', False):
                    # For parent groups, create detailed usage info
                    group_usage = self._get_group_usage_details(bookmark_id, bookmark_info, bookmark_analysis)
                    page_exists_used.append((bookmark_name, group_usage, True))  # True = is_group
                else:
                    # Regular bookmark
                    button_count, navigator_count = self._count_bookmark_usage(bookmark_id, bookmark_name, bookmark_usage)
                    usage_detail = self._format_usage_count(button_count, navigator_count)
                    page_exists_used.append((bookmark_name, usage_detail, False))  # False = not_group
            else:
                page_exists_unused.append(bookmark_name)
        
        # Log categorized results with hierarchical display
        if page_missing:
            self.log_callback("  📍 Page Missing:")
            for bookmark_name in page_missing:
                self.log_callback(f"    ❌ {bookmark_name}")
        
        if page_exists_used:
            self.log_callback("  📍 Page Exists - With Button/Navigation:")
            for bookmark_name, usage_detail, is_group in page_exists_used:
                if is_group:
                    self.log_callback(f"    ✅ {bookmark_name} (bookmark group) {usage_detail}")
                    # Show indented children
                    group_id = None
                    for bid, binfo in bookmarks.items():
                        if binfo['display_name'] == bookmark_name and binfo.get('is_parent', False):
                            group_id = bid
                            break
                    
                    if group_id and group_id in parent_groups:
                        children = parent_groups[group_id]['children']
                        for child in children:
                            child_name = child['info']['display_name']
                            child_used = child['is_used']
                            status = "✅" if child_used else "⚠️"
                            # Show child usage details
                            button_count, navigator_count = self._count_bookmark_usage(
                                child['id'], child_name, bookmark_usage
                            )
                            child_usage = self._format_usage_count(button_count, navigator_count)
                            self.log_callback(f"      {status} {child_name} {child_usage}")
                else:
                    self.log_callback(f"    ✅ {bookmark_name} {usage_detail}")
        
        if page_exists_unused:
            self.log_callback("  📍 Page Exists - No Button/Navigation:")
            for bookmark_name in page_exists_unused:
                # Check if this is a parent group
                is_group = False
                group_id = None
                for bid, binfo in bookmarks.items():
                    if binfo['display_name'] == bookmark_name and binfo.get('is_parent', False):
                        is_group = True
                        group_id = bid
                        break
                
                if is_group and group_id:
                    # Show parent group with details
                    group_usage = self._get_group_usage_details(group_id, bookmarks[group_id], bookmark_analysis)
                    self.log_callback(f"    ⚠️ {bookmark_name} (bookmark group) {group_usage}")
                    
                    # Show indented children for unused groups too
                    if group_id in parent_groups:
                        children = parent_groups[group_id]['children']
                        for child in children:
                            child_name = child['info']['display_name']
                            child_used = child['is_used']
                            status = "✅" if child_used else "⚠️"
                            # Show child usage details
                            button_count, navigator_count = self._count_bookmark_usage(
                                child['id'], child_name, bookmark_usage
                            )
                            child_usage = self._format_usage_count(button_count, navigator_count)
                            self.log_callback(f"      {status} {child_name} {child_usage}")
                else:
                    # Regular bookmark (not a group)
                    self.log_callback(f"    ⚠️ {bookmark_name}")
    
    def _get_group_usage_details(self, group_id: str, group_info: Dict, bookmark_analysis: Dict) -> str:
        """Get detailed usage statistics for a bookmark group"""
        bookmark_usage = bookmark_analysis.get('bookmark_usage', {})
        all_bookmarks = bookmark_analysis.get('bookmarks', {})
        
        # Count navigator usage for the group itself
        group_navigator_count = 0
        bookmark_navigators = bookmark_usage.get('bookmark_navigators', [])
        for page_name, visual_type, bookmark_refs in bookmark_navigators:
            if group_id in bookmark_refs:
                group_navigator_count += 1
        
        # Count children and their usage
        children_used = 0
        children_total = 0
        
        for child_id, child_info in all_bookmarks.items():
            if child_info.get('parent_id') == group_id:
                children_total += 1
                
                # Check if child is used (either directly or via parent group)
                child_name = child_info.get('display_name', '')
                if self._is_bookmark_used_including_children(child_id, child_name, child_info, bookmark_analysis):
                    children_used += 1
        
        # Format the detailed statistics
        parts = []
        
        if group_navigator_count > 0:
            nav_text = "navigator" if group_navigator_count == 1 else "navigators"
            parts.append(f"{group_navigator_count} {nav_text}")
        else:
            parts.append("0 navigators")
        
        if children_total > 0:
            parts.append(f"{children_used}/{children_total} bookmarks used")
        
        return f"({', '.join(parts)})"
    
    def _count_bookmark_usage(self, bookmark_id: str, bookmark_name: str, bookmark_usage: Dict) -> tuple:
        """Count how many buttons/navigators use this bookmark"""
        button_count = 0
        navigator_count = 0
        
        # Count in bookmark navigators - check by both ID and name
        navigators = bookmark_usage.get('bookmark_navigators', [])
        for page_name, visual_type, bookmark_refs in navigators:
            # bookmark_refs is a set, so check if our ID or name is in it
            if bookmark_id in bookmark_refs or bookmark_name in bookmark_refs:
                navigator_count += 1
        
        # Count in bookmark buttons - check by both ID and name  
        buttons = bookmark_usage.get('bookmark_buttons', [])
        for page_name, visual_type, bookmark_refs in buttons:
            # bookmark_refs is a set, so check if our ID or name is in it
            if bookmark_id in bookmark_refs or bookmark_name in bookmark_refs:
                button_count += 1
        
        return button_count, navigator_count
    
    def _is_bookmark_used_including_children(self, bookmark_id: str, bookmark_name: str, bookmark_info: Dict, bookmark_analysis: Dict) -> bool:
        """Check if a bookmark (or its children if it's a parent) is used"""
        bookmark_usage = bookmark_analysis.get('bookmark_usage', {})
        used_bookmarks = bookmark_usage.get('used_bookmarks', set())
        
        # Direct usage check - check both ID and name
        if bookmark_id in used_bookmarks or bookmark_name in used_bookmarks:
            return True
        
        # For parent groups, check if the group itself is directly referenced by navigators
        if bookmark_info.get('is_parent', False):
            # Check if any bookmark navigator references this parent group directly
            bookmark_navigators = bookmark_usage.get('bookmark_navigators', [])
            for page_name, visual_type, bookmark_refs in bookmark_navigators:
                if bookmark_id in bookmark_refs:
                    return True
            
            # Also check if any children are used
            all_bookmarks = bookmark_analysis.get('bookmarks', {})
            for child_id, child_info in all_bookmarks.items():
                if child_info.get('parent_id') == bookmark_id:
                    child_name = child_info.get('display_name', '')
                    if child_id in used_bookmarks or child_name in used_bookmarks:
                        return True
        
        # For child bookmarks, check if their parent group is used by a navigator
        parent_id = bookmark_info.get('parent_id')
        if parent_id:
            all_bookmarks = bookmark_analysis.get('bookmarks', {})
            parent_info = all_bookmarks.get(parent_id, {})
            if parent_info:
                # Check if parent group is used by any navigator
                bookmark_navigators = bookmark_usage.get('bookmark_navigators', [])
                for page_name, visual_type, bookmark_refs in bookmark_navigators:
                    if parent_id in bookmark_refs:
                        return True
        
        return False
    
    def _format_usage_count(self, button_count: int, navigator_count: int) -> str:
        """Format the usage count string based on buttons and navigators"""
        parts = []
        
        if button_count > 0:
            if button_count == 1:
                parts.append("1 button")
            else:
                parts.append(f"{button_count} buttons")
        
        if navigator_count > 0:
            if navigator_count == 1:
                parts.append("1 navigator")
            else:
                parts.append(f"{navigator_count} navigators")
        
        if parts:
            return f"({', '.join(parts)})"
        else:
            return "(unused)"
    
    def _scan_pages_for_bookmark_usage(self, report_dir: Path) -> Dict[str, Any]:
        """Scan all pages to find bookmark usage in visuals"""
        bookmark_usage = {
            'used_bookmarks': set(),
            'pages_with_navigation': set(),
            'bookmark_navigators': [],
            'bookmark_buttons': []
        }
        
        pages_dir = report_dir / "definition" / "pages"
        if not pages_dir.exists():
            return bookmark_usage
        
        for page_dir in pages_dir.iterdir():
            if not page_dir.is_dir():
                continue
                
            page_name = page_dir.name
            visuals_dir = page_dir / "visuals"
            
            if not visuals_dir.exists():
                continue
            
            page_has_navigation = False
            
            for visual_dir in visuals_dir.iterdir():
                if not visual_dir.is_dir():
                    continue
                    
                visual_json = visual_dir / "visual.json"
                if not visual_json.exists():
                    continue
                    
                try:
                    with open(visual_json, 'r', encoding='utf-8') as f:
                        visual_data = json.load(f)
                    
                    # Check for bookmark navigation in various places
                    bookmark_refs = self._extract_bookmark_references(visual_data)
                    
                    if bookmark_refs:
                        page_has_navigation = True
                        bookmark_usage['used_bookmarks'].update(bookmark_refs)
                        
                        # Determine visual type for logging
                        visual_config = visual_data.get('visual', {})
                        visual_type = visual_config.get('visualType', 'unknown')
                        
                        # Log the bookmark references found
                        for bookmark_ref in bookmark_refs:
                            if 'bookmarkNavigator' in visual_type.lower() or visual_type == 'bookmarkNavigator':
                                bookmark_usage['bookmark_navigators'].append((page_name, visual_type, {bookmark_ref}))
                            elif 'button' in visual_type.lower() or 'actionButton' in visual_type:
                                bookmark_usage['bookmark_buttons'].append((page_name, visual_type, {bookmark_ref}))
                            else:
                                # Other visual types that might reference bookmarks
                                pass
                        
                except Exception:
                    continue
            
            if page_has_navigation:
                bookmark_usage['pages_with_navigation'].add(page_name)
        
        return bookmark_usage
    
    def _extract_bookmark_references(self, visual_data: Dict) -> Set[str]:
        """Extract bookmark references from visual data"""
        bookmark_refs = set()
        
        def search_dict_for_bookmarks(obj, path=""):
            """Recursively search for bookmark references in nested dictionaries"""
            if isinstance(obj, dict):
                for key, value in obj.items():
                    current_path = f"{path}.{key}" if path else key
                    
                    # Look for bookmark-related keys
                    if 'bookmark' in key.lower():
                        if isinstance(value, str) and value:
                            bookmark_refs.add(value)
                        elif isinstance(value, dict) and 'name' in value:
                            bookmark_refs.add(value['name'])
                        elif isinstance(value, dict):
                            # Check for nested bookmark structure like bookmark.expr.Literal.Value
                            literal_value = self._extract_literal_value(value)
                            if literal_value:
                                bookmark_refs.add(literal_value)
                    
                    # Check for navigation actions
                    elif key.lower() in ['navigation', 'action', 'bookmarkstate', 'bookmarkname']:
                        if isinstance(value, str) and value:
                            bookmark_refs.add(value)
                        elif isinstance(value, dict):
                            search_dict_for_bookmarks(value, current_path)
                    
                    # Always recursively search nested objects to find deep bookmark references
                    if isinstance(value, (dict, list)):
                        search_dict_for_bookmarks(value, current_path)
            
            elif isinstance(obj, list):
                for i, item in enumerate(obj):
                    search_dict_for_bookmarks(item, f"{path}[{i}]")
        
        search_dict_for_bookmarks(visual_data)
        return bookmark_refs
    
    def _extract_literal_value(self, obj: Dict) -> str:
        """Extract value from Power BI's expr.Literal.Value structure"""
        if isinstance(obj, dict):
            expr = obj.get('expr', {})
            if isinstance(expr, dict):
                literal = expr.get('Literal', {})
                if isinstance(literal, dict):
                    value = literal.get('Value', '')
                    if isinstance(value, str):
                        # Remove surrounding quotes if present
                        return value.strip("'\"")
        return ""
    
    def _find_bookmark_opportunities(self, bookmark_analysis: Dict[str, Any]) -> List[CleanupOpportunity]:
        """Find bookmark cleanup opportunities"""
        opportunities = []
        
        bookmarks = bookmark_analysis.get('bookmarks', {})
        bookmark_usage = bookmark_analysis.get('bookmark_usage', {})
        used_bookmarks = bookmark_usage.get('used_bookmarks', set())
        
        # Track which children will be removed for each parent group
        parent_group_removals = {}  # parent_id -> {total_children, children_to_remove}
        
        for bookmark_id, bookmark_info in bookmarks.items():
            bookmark_name = bookmark_info['display_name']
            page_id = bookmark_info['page_id']
            page_exists = bookmark_info['page_exists']
            parent_id = bookmark_info.get('parent_id')
            
            # Initialize parent group tracking
            if parent_id and parent_id not in parent_group_removals:
                parent_group_removals[parent_id] = {'total_children': 0, 'children_to_remove': 0}
            
            # Category A: Guaranteed unused (page doesn't exist)
            if page_id and not page_exists:
                opportunity = CleanupOpportunity(
                    item_type='bookmark_guaranteed_unused',
                    item_name=bookmark_name,
                    location='Bookmarks',
                    reason=f"Target page '{page_id[:12]}...' no longer exists",
                    safety_level='safe',
                    size_bytes=0,  # Bookmarks don't take significant space
                    bookmark_id=bookmark_id
                )
                opportunities.append(opportunity)
                
                # Track removal for parent group
                if parent_id:
                    parent_group_removals[parent_id]['children_to_remove'] += 1
            
            # Category B: Likely unused (no navigation found) - but skip parent groups
            elif not bookmark_info.get('is_parent', False) and not self._is_bookmark_used_including_children(bookmark_id, bookmark_name, bookmark_info, bookmark_analysis):
                opportunity = CleanupOpportunity(
                    item_type='bookmark_likely_unused',
                    item_name=bookmark_name,
                    location='Bookmarks',
                    reason="No navigation buttons found, but could be used via bookmark pane in service",
                    safety_level='warning',
                    size_bytes=0,  # Bookmarks don't take significant space
                    bookmark_id=bookmark_id
                )
                opportunities.append(opportunity)
                
                # Track removal for parent group
                if parent_id:
                    parent_group_removals[parent_id]['children_to_remove'] += 1
        
        # Count total children for each parent group
        for bookmark_id, bookmark_info in bookmarks.items():
            parent_id = bookmark_info.get('parent_id')
            if parent_id and parent_id in parent_group_removals:
                parent_group_removals[parent_id]['total_children'] += 1
        
        # Check if any parent groups should be removed (all children being removed)
        for parent_id, removal_info in parent_group_removals.items():
            if removal_info['children_to_remove'] == removal_info['total_children'] and removal_info['total_children'] > 0:
                # All children of this parent group are being removed, so remove the parent too
                parent_info = bookmarks.get(parent_id, {})
                if parent_info:
                    parent_name = parent_info.get('display_name', parent_id)
                    opportunity = CleanupOpportunity(
                        item_type='bookmark_empty_group',
                        item_name=parent_name,
                        location='Bookmarks',
                        reason=f"Parent group is empty after removing all {removal_info['total_children']} unused child bookmark(s)",
                        safety_level='safe',
                        size_bytes=0,
                        bookmark_id=parent_id
                    )
                    opportunities.append(opportunity)
        
        # Also check for parent groups that are unused but don't have any children being removed
        # (groups that are already empty or have no navigation)
        for bookmark_id, bookmark_info in bookmarks.items():
            if bookmark_info.get('is_parent', False):
                # Check if this parent group is unused
                if not self._is_bookmark_used_including_children(bookmark_id, bookmark_info['display_name'], bookmark_info, bookmark_analysis):
                    # This parent group is not used - check if we already added it
                    already_added = any(opp.bookmark_id == bookmark_id for opp in opportunities)
                    if not already_added:
                        opportunity = CleanupOpportunity(
                            item_type='bookmark_empty_group',
                            item_name=bookmark_info['display_name'],
                            location='Bookmarks',
                            reason="Bookmark group is not used by any navigators and has no active bookmarks",
                            safety_level='safe',
                            size_bytes=0,
                            bookmark_id=bookmark_id
                        )
                        opportunities.append(opportunity)
        
        return opportunities
    
    def _analyze_visual_filters(self, report_dir: Path) -> Dict[str, Any]:
        """Analyze visual-level filters across all pages and visuals"""
        self.log_callback("🎯 Analyzing visual-level filters...")
        
        filter_analysis = {
            'pages_with_visual_filters': {},
            'total_visual_filters': 0,
            'total_visuals_with_filters': 0,
            'total_pages_scanned': 0
        }
        
        pages_dir = report_dir / "definition" / "pages"
        if not pages_dir.exists():
            self.log_callback("  ⚠️ No pages directory found")
            return filter_analysis
        
        page_dirs = [d for d in pages_dir.iterdir() if d.is_dir() and d.name != "pages.json"]
        filter_analysis['total_pages_scanned'] = len(page_dirs)
        
        self.log_callback(f"  📄 Scanning {len(page_dirs)} pages for visual-level filters...")
        
        for page_dir in page_dirs:
            page_name = page_dir.name
            
            # Get page display name if available
            page_json = page_dir / "page.json"
            page_display_name = page_name
            if page_json.exists():
                try:
                    with open(page_json, 'r', encoding='utf-8') as f:
                        page_data = json.load(f)
                    page_display_name = page_data.get('displayName', page_name)
                except:
                    pass
            
            visuals_dir = page_dir / "visuals"
            if not visuals_dir.exists():
                continue
            
            page_visual_filters = []
            visual_dirs = [d for d in visuals_dir.iterdir() if d.is_dir()]
            
            for visual_dir in visual_dirs:
                visual_json = visual_dir / "visual.json"
                if not visual_json.exists():
                    continue
                
                try:
                    with open(visual_json, 'r', encoding='utf-8') as f:
                        visual_data = json.load(f)
                    
                    # Check for visual-level filters
                    filter_config = visual_data.get('filterConfig', {})
                    filters = filter_config.get('filters', [])
                    
                    if filters:
                        visual_name = visual_data.get('name', visual_dir.name)
                        visual_type = visual_data.get('visual', {}).get('visualType', 'unknown')
                        
                        # Count total and visible filters
                        total_filters = len(filters)
                        visible_filters = len([f for f in filters if not f.get('isHiddenInViewMode', False)])
                        hidden_filters = total_filters - visible_filters
                        
                        page_visual_filters.append({
                            'visual_id': visual_name,
                            'visual_type': visual_type,
                            'total_filters': total_filters,
                            'visible_filters': visible_filters,
                            'hidden_filters': hidden_filters,
                            'filters': filters,
                            'visual_path': visual_json
                        })
                        
                        filter_analysis['total_visual_filters'] += visible_filters
                        filter_analysis['total_visuals_with_filters'] += 1
                        
                except Exception as e:
                    # Skip problematic visuals
                    continue
            
            if page_visual_filters:
                filter_analysis['pages_with_visual_filters'][page_name] = {
                    'page_display_name': page_display_name,
                    'visuals_with_filters': page_visual_filters
                }
        
        # Log summary
        pages_with_filters = len(filter_analysis['pages_with_visual_filters'])
        total_filters = filter_analysis['total_visual_filters']
        total_visuals = filter_analysis['total_visuals_with_filters']
        
        if total_filters > 0:
            self.log_callback(f"  🎯 Found {total_filters} visible visual-level filters across {total_visuals} visuals on {pages_with_filters} pages")
            
            # Log details by page (only page summaries, not individual visuals)
            for page_id, page_info in filter_analysis['pages_with_visual_filters'].items():
                page_display = page_info['page_display_name']
                visuals = page_info['visuals_with_filters']
                page_filter_count = sum(v['visible_filters'] for v in visuals)
                
                self.log_callback(f"    📄 {page_display}: {page_filter_count} filters on {len(visuals)} visuals")
        else:
            self.log_callback("  ✅ No visible visual-level filters found")
        
        return filter_analysis
    
    def _find_visual_filter_opportunities(self, filter_analysis: Dict[str, Any]) -> List[CleanupOpportunity]:
        """Find opportunities to hide visual-level filters"""
        opportunities = []
        
        total_visible_filters = filter_analysis.get('total_visual_filters', 0)
        
        if total_visible_filters > 0:
            # Create a single opportunity representing all visual filters
            opportunity = CleanupOpportunity(
                item_type='visual_filter',
                item_name=f"All visual-level filters ({total_visible_filters} filters)",
                location='Visual filterConfig across all pages',
                reason=f"Hide {total_visible_filters} visual-level filters to clean up interface",
                safety_level='safe',
                size_bytes=0,  # Hiding filters doesn't save space
                filter_count=total_visible_filters
            )
            opportunities.append(opportunity)
        
        return opportunities
    
    def _format_bytes(self, bytes_count: int) -> str:
        """Format bytes into human readable format"""
        if bytes_count == 0:
            return "0 B"
        
        for unit in ['B', 'KB', 'MB', 'GB']:
            if bytes_count < 1024:
                return f"{bytes_count:.1f} {unit}"
            bytes_count /= 1024
        return f"{bytes_count:.1f} TB"
