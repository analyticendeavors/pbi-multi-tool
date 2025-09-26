"""
Table Column Widths Core Engine - Enhanced Version
Built by Reid Havens of Analytic Endeavors

Core logic for analyzing and standardizing column widths in Power BI visuals.
Enhanced with matrix improvements and better content-aware logic.
Defaults to Auto-fit for intelligent width calculation.
"""

import json
import re
from pathlib import Path
from typing import Dict, List, Any, Optional, Tuple, Set
from dataclasses import dataclass
from enum import Enum
import math


class DataScale(Enum):
    """Enumeration for data scales"""
    ONES = "ones"           # 1-999
    TENS = "tens"           # 10-9,999
    HUNDREDS = "hundreds"   # 100-99,999
    THOUSANDS = "thousands" # 1K-999K
    MILLIONS = "millions"   # 1M-999M
    BILLIONS = "billions"   # 1B-999B
    TRILLIONS = "trillions" # 1T+


class DataMagnitude(Enum):
    """Enumeration for data magnitude"""
    SMALL = "small"      # <1000
    MEDIUM = "medium"    # 1K-999K
    LARGE = "large"      # 1M-999M
    XLARGE = "xlarge"    # 1B+


@dataclass
class ScaleConfiguration:
    """Configuration for data scale awareness"""
    typical_scale: DataScale = DataScale.THOUSANDS
    magnitude: DataMagnitude = DataMagnitude.MEDIUM
    use_abbreviations: bool = True  # K, M, B, T
    decimal_places: int = 1         # 1.2K vs 1,234
    currency_symbol: bool = True    # $1.2K vs 1.2K
    
    def get_sample_display_length(self) -> int:
        """Get typical display length for this scale configuration"""
        # Base length calculations for different scales
        base_lengths = {
            DataScale.ONES: 3,       # "999"
            DataScale.TENS: 4,       # "9,999" 
            DataScale.HUNDREDS: 6,   # "99,999"
            DataScale.THOUSANDS: 5,  # "999K" or "999,999"
            DataScale.MILLIONS: 5,   # "999M" or "999,999,999"
            DataScale.BILLIONS: 5,   # "999B"
            DataScale.TRILLIONS: 5   # "999T"
        }
        
        base_length = base_lengths.get(self.typical_scale, 5)
        
        # Add currency symbol
        if self.currency_symbol:
            base_length += 1
        
        # Add decimal places for abbreviated formats
        if self.use_abbreviations and self.decimal_places > 0:
            base_length += self.decimal_places + 1  # +1 for decimal point
        
        # Add comma separators for non-abbreviated large numbers
        if not self.use_abbreviations:
            if self.typical_scale in [DataScale.THOUSANDS, DataScale.MILLIONS, DataScale.BILLIONS]:
                base_length += 2  # Typical comma separators
        
        return base_length


class FieldType(Enum):
    """Enumeration for field types"""
    CATEGORICAL = "categorical"
    MEASURE = "measure"
    UNKNOWN = "unknown"


class VisualType(Enum):
    """Enumeration for visual types"""
    TABLE = "tableEx"
    MATRIX = "pivotTable"
    UNKNOWN = "unknown"


@dataclass
class FieldInfo:
    """Information about a field in a visual"""
    name: str
    display_name: str
    field_type: FieldType
    metadata_key: str
    current_width: Optional[float] = None
    suggested_width: Optional[float] = None
    is_overridden: bool = False
    scale_config: Optional[ScaleConfiguration] = None  # For measures


@dataclass
class FontInfo:
    """Font information from visual configuration"""
    family: str = "Segoe UI"
    size: int = 11
    weight: str = "normal"
    
    def get_char_width(self) -> float:
        """Calculate average character width in pixels"""
        # Base calculations for Segoe UI at different sizes
        base_char_width = {
            8: 5.5, 9: 6.0, 10: 6.5, 11: 7.0, 12: 7.5, 
            13: 8.0, 14: 8.5, 15: 9.0, 16: 9.5
        }
        
        char_width = base_char_width.get(self.size, 7.0)
        
        # Adjust for bold text
        if "bold" in self.weight.lower():
            char_width *= 1.15
        
        return char_width


@dataclass
class VisualInfo:
    """Information about a visual containing tables/matrices"""
    visual_id: str
    visual_name: str
    visual_type: VisualType
    page_name: str
    page_id: str
    fields: List[FieldInfo]
    font_info: FontInfo
    layout_type: Optional[str] = None  # For matrices: "Compact" or "Outline"
    current_width: int = 0
    current_height: int = 0


class WidthPreset(Enum):
    """Predefined width presets"""
    NARROW = "narrow"
    MEDIUM = "medium" 
    WIDE = "wide"
    AUTO_FIT = "auto_fit"
    FIT_TO_TOTALS = "fit_to_totals"
    CUSTOM = "custom"


@dataclass
class WidthConfiguration:
    """Configuration for column widths - defaults to Auto-fit for intelligent sizing"""
    categorical_preset: WidthPreset = WidthPreset.AUTO_FIT  # Changed to auto-fit default
    categorical_custom: int = 105
    measure_preset: WidthPreset = WidthPreset.AUTO_FIT      # Changed to auto-fit default
    measure_custom: int = 95
    max_width: int = 300
    min_width: int = 50
    # New scale configuration
    default_scale_config: ScaleConfiguration = None
    
    def __post_init__(self):
        if self.default_scale_config is None:
            self.default_scale_config = ScaleConfiguration()


class TableColumnWidthsEngine:
    """
    Core engine for table column width operations with enhanced content-aware logic
    """
    
    # Simplified preset width definitions (in pixels)
    PRESET_WIDTHS = {
        WidthPreset.NARROW: {"categorical": 60, "measure": 70},
        WidthPreset.MEDIUM: {"categorical": 105, "measure": 95},      
        WidthPreset.WIDE: {"categorical": 165, "measure": 145},
    }
    
    def __init__(self, pbip_path: str):
        self.pbip_path = Path(pbip_path)
        self.report_dir = self.pbip_path.parent / f"{self.pbip_path.stem}.Report"
        self.visuals_info: List[VisualInfo] = []
        
        if not self.report_dir.exists():
            raise ValueError(f"Report directory not found: {self.report_dir}")
    
    def scan_visuals(self) -> List[VisualInfo]:
        """Scan the report for table and matrix visuals"""
        self.visuals_info.clear()
        
        # Find all page directories
        pages_dir = self.report_dir / "definition" / "pages"
        if not pages_dir.exists():
            raise ValueError("Pages directory not found in report")
        
        for page_dir in pages_dir.iterdir():
            if page_dir.is_dir() and page_dir.name != "pages.json":
                self._scan_page_visuals(page_dir)
        
        return self.visuals_info
    
    def _scan_page_visuals(self, page_dir: Path):
        """Scan visuals in a specific page"""
        page_id = page_dir.name
        
        # Read page info
        page_info = self._read_page_info(page_dir)
        page_name = page_info.get("displayName", f"Page {page_id[:8]}")
        
        # Find visuals directory
        visuals_dir = page_dir / "visuals"
        if not visuals_dir.exists():
            return
        
        # Process each visual
        for visual_dir in visuals_dir.iterdir():
            if visual_dir.is_dir():
                visual_file = visual_dir / "visual.json"
                if visual_file.exists():
                    try:
                        visual_info = self._analyze_visual(visual_file, page_id, page_name)
                        if visual_info:
                            self.visuals_info.append(visual_info)
                    except Exception as e:
                        # Log error but continue processing other visuals
                        print(f"Warning: Could not analyze visual {visual_dir.name}: {e}")
                        continue
    
    def _read_page_info(self, page_dir: Path) -> Dict[str, Any]:
        """Read page information"""
        page_file = page_dir / "page.json"
        if page_file.exists():
            try:
                with open(page_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except Exception:
                pass
        return {}
    
    def _analyze_visual(self, visual_file: Path, page_id: str, page_name: str) -> Optional[VisualInfo]:
        """Analyze a single visual file"""
        try:
            with open(visual_file, 'r', encoding='utf-8') as f:
                visual_data = json.load(f)
            
            visual_config = visual_data.get("visual", {})
            visual_type_str = visual_config.get("visualType", "")
            
            # Check if it's a table or matrix
            if visual_type_str == "tableEx":
                visual_type = VisualType.TABLE
            elif visual_type_str == "pivotTable":
                visual_type = VisualType.MATRIX
            else:
                return None  # Not a table or matrix
            
            # Extract basic info
            visual_id = visual_data.get("name", "")
            visual_name = self._extract_visual_name(visual_data, visual_id)
            
            # Extract position/size info
            position = visual_data.get("position", {})
            current_width = position.get("width", 0)
            current_height = position.get("height", 0)
            
            # Extract font information
            font_info = self._extract_font_info(visual_config)
            
            # Extract layout type for matrices
            layout_type = None
            if visual_type == VisualType.MATRIX:
                layout_type = self._extract_matrix_layout(visual_config)
            
            # Extract field information
            fields = self._extract_field_info(visual_config, visual_type)
            
            return VisualInfo(
                visual_id=visual_id,
                visual_name=visual_name,
                visual_type=visual_type,
                page_name=page_name,
                page_id=page_id,
                fields=fields,
                font_info=font_info,
                layout_type=layout_type,
                current_width=current_width,
                current_height=current_height
            )
            
        except Exception as e:
            # Log error but continue processing other visuals
            print(f"Error analyzing visual {visual_file}: {e}")
            return None
    
    def _extract_visual_name(self, visual_data: Dict[str, Any], visual_id: str) -> str:
        """Extract a readable name for the visual"""
        # Try to find visual title/name in multiple locations
        visual_config = visual_data.get("visual", {})
        
        # Method 1: Check for title in visual objects
        objects = visual_config.get("objects", {})
        if "general" in objects:
            for obj in objects["general"]:
                if "properties" in obj and "title" in obj["properties"]:
                    title_expr = obj["properties"]["title"].get("expr", {})
                    if "Literal" in title_expr and "Value" in title_expr["Literal"]:
                        title = title_expr["Literal"]["Value"].strip("'\"")
                        if title and title.lower() != "title":  # Skip default "Title" text
                            return title
        
        # Method 2: Check in visual data root level
        if "title" in visual_data:
            title = visual_data["title"]
            if isinstance(title, str) and title.strip():
                return title.strip()
        
        # Method 3: Check visual config for name/title
        if "title" in visual_config:
            title = visual_config["title"]
            if isinstance(title, str) and title.strip():
                return title.strip()
        
        # Method 4: Check for displayName
        if "displayName" in visual_data:
            name = visual_data["displayName"]
            if isinstance(name, str) and name.strip():
                return name.strip()
        
        # Method 5: Generate friendly name based on visual type
        visual_type = visual_config.get("visualType", "Visual")
        if visual_type == "tableEx":
            return "Table"
        elif visual_type == "pivotTable":
            # Try to determine matrix layout for better naming
            layout_type = self._extract_matrix_layout(visual_config)
            return f"{layout_type} Matrix"
        
        # Fallback: Use visual type + short ID
        return f"{visual_type} ({visual_id[:8]}...)"
    
    def _extract_font_info(self, visual_config: Dict[str, Any]) -> FontInfo:
        """Extract font information from visual configuration"""
        font_info = FontInfo()
        
        objects = visual_config.get("objects", {})
        
        # Look for font settings in various object types
        font_objects = ["columnHeaders", "rowHeaders", "values", "general"]
        
        for obj_type in font_objects:
            if obj_type in objects:
                for obj in objects[obj_type]:
                    properties = obj.get("properties", {})
                    
                    # Extract font family
                    if "fontFamily" in properties:
                        font_expr = properties["fontFamily"].get("expr", {})
                        if "Literal" in font_expr and "Value" in font_expr["Literal"]:
                            font_info.family = font_expr["Literal"]["Value"].strip("'\"")
                    
                    # Extract font size
                    if "fontSize" in properties:
                        size_expr = properties["fontSize"].get("expr", {})
                        if "Literal" in size_expr and "Value" in size_expr["Literal"]:
                            try:
                                font_info.size = int(float(size_expr["Literal"]["Value"]))
                            except ValueError:
                                pass
                    
                    # Extract font weight
                    if "fontWeight" in properties:
                        weight_expr = properties["fontWeight"].get("expr", {})
                        if "Literal" in weight_expr and "Value" in weight_expr["Literal"]:
                            font_info.weight = weight_expr["Literal"]["Value"].strip("'\"")
                    
                    # If we found font settings, we can break
                    if any(key in properties for key in ["fontFamily", "fontSize", "fontWeight"]):
                        break
        
        return font_info
    
    def _extract_matrix_layout(self, visual_config: Dict[str, Any]) -> Optional[str]:
        """Extract matrix layout type (Compact, Outline, or Tabular)"""
        objects = visual_config.get("objects", {})
        
        # Check for layout in general objects
        if "general" in objects:
            for obj in objects["general"]:
                properties = obj.get("properties", {})
                if "layout" in properties:
                    layout_expr = properties["layout"].get("expr", {})
                    if "Literal" in layout_expr and "Value" in layout_expr["Literal"]:
                        layout_value = layout_expr["Literal"]["Value"].strip("'\"")
                        # Map Power BI layout values to readable names
                        if layout_value.lower() == "outline":
                            return "Outline"
                        elif layout_value.lower() == "compact":
                            return "Compact"
                        elif layout_value.lower() == "tabular":
                            return "Tabular"
        
        # Check for matrix style indicators in row headers
        if "rowHeaders" in objects:
            for obj in objects["rowHeaders"]:
                properties = obj.get("properties", {})
                # Look for stepped layout property (indicates tabular)
                if "stepped" in properties:
                    stepped_expr = properties["stepped"].get("expr", {})
                    if "Literal" in stepped_expr and "Value" in stepped_expr["Literal"]:
                        is_stepped = stepped_expr["Literal"]["Value"].lower() == "true"
                        if is_stepped:
                            return "Tabular"
        
        # Fallback: try to detect from query structure
        query = visual_config.get("query", {})
        query_state = query.get("queryState", {})
        rows = query_state.get("Rows", {}).get("projections", [])
        
        # If multiple hierarchy levels, likely compact or tabular
        hierarchy_count = sum(1 for proj in rows if "HierarchyLevel" in proj.get("field", {}))
        
        if hierarchy_count > 2:
            # Multiple hierarchies typically use Compact by default
            return "Compact"
        elif hierarchy_count > 0:
            # Single hierarchy might be Tabular
            return "Tabular"
        
    def _detect_hierarchy_levels(self, visual_info: VisualInfo) -> int:
        """Detect the number of hierarchy levels in the matrix for intelligent width calculation"""
        if visual_info.visual_type != VisualType.MATRIX:
            return 0
        
        # Count unique hierarchy levels from categorical fields
        hierarchy_levels = set()
        
        for field in visual_info.fields:
            if field.field_type == FieldType.CATEGORICAL:
                content_type = self._analyze_content_type(field.display_name)
                if content_type in ["hierarchy_parent", "hierarchy_child"]:
                    hierarchy_levels.add(field.display_name.lower())
        
        # Estimate hierarchy depth based on field patterns
        level_indicators = {
            "year": 1,
            "quarter": 2, 
            "month": 3,
            "week": 4,
            "day": 5,
            "date": 5,
            "eow": 4,  # End of Week
            "bow": 4   # Beginning of Week
        }
        
        max_level = 0
        for field in visual_info.fields:
            if field.field_type == FieldType.CATEGORICAL:
                field_lower = field.display_name.lower()
                for indicator, level in level_indicators.items():
                    if indicator in field_lower:
                        max_level = max(max_level, level)
        
        # Return the higher of detected levels or field count
        return max(max_level, len([f for f in visual_info.fields if f.field_type == FieldType.CATEGORICAL]))
    
    def _extract_field_info(self, visual_config: Dict[str, Any], visual_type: VisualType) -> List[FieldInfo]:
        """Extract field information from visual query"""
        fields = []
        
        query = visual_config.get("query", {})
        query_state = query.get("queryState", {})
        
        # Get current column widths
        current_widths = self._extract_current_widths(visual_config)
        
        # Extract fields based on visual type
        if visual_type == VisualType.TABLE:
            # For tables, fields are in Values
            projections = query_state.get("Values", {}).get("projections", [])
        else:
            # For matrices, combine Rows and Values
            projections = []
            projections.extend(query_state.get("Rows", {}).get("projections", []))
            projections.extend(query_state.get("Values", {}).get("projections", []))
        
        for projection in projections:
            field_info = self._parse_field_projection(projection, current_widths)
            if field_info:
                fields.append(field_info)
        
        return fields
    
    def _extract_current_widths(self, visual_config: Dict[str, Any]) -> Dict[str, float]:
        """Extract current column widths from visual objects"""
        current_widths = {}
        
        objects = visual_config.get("objects", {})
        column_widths = objects.get("columnWidth", [])
        
        for width_obj in column_widths:
            selector = width_obj.get("selector", {})
            metadata = selector.get("metadata", "")
            
            properties = width_obj.get("properties", {})
            value_expr = properties.get("value", {}).get("expr", {})
            
            if "Literal" in value_expr and "Value" in value_expr["Literal"]:
                try:
                    width_str = value_expr["Literal"]["Value"]
                    # Remove 'D' suffix if present
                    width_value = float(width_str.rstrip('D'))
                    current_widths[metadata] = width_value
                except ValueError:
                    pass
        
        return current_widths
    
    def _parse_field_projection(self, projection: Dict[str, Any], current_widths: Dict[str, float]) -> Optional[FieldInfo]:
        """Parse a field projection to extract field information"""
        field = projection.get("field", {})
        query_ref = projection.get("queryRef", "")
        native_query_ref = projection.get("nativeQueryRef", "")
        
        # Determine field type and extract name
        field_type = FieldType.UNKNOWN
        display_name = native_query_ref or query_ref
        metadata_key = query_ref
        
        if "Measure" in field:
            field_type = FieldType.MEASURE
            measure_info = field["Measure"]
            property_name = measure_info.get("Property", "")
            display_name = property_name or display_name
            
        elif "HierarchyLevel" in field:
            field_type = FieldType.CATEGORICAL
            hierarchy_info = field["HierarchyLevel"]
            level = hierarchy_info.get("Level", "")
            display_name = level or display_name
            
        elif "Column" in field:
            field_type = FieldType.CATEGORICAL
            column_info = field["Column"]
            property_name = column_info.get("Property", "")
            display_name = property_name or display_name
        
        # Get current width
        current_width = current_widths.get(metadata_key)
        
        return FieldInfo(
            name=metadata_key,
            display_name=display_name,
            field_type=field_type,
            metadata_key=metadata_key,
            current_width=current_width
        )
    
    def apply_scale_configuration(self, scale_config: ScaleConfiguration) -> None:
        """Apply scale configuration to all measure fields"""
        for visual_info in self.visuals_info:
            for field in visual_info.fields:
                if field.field_type == FieldType.MEASURE:
                    field.scale_config = scale_config
    
    def _calculate_auto_fit_width(self, text: str, font_info: FontInfo, max_width: int, min_width: int, 
                                 field_type: FieldType = FieldType.UNKNOWN, 
                                 visual_type: VisualType = VisualType.UNKNOWN,
                                 scale_config: Optional[ScaleConfiguration] = None,
                                 layout_type: Optional[str] = None,
                                 hierarchy_levels: int = 0) -> float:
        """Calculate optimal width for text with enhanced content-aware logic"""
        char_width = font_info.get_char_width()
        text_length = len(text)
        
        # Content-aware analysis
        content_type = self._analyze_content_type(text)
        
        # Calculate base width for single line
        single_line_width = text_length * char_width
        
        # Apply content-specific adjustments
        adjusted_width = self._apply_content_adjustments(single_line_width, content_type, text, field_type, visual_type)
        
        # Apply scale-aware adjustments for measures
        if field_type == FieldType.MEASURE and scale_config:
            adjusted_width = self._apply_scale_adjustments(adjusted_width, scale_config, visual_type)
        
        # Apply visual-type specific adjustments (ENHANCED)
        if visual_type == VisualType.MATRIX:
            adjusted_width = self._apply_matrix_adjustments(adjusted_width, field_type, content_type, layout_type, hierarchy_levels)
        
        # Add base padding (approximately 20px total)
        padding = 20
        total_width = adjusted_width + padding
        
        # Smart minimum width based on content and field type
        smart_min_width = self._calculate_smart_minimum(text, field_type, content_type, visual_type, layout_type, hierarchy_levels)
        effective_min_width = max(min_width, smart_min_width)
        
        # If it fits in one line, use that width
        if total_width <= max_width:
            return max(effective_min_width, total_width)
        
        # Calculate width for 3 lines max
        chars_per_line = max_width / char_width
        max_chars_for_3_lines = chars_per_line * 3
        
        if text_length <= max_chars_for_3_lines:
            # Text will fit in 3 lines or less
            return max(effective_min_width, max_width)
    
    def _calculate_fit_to_totals_width(self, field: FieldInfo, visual_info: VisualInfo, font_info: FontInfo, 
                                      max_width: int, min_width: int) -> float:
        """Calculate optimal width for measure columns to accommodate totals/subtotals"""
        char_width = font_info.get_char_width()
        
        # Analyze the field name for total-related patterns
        field_name_lower = field.display_name.lower()
        
        # Estimate detail value characteristics
        detail_length = len(field.display_name)  # Use field name as proxy for typical values
        
        # Base estimation for total magnitude
        total_multiplier = 1.0
        
        # Estimate based on visual type and hierarchy levels
        if visual_info.visual_type == VisualType.MATRIX:
            hierarchy_levels = self._detect_hierarchy_levels(visual_info)
            
            # More hierarchy levels = more data aggregation = larger totals
            if hierarchy_levels >= 4:  # Year > Quarter > Month > Week
                total_multiplier = 3.5  # Totals ~3.5x longer than details
            elif hierarchy_levels >= 3:  # Year > Quarter > Month  
                total_multiplier = 2.8  # Totals ~2.8x longer
            elif hierarchy_levels >= 2:  # Year > Quarter
                total_multiplier = 2.2  # Totals ~2.2x longer
            else:
                total_multiplier = 1.8  # Basic aggregation
        else:
            # Table visuals typically have simpler totals
            total_multiplier = 2.0
        
        # Adjust based on content type patterns
        content_type = self._analyze_content_type(field.display_name)
        
        if content_type == "currency":
            # Currency totals often have more digits + formatting
            total_multiplier += 0.5
            base_padding = 25  # Extra space for $ and commas
        elif content_type == "large_number":
            # Large numbers get comma separators
            total_multiplier += 0.3
            base_padding = 20
        elif content_type == "percentage":
            # Percentages usually don't get much larger in totals
            total_multiplier = min(total_multiplier, 1.5)
            base_padding = 15
        else:
            base_padding = 20
        
        # Common total value patterns and their typical lengths
        total_length_estimates = {
            # Assuming $861,060 detail values
            "currency_millions": 12,    # "$19,932,899"
            "currency_thousands": 10,   # "$9,336,498"
            "number_millions": 10,      # "19,932,899"
            "number_thousands": 8,      # "9,336,498"
            "percentage": 6,            # "100.0%"
        }
        
        # Estimate total display length
        if content_type == "currency":
            if total_multiplier > 3.0:
                estimated_total_length = total_length_estimates["currency_millions"]
            else:
                estimated_total_length = total_length_estimates["currency_thousands"]
        elif content_type == "percentage":
            estimated_total_length = total_length_estimates["percentage"]
        else:
            if total_multiplier > 3.0:
                estimated_total_length = total_length_estimates["number_millions"]
            else:
                estimated_total_length = total_length_estimates["number_thousands"]
        
        # Calculate width based on estimated total length
        total_width = estimated_total_length * char_width + base_padding
        
        # Apply matrix-specific adjustments
        if visual_info.visual_type == VisualType.MATRIX:
            hierarchy_levels = self._detect_hierarchy_levels(visual_info)
            total_width = self._apply_matrix_adjustments(
                total_width, FieldType.MEASURE, content_type, 
                visual_info.layout_type, hierarchy_levels
            )
        
        # Ensure reasonable bounds
        smart_min = self._calculate_smart_minimum(
            field.display_name, FieldType.MEASURE, content_type, 
            visual_info.visual_type, visual_info.layout_type, 
            self._detect_hierarchy_levels(visual_info) if visual_info.visual_type == VisualType.MATRIX else 0
        )
        
        # Add extra buffer for totals (they're critical to display properly)
        total_buffer = 15  # Increased from 10 to 15 for better wrap prevention
        final_width = total_width + total_buffer
        
        return max(smart_min, min(final_width, max_width))
        
        # For very long text, use max width but warn about potential wrapping
        return max(effective_min_width, max_width)
    
    def _analyze_content_type(self, text: str) -> str:
        """Analyze content type for enhanced width calculation"""
        text_lower = text.lower()
        
        # Date patterns (ENHANCED)
        if re.search(r'date|time|\d{1,2}/\d{1,2}/\d{2,4}|\d{4}-\d{2}-\d{2}', text_lower):
            # Check for complex date formats (ENHANCED)
            if re.search(r'eow:|end of week|bow:|beginning of week', text_lower):
                return "complex_date"
            return "date"
        
        # Currency patterns
        if re.search(r'\$|usd|currency|amount|revenue|cost|price', text_lower):
            return "currency"
        
        # Percentage patterns
        if re.search(r'%|percent|rate|ratio', text_lower):
            return "percentage"
        
        # Large number patterns
        if re.search(r'total|sum|count|quantity|volume', text_lower):
            return "large_number"
        
        # Short categorical values
        if len(text) <= 4 and not re.search(r'\d', text):
            return "short_categorical"
        
        # Hierarchical levels (Year, Quarter, etc.) (ENHANCED)
        if re.search(r'^(year|quarter|month|week|day)$', text_lower):
            return "hierarchy_parent"
        elif re.search(r'^(q[1-4]|jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)$', text_lower):
            return "hierarchy_child"
        
        return "standard"
    
    def _apply_content_adjustments(self, base_width: float, content_type: str, text: str, 
                                 field_type: FieldType, visual_type: VisualType) -> float:
        """Apply content-specific width adjustments (ENHANCED)"""
        adjusted_width = base_width
        
        if content_type == "complex_date":
            # ENHANCED: Add extra padding for complex date formats like "EOW: 01/07/24"
            adjusted_width += 35  # Increased from 25
        elif content_type == "date":
            # Standard date padding
            adjusted_width += 15
        elif content_type == "currency":
            # ENHANCED: Matrix-specific currency handling
            if visual_type == VisualType.MATRIX:
                adjusted_width += 20  # Increased from 15 for matrices
            else:
                adjusted_width += 15
            # Extra space for large currency values
            if len(text) > 8:
                adjusted_width += 10
        elif content_type == "percentage":
            # Add space for % symbol
            adjusted_width += 10
        elif content_type == "large_number":
            # Add space for comma separators in large numbers
            if len(text) > 6:
                adjusted_width += 15
        elif content_type == "short_categorical":
            # Ensure minimum readable width for short values
            adjusted_width = max(adjusted_width, 45)
        elif content_type == "hierarchy_parent":
            # ENHANCED: Give parent hierarchy levels more width
            adjusted_width += 15  # Increased from 10
        elif content_type == "hierarchy_child":
            # NEW: Child hierarchy levels get moderate extra width
            adjusted_width += 8
        
        return adjusted_width
    
    def _apply_scale_adjustments(self, width: float, scale_config: ScaleConfiguration, 
                               visual_type: VisualType) -> float:
        """Apply scale-aware adjustments for measure columns"""
        # Get expected display length for this scale
        expected_length = scale_config.get_sample_display_length()
        
        # Calculate character-based width requirement
        char_width = 7.0  # Average character width
        scale_based_width = expected_length * char_width
        
        # Use the larger of content-based or scale-based width
        adjusted_width = max(width, scale_based_width)
        
        # Add extra padding for different magnitude levels
        magnitude_padding = {
            DataMagnitude.SMALL: 5,
            DataMagnitude.MEDIUM: 10,
            DataMagnitude.LARGE: 15,
            DataMagnitude.XLARGE: 20
        }
        
        adjusted_width += magnitude_padding.get(scale_config.magnitude, 10)
        
        return adjusted_width
    
    def _apply_matrix_adjustments(self, width: float, field_type: FieldType, content_type: str, layout_type: Optional[str] = None, hierarchy_levels: int = 0) -> float:
        """Apply matrix-specific width adjustments (ENHANCED for Compact Matrix with level detection)"""
        if field_type == FieldType.CATEGORICAL:
            # COMPACT MATRIX: Intelligent width based on hierarchy levels
            if layout_type == "Compact":
                # Base adjustment for compact matrices
                base_adjustment = 1.4  # 40% wider baseline
                
                # Add width per hierarchy level (invisible and intelligent)
                level_bonus = min(hierarchy_levels * 0.3, 1.5)  # 30% per level, max 150% bonus
                base_adjustment += level_bonus
                
                # Content-specific adjustments
                if content_type == "hierarchy_parent":
                    base_adjustment += 0.2  # Additional 20% for parent levels
                elif content_type == "hierarchy_child":
                    base_adjustment += 0.1  # Additional 10% for child levels
                elif content_type in ["date", "complex_date"]:
                    base_adjustment += 0.3  # Additional 30% for date fields
                
                # Cap the maximum adjustment to prevent excessive width
                base_adjustment = min(base_adjustment, 3.0)  # Max 300% of original
                
            elif layout_type == "Tabular":
                # Tabular matrices need moderate adjustment
                base_adjustment = 1.2  # 20% wider baseline
                
                # Smaller level bonus for tabular
                level_bonus = min(hierarchy_levels * 0.15, 0.6)  # 15% per level, max 60% bonus
                base_adjustment += level_bonus
                
            else:
                # Outline matrix (original logic)
                base_adjustment = 0.95
                
                if content_type == "hierarchy_parent":
                    base_adjustment = 1.0  # No compression for parent levels
                elif content_type == "hierarchy_child":
                    base_adjustment = 0.92
            
            return width * base_adjustment
        else:
            # MEASURE COLUMNS: Enhanced for subtotals/totals
            base_adjustment = 1.05
            
            # Add extra space for measures to accommodate subtotals better
            if layout_type == "Compact":
                # Compact matrices often have more complex totals
                base_adjustment = 1.15  # 15% wider for measures in compact matrices
                
                # Small bonus for complex hierarchies (more subtotal levels)
                if hierarchy_levels > 3:
                    base_adjustment += 0.05  # Additional 5% for complex hierarchies
                    
            elif layout_type == "Tabular":
                # Tabular matrices also need space for subtotals
                base_adjustment = 1.1  # 10% wider for tabular measures
            
            return width * base_adjustment
    
    def _calculate_smart_minimum(self, text: str, field_type: FieldType, content_type: str, 
                               visual_type: VisualType, layout_type: Optional[str] = None, hierarchy_levels: int = 0) -> float:
        """Calculate smart minimum width based on content analysis (ENHANCED for Compact Matrix with level detection)"""
        # Base minimums
        if field_type == FieldType.CATEGORICAL:
            if content_type == "short_categorical":
                # ENHANCED: Compact matrix-specific minimums with level bonus
                if visual_type == VisualType.MATRIX and layout_type == "Compact":
                    base_min = 75 + (hierarchy_levels * 15)  # 15px per level
                    return min(base_min, 150)  # Cap at 150px
                elif visual_type == VisualType.MATRIX and layout_type == "Tabular":
                    base_min = 65 + (hierarchy_levels * 8)  # 8px per level
                    return min(base_min, 120)  # Cap at 120px
                return 55 if visual_type == VisualType.MATRIX else 50
            elif content_type == "hierarchy_parent":
                # ENHANCED: Parent hierarchy levels with intelligent level detection
                if visual_type == VisualType.MATRIX and layout_type == "Compact":
                    base_min = 90 + (hierarchy_levels * 20)  # 20px per level for parents
                    return min(base_min, 200)  # Cap at 200px
                elif visual_type == VisualType.MATRIX and layout_type == "Tabular":
                    base_min = 70 + (hierarchy_levels * 10)  # 10px per level
                    return min(base_min, 140)  # Cap at 140px
                return 60 if visual_type == VisualType.MATRIX else 55
            elif content_type == "hierarchy_child":
                # NEW: Child hierarchy levels with level bonus
                if visual_type == VisualType.MATRIX and layout_type == "Compact":
                    base_min = 80 + (hierarchy_levels * 15)  # 15px per level for children
                    return min(base_min, 180)  # Cap at 180px
                elif visual_type == VisualType.MATRIX and layout_type == "Tabular":
                    base_min = 60 + (hierarchy_levels * 8)  # 8px per level
                    return min(base_min, 120)  # Cap at 120px
                return 50 if visual_type == VisualType.MATRIX else 45
            elif content_type in ["date", "complex_date"]:
                # ENHANCED: Date fields with level consideration
                if visual_type == VisualType.MATRIX and layout_type == "Compact":
                    base_min = 100 + (hierarchy_levels * 10)  # 10px per level for dates
                    return min(base_min, 160)  # Cap at 160px
                elif visual_type == VisualType.MATRIX and layout_type == "Tabular":
                    base_min = 85 + (hierarchy_levels * 5)  # 5px per level
                    return min(base_min, 120)  # Cap at 120px
                return 85 if visual_type == VisualType.MATRIX else 80
            else:
                # ENHANCED: Standard categorical with level consideration
                if visual_type == VisualType.MATRIX and layout_type == "Compact":
                    base_min = 85 + (hierarchy_levels * 12)  # 12px per level
                    return min(base_min, 160)  # Cap at 160px
                elif visual_type == VisualType.MATRIX and layout_type == "Tabular":
                    base_min = 70 + (hierarchy_levels * 6)  # 6px per level
                    return min(base_min, 120)  # Cap at 120px
                return 65 if visual_type == VisualType.MATRIX else 60
        else:  # Measures
            if content_type == "currency":
                # ENHANCED: Matrix currency measures with subtotal consideration
                if visual_type == VisualType.MATRIX and layout_type == "Compact":
                    base_min = 100 + (hierarchy_levels * 5)  # 5px per level for subtotals
                    return min(base_min, 130)  # Cap at 130px
                elif visual_type == VisualType.MATRIX and layout_type == "Tabular":
                    base_min = 90 + (hierarchy_levels * 3)  # 3px per level
                    return min(base_min, 110)  # Cap at 110px
                return 85 if visual_type == VisualType.MATRIX else 75
            elif content_type == "large_number":
                # ENHANCED: Large numbers with subtotal space
                if visual_type == VisualType.MATRIX and layout_type == "Compact":
                    base_min = 105 + (hierarchy_levels * 5)  # 5px per level
                    return min(base_min, 140)  # Cap at 140px
                elif visual_type == VisualType.MATRIX and layout_type == "Tabular":
                    base_min = 95 + (hierarchy_levels * 3)  # 3px per level
                    return min(base_min, 120)  # Cap at 120px
                return 90 if visual_type == VisualType.MATRIX else 80
            else:
                # ENHANCED: Standard measures with level consideration
                if visual_type == VisualType.MATRIX and layout_type == "Compact":
                    base_min = 90 + (hierarchy_levels * 4)  # 4px per level
                    return min(base_min, 120)  # Cap at 120px
                elif visual_type == VisualType.MATRIX and layout_type == "Tabular":
                    base_min = 80 + (hierarchy_levels * 2)  # 2px per level
                    return min(base_min, 100)  # Cap at 100px
                return 75 if visual_type == VisualType.MATRIX else 65
    
    def _apply_preset_matrix_adjustments(self, base_width: int, field: FieldInfo, layout_type: Optional[str] = None, hierarchy_levels: int = 0) -> int:
        """Apply matrix-specific adjustments to preset widths with level detection"""
        if field.field_type == FieldType.CATEGORICAL:
            content_type = self._analyze_content_type(field.display_name)
            
            # COMPACT MATRIX: Much wider categorical columns with level bonuses
            if layout_type == "Compact":
                base_adjustment = 1.4  # 40% wider baseline
                
                # Add level bonus
                level_bonus = min(hierarchy_levels * 0.2, 1.0)  # 20% per level, max 100% bonus
                base_adjustment += level_bonus
                
                if content_type == "hierarchy_parent":
                    # Parent levels need significant extra width
                    base_adjustment += 0.2  # Additional 20%
                elif content_type == "hierarchy_child":
                    # Child levels need moderate extra width
                    base_adjustment += 0.1  # Additional 10%
                    
                # Cap maximum adjustment
                base_adjustment = min(base_adjustment, 2.5)  # Max 250% of original
                
            elif layout_type == "Tabular":
                # Tabular matrices need moderate adjustment with level bonus
                base_adjustment = 1.2  # 20% wider baseline
                
                # Smaller level bonus for tabular
                level_bonus = min(hierarchy_levels * 0.1, 0.5)  # 10% per level, max 50% bonus
                base_adjustment += level_bonus
                
            else:
                # Outline matrix (less aggressive adjustment)
                if content_type == "hierarchy_parent":
                    base_adjustment = 1.0  # No compression for parent levels
                else:
                    base_adjustment = 0.95  # Slight compression for others
            
            return int(base_width * base_adjustment)
        else:
            # MEASURES: Enhanced for subtotals with level consideration
            if layout_type == "Compact":
                base_adjustment = 1.15  # 15% wider baseline
                
                # Small bonus for complex hierarchies (more subtotal levels)
                if hierarchy_levels > 3:
                    base_adjustment += 0.05  # Additional 5% for complex hierarchies
                    
            elif layout_type == "Tabular":
                base_adjustment = 1.1  # 10% wider for tabular measures
                
                # Smaller bonus for tabular
                if hierarchy_levels > 3:
                    base_adjustment += 0.03  # Additional 3% for complex hierarchies
            else:
                # Standard matrix measure adjustment
                base_adjustment = 1.05  # 5% wider
            
            return int(base_width * base_adjustment)
    
    def calculate_optimal_widths(self, visual_info: VisualInfo, config: WidthConfiguration) -> None:
        """Calculate optimal widths for all fields in a visual with enhanced logic"""
        # Detect hierarchy levels for intelligent matrix width calculation
        hierarchy_levels = self._detect_hierarchy_levels(visual_info) if visual_info.visual_type == VisualType.MATRIX else 0
        
        for field in visual_info.fields:
            if field.is_overridden:
                continue  # Skip if manually overridden
            
            # Determine which preset to use
            if field.field_type == FieldType.CATEGORICAL:
                preset = config.categorical_preset
                custom_width = config.categorical_custom
            else:
                preset = config.measure_preset
                custom_width = config.measure_custom
            
            # Calculate width based on preset
            if preset == WidthPreset.AUTO_FIT:
                field.suggested_width = self._calculate_auto_fit_width(
                    field.display_name, visual_info.font_info, config.max_width, config.min_width,
                    field.field_type, visual_info.visual_type, field.scale_config,
                    visual_info.layout_type, hierarchy_levels  # Pass hierarchy levels
                )
            elif preset == WidthPreset.FIT_TO_TOTALS:
                # NEW: Fit to Totals - only for measures
                if field.field_type == FieldType.MEASURE:
                    field.suggested_width = self._calculate_fit_to_totals_width(
                        field, visual_info, visual_info.font_info, config.max_width, config.min_width
                    )
                else:
                    # Fallback to auto-fit for categorical fields
                    field.suggested_width = self._calculate_auto_fit_width(
                        field.display_name, visual_info.font_info, config.max_width, config.min_width,
                        field.field_type, visual_info.visual_type, field.scale_config,
                        visual_info.layout_type, hierarchy_levels
                    )
            elif preset == WidthPreset.CUSTOM:
                field.suggested_width = max(config.min_width, min(custom_width, config.max_width))
            else:
                # Use preset width with visual-type awareness (ENHANCED)
                category = "categorical" if field.field_type == FieldType.CATEGORICAL else "measure"
                base_width = self.PRESET_WIDTHS[preset][category]
                
                # ENHANCED: Apply matrix-specific adjustments
                if visual_info.visual_type == VisualType.MATRIX:
                    base_width = self._apply_preset_matrix_adjustments(
                        base_width, field, visual_info.layout_type, hierarchy_levels  # Pass hierarchy levels
                    )
                
                field.suggested_width = max(config.min_width, min(base_width, config.max_width))
    
    def apply_width_changes(self, visual_ids: List[str], configs: Dict[str, WidthConfiguration] = None, global_config: WidthConfiguration = None) -> Dict[str, Any]:
        """Apply width changes to specified visuals with per-visual or global configuration"""
        results = {
            "success": True,
            "visuals_updated": 0,
            "fields_updated": 0,
            "errors": []
        }
        
        for visual_info in self.visuals_info:
            if visual_info.visual_id in visual_ids:
                try:
                    # Determine which configuration to use
                    if configs and visual_info.visual_id in configs:
                        # Use per-visual configuration
                        config = configs[visual_info.visual_id]
                    elif global_config:
                        # Use global configuration
                        config = global_config
                    else:
                        # Fallback to default configuration
                        config = WidthConfiguration()
                    
                    # Calculate optimal widths
                    self.calculate_optimal_widths(visual_info, config)
                    
                    # Apply changes to the visual file
                    updated_fields = self._update_visual_file(visual_info)
                    
                    results["visuals_updated"] += 1
                    results["fields_updated"] += updated_fields
                    
                except Exception as e:
                    results["errors"].append(f"Visual {visual_info.visual_name}: {str(e)}")
                    results["success"] = False
        
        return results
    
    def _update_visual_file(self, visual_info: VisualInfo) -> int:
        """Update the visual file with new column widths"""
        # Construct path to visual file
        visual_file = (self.report_dir / "definition" / "pages" / visual_info.page_id / 
                      "visuals" / visual_info.visual_id / "visual.json")
        
        if not visual_file.exists():
            raise ValueError(f"Visual file not found: {visual_file}")
        
        # Read current visual data
        with open(visual_file, 'r', encoding='utf-8') as f:
            visual_data = json.load(f)
        
        # Update column widths
        updated_count = self._update_column_widths(visual_data, visual_info.fields)
        
        # Write updated data back
        with open(visual_file, 'w', encoding='utf-8') as f:
            json.dump(visual_data, f, indent=2, ensure_ascii=False)
        
        return updated_count
    
    def _update_column_widths(self, visual_data: Dict[str, Any], fields: List[FieldInfo]) -> int:
        """Update column width objects in visual data and turn off auto-size"""
        visual_config = visual_data.setdefault("visual", {})
        objects = visual_config.setdefault("objects", {})
        column_widths = objects.setdefault("columnWidth", [])
        
        # Turn off auto-size columns to preserve our custom settings
        self._disable_auto_size_columns(objects)
        
        # Create a map of existing column widths
        existing_widths = {}
        for i, width_obj in enumerate(column_widths):
            selector = width_obj.get("selector", {})
            metadata = selector.get("metadata", "")
            if metadata:
                existing_widths[metadata] = i
        
        updated_count = 0
        
        for field in fields:
            if field.suggested_width is None:
                continue
            
            # Create the width value with 'D' suffix
            width_value = f"{field.suggested_width}D"
            
            width_obj = {
                "properties": {
                    "value": {
                        "expr": {
                            "Literal": {
                                "Value": width_value
                            }
                        }
                    }
                },
                "selector": {
                    "metadata": field.metadata_key
                }
            }
            
            if field.metadata_key in existing_widths:
                # Update existing width
                column_widths[existing_widths[field.metadata_key]] = width_obj
            else:
                # Add new width
                column_widths.append(width_obj)
            
            updated_count += 1
        
        return updated_count
    
    def _disable_auto_size_columns(self, objects: Dict[str, Any]) -> None:
        """Disable auto-size columns to preserve custom width settings"""
        # Turn off auto-size for column headers
        column_headers = objects.setdefault("columnHeaders", [])
        
        # Look for existing auto-size setting
        auto_size_obj = None
        for obj in column_headers:
            if "properties" in obj and "autoSizeColumnWidth" in obj["properties"]:
                auto_size_obj = obj
                break
        
        # Create or update auto-size setting
        if auto_size_obj is None:
            auto_size_obj = {
                "properties": {},
                "selector": {"metadata": ""}
            }
            column_headers.append(auto_size_obj)
        
        # Set auto-size to false
        auto_size_obj["properties"]["autoSizeColumnWidth"] = {
            "expr": {
                "Literal": {
                    "Value": "false"
                }
            }
        }
    
    def get_visual_summary(self) -> Dict[str, Any]:
        """Get summary of scanned visuals"""
        total_visuals = len(self.visuals_info)
        table_count = sum(1 for v in self.visuals_info if v.visual_type == VisualType.TABLE)
        matrix_count = sum(1 for v in self.visuals_info if v.visual_type == VisualType.MATRIX)
        
        total_fields = sum(len(v.fields) for v in self.visuals_info)
        categorical_fields = sum(1 for v in self.visuals_info for f in v.fields if f.field_type == FieldType.CATEGORICAL)
        measure_fields = sum(1 for v in self.visuals_info for f in v.fields if f.field_type == FieldType.MEASURE)
        
        return {
            "total_visuals": total_visuals,
            "table_count": table_count,
            "matrix_count": matrix_count,
            "total_fields": total_fields,
            "categorical_fields": categorical_fields,
            "measure_fields": measure_fields,
            "pages": list(set(v.page_name for v in self.visuals_info))
        }
    
    def create_backup(self) -> str:
        """Create a backup of the original PBIP file"""
        import shutil
        from datetime import datetime
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_path = self.pbip_path.parent / f"{self.pbip_path.stem}_backup_{timestamp}.pbip"
        
        shutil.copy2(self.pbip_path, backup_path)
        return str(backup_path)
    
    def get_scale_recommendations(self) -> Dict[str, Any]:
        """Get scale recommendations based on scanned measure fields"""
        measure_fields = []
        for visual_info in self.visuals_info:
            for field in visual_info.fields:
                if field.field_type == FieldType.MEASURE:
                    measure_fields.append(field.display_name)
        
        # Analyze field names for scale clues
        scale_hints = {
            DataScale.THOUSANDS: ["revenue", "sales", "amount", "value", "cost"],
            DataScale.MILLIONS: ["total", "sum", "aggregate", "volume"],
            DataScale.ONES: ["count", "quantity", "number", "index", "id"],
            DataScale.BILLIONS: ["market", "global", "enterprise", "company"]
        }
        
        recommended_scale = DataScale.THOUSANDS  # Default
        confidence = 0.5
        
        for scale, keywords in scale_hints.items():
            matches = sum(1 for field in measure_fields 
                         for keyword in keywords 
                         if keyword.lower() in field.lower())
            
            field_confidence = matches / max(len(measure_fields), 1)
            if field_confidence > confidence:
                recommended_scale = scale
                confidence = field_confidence
        
        return {
            "recommended_scale": recommended_scale,
            "confidence": confidence,
            "measure_fields_count": len(measure_fields),
            "sample_fields": measure_fields[:5]  # First 5 for preview
        }
