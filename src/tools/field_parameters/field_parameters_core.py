"""
Field Parameters Core Logic
Built by Reid Havens of Analytic Endeavors

Handles parsing existing field parameters and generating TMDL code.
"""

import re
from typing import List, Dict, Optional, Tuple, Any
from dataclasses import dataclass, field
from pathlib import Path


@dataclass(slots=True)
class FieldItem:
    """Represents a single field in a field parameter"""
    display_name: str
    field_reference: str  # NAMEOF('Table'[Field])
    table_name: str  # Extracted from NAMEOF
    field_name: str  # Extracted from NAMEOF
    order_within_group: int = 1
    original_order_within_group: int = 1  # Stores original value for custom sort preservation
    categories: List[Tuple[int, str]] = field(default_factory=list)  # [(sort, name), ...]

    def __post_init__(self):
        """Extract table and field names from NAMEOF reference"""
        if not self.table_name or not self.field_name:
            # Parse NAMEOF('Table'[Field])
            match = re.search(r"NAMEOF\('([^']+)'\[([^\]]+)\]\)", self.field_reference)
            if match:
                self.table_name = match.group(1)
                self.field_name = match.group(2)


@dataclass(slots=True)
class CategoryLevel:
    """Represents a category column in the field parameter"""
    name: str  # Display name e.g., "Geography", "Department"
    sort_order: int  # Order of this column (1st, 2nd, 3rd category column)
    column_name: str  # TMDL column name e.g., "Pipeline Category"
    labels: List[str] = field(default_factory=list)  # Allowed values e.g., ["Financial", "Operational"]
    is_calculated: bool = False  # True if this is a DAX calculated column (read-only)


@dataclass
class FieldParameter:
    """Complete field parameter structure"""
    table_name: str
    parameter_name: str  # Display column name (e.g., "Pipeline Columns")
    fields: List[FieldItem] = field(default_factory=list)
    category_levels: List[CategoryLevel] = field(default_factory=list)
    keep_lineage_tags: bool = True
    lineage_tags: Dict[str, str] = field(default_factory=dict)  # column_name -> lineage_tag
    pbi_id: Optional[str] = None
    extra_columns: List[Dict[str, Any]] = field(default_factory=list)  # Non-standard columns
    # Tuple format: "sort_first" = (1, "Label"), "label_first" = ("Label", 1)
    # Default to "sort_first" as that's Power BI's native format
    category_tuple_format: str = "sort_first"
    # Output formatting options
    group_by_category: bool = False  # Add blank lines between category groups
    group_by_category_index: int = 0  # Which category column to group by (0 = first)
    update_field_order: bool = True  # When True, use current positions; False, use original order values
    show_unassigned_as_blank: bool = False  # When True, show "" instead of "Uncategorized" in output


class FieldParameterParser:
    """Parse TMDL field parameter definitions"""
    
    @staticmethod
    def parse_tmdl(tmdl_content: str) -> FieldParameter:
        """Parse TMDL content and extract field parameter structure"""
        
        # Extract table name - handles quoted and unquoted names
        # Matches: table 'Name With Spaces' or table "Name" or table SimpleName
        table_match = re.search(r"^(?:createOrReplace\s+)?table\s+(?:'([^']+)'|\"([^\"]+)\"|(\S+))", 
                                tmdl_content, re.MULTILINE)
        if not table_match:
            raise ValueError("Could not find table name in TMDL")
        
        # Get whichever group matched (single-quoted, double-quoted, or unquoted)
        table_name = table_match.group(1) or table_match.group(2) or table_match.group(3)
        
        # Find parameter name (the main display column)
        param_name_match = re.search(r'column\s+[\'"]?([^\'"\n]+)[\'"]?\s+.*?relatedColumnDetails\s+groupByColumn',
                                      tmdl_content, re.DOTALL)
        if not param_name_match:
            raise ValueError("Could not find parameter name")
        
        parameter_name = param_name_match.group(1).strip().strip("'\"")
        
        # Extract lineage tags
        lineage_tags = {}
        table_lineage = re.search(r'^table.*?lineageTag:\s*([a-f0-9\-]+)', tmdl_content, re.MULTILINE | re.DOTALL)
        if table_lineage:
            lineage_tags['_table'] = table_lineage.group(1)
        
        column_lineage_matches = re.finditer(
            r'column\s+[\'"]?([^\'"\n]+)[\'"]?\s+.*?lineageTag:\s*([a-f0-9\-]+)',
            tmdl_content,
            re.DOTALL
        )
        for match in column_lineage_matches:
            col_name = match.group(1).strip().strip("'\"")
            lineage_tags[col_name] = match.group(2)
        
        # Extract PBI_Id
        pbi_id = None
        pbi_match = re.search(r'annotation\s+PBI_Id\s*=\s*([a-f0-9]+)', tmdl_content)
        if pbi_match:
            pbi_id = pbi_match.group(1)
        
        # Parse partition source to get fields
        source_match = re.search(r'source\s*=\s*\{(.*?)\}', tmdl_content, re.DOTALL)
        if not source_match:
            raise ValueError("Could not find partition source")
        
        source_content = source_match.group(1)
        
        # Parse tuples - handle multi-line with comments
        # Format: ("Display", NAMEOF('Table'[Field]), order, cat1_sort, "Cat1", cat2_sort, "Cat2", ...)
        tuple_pattern = r'\("([^"]+)",\s*NAMEOF\(\'([^\']+)\'\[([^\]]+)\]\),\s*(\d+)(?:,\s*(\d+),\s*"([^"]+)")*(?:,\s*(\d+),\s*"([^"]+)")*\)'
        
        fields = []
        category_names_found = set()
        
        for line in source_content.split('\n'):
            # Skip comments
            if '--' in line:
                continue
            
            # Find tuples
            for match in re.finditer(tuple_pattern, line):
                display_name = match.group(1)
                table_name_ref = match.group(2)
                field_name_ref = match.group(3)
                order = int(match.group(4))
                
                # Extract categories (pairs of sort, name)
                categories = []
                remaining = line[match.end():]
                cat_matches = re.findall(r',\s*(\d+),\s*"([^"]+)"', remaining)
                for cat_sort, cat_name in cat_matches:
                    categories.append((int(cat_sort), cat_name))
                    category_names_found.add(cat_name)
                
                field_item = FieldItem(
                    display_name=display_name,
                    field_reference=f"NAMEOF('{table_name_ref}'[{field_name_ref}])",
                    table_name=table_name_ref,
                    field_name=field_name_ref,
                    order_within_group=order,
                    original_order_within_group=order,  # Preserve original for custom sort
                    categories=categories
                )
                fields.append(field_item)
        
        # Detect category levels from columns
        category_levels = []
        category_col_pattern = r'column\s+[\'"]?([^\'"\n]+)[\'"]?\s+.*?sourceColumn:\s*\[Value(\d+)\]'
        
        # Find which Value index corresponds to categories
        # Standard: Value1=Display, Value2=Fields, Value3=Order, Value4+=Categories
        for match in re.finditer(category_col_pattern, tmdl_content, re.DOTALL):
            col_name = match.group(1).strip().strip("'\"")
            value_idx = int(match.group(2))
            
            # Category columns are Value4, Value6, Value8, etc. (sort columns)
            # Display categories are Value5, Value7, Value9, etc.
            if value_idx >= 5 and value_idx % 2 == 1:  # Odd numbers 5, 7, 9...
                # This is a category display column
                category_levels.append(CategoryLevel(
                    name="",  # Will be filled from actual category names
                    sort_order=0,  # Will be determined from data
                    column_name=col_name
                ))
        
        # Detect extra columns (not part of standard parameter structure)
        extra_columns = []
        # Standard columns: Display, Fields, Order, [Category Sort, Category Name] pairs
        standard_prefixes = [parameter_name, f"{parameter_name} Fields", f"{parameter_name} Order"]
        
        all_columns = re.findall(r'column\s+([^\n]+)', tmdl_content)
        for col_def in all_columns:
            col_name = col_def.strip().strip("'\"").split()[0]
            is_standard = any(col_name.startswith(prefix) for prefix in standard_prefixes)
            is_standard = is_standard or "Sort" in col_name or col_name in [cl.column_name for cl in category_levels]
            
            if not is_standard and "partition" not in col_def.lower():
                # This is an extra column
                extra_columns.append({"definition": col_def})
        
        return FieldParameter(
            table_name=table_name,
            parameter_name=parameter_name,
            fields=fields,
            category_levels=category_levels,
            keep_lineage_tags=True,
            lineage_tags=lineage_tags,
            pbi_id=pbi_id,
            extra_columns=extra_columns
        )


class FieldParameterGenerator:
    """Generate TMDL code for field parameters"""
    
    @staticmethod
    def generate_tmdl(param: FieldParameter) -> str:
        """Generate complete TMDL definition"""
        
        lines = []
        
        # Header
        lines.append("createOrReplace\n")
        lines.append(f"\ttable '{param.table_name}'")
        
        # Table lineage tag
        if param.keep_lineage_tags and '_table' in param.lineage_tags:
            lines.append(f"\t\tlineageTag: {param.lineage_tags['_table']}")
        lines.append("")
        
        # Display column
        lines.append(f"\t\tcolumn '{param.parameter_name}'")
        if param.keep_lineage_tags and param.parameter_name in param.lineage_tags:
            lines.append(f"\t\t\tlineageTag: {param.lineage_tags[param.parameter_name]}")
        lines.append("\t\t\tsummarizeBy: none")
        lines.append("\t\t\tsourceColumn: [Value1]")
        lines.append(f"\t\t\tsortByColumn: '{param.parameter_name} Order'")
        lines.append("")
        lines.append("\t\t\trelatedColumnDetails")
        lines.append(f"\t\t\t\tgroupByColumn: '{param.parameter_name} Fields'")
        lines.append("")
        lines.append("\t\t\tannotation SummarizationSetBy = Automatic")
        lines.append("")
        
        # Fields column
        fields_col = f"{param.parameter_name} Fields"
        lines.append(f"\t\tcolumn '{fields_col}'")
        lines.append("\t\t\tisHidden")
        if param.keep_lineage_tags and fields_col in param.lineage_tags:
            lines.append(f"\t\t\tlineageTag: {param.lineage_tags[fields_col]}")
        lines.append("\t\t\tsummarizeBy: none")
        lines.append("\t\t\tsourceColumn: [Value2]")
        lines.append(f"\t\t\tsortByColumn: '{param.parameter_name} Order'")
        lines.append("")
        lines.append('\t\t\textendedProperty ParameterMetadata = {"version":3,"kind":2}')
        lines.append("")
        lines.append("\t\t\tannotation SummarizationSetBy = Automatic")
        lines.append("")
        
        # Order column
        order_col = f"{param.parameter_name} Order"
        lines.append(f"\t\tcolumn '{order_col}'")
        lines.append("\t\t\tisHidden")
        lines.append("\t\t\tformatString: 0")
        if param.keep_lineage_tags and order_col in param.lineage_tags:
            lines.append(f"\t\t\tlineageTag: {param.lineage_tags[order_col]}")
        lines.append("\t\t\tsummarizeBy: sum")
        lines.append("\t\t\tsourceColumn: [Value3]")
        lines.append("")
        lines.append("\t\t\tannotation SummarizationSetBy = Automatic")
        lines.append("")
        
        # Category columns (if any)
        value_idx = 4
        for level_idx, cat_level in enumerate(param.category_levels):
            # Category sort column
            sort_col = f"{cat_level.column_name} Sort"
            lines.append(f"\t\tcolumn '{sort_col}'")
            lines.append("\t\t\tisHidden")
            lines.append("\t\t\tformatString: 0")
            if param.keep_lineage_tags and sort_col in param.lineage_tags:
                lines.append(f"\t\t\tlineageTag: {param.lineage_tags[sort_col]}")
            lines.append("\t\t\tsummarizeBy: sum")
            lines.append(f"\t\t\tsourceColumn: [Value{value_idx}]")
            lines.append("")
            lines.append("\t\t\tannotation SummarizationSetBy = Automatic")
            lines.append("")
            value_idx += 1
            
            # Category display column
            lines.append(f"\t\tcolumn '{cat_level.column_name}'")
            if param.keep_lineage_tags and cat_level.column_name in param.lineage_tags:
                lines.append(f"\t\t\tlineageTag: {param.lineage_tags[cat_level.column_name]}")
            lines.append("\t\t\tsummarizeBy: none")
            lines.append(f"\t\t\tsourceColumn: [Value{value_idx}]")
            lines.append(f"\t\t\tsortByColumn: '{sort_col}'")
            lines.append("")
            lines.append("\t\t\tannotation SummarizationSetBy = Automatic")
            lines.append("")
            value_idx += 1
        
        # Extra columns (preserve as-is)
        for extra_col in param.extra_columns:
            lines.append(f"\t\t{extra_col['definition']}")
            lines.append("")
        
        # Partition source
        lines.append(f"\t\tpartition '{param.table_name}' = calculated")
        lines.append("\t\t\tmode: import")
        lines.append("\t\t\tsource = ```")
        lines.append("\t\t\t\t{")

        # Helper to build tuple for a field item
        def build_tuple(field_item: FieldItem) -> str:
            # Use original order if update_field_order is False (preserve custom sort)
            order_value = field_item.order_within_group if param.update_field_order else field_item.original_order_within_group
            tuple_parts = [
                f'"{field_item.display_name}"',
                field_item.field_reference,
                str(order_value)
            ]

            # Add category values - always include all category columns for tuple consistency
            for cat_entry in field_item.categories:
                c_sort, c_label = cat_entry
                # Handle show_unassigned_as_blank toggle
                # Empty label = uncategorized field
                if param.show_unassigned_as_blank and not c_label:
                    output_label = ""
                else:
                    output_label = c_label if c_label else "Uncategorized"
                # Format sort value (use int if whole number, else float)
                sort_str = str(int(c_sort)) if c_sort == int(c_sort) else str(c_sort)
                if param.category_tuple_format == "label_first":
                    # Label first, then sort: "Label", 1
                    tuple_parts.append(f'"{output_label}"')
                    tuple_parts.append(sort_str)
                else:
                    # Sort first, then label: 1, "Label" (default, Power BI native)
                    tuple_parts.append(sort_str)
                    tuple_parts.append(f'"{output_label}"')

            return f"({', '.join(tuple_parts)})"

        # Output fields - with optional grouping by category
        total_fields = len(param.fields)

        if param.group_by_category and param.category_levels:
            # Group fields by the selected category column
            cat_idx = param.group_by_category_index
            fields_by_category = {}

            for field_item in param.fields:
                # Get category key at the specified index
                if field_item.categories and len(field_item.categories) > cat_idx:
                    cat_entry = field_item.categories[cat_idx]
                    if cat_entry[1]:  # label is not empty
                        cat_key = (cat_entry[0], cat_entry[1])
                    else:
                        cat_key = (999, "Uncategorized")
                else:
                    cat_key = (999, "Uncategorized")

                if cat_key not in fields_by_category:
                    fields_by_category[cat_key] = []
                fields_by_category[cat_key].append(field_item)

            # Sort categories by sort order
            sorted_categories = sorted(fields_by_category.keys(), key=lambda x: x[0])

            field_counter = 0
            for cat_group_idx, (group_sort, group_name) in enumerate(sorted_categories):
                category_fields = fields_by_category[(group_sort, group_name)]

                for field_item in category_fields:
                    field_counter += 1
                    is_last = (field_counter == total_fields)
                    comma = "" if is_last else ","
                    lines.append(f"\t\t\t\t\t{build_tuple(field_item)}{comma}")

                # Blank line between category groups (except after last)
                if cat_group_idx < len(sorted_categories) - 1:
                    lines.append("")
        else:
            # Simple list - no grouping
            for field_idx, field_item in enumerate(param.fields):
                is_last = field_idx == len(param.fields) - 1
                comma = "" if is_last else ","
                lines.append(f"\t\t\t\t\t{build_tuple(field_item)}{comma}")

        lines.append("\t\t\t\t}")
        lines.append("\t\t\t\t```")
        lines.append("")

        # PBI_Id annotation
        if param.pbi_id:
            lines.append(f"\t\tannotation PBI_Id = {param.pbi_id}")
            lines.append("")

        # AE Multi-Tool attribution annotation
        lines.append("\t\tannotation AE_MultiTool = Edited with Analytic Endeavors PBI Multi-Tool")
        lines.append("")

        return "\n".join(lines)

    @staticmethod
    def generate_dax_expression(param: FieldParameter) -> str:
        """Generate just the DAX table expression for TOM writeback.

        This produces properly formatted DAX with leading newline so the { starts on line 2:

        {
            ("Display", NAMEOF('Table'[Field]), 1),
            ("Display2", NAMEOF('Table'[Field2]), 2)
        }
        """
        lines = []
        # Start with empty line so { appears on line 2 (after "TableName = ")
        lines.append("")
        lines.append("{")

        # Helper to build tuple for a field item
        def build_tuple(field_item: FieldItem) -> str:
            # Use original order if update_field_order is False (preserve custom sort)
            order_value = field_item.order_within_group if param.update_field_order else field_item.original_order_within_group
            tuple_parts = [
                f'"{field_item.display_name}"',
                field_item.field_reference,
                str(order_value)
            ]

            # Add category values - always include all category columns for tuple consistency
            for cat_entry in field_item.categories:
                c_sort, c_label = cat_entry
                # Handle show_unassigned_as_blank toggle
                # Empty label = uncategorized field
                if param.show_unassigned_as_blank and not c_label:
                    output_label = ""
                else:
                    output_label = c_label if c_label else "Uncategorized"
                # Format sort value (use int if whole number, else float)
                sort_str = str(int(c_sort)) if c_sort == int(c_sort) else str(c_sort)
                if param.category_tuple_format == "label_first":
                    tuple_parts.append(f'"{output_label}"')
                    tuple_parts.append(sort_str)
                else:
                    tuple_parts.append(sort_str)
                    tuple_parts.append(f'"{output_label}"')

            return f"({', '.join(tuple_parts)})"

        # Output fields - with optional grouping by category
        total_fields = len(param.fields)

        if param.group_by_category and param.category_levels:
            # Group fields by the selected category column
            cat_idx = param.group_by_category_index
            fields_by_category = {}

            for field_item in param.fields:
                if field_item.categories and len(field_item.categories) > cat_idx:
                    cat_entry = field_item.categories[cat_idx]
                    if cat_entry[1]:
                        cat_key = (cat_entry[0], cat_entry[1])
                    else:
                        cat_key = (999, "Uncategorized")
                else:
                    cat_key = (999, "Uncategorized")

                if cat_key not in fields_by_category:
                    fields_by_category[cat_key] = []
                fields_by_category[cat_key].append(field_item)

            sorted_categories = sorted(fields_by_category.keys(), key=lambda x: x[0])

            field_counter = 0
            for cat_group_idx, (group_sort, group_name) in enumerate(sorted_categories):
                category_fields = fields_by_category[(group_sort, group_name)]

                for field_item in category_fields:
                    field_counter += 1
                    is_last = (field_counter == total_fields)
                    comma = "" if is_last else ","
                    lines.append(f"\t{build_tuple(field_item)}{comma}")

                # Blank line between category groups (except after last)
                if cat_group_idx < len(sorted_categories) - 1:
                    lines.append("")
        else:
            # Simple list - no grouping
            for field_idx, field_item in enumerate(param.fields):
                is_last = field_idx == len(param.fields) - 1
                comma = "" if is_last else ","
                lines.append(f"\t{build_tuple(field_item)}{comma}")

        lines.append("}")

        return "\n".join(lines)
