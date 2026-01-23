"""
AvailableFieldsPanel
Panel component for the Field Parameters tool.
Thin wrapper around the shared FieldNavigator component.

Built by Reid Havens of Analytic Endeavors
"""

import tkinter as tk
from tkinter import ttk
from typing import Dict, List, Optional, TYPE_CHECKING
import logging

from core.theme_manager import get_theme_manager
from core.widgets import FieldNavigator
from core.pbi_connector import FieldInfo

if TYPE_CHECKING:
    from tools.field_parameters.field_parameters_ui import FieldParametersTab


class AvailableFieldsPanel(tk.Frame):
    """
    Panel showing available measures and columns from the model.

    Wraps the shared FieldNavigator component and bridges its callbacks
    to the Field Parameters tool's main tab.

    Note: This is a plain Frame, not a LabelFrame, because FieldNavigator
    already handles its own section header styling.
    """

    def __init__(self, parent, main_tab: 'FieldParametersTab'):
        tk.Frame.__init__(self, parent)
        self.main_tab = main_tab
        self.logger = logging.getLogger(__name__)
        self._theme_manager = get_theme_manager()

        # Store table metadata for field lookups (used by duplicate checker)
        self.tables_data = {}  # table_name -> TableFieldsInfo

        self._setup_ui()

    def _setup_ui(self):
        """Setup the FieldNavigator component"""
        # Create FieldNavigator with appropriate callbacks
        # Note: drop_target set later via set_drop_target() after builder_panel exists
        self._navigator = FieldNavigator(
            parent=self,
            theme_manager=self._theme_manager,
            on_fields_selected=self._handle_fields_selected,
            drop_target=None,  # Set later via set_drop_target()
            section_title="Available Fields",
            section_icon="table",
            show_columns=True,
            show_add_button=True,
            duplicate_checker=self._is_field_duplicate,
            show_duplicate_dialogs=True,
            placeholder_text="No tables or fields loaded.\n\nConnect to a model to view\navailable tables and fields.",
            can_add_validator=self._validate_can_add_fields,
        )
        self._navigator.pack(fill=tk.BOTH, expand=True)

    def set_drop_target(self, drop_target):
        """
        Set the drop target for drag-and-drop operations.

        Call this after the target widget (e.g., builder_panel) is created.

        Args:
            drop_target: Widget implementing DropTargetProtocol
        """
        self._navigator.set_drop_target(drop_target)

    def _handle_fields_selected(self, fields: List[FieldInfo], position: Optional[int]):
        """
        Handle fields selected in the navigator.
        Bridges the generic callback to main_tab.on_add_field().

        Args:
            fields: List of FieldInfo objects to add
            position: Insert position (None = end, 0 = top, etc.)
        """
        for i, field_info in enumerate(fields):
            if position is not None:
                # Insert at specific position, incrementing for each field
                self.main_tab.on_add_field(
                    field_info.table_name,
                    field_info.name,
                    position=position + i
                )
            else:
                # Add to end
                self.main_tab.on_add_field(field_info.table_name, field_info.name)

    def _is_field_duplicate(self, table_name: str, field_name: str) -> bool:
        """
        Check if a field is already in the current parameter.
        Used as duplicate_checker callback for FieldNavigator.
        """
        if not hasattr(self.main_tab, 'current_parameter') or not self.main_tab.current_parameter:
            return False

        for existing in self.main_tab.current_parameter.fields:
            if existing.table_name == table_name and existing.field_name == field_name:
                return True
        return False

    def _validate_can_add_fields(self) -> bool:
        """
        Validate that we can add fields (parameter exists).
        Shows warning if validation fails.
        Returns True if OK to proceed, False otherwise.
        """
        # Use main_tab's validation method (shows warning once)
        if hasattr(self.main_tab, 'validate_can_add_field'):
            return self.main_tab.validate_can_add_field(show_warning=True)

        # Fallback check
        if hasattr(self.main_tab, 'current_parameter') and not self.main_tab.current_parameter:
            from core.ui_base import ThemedMessageBox
            ThemedMessageBox.showwarning(
                self,
                "No Parameter",
                "Please create or select a parameter first"
            )
            return False

        return True

    # =========================================================================
    # PUBLIC API - Methods called by main_tab and other panels
    # =========================================================================

    def update_available_fields(self, tables: Dict):
        """
        Update tree with available tables and fields from TableFieldsInfo structure.

        Args:
            tables: Dict mapping table_name -> TableFieldsInfo
        """
        # Store table data for duplicate checker lookups
        self.tables_data = tables

        # Delegate to navigator
        self._navigator.set_fields(tables)

    def set_enabled(self, enabled: bool):
        """Enable/disable panel controls"""
        self._navigator.set_enabled(enabled)

    def clear(self):
        """Clear panel state"""
        self.tables_data = {}
        self._navigator.clear()

    def on_theme_changed(self):
        """Update widget colors when theme changes"""
        self._navigator.on_theme_changed()
