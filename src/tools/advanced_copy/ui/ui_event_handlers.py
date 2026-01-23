"""
Event Handlers Mixin for Advanced Copy UI
Built by Reid Havens of Analytic Endeavors

Handles all event binding and UI state management.
"""

import tkinter as tk

from tools.advanced_copy.ui.ui_helpers import HelpersMixin


class EventHandlersMixin:
    """
    Mixin for event handlers and UI state management.
    
    Methods extracted from AdvancedCopyTab:
    - _setup_events()
    - _on_content_mode_change()
    - _on_destination_mode_change()
    - _on_path_change()
    - _update_ui_state()
    - _on_tree_click()
    - _on_bookmark_tree_select()
    - _on_target_listbox_click()
    - _on_target_listbox_select()
    - reset_tab()
    """
    
    def __init__(self, *args, **kwargs):
        """Initialize event handler state tracking"""
        super().__init__(*args, **kwargs)
        self._last_destination_mode = None  # Track last destination mode to detect actual changes
        self._last_content_mode = None  # Track last content mode to detect actual changes
    
    def _setup_events(self):
        """Setup event handlers"""
        self.report_path.trace('w', lambda *args: self._on_path_change())
        
        # Initialize mode tracking to match default selections
        self._last_content_mode = self.copy_content_mode.get()
        self._last_destination_mode = self.copy_destination_mode.get()
    
    def _on_content_mode_change(self):
        """Handle copy content mode changes"""
        current_mode = self.copy_content_mode.get()
        
        # Only process if the mode actually changed (not just clicking the same option)
        if current_mode == self._last_content_mode:
            return  # No change, do nothing
        
        # Update guide text based on new mode
        self._update_guide_text()
        
        # Reset analysis when content mode changes
        if self.analysis_results:
            self.analysis_results = None
            self._hide_page_selection_ui()
            content_mode = 'Full Page' if current_mode == 'full_page' else 'Bookmark + Visuals'
            self.log_message(f"Copy content changed to: {content_mode}")
            self.log_message("   Please re-analyze the report")
        
        # Update tracking variable
        self._last_content_mode = current_mode
    
    def _on_destination_mode_change(self):
        """Handle copy destination mode changes"""
        current_mode = self.copy_destination_mode.get()
        
        # Only process if the mode actually changed (not just clicking the same option)
        if current_mode == self._last_destination_mode:
            return  # No change, do nothing
        
        # Update guide text
        self._update_guide_text()
        
        # Show/hide target PBIP input based on destination
        if current_mode == "cross_pbip":
            self._show_target_pbip_input()
        else:
            self._hide_target_pbip_input()
        
        # Clear target analysis when switching modes (will need to re-analyze)
        if self.target_analysis_results:
            self.target_analysis_results = None
            self.log_message("   Target analysis cleared - please re-analyze if needed")
        
        # Log the change
        dest_mode = 'Same PBIP' if current_mode == 'same_pbip' else 'Cross-PBIP'
        self.log_message(f"Copy destination changed to: {dest_mode}")

        # Update tracking variable
        self._last_destination_mode = current_mode

        # Update analyze button state (cross-PBIP requires target path)
        self._update_ui_state()
    
    def _on_path_change(self):
        """Handle path changes"""
        self._update_ui_state()
    
    def _update_ui_state(self):
        """Update UI state"""
        has_source = bool(self.report_path.get())

        # In cross-PBIP destination mode, also need target file
        if self.copy_destination_mode.get() == "cross_pbip":
            has_target = bool(self.target_pbip_path.get())
            has_all_files = has_source and has_target
        else:
            has_all_files = has_source

        # Update analyze button state using set_enabled (RoundedButton)
        if self.analyze_button and hasattr(self.analyze_button, 'set_enabled'):
            self.analyze_button.set_enabled(has_all_files)
            self._analyze_button_enabled = has_all_files
        elif self.analyze_button:
            # Fallback for ttk.Button
            self.analyze_button.config(state=tk.NORMAL if has_all_files else tk.DISABLED)
    
    def _on_tree_click(self, event):
        """Handle tree click - arrow expands/collapses, group text selects children"""
        if self._updating_ui:
            return 'break'  # Stop event propagation
        
        # Get the item that was clicked
        item_id = self.bookmarks_treeview.identify('item', event.x, event.y)
        if not item_id or item_id not in self._bookmark_tree_mapping:
            return  # Let default behavior handle this
        
        # Check what part of the tree was clicked
        region = self.bookmarks_treeview.identify_region(event.x, event.y)
        element = self.bookmarks_treeview.identify_element(event.x, event.y)
        
        # If clicked on the tree expand/collapse indicator (arrow), allow default behavior
        if element in ('Treeitem.indicator', 'indicator'):
            return  # Let tkinter handle expand/collapse
        
        item_info = self._bookmark_tree_mapping[item_id]
        
        # Prevent selection of disabled bookmarks (only for direct bookmark clicks)
        if item_info['type'] == 'bookmark' and not item_info.get('selectable', True):
            # Show a brief message
            self.log_message("   ⚠️ Cannot select 'All visuals' bookmarks - only 'Selected visuals' supported")
            return 'break'  # Stop event propagation to prevent selection
        
        # If clicked on a group TEXT/ICON (not the arrow), select/deselect all its children
        if item_info['type'] == 'group':
            # Set flag to prevent recursive updates during this operation
            self._updating_ui = True
            try:
                children = self.bookmarks_treeview.get_children(item_id)
                
                if not children:
                    return 'break'  # No children to select
                
                # Get only selectable children
                selectable_children = []
                for child in children:
                    child_info = self._bookmark_tree_mapping.get(child, {})
                    if child_info.get('selectable', True):
                        selectable_children.append(child)
                
                if not selectable_children:
                    self.log_message("   ⚠️ No selectable bookmarks in this group")
                    return 'break'
                
                # Check if all selectable children are already selected
                current_selection = set(self.bookmarks_treeview.selection())
                selectable_children_set = set(selectable_children)
                
                # Clear current selection first
                self.bookmarks_treeview.selection_set(())
                
                if selectable_children_set.issubset(current_selection):
                    # All selectable children were selected, deselect them (already cleared)
                    pass
                else:
                    # Not all selectable children selected, select them all
                    self.bookmarks_treeview.selection_set(selectable_children)
                
            finally:
                self._updating_ui = False
            
            # Trigger update after a short delay to allow selection to complete
            self.frame.after(10, self._update_selection_status)
            return 'break'  # Prevent default tree behavior for group selection
    
    def _on_bookmark_tree_select(self, event=None):
        """Handle bookmark tree selection"""
        if self._updating_ui:
            return

        # Remove any disabled items from selection (they shouldn't be selectable)
        if hasattr(self, 'bookmarks_treeview') and self.bookmarks_treeview:
            current_selection = list(self.bookmarks_treeview.selection())
            for item_id in current_selection:
                # Check if item is marked as not selectable in the mapping
                if hasattr(self, '_bookmark_tree_mapping') and item_id in self._bookmark_tree_mapping:
                    item_info = self._bookmark_tree_mapping[item_id]
                    if not item_info.get('selectable', True):
                        self.bookmarks_treeview.selection_remove(item_id)

        # Cancel any pending update
        if self._pending_update:
            self.frame.after_cancel(self._pending_update)
        # Schedule new update with a small delay to allow selection to complete
        self._pending_update = self.frame.after(50, self._update_selection_status)
    
    def _on_target_listbox_click(self, event):
        """Handle target pages listbox clicks to detect Select All"""
        if self._updating_ui:
            return 'break'
        
        # Get the clicked index
        try:
            index = self.target_pages_listbox.nearest(event.y)
        except:
            return
        
        # Check if clicked on "Select All" (index 0)
        if index == 0:
            # Set flag to prevent recursive updates
            self._updating_ui = True
            try:
                # Get current selection
                current_selection = set(self.target_pages_listbox.curselection())
                
                # Get all indices except index 0 (the Select All item)
                all_indices = list(range(1, self.target_pages_listbox.size()))
                
                if not all_indices:
                    return 'break'
                
                # Clear selection first
                self.target_pages_listbox.selection_clear(0, tk.END)
                
                # Check if all pages are already selected
                all_selected = all(i in current_selection for i in all_indices)
                
                if all_selected:
                    # Deselect all (already cleared above)
                    pass
                else:
                    # Select all pages (except the Select All item itself)
                    for i in all_indices:
                        self.target_pages_listbox.selection_set(i)
            finally:
                self._updating_ui = False
            
            # Trigger update after a short delay
            self.frame.after(10, self._update_selection_status)
            return 'break'
    
    def _on_target_listbox_select(self, event=None):
        """Handle target pages listbox selection"""
        if self._updating_ui:
            return
        # Cancel any pending update
        if self._pending_update:
            self.frame.after_cancel(self._pending_update)
        # Schedule new update with a small delay to allow selection to complete
        self._pending_update = self.frame.after(50, self._update_selection_status)
    
    def reset_tab(self) -> None:
        """Reset the tab to initial state"""
        if self.is_busy:
            if not self.ask_yes_no("Confirm Reset", "An operation is in progress. Stop and reset?",
                                   icon_path=self._tool_icon_path):
                return

        # Clear state
        self.report_path.set("")
        self.target_pbip_path.set("")  # Clear target PBIP path
        self.analysis_results = None
        self.target_analysis_results = None  # Clear target analysis
        self.available_pages = []
        self.all_report_pages = []
        self._updating_ui = False  # Reset guard flag
        self._pending_update = None  # Reset pending update

        # Reset progress persistence
        self.progress_persist = False
        if self.progress_frame:
            self.progress_frame.pack_forget()

        # Reset button states using set_enabled (RoundedButton)
        if self.analyze_button and hasattr(self.analyze_button, 'set_enabled'):
            self.analyze_button.set_enabled(False)
            self._analyze_button_enabled = False
        elif self.analyze_button:
            self.analyze_button.config(state=tk.DISABLED)

        if self.copy_button and hasattr(self.copy_button, 'set_enabled'):
            self.copy_button.set_enabled(False)
            self._copy_button_enabled = False
        elif self.copy_button:
            self.copy_button.config(state=tk.DISABLED)

        if self.pages_listbox:
            self.pages_listbox.config(state=tk.NORMAL)
        self._hide_page_selection_ui()

        # Clear log and show welcome
        if hasattr(self, 'log_section') and self.log_section:
            log_text = self.log_section.log_text
            log_text.config(state=tk.NORMAL)
            log_text.delete(1.0, tk.END)
            log_text.config(state=tk.DISABLED)
        elif self.log_text:
            self.log_text.config(state=tk.NORMAL)
            self.log_text.delete(1.0, tk.END)
            self.log_text.config(state=tk.DISABLED)

        # Explicitly call HelpersMixin's method (BaseToolTab's version only has 2 lines)
        HelpersMixin._show_welcome_message(self)
        self.log_message("Advanced Copy reset successfully!")
