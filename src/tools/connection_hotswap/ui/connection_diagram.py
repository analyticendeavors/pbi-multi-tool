"""
Connection Diagram - Visual representation of connection mappings
Built by Reid Havens of Analytic Endeavors

Canvas-based visualization showing connection flow from source to target.
"""

import tkinter as tk
from tkinter import ttk
from typing import List, Optional, Callable, Dict, Tuple
from dataclasses import dataclass
import math

from core.theme_manager import get_theme_manager
from tools.connection_hotswap.models import ConnectionMapping, SwapStatus


@dataclass
class NodePosition:
    """Position and dimensions of a node in the diagram."""
    x: int
    y: int
    width: int
    height: int
    canvas_id: int = 0
    text_id: int = 0


class ConnectionDiagram(tk.Canvas):
    """
    Canvas-based visualization of connection mappings.

    Shows source connections on the left, target connections on the right,
    with arrows indicating the mapping relationship.
    """

    # Layout constants
    NODE_WIDTH = 160
    NODE_HEIGHT = 40
    NODE_PADDING = 15
    ARROW_MARGIN = 20
    HORIZONTAL_GAP = 80

    def __init__(
        self,
        parent,
        on_node_click: Optional[Callable[[int], None]] = None,
        **kwargs
    ):
        """
        Initialize the connection diagram.

        Args:
            parent: Parent widget
            on_node_click: Callback when a node is clicked (receives mapping index)
        """
        self._theme_manager = get_theme_manager()
        colors = self._theme_manager.colors
        is_dark = self._theme_manager.is_dark

        bg_color = colors.get('surface', '#1e1e2e' if is_dark else '#ffffff')

        super().__init__(
            parent,
            bg=bg_color,
            highlightthickness=0,
            **kwargs
        )

        self.on_node_click = on_node_click
        self.mappings: List[ConnectionMapping] = []

        # Track node positions for hit detection
        self._source_nodes: Dict[int, NodePosition] = {}
        self._target_nodes: Dict[int, NodePosition] = {}
        self._selected_index: Optional[int] = None

        # Bind events
        self.bind("<Button-1>", self._on_click)
        self.bind("<Configure>", self._on_resize)

    def update_mappings(self, mappings: List[ConnectionMapping]):
        """Update the diagram with new mappings."""
        self.mappings = mappings
        self._draw_diagram()

    def set_selected(self, index: Optional[int]):
        """Set the selected mapping index."""
        self._selected_index = index
        self._draw_diagram()

    def _draw_diagram(self):
        """Redraw the entire diagram."""
        self.delete("all")
        self._source_nodes.clear()
        self._target_nodes.clear()

        if not self.mappings:
            self._draw_empty_message()
            return

        colors = self._theme_manager.colors
        is_dark = self._theme_manager.is_dark

        # Calculate canvas dimensions
        canvas_width = self.winfo_width()
        canvas_height = self.winfo_height()

        if canvas_width < 100 or canvas_height < 50:
            # Canvas not ready yet
            return

        # Calculate positions
        num_mappings = len(self.mappings)
        total_height = num_mappings * (self.NODE_HEIGHT + self.NODE_PADDING) - self.NODE_PADDING
        start_y = max(20, (canvas_height - total_height) // 2)

        # Source nodes on left, targets on right
        source_x = self.ARROW_MARGIN
        target_x = canvas_width - self.NODE_WIDTH - self.ARROW_MARGIN

        # Ensure minimum gap
        if target_x - source_x - self.NODE_WIDTH < self.HORIZONTAL_GAP:
            target_x = source_x + self.NODE_WIDTH + self.HORIZONTAL_GAP

        # Draw header labels
        self._draw_header(source_x, "Source Connections", colors)
        self._draw_header(target_x, "Target Connections", colors)

        # Draw each mapping
        y = start_y + 30  # Account for header

        for i, mapping in enumerate(self.mappings):
            is_selected = i == self._selected_index

            # Draw source node
            source_node = self._draw_node(
                source_x, y,
                mapping.source.display_name,
                self._get_source_color(mapping, colors, is_dark),
                is_selected,
                colors
            )
            self._source_nodes[i] = source_node

            # Draw target node
            target_name = mapping.target.display_name if mapping.target else "(not configured)"
            target_color = self._get_target_color(mapping, colors, is_dark)

            target_node = self._draw_node(
                target_x, y,
                target_name,
                target_color,
                is_selected,
                colors
            )
            self._target_nodes[i] = target_node

            # Draw arrow between nodes
            self._draw_arrow(
                source_x + self.NODE_WIDTH,
                y + self.NODE_HEIGHT // 2,
                target_x,
                y + self.NODE_HEIGHT // 2,
                mapping,
                colors,
                is_dark
            )

            # Draw status indicator
            self._draw_status_indicator(
                target_x + self.NODE_WIDTH + 10,
                y + self.NODE_HEIGHT // 2,
                mapping.status,
                colors
            )

            y += self.NODE_HEIGHT + self.NODE_PADDING

    def _draw_empty_message(self):
        """Draw message when no mappings exist."""
        colors = self._theme_manager.colors
        canvas_width = self.winfo_width()
        canvas_height = self.winfo_height()

        self.create_text(
            canvas_width // 2,
            canvas_height // 2,
            text="Connect to a model to see connection diagram",
            fill=colors.get('text_muted', '#6b7280'),
            font=("Segoe UI", 10, "italic")
        )

    def _draw_header(self, x: int, text: str, colors: dict):
        """Draw a section header."""
        self.create_text(
            x + self.NODE_WIDTH // 2,
            10,
            text=text,
            fill=colors.get('text_muted', '#6b7280'),
            font=("Segoe UI", 9, "bold"),
            anchor="n"
        )

    def _draw_node(
        self,
        x: int,
        y: int,
        text: str,
        fill_color: str,
        is_selected: bool,
        colors: dict
    ) -> NodePosition:
        """Draw a connection node."""
        is_dark = self._theme_manager.is_dark

        # Selection highlight
        border_color = colors.get('primary', '#4a6cf5') if is_selected else colors.get('border', '#3a3a4a')
        border_width = 3 if is_selected else 1

        # Draw rounded rectangle
        rect_id = self._draw_rounded_rect(
            x, y,
            x + self.NODE_WIDTH, y + self.NODE_HEIGHT,
            radius=8,
            fill=fill_color,
            outline=border_color,
            width=border_width
        )

        # Truncate text if too long
        display_text = text
        if len(display_text) > 20:
            display_text = display_text[:17] + "..."

        # Draw text
        text_id = self.create_text(
            x + self.NODE_WIDTH // 2,
            y + self.NODE_HEIGHT // 2,
            text=display_text,
            fill=colors.get('text_primary', '#e0e0e0' if is_dark else '#333333'),
            font=("Segoe UI", 9),
            anchor="center"
        )

        return NodePosition(x, y, self.NODE_WIDTH, self.NODE_HEIGHT, rect_id, text_id)

    def _draw_rounded_rect(
        self,
        x1: int, y1: int,
        x2: int, y2: int,
        radius: int = 10,
        **kwargs
    ) -> int:
        """Draw a rounded rectangle on the canvas."""
        points = [
            x1 + radius, y1,
            x2 - radius, y1,
            x2, y1,
            x2, y1 + radius,
            x2, y2 - radius,
            x2, y2,
            x2 - radius, y2,
            x1 + radius, y2,
            x1, y2,
            x1, y2 - radius,
            x1, y1 + radius,
            x1, y1,
        ]
        return self.create_polygon(points, smooth=True, **kwargs)

    def _draw_arrow(
        self,
        x1: int, y1: int,
        x2: int, y2: int,
        mapping: ConnectionMapping,
        colors: dict,
        is_dark: bool
    ):
        """Draw an arrow between source and target nodes."""
        # Arrow color based on status
        if mapping.status == SwapStatus.SUCCESS:
            arrow_color = colors.get('success', '#10b981')
        elif mapping.status == SwapStatus.ERROR:
            arrow_color = colors.get('error', '#ef4444')
        elif mapping.status in (SwapStatus.READY, SwapStatus.MATCHED):
            arrow_color = colors.get('primary', '#4a6cf5')
        else:
            arrow_color = colors.get('text_muted', '#6b7280')

        # Draw the line
        self.create_line(
            x1 + 5, y1,
            x2 - 5, y2,
            fill=arrow_color,
            width=2,
            arrow=tk.LAST,
            arrowshape=(10, 12, 5)
        )

    def _draw_status_indicator(
        self,
        x: int, y: int,
        status: SwapStatus,
        colors: dict
    ):
        """Draw a status indicator circle."""
        radius = 6

        # Color based on status
        fill_color = {
            SwapStatus.PENDING: colors.get('text_muted', '#6b7280'),
            SwapStatus.MATCHED: colors.get('warning', '#f59e0b'),
            SwapStatus.READY: colors.get('primary', '#4a6cf5'),
            SwapStatus.SWAPPING: colors.get('warning', '#f59e0b'),
            SwapStatus.SUCCESS: colors.get('success', '#10b981'),
            SwapStatus.ERROR: colors.get('error', '#ef4444'),
        }.get(status, colors.get('text_muted'))

        self.create_oval(
            x - radius, y - radius,
            x + radius, y + radius,
            fill=fill_color,
            outline=""
        )

    def _get_source_color(
        self,
        mapping: ConnectionMapping,
        colors: dict,
        is_dark: bool
    ) -> str:
        """Get the fill color for a source node."""
        if mapping.source.is_cloud:
            return '#2563eb' if is_dark else '#3b82f6'  # Blue for cloud
        else:
            return '#059669' if is_dark else '#10b981'  # Green for local

    def _get_target_color(
        self,
        mapping: ConnectionMapping,
        colors: dict,
        is_dark: bool
    ) -> str:
        """Get the fill color for a target node."""
        if not mapping.target:
            return colors.get('surface', '#2a2a3a' if is_dark else '#f5f5fa')

        if mapping.target.target_type == "cloud":
            return '#1d4ed8' if is_dark else '#60a5fa'  # Lighter blue
        else:
            return '#047857' if is_dark else '#34d399'  # Lighter green

    def _on_click(self, event):
        """Handle click events to detect node clicks."""
        # Check if click is on any source or target node
        for i, node in self._source_nodes.items():
            if self._point_in_node(event.x, event.y, node):
                self._selected_index = i
                if self.on_node_click:
                    self.on_node_click(i)
                self._draw_diagram()
                return

        for i, node in self._target_nodes.items():
            if self._point_in_node(event.x, event.y, node):
                self._selected_index = i
                if self.on_node_click:
                    self.on_node_click(i)
                self._draw_diagram()
                return

    def _point_in_node(self, px: int, py: int, node: NodePosition) -> bool:
        """Check if a point is within a node's bounds."""
        return (node.x <= px <= node.x + node.width and
                node.y <= py <= node.y + node.height)

    def _on_resize(self, event):
        """Handle canvas resize."""
        # Redraw on resize
        self.after(10, self._draw_diagram)

    def apply_theme(self):
        """Apply theme changes to the diagram."""
        colors = self._theme_manager.colors
        is_dark = self._theme_manager.is_dark
        bg_color = colors.get('surface', '#1e1e2e' if is_dark else '#ffffff')
        self.configure(bg=bg_color)
        self._draw_diagram()
