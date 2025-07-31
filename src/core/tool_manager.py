"""
Tool Manager - Centralized tool discovery and management system
Built by Reid Havens of Analytic Endeavors

This module provides a plugin-like architecture for registering and managing
Power BI tools within the application.
"""

import importlib
import pkgutil
from pathlib import Path
from typing import Dict, List, Type, Optional, Any
from abc import ABC, abstractmethod
import logging

from core.constants import AppConstants


class ToolRegistrationError(Exception):
    """Raised when tool registration fails."""
    pass


class BaseTool(ABC):
    """
    Base class for all Power BI tools.
    Defines the interface that all tools must implement.
    """
    
    def __init__(self, tool_id: str, name: str, description: str, version: str = "1.0.0"):
        self.tool_id = tool_id
        self.name = name
        self.description = description
        self.version = version
        self.enabled = True
        self._logger = None
    
    @property
    def logger(self):
        """Get logger for this tool"""
        if self._logger is None:
            self._logger = logging.getLogger(f"tool.{self.tool_id}")
        return self._logger
    
    @abstractmethod
    def create_ui_tab(self, parent, main_app) -> 'BaseToolTab':
        """
        Create the UI tab for this tool.
        
        Args:
            parent: The parent widget (notebook)
            main_app: Reference to the main application
            
        Returns:
            BaseToolTab: The tab UI component
        """
        pass
    
    @abstractmethod
    def get_tab_title(self) -> str:
        """Get the display title for the tab (with emoji)"""
        pass
    
    @abstractmethod
    def get_help_content(self) -> Dict[str, Any]:
        """Get help content specific to this tool"""
        pass
    
    def can_run(self) -> bool:
        """Check if tool can run (dependencies, etc.)"""
        return True
    
    def get_metadata(self) -> Dict[str, Any]:
        """Get tool metadata"""
        return {
            'id': self.tool_id,
            'name': self.name,
            'description': self.description,
            'version': self.version,
            'enabled': self.enabled
        }


class ToolManager:
    """
    Manages registration, discovery, and lifecycle of Power BI tools.
    """
    
    def __init__(self, logger_callback: Optional[callable] = None):
        self._tools: Dict[str, BaseTool] = {}
        self._tool_tabs: Dict[str, 'BaseToolTab'] = {}
        self.logger_callback = logger_callback or self._default_log
        
        # Initialize logging
        self.logger = logging.getLogger("ToolManager")
    
    def _default_log(self, message: str) -> None:
        """Default logging implementation"""
        print(f"[ToolManager] {message}")
    
    def register_tool(self, tool: BaseTool) -> None:
        """
        Register a tool with the manager.
        
        Args:
            tool: The tool instance to register
            
        Raises:
            ToolRegistrationError: If registration fails
        """
        if not isinstance(tool, BaseTool):
            raise ToolRegistrationError(f"Tool must inherit from BaseTool: {type(tool)}")
        
        if tool.tool_id in self._tools:
            raise ToolRegistrationError(f"Tool '{tool.tool_id}' is already registered")
        
        # Validate tool can run
        if not tool.can_run():
            raise ToolRegistrationError(f"Tool '{tool.tool_id}' cannot run (dependencies not met)")
        
        self._tools[tool.tool_id] = tool
        self.logger_callback(f"✅ Registered tool: {tool.name} (v{tool.version})")
    
    def get_tool(self, tool_id: str) -> Optional[BaseTool]:
        """Get a registered tool by ID"""
        return self._tools.get(tool_id)
    
    def get_all_tools(self) -> List[BaseTool]:
        """Get all registered tools"""
        return list(self._tools.values())
    
    def get_enabled_tools(self) -> List[BaseTool]:
        """Get only enabled tools"""
        return [tool for tool in self._tools.values() if tool.enabled]
    
    def create_tool_tabs(self, notebook_parent, main_app) -> Dict[str, 'BaseToolTab']:
        """
        Create UI tabs for all enabled tools.
        
        Args:
            notebook_parent: The notebook widget to add tabs to
            main_app: Reference to the main application
            
        Returns:
            Dict mapping tool_id to tab instance
        """
        self._tool_tabs.clear()
        
        for tool in self.get_enabled_tools():
            try:
                tab = tool.create_ui_tab(notebook_parent, main_app)
                self._tool_tabs[tool.tool_id] = tab
                
                # Add tab to notebook
                notebook_parent.add(tab.get_frame(), text=tool.get_tab_title())
                
                self.logger_callback(f"✅ Created tab for: {tool.name}")
                
            except Exception as e:
                self.logger_callback(f"❌ Failed to create tab for {tool.name}: {e}")
                continue
        
        return self._tool_tabs
    
    def get_tool_tab(self, tool_id: str) -> Optional['BaseToolTab']:
        """Get a tool's UI tab by tool ID"""
        return self._tool_tabs.get(tool_id)
    
    def discover_and_register_tools(self, tools_package_path: str = "tools") -> int:
        """
        Automatically discover and register tools from the tools package.
        
        Args:
            tools_package_path: Path to the tools package
            
        Returns:
            Number of tools registered
        """
        registered_count = 0
        
        try:
            # Import the tools package
            tools_package = importlib.import_module(tools_package_path)
            package_path = Path(tools_package.__file__).parent
            
            # Discover tool modules
            for finder, name, ispkg in pkgutil.iter_modules([str(package_path)]):
                if ispkg:  # Tool packages are directories
                    try:
                        # Try to import the tool module
                        tool_module_name = f"{tools_package_path}.{name}"
                        tool_module = importlib.import_module(tool_module_name)
                        
                        # Look for tool class (should end with 'Tool')
                        tool_class = self._find_tool_class(tool_module)
                        if tool_class:
                            # Create and register tool instance
                            tool_instance = tool_class()
                            self.register_tool(tool_instance)
                            registered_count += 1
                        
                    except Exception as e:
                        self.logger_callback(f"⚠️ Failed to load tool '{name}': {e}")
                        continue
            
            self.logger_callback(f"🔍 Tool discovery complete: {registered_count} tools registered")
            
        except Exception as e:
            self.logger_callback(f"❌ Tool discovery failed: {e}")
        
        return registered_count
    
    def _find_tool_class(self, module) -> Optional[Type[BaseTool]]:
        """Find the tool class in a module"""
        for item_name in dir(module):
            item = getattr(module, item_name)
            
            # Check if it's a class that inherits from BaseTool
            if (isinstance(item, type) and 
                issubclass(item, BaseTool) and 
                item is not BaseTool and
                item_name.endswith('Tool')):
                return item
        
        return None
    
    def get_tool_help_content(self, tool_id: str) -> Dict[str, Any]:
        """Get help content for a specific tool"""
        tool = self.get_tool(tool_id)
        if tool:
            return tool.get_help_content()
        return {}
    
    def disable_tool(self, tool_id: str) -> bool:
        """Disable a tool"""
        tool = self.get_tool(tool_id)
        if tool:
            tool.enabled = False
            self.logger_callback(f"🔇 Disabled tool: {tool.name}")
            return True
        return False
    
    def enable_tool(self, tool_id: str) -> bool:
        """Enable a tool"""
        tool = self.get_tool(tool_id)
        if tool:
            tool.enabled = True
            self.logger_callback(f"🔊 Enabled tool: {tool.name}")
            return True
        return False
    
    def get_status_summary(self) -> Dict[str, Any]:
        """Get summary of tool manager status"""
        enabled_tools = self.get_enabled_tools()
        
        return {
            'total_tools': len(self._tools),
            'enabled_tools': len(enabled_tools),
            'disabled_tools': len(self._tools) - len(enabled_tools),
            'active_tabs': len(self._tool_tabs),
            'tools': [tool.get_metadata() for tool in self._tools.values()]
        }


# Global tool manager instance
_tool_manager: Optional[ToolManager] = None


def get_tool_manager() -> ToolManager:
    """Get the global tool manager instance"""
    global _tool_manager
    if _tool_manager is None:
        _tool_manager = ToolManager()
    return _tool_manager


def reset_tool_manager() -> None:
    """Reset the global tool manager (useful for testing)"""
    global _tool_manager
    _tool_manager = None
