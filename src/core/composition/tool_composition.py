"""
Tool Composition Framework - Base classes for composable tool architecture
Built by Reid Havens of Analytic Endeavors

This module provides a composition-based architecture where tools are composed
of various components (UI, validation, processing, etc.) while maintaining
inheritance from a base tool class for common functionality.
"""

import threading
import traceback
from abc import ABC, abstractmethod
from typing import Callable, Optional, Any, Dict, List
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox


class ToolComponent(ABC):
    """
    Abstract base class for tool components.
    Components are composable parts that can be mixed and matched.
    """
    
    def __init__(self, tool_context: 'BaseComposableTool'):
        self.tool_context = tool_context
        self._initialized = False
    
    @abstractmethod
    def initialize(self) -> bool:
        """Initialize the component. Return True if successful."""
        pass
    
    @abstractmethod
    def cleanup(self) -> None:
        """Cleanup resources when component is destroyed."""
        pass
    
    def is_initialized(self) -> bool:
        """Check if component is initialized."""
        return self._initialized
    
    def mark_initialized(self) -> None:
        """Mark component as initialized."""
        self._initialized = True


class ValidationComponent(ToolComponent):
    """
    Component for input validation functionality.
    Provides composable validation methods.
    """
    
    def initialize(self) -> bool:
        """Initialize validation component."""
        self.mark_initialized()
        return True
    
    def cleanup(self) -> None:
        """Cleanup validation component."""
        pass
    
    def validate_file_exists(self, file_path: str, file_description: str = "File") -> None:
        """Validate that a file exists"""
        if not file_path:
            raise ValueError(f"{file_description} path is required")
        
        path_obj = Path(file_path)
        if not path_obj.exists():
            raise ValueError(f"{file_description} not found: {file_path}")
        
        if not path_obj.is_file():
            raise ValueError(f"{file_description} path must point to a file: {file_path}")
    
    def validate_pbip_file(self, file_path: str, file_description: str = "File") -> None:
        """Validate that a file is a valid PBIP file"""
        self.validate_file_exists(file_path, file_description)
        
        if not file_path.lower().endswith('.pbip'):
            raise ValueError(f"{file_description} must be a .pbip file")
        
        # Check for corresponding .Report directory
        path_obj = Path(file_path)
        report_dir = path_obj.parent / f"{path_obj.stem}.Report"
        if not report_dir.exists():
            raise ValueError(f"{file_description} missing corresponding .Report directory")
    
    def validate_output_path(self, output_path: str) -> None:
        """Validate output path"""
        if not output_path:
            raise ValueError("Output path is required")
        
        path_obj = Path(output_path)
        
        # Check if parent directory exists
        if not path_obj.parent.exists():
            raise ValueError(f"Output directory does not exist: {path_obj.parent}")
        
        # Check if we can write to the directory
        if not path_obj.parent.is_dir():
            raise ValueError(f"Output parent must be a directory: {path_obj.parent}")


class FileInputComponent(ToolComponent):
    """
    Component for file input functionality.
    Provides composable file handling methods.
    """
    
    def initialize(self) -> bool:
        """Initialize file input component."""
        self.mark_initialized()
        return True
    
    def cleanup(self) -> None:
        """Cleanup file input component."""
        pass
    
    def clean_file_path(self, path: str) -> str:
        """Clean file path by removing quotes and normalizing"""
        if not path:
            return path
        
        cleaned = path.strip()
        
        # Remove surrounding quotes
        if len(cleaned) >= 2:
            if (cleaned.startswith('"') and cleaned.endswith('"')) or \
               (cleaned.startswith("'") and cleaned.endswith("'")):
                cleaned = cleaned[1:-1]
        
        # Normalize path separators
        return str(Path(cleaned)) if cleaned else cleaned
    
    def setup_path_cleaning(self, path_var: tk.StringVar):
        """Setup automatic path cleaning on a StringVar"""
        def on_path_change(*args):
            current = path_var.get()
            cleaned = self.clean_file_path(current)
            if cleaned != current:
                path_var.set(cleaned)
        
        path_var.trace('w', on_path_change)


class ThreadingComponent(ToolComponent):
    """
    Component for background threading functionality.
    Provides safe background execution with proper error handling.
    """
    
    def __init__(self, tool_context: 'BaseComposableTool'):
        super().__init__(tool_context)
        self._active_threads = []
    
    def initialize(self) -> bool:
        """Initialize threading component."""
        self.mark_initialized()
        return True
    
    def cleanup(self) -> None:
        """Cleanup threading component - wait for threads to complete."""
        for thread in self._active_threads:
            if thread.is_alive():
                # Don't wait indefinitely, but give threads a chance to finish
                thread.join(timeout=1.0)
        self._active_threads.clear()
    
    def run_in_background(self, target_func: Callable, 
                         success_callback: Optional[Callable] = None,
                         error_callback: Optional[Callable] = None,
                         finally_callback: Optional[Callable] = None) -> threading.Thread:
        """
        Run a function in background thread with proper error handling.
        
        Args:
            target_func: Function to run in background
            success_callback: Called on success with result
            error_callback: Called on error with exception
            finally_callback: Called after success or error
            
        Returns:
            The thread object
        """
        
        def thread_target():
            """Thread target with proper error handling and closures."""
            result = None
            exception = None
            
            try:
                # Execute the target function
                result = target_func()
                
            except Exception as e:
                exception = e
                # Log the full traceback for debugging
                self.tool_context.log_message(f"❌ Background operation failed: {e}")
                self.tool_context.log_message(f"📋 Traceback: {traceback.format_exc()}")
            
            # Schedule callbacks on main thread with proper closures
            def schedule_callbacks():
                try:
                    if exception is not None:
                        # Handle error
                        if error_callback:
                            error_callback(exception)
                        else:
                            self._default_error_handler(exception)
                    else:
                        # Handle success
                        if success_callback:
                            success_callback(result)
                    
                except Exception as callback_error:
                    # Handle callback errors
                    self.tool_context.log_message(f"❌ Callback error: {callback_error}")
                    self._default_error_handler(callback_error)
                
                finally:
                    # Always call finally callback
                    if finally_callback:
                        try:
                            finally_callback()
                        except Exception as finally_error:
                            self.tool_context.log_message(f"❌ Finally callback error: {finally_error}")
            
            # Schedule on main thread
            if hasattr(self.tool_context, 'frame') and self.tool_context.frame:
                self.tool_context.frame.after(0, schedule_callbacks)
            else:
                # Fallback if no frame available
                schedule_callbacks()
        
        # Create and start thread
        thread = threading.Thread(target=thread_target, daemon=True)
        self._active_threads.append(thread)
        thread.start()
        
        return thread
    
    def _default_error_handler(self, error: Exception):
        """Default error handler"""
        self.tool_context.log_message(f"❌ Error: {error}")
        if hasattr(self.tool_context, 'show_error'):
            self.tool_context.show_error("Operation Error", str(error))
        else:
            messagebox.showerror("Operation Error", str(error))


class ProgressComponent(ToolComponent):
    """
    Component for progress indication functionality.
    Provides composable progress tracking methods.
    """
    
    def __init__(self, tool_context: 'BaseComposableTool'):
        super().__init__(tool_context)
        self.progress_bar = None
        self.is_busy = False
    
    def initialize(self) -> bool:
        """Initialize progress component."""
        self.mark_initialized()
        return True
    
    def cleanup(self) -> None:
        """Cleanup progress component."""
        if self.progress_bar:
            self.hide_progress()
    
    def create_progress_bar(self, parent: ttk.Widget) -> ttk.Progressbar:
        """Create a progress bar widget"""
        self.progress_bar = ttk.Progressbar(parent, mode='indeterminate')
        return self.progress_bar
    
    def show_progress(self):
        """Show progress indication"""
        if self.progress_bar:
            self.progress_bar.grid(sticky=(tk.W, tk.E), pady=(10, 0))
            self.progress_bar.start()
            self.is_busy = True
    
    def hide_progress(self):
        """Hide progress indication"""
        if self.progress_bar:
            self.progress_bar.stop()
            self.progress_bar.grid_remove()
            self.is_busy = False
    
    def run_with_progress(self, target_func: Callable,
                         success_callback: Optional[Callable] = None,
                         error_callback: Optional[Callable] = None) -> threading.Thread:
        """
        Run function with automatic progress indication.
        Requires ThreadingComponent to be available.
        """
        if not hasattr(self.tool_context, 'threading'):
            raise RuntimeError("ThreadingComponent required for run_with_progress")
        
        def progress_finally():
            self.hide_progress()
        
        # Show progress before starting
        self.show_progress()
        
        # Run in background with progress cleanup
        return self.tool_context.threading.run_in_background(
            target_func=target_func,
            success_callback=success_callback,
            error_callback=error_callback,
            finally_callback=progress_finally
        )


class BaseComposableTool(ABC):
    """
    Base class for composable tools.
    Tools inherit from this and compose functionality using components.
    """
    
    def __init__(self, tool_id: str, tool_name: str):
        self.tool_id = tool_id
        self.tool_name = tool_name
        self._components: Dict[str, ToolComponent] = {}
        self._initialized = False
        
        # Initialize default components
        self._initialize_default_components()
    
    def _initialize_default_components(self):
        """Initialize default components that most tools need."""
        self.add_component('validation', ValidationComponent(self))
        self.add_component('file_input', FileInputComponent(self))
        self.add_component('threading', ThreadingComponent(self))
        self.add_component('progress', ProgressComponent(self))
    
    def add_component(self, name: str, component: ToolComponent):
        """Add a component to the tool."""
        self._components[name] = component
        # Create convenient property access
        setattr(self, name, component)
    
    def get_component(self, name: str) -> Optional[ToolComponent]:
        """Get a component by name."""
        return self._components.get(name)
    
    def initialize_components(self) -> bool:
        """Initialize all components."""
        success = True
        for name, component in self._components.items():
            try:
                if not component.initialize():
                    self.log_message(f"❌ Failed to initialize component: {name}")
                    success = False
            except Exception as e:
                self.log_message(f"❌ Error initializing component {name}: {e}")
                success = False
        
        self._initialized = success
        return success
    
    def cleanup_components(self):
        """Cleanup all components."""
        for component in self._components.values():
            try:
                component.cleanup()
            except Exception as e:
                self.log_message(f"❌ Error cleaning up component: {e}")
    
    def is_initialized(self) -> bool:
        """Check if tool is initialized."""
        return self._initialized
    
    @abstractmethod
    def log_message(self, message: str):
        """Log message - must be implemented by subclasses."""
        pass
    
    @abstractmethod
    def show_error(self, title: str, message: str):
        """Show error dialog - must be implemented by subclasses."""
        pass
    
    def __del__(self):
        """Cleanup when tool is destroyed."""
        try:
            self.cleanup_components()
        except:
            pass  # Ignore errors during cleanup


class UIComposableTool(BaseComposableTool):
    """
    Base class for composable tools with UI.
    Extends BaseComposableTool with UI-specific functionality.
    """
    
    def __init__(self, parent, main_app, tool_id: str, tool_name: str):
        super().__init__(tool_id, tool_name)
        
        # UI references
        self.parent = parent
        self.main_app = main_app
        
        # Create main frame
        self.frame = ttk.Frame(parent, padding="20")
        
        # UI state
        self.log_text = None
    
    def log_message(self, message: str):
        """Log message to UI log area."""
        if self.log_text:
            self.log_text.config(state=tk.NORMAL)
            self.log_text.insert(tk.END, message + "\n")
            self.log_text.config(state=tk.DISABLED)
            self.log_text.see(tk.END)
            if hasattr(self, 'frame'):
                self.frame.update_idletasks()
        else:
            # Fallback to print if no UI log available
            print(f"[{self.tool_name}] {message}")
    
    def show_error(self, title: str, message: str):
        """Show error dialog."""
        messagebox.showerror(title, message)
    
    def show_info(self, title: str, message: str):
        """Show info dialog."""
        messagebox.showinfo(title, message)
    
    def show_warning(self, title: str, message: str):
        """Show warning dialog."""
        messagebox.showwarning(title, message)
    
    def ask_yes_no(self, title: str, message: str) -> bool:
        """Show yes/no dialog"""
        return messagebox.askyesno(title, message)
    
    def get_frame(self) -> ttk.Frame:
        """Return the main frame for this tool."""
        return self.frame
    
    @abstractmethod
    def setup_ui(self) -> None:
        """Setup the UI - must be implemented by subclasses."""
        pass
    
    @abstractmethod
    def reset_tool(self) -> None:
        """Reset the tool to initial state - must be implemented by subclasses."""
        pass
    
    @abstractmethod
    def show_help_dialog(self) -> None:
        """Show help dialog - must be implemented by subclasses."""
        pass


# Convenience functions for creating common component combinations

def create_file_processing_tool(parent, main_app, tool_id: str, tool_name: str) -> UIComposableTool:
    """
    Create a tool with common file processing components.
    Returns a UIComposableTool with validation, file input, threading, and progress components.
    """
    tool = UIComposableTool(parent, main_app, tool_id, tool_name)
    
    # Default components are already added by BaseComposableTool
    # Additional components can be added here if needed
    
    return tool


def create_analysis_tool(parent, main_app, tool_id: str, tool_name: str) -> UIComposableTool:
    """
    Create a tool optimized for analysis operations.
    Returns a UIComposableTool with enhanced logging and progress tracking.
    """
    tool = create_file_processing_tool(parent, main_app, tool_id, tool_name)
    
    # Could add analysis-specific components here
    # tool.add_component('analysis', AnalysisComponent(tool))
    
    return tool
