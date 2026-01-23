"""
Process Control Utilities for Power BI Desktop
Built by Reid Havens of Analytic Endeavors

Utilities for controlling Power BI Desktop process during thin report file swapping:
- Save file (Ctrl+S via pyautogui)
- Close Power BI gracefully (WM_CLOSE)
- Wait for file to be unlocked
- Reopen file after modification
"""

import ctypes
import logging
import os
import subprocess
import time
from dataclasses import dataclass
from typing import List, Optional, Tuple

# Windows constants
WM_CLOSE = 0x0010
WM_QUIT = 0x0012
SW_SHOW = 5
SW_RESTORE = 9


@dataclass
class ProcessControlResult:
    """Result of a process control operation."""
    success: bool
    message: str


class PowerBIProcessController:
    """
    Controls Power BI Desktop process for file operations.

    Used during thin report connection swapping to:
    1. Save the current file (if user chooses)
    2. Close Power BI Desktop
    3. Wait for file to be unlocked
    4. Allow file modification
    5. Reopen the file
    """

    def __init__(self):
        self.logger = logging.getLogger(__name__)
        self._has_pyautogui = None

    def _check_pyautogui(self) -> bool:
        """Check if pyautogui is available."""
        if self._has_pyautogui is None:
            try:
                import pyautogui
                self._has_pyautogui = True
            except ImportError:
                self._has_pyautogui = False
                self.logger.warning("pyautogui not installed - save automation unavailable")
        return self._has_pyautogui

    def get_window_handles(self, process_id: int) -> List[int]:
        """
        Get all window handles for a process.

        Args:
            process_id: The process ID to find windows for

        Returns:
            List of window handles (HWNDs)
        """
        handles = []

        try:
            # Define callback type
            EnumWindowsProc = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)

            def callback(hwnd, lparam):
                # Get process ID for this window
                pid = ctypes.c_ulong()
                ctypes.windll.user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))

                if pid.value == process_id:
                    # Check if window is visible
                    if ctypes.windll.user32.IsWindowVisible(hwnd):
                        handles.append(hwnd)

                return True  # Continue enumeration

            ctypes.windll.user32.EnumWindows(EnumWindowsProc(callback), 0)

        except Exception as e:
            self.logger.debug(f"Error enumerating windows: {e}")

        return handles

    def get_main_window_handle(self, process_id: int) -> Optional[int]:
        """
        Get the main window handle for a Power BI Desktop process.

        Args:
            process_id: The PBIDesktop.exe process ID

        Returns:
            Window handle or None
        """
        try:
            import psutil
            process = psutil.Process(process_id)

            # Try to get main window handle via psutil
            # Some versions expose this, some don't
            if hasattr(process, 'win32_main_window_handle'):
                hwnd = process.win32_main_window_handle()
                if hwnd:
                    return hwnd

        except Exception:
            pass

        # Fallback: enumerate windows and find one with a title
        handles = self.get_window_handles(process_id)

        for hwnd in handles:
            # Get window title length
            length = ctypes.windll.user32.GetWindowTextLengthW(hwnd)
            if length > 0:
                # Get window title
                buffer = ctypes.create_unicode_buffer(length + 1)
                ctypes.windll.user32.GetWindowTextW(hwnd, buffer, length + 1)
                title = buffer.value

                # Power BI Desktop main window has specific title format
                if ' - Power BI Desktop' in title or 'Power BI Desktop' in title:
                    self.logger.debug(f"Found main window: hwnd={hwnd}, title='{title}'")
                    return hwnd

        # Return first visible window if any
        if handles:
            return handles[0]

        return None

    def bring_window_to_foreground(self, hwnd: int) -> bool:
        """
        Bring a window to the foreground.

        Args:
            hwnd: Window handle

        Returns:
            True if successful
        """
        try:
            # Restore if minimized
            if ctypes.windll.user32.IsIconic(hwnd):
                ctypes.windll.user32.ShowWindow(hwnd, SW_RESTORE)

            # Bring to foreground
            ctypes.windll.user32.SetForegroundWindow(hwnd)
            time.sleep(0.2)  # Give time for window to come to front

            return True

        except Exception as e:
            self.logger.debug(f"Error bringing window to foreground: {e}")
            return False

    def save_file(self, process_id: int, timeout: float = 5.0) -> ProcessControlResult:
        """
        Save the currently open file in Power BI Desktop using Ctrl+S.

        Uses pyautogui if available, otherwise falls back to Windows-native
        SendKeys via PowerShell.

        Args:
            process_id: Power BI Desktop process ID
            timeout: Maximum time to wait for save to complete

        Returns:
            ProcessControlResult with success status
        """
        # Get main window
        hwnd = self.get_main_window_handle(process_id)
        if not hwnd:
            return ProcessControlResult(
                success=False,
                message="Could not find Power BI Desktop window"
            )

        # Bring to foreground
        if not self.bring_window_to_foreground(hwnd):
            return ProcessControlResult(
                success=False,
                message="Could not bring window to foreground"
            )

        # Try pyautogui first if available
        if self._check_pyautogui():
            try:
                import pyautogui

                # Send Ctrl+S
                self.logger.info("Sending Ctrl+S to save file (pyautogui)...")
                pyautogui.hotkey('ctrl', 's')

                # Wait for save to complete
                time.sleep(min(timeout, 3.0))

                return ProcessControlResult(
                    success=True,
                    message="Save command sent"
                )

            except Exception as e:
                self.logger.warning(f"pyautogui save failed: {e}, trying fallback...")

        # Fallback: Use Windows SendKeys via PowerShell
        try:
            self.logger.info("Sending Ctrl+S to save file (Windows SendKeys)...")

            # PowerShell script to send Ctrl+S using .NET SendKeys
            ps_script = '''
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.SendKeys]::SendWait("^s")
'''
            result = subprocess.run(
                ['powershell', '-NoProfile', '-NonInteractive', '-Command', ps_script],
                capture_output=True,
                text=True,
                timeout=5.0
            )

            # Wait for save to complete
            time.sleep(min(timeout, 3.0))

            if result.returncode == 0:
                return ProcessControlResult(
                    success=True,
                    message="Save command sent (Windows SendKeys)"
                )
            else:
                self.logger.warning(f"SendKeys failed: {result.stderr}")
                return ProcessControlResult(
                    success=False,
                    message=f"SendKeys failed: {result.stderr}"
                )

        except subprocess.TimeoutExpired:
            return ProcessControlResult(
                success=False,
                message="SendKeys timed out"
            )
        except Exception as e:
            self.logger.error(f"Error saving file: {e}")
            return ProcessControlResult(
                success=False,
                message=f"Error sending save command: {e}"
            )

    def close_gracefully(self, process_id: int, timeout: float = 10.0) -> ProcessControlResult:
        """
        Close Power BI Desktop gracefully by sending WM_CLOSE.

        This triggers the normal close behavior, including "Save changes?" dialog
        if the file has unsaved changes.

        Args:
            process_id: Power BI Desktop process ID
            timeout: Maximum time to wait for process to exit

        Returns:
            ProcessControlResult with success status
        """
        try:
            import psutil

            # Verify process exists
            try:
                process = psutil.Process(process_id)
            except psutil.NoSuchProcess:
                return ProcessControlResult(
                    success=True,
                    message="Process already closed"
                )

            # Get main window
            hwnd = self.get_main_window_handle(process_id)
            if not hwnd:
                return ProcessControlResult(
                    success=False,
                    message="Could not find Power BI Desktop window to close"
                )

            # Send WM_CLOSE to main window
            self.logger.info(f"Sending WM_CLOSE to window {hwnd}...")
            ctypes.windll.user32.PostMessageW(hwnd, WM_CLOSE, 0, 0)

            # Wait for process to exit
            try:
                process.wait(timeout=timeout)
                self.logger.info("Power BI Desktop closed successfully")
                return ProcessControlResult(
                    success=True,
                    message="Power BI Desktop closed"
                )
            except psutil.TimeoutExpired:
                # Process didn't exit - might be waiting for user input (save dialog)
                return ProcessControlResult(
                    success=False,
                    message="Process did not exit (may be waiting for user input)"
                )

        except Exception as e:
            self.logger.error(f"Error closing process: {e}")
            return ProcessControlResult(
                success=False,
                message=f"Error closing Power BI Desktop: {e}"
            )

    def force_close(self, process_id: int) -> ProcessControlResult:
        """
        Force-close Power BI Desktop by terminating the process.

        WARNING: This will lose any unsaved changes!

        Args:
            process_id: Power BI Desktop process ID

        Returns:
            ProcessControlResult with success status
        """
        try:
            import psutil

            try:
                process = psutil.Process(process_id)
                process.terminate()

                # Wait briefly for termination
                try:
                    process.wait(timeout=5.0)
                except psutil.TimeoutExpired:
                    # Force kill if terminate didn't work
                    process.kill()
                    process.wait(timeout=3.0)

                return ProcessControlResult(
                    success=True,
                    message="Power BI Desktop terminated"
                )

            except psutil.NoSuchProcess:
                return ProcessControlResult(
                    success=True,
                    message="Process already closed"
                )

        except Exception as e:
            self.logger.error(f"Error force-closing process: {e}")
            return ProcessControlResult(
                success=False,
                message=f"Error terminating process: {e}"
            )

    def wait_for_file_unlock(self, file_path: str, timeout: float = 30.0,
                             poll_interval: float = 0.5) -> ProcessControlResult:
        """
        Wait until a file is no longer locked.

        Args:
            file_path: Path to the file to check
            timeout: Maximum time to wait in seconds
            poll_interval: Time between checks in seconds

        Returns:
            ProcessControlResult with success status
        """
        start_time = time.time()

        while time.time() - start_time < timeout:
            try:
                # Try to open file exclusively
                with open(file_path, 'r+b'):
                    self.logger.info(f"File is now unlocked: {file_path}")
                    return ProcessControlResult(
                        success=True,
                        message="File is unlocked and ready for modification"
                    )
            except (IOError, PermissionError):
                # File is still locked
                time.sleep(poll_interval)
            except FileNotFoundError:
                return ProcessControlResult(
                    success=False,
                    message=f"File not found: {file_path}"
                )

        return ProcessControlResult(
            success=False,
            message=f"File still locked after {timeout} seconds"
        )

    def reopen_file(self, file_path: str) -> ProcessControlResult:
        """
        Reopen a file in Power BI Desktop.

        Uses the default Windows file handler (os.startfile).

        Args:
            file_path: Path to the .pbix or .pbip file

        Returns:
            ProcessControlResult with success status
        """
        try:
            if not os.path.exists(file_path):
                return ProcessControlResult(
                    success=False,
                    message=f"File not found: {file_path}"
                )

            self.logger.info(f"Opening file: {file_path}")
            os.startfile(file_path)

            return ProcessControlResult(
                success=True,
                message="File opened in Power BI Desktop"
            )

        except Exception as e:
            self.logger.error(f"Error opening file: {e}")
            return ProcessControlResult(
                success=False,
                message=f"Error opening file: {e}"
            )

    def save_close_and_reopen(self, process_id: int, file_path: str,
                               save_first: bool = True,
                               modify_callback=None) -> ProcessControlResult:
        """
        Complete workflow: save (optionally), close, modify, and reopen.

        This is the main method for PBIX file swapping where the file must
        be closed before modification.

        Args:
            process_id: Power BI Desktop process ID
            file_path: Path to the file
            save_first: Whether to save before closing
            modify_callback: Optional function to call after file is unlocked
                           (receives file_path, should return (success, message))

        Returns:
            ProcessControlResult with overall success status
        """
        try:
            # Step 1: Save if requested
            if save_first:
                save_result = self.save_file(process_id)
                if not save_result.success:
                    self.logger.warning(f"Save failed: {save_result.message}")
                    # Continue anyway - user may have already saved

            # Step 2: Close Power BI Desktop
            close_result = self.close_gracefully(process_id, timeout=15.0)
            if not close_result.success:
                # Try force close if graceful close failed
                self.logger.warning("Graceful close failed, trying force close...")
                close_result = self.force_close(process_id)
                if not close_result.success:
                    return ProcessControlResult(
                        success=False,
                        message=f"Could not close Power BI Desktop: {close_result.message}"
                    )

            # Step 3: Wait for file to be unlocked
            unlock_result = self.wait_for_file_unlock(file_path, timeout=30.0)
            if not unlock_result.success:
                return ProcessControlResult(
                    success=False,
                    message=f"File did not unlock: {unlock_result.message}"
                )

            # Step 4: Call modify callback if provided
            if modify_callback:
                try:
                    mod_success, mod_message = modify_callback(file_path)
                    if not mod_success:
                        return ProcessControlResult(
                            success=False,
                            message=f"File modification failed: {mod_message}"
                        )
                except Exception as e:
                    return ProcessControlResult(
                        success=False,
                        message=f"Error during file modification: {e}"
                    )

            # Step 5: Reopen the file
            reopen_result = self.reopen_file(file_path)
            if not reopen_result.success:
                return ProcessControlResult(
                    success=False,
                    message=f"Could not reopen file: {reopen_result.message}"
                )

            return ProcessControlResult(
                success=True,
                message="File swapped and reopened successfully"
            )

        except Exception as e:
            self.logger.error(f"Error in save_close_and_reopen: {e}")
            return ProcessControlResult(
                success=False,
                message=f"Unexpected error: {e}"
            )


# Module-level instance for convenience
_controller: Optional[PowerBIProcessController] = None


def get_controller() -> PowerBIProcessController:
    """Get the global PowerBIProcessController instance."""
    global _controller
    if _controller is None:
        _controller = PowerBIProcessController()
    return _controller
