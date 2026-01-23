# Custom Patterns in Standalone .exe

## Overview

The Sensitivity Scanner supports custom pattern rules that users can create, edit, and save. This document explains how custom patterns work in both development and standalone .exe environments.

## The Problem

When running as a standalone `.exe` (built with PyInstaller):
- ❌ Files are extracted to a **temporary read-only folder** at runtime
- ❌ You **cannot write** to the bundled `data/` directory
- ❌ Custom pattern saves would fail

## The Solution

We use a **two-tier pattern storage system**:

### 1. **Default Patterns (Read-Only)**
- **Development:** `src/data/sensitivity_patterns.json`
- **Standalone .exe:** Extracted to `sys._MEIPASS/data/sensitivity_patterns.json`
- These are the factory default patterns bundled with the application
- **Never modified** by the application

### 2. **Custom Patterns (Writable)**
- **Development:** `src/data/sensitivity_patterns_custom.json`
- **Standalone .exe:** `%APPDATA%\AE Power BI Multi-Tool\Sensitivity Scanner\sensitivity_patterns_custom.json`
- Created when users add/edit/delete patterns
- Persists across application restarts
- **Takes precedence** over default patterns when present

## How It Works

### Pattern Loading Priority

```
1. Check for custom patterns in writable location
   ├─ If found → Use custom patterns
   └─ If not found → Use default patterns
```

### File Locations by Environment

| Environment | Default Patterns | Custom Patterns |
|-------------|-----------------|-----------------|
| **Development** | `src/data/sensitivity_patterns.json` | `src/data/sensitivity_patterns_custom.json` |
| **Standalone .exe** | `{temp_dir}/data/sensitivity_patterns.json` | `%APPDATA%/AE Power BI Multi-Tool/Sensitivity Scanner/sensitivity_patterns_custom.json` |

## User Experience

When users click **"Manage Rules"** in the Sensitivity Scanner:

1. ✅ **View patterns** - Loads from custom file if exists, otherwise default
2. ✅ **Add pattern** - Saves to custom file in AppData
3. ✅ **Edit pattern** - Updates custom file in AppData
4. ✅ **Delete pattern** - Updates custom file in AppData
5. ✅ **Reset to Defaults** - Deletes custom file, reverts to bundled defaults
6. ✅ **Export patterns** - Exports current patterns to user-chosen location

## Technical Implementation

### Code Changes Made

#### `pattern_manager.py`
- Added `_get_bundled_patterns_path()` - Finds default patterns (supports .exe)
- Added `_get_writable_patterns_path()` - Gets writable AppData location
- Uses `sys._MEIPASS` to locate bundled data in standalone .exe
- Uses `os.environ['APPDATA']` for writable storage

#### `pattern_detector.py`
- Updated `_get_default_patterns_file()` to use same logic
- Checks custom patterns in AppData first when running as .exe
- Falls back to bundled defaults if no custom patterns exist

#### `build_ae_pbi_multi_tool.bat`
- Added `--add-data "data;data"` to both build commands
- Ensures `sensitivity_patterns.json` is included in the .exe

## Benefits

✅ **Write permissions** - AppData is always writable  
✅ **Persistence** - Custom patterns survive application updates  
✅ **User isolation** - Each Windows user has their own custom patterns  
✅ **Clean reset** - Deleting custom file reverts to defaults  
✅ **No admin rights** - AppData doesn't require administrator privileges  
✅ **Standard practice** - Follows Windows application guidelines  

## AppData Location

Custom patterns are stored at:
```
C:\Users\{username}\AppData\Roaming\AE Power BI Multi-Tool\Sensitivity Scanner\sensitivity_patterns_custom.json
```

This is the **standard Windows location** for user-specific application data.

## For Developers

If you need to debug custom patterns:

```python
import sys
import os
from pathlib import Path

# Check if running as .exe
if getattr(sys, 'frozen', False):
    print("Running as standalone .exe")
    print(f"Bundled data: {sys._MEIPASS}")
    appdata = Path(os.environ.get('APPDATA'))
    print(f"Custom patterns: {appdata / 'AE Power BI Multi-Tool' / 'Sensitivity Scanner'}")
else:
    print("Running in development")
```

## Migration Notes

If users have existing custom patterns from an older version:
- Old location: `src/data/sensitivity_patterns_custom.json`
- New location: `%APPDATA%\AE Power BI Multi-Tool\Sensitivity Scanner\sensitivity_patterns_custom.json`
- No automatic migration (fresh start with defaults)
- Users can export/import to transfer patterns

---

Built by Reid Havens of Analytic Endeavors
