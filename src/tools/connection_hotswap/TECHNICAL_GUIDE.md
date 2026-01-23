# Connection Hot-Swap - Technical Guide

**Tool Version:** 2.2.0
**Last Updated:** January 2026
**Built by:** Reid Havens of Analytic Endeavors

---

## Table of Contents

1. [Overview](#overview)
2. [Architecture](#architecture)
3. [Core Data Models](#core-data-models)
4. [Core Capabilities](#core-capabilities)
5. [Swap Workflows](#swap-workflows)
6. [Safety Mechanisms & Failsafes](#safety-mechanisms--failsafes)
7. [Preset & Configuration System](#preset--configuration-system)
8. [Cloud Browser Features](#cloud-browser-features)
9. [Technical Requirements](#technical-requirements)
10. [Connection String Formats](#connection-string-formats)
11. [Error Handling](#error-handling)
12. [Best Practices](#best-practices)
13. [Troubleshooting](#troubleshooting)
14. [Version History](#version-history)

---

## Overview

### What is Connection Hot-Swap?

Connection Hot-Swap is a professional-grade connection management tool for Power BI development workflows. It enables dynamically switching Power BI data source connections between environments without republishing reports - seamless transitions between development, testing, and production datasets, whether local Power BI Desktop models or cloud-hosted Power BI Service datasets.

### Key Capabilities

- **Composite Models**: Hot-swap connections while the report is open in Desktop (TOM-based)
- **Thin Reports**: Swap via PBIP file modification (PBIP can be modified while open)
- **Cloud-to-Cloud**: Switch between different cloud datasets, workspaces, or connection types
- **Universal Perspectives**: Connect to perspectives in ANY workspace type (Pro, Premium, PPU, Fabric)

### Key Use Cases

- **Environment Switching**: Quickly swap between Dev/Test/Prod datasets
- **Thin Report Development**: Connect cloud-published reports to local models for development
- **Cloud Migration**: Move reports between workspaces or tenants
- **Offline Development**: Work with local copies of cloud models
- **Team Collaboration**: Share preset configurations across development teams
- **Cloud-to-Cloud Switching**: Change connection types (Semantic Model to XMLA, etc.)

### Supported Scenarios

| Model Type | Swap Direction | Method | Notes |
|------------|----------------|--------|-------|
| Composite Model | Cloud to Local | TOM modification | Hot-swap while open in Desktop |
| Composite Model | Local to Cloud | TOM modification | Hot-swap while open in Desktop |
| Composite Model | Cloud to Cloud | TOM modification | Hot-swap while open in Desktop |
| Thin Report | Cloud to Local | File modification | Requires PBIP format |
| Thin Report | Local to Cloud | File modification | Requires PBIP format |

**Important Distinction:**
- **Composite Models**: True hot-swap via TOM - report stays open and works immediately
- **Thin Reports**: Require close/reopen after swap; PBIP allows file edit while open (no lock), PBIX requires closing Desktop first

### What It Does NOT Do

- Does not modify report visuals or data model structure
- Does not transfer data between environments
- Does not work with PBIX thin reports while Desktop has them open (file locked)
- Is NOT officially supported by Microsoft

---

## Architecture

### Directory Structure

```
connection_hotswap/
├── connection_hotswap_tool.py      # Tool registration and metadata
├── models.py                        # Data classes and enums
│
├── logic/                           # Business logic layer
│   ├── connection_detector.py       # TOM connection enumeration
│   ├── connection_swapper.py        # TOM modification engine
│   ├── local_model_matcher.py       # Local model discovery & matching
│   ├── cloud_workspace_browser.py   # Power BI Service integration
│   ├── pbix_modifier.py             # Thin report file modification
│   ├── preset_manager.py            # Configuration persistence
│   ├── health_checker.py            # Target health monitoring
│   ├── schema_validator.py          # Compatibility validation
│   ├── connection_cache.py          # Persistent cloud connection cache
│   └── process_control.py           # Power BI Desktop automation
│
├── ui/                              # User interface layer
│   ├── connection_hotswap_tab.py    # Main UI tab
│   ├── connection_diagram.py        # Visual flow diagram
│   ├── components/
│   │   └── inline_target_picker.py  # Target selection dropdown
│   └── dialogs/
│       ├── cloud_browser_dialog.py  # Cloud workspace browser
│       ├── local_selector_dialog.py # Local model selector
│       ├── thin_report_dialog.py    # Thin report swap helper
│       └── schema_mismatch_dialog.py # Schema validation warnings
│
└── __init__.py
```

### Layer Separation

The tool follows a clean three-layer architecture:

1. **Models Layer** (`models.py`)
   - Pure data classes with no business logic
   - Serialization/deserialization methods
   - Type definitions and enums

2. **Logic Layer** (`logic/`)
   - All business operations
   - No UI dependencies
   - Testable in isolation

3. **UI Layer** (`ui/`)
   - Tkinter-based interface
   - Calls into logic layer
   - Event handling and display

---

## Core Data Models

### Connection Types

```python
class ConnectionType(Enum):
    LIVE_CONNECTION = "LiveConnection"      # Direct live connection
    DIRECT_QUERY = "DirectQuery"            # DirectQuery partition
    IMPORT = "Import"                       # Import mode (not swappable)
    DUAL = "Dual"                           # Dual storage mode
    COMPOSITE = "Composite"                 # Mixed mode model
    UNKNOWN = "Unknown"

class CloudConnectionType(Enum):
    PBI_SEMANTIC_MODEL = "pbiServiceLive"           # Power BI Semantic Model connector (v2)
    AAS_XMLA = "analysisServicesDatabaseLive"       # XMLA endpoint connection (v1)
    PBI_XMLA_STYLE = "pbiServiceXmlaStyleLive"      # Pro workspace perspectives (v4)
    UNKNOWN = "unknown"
```

### Connection Type Selection

The tool automatically selects the optimal connection type:

| Scenario | Connection Type | Version | Notes |
|----------|----------------|---------|-------|
| Standard semantic model (no perspective) | `pbiServiceLive` | v2 | Works with all workspace types |
| XMLA endpoint (Premium/PPU/Fabric) | `analysisServicesDatabaseLive` | v1 | Full XMLA access, perspectives supported |
| Perspective on Pro workspace | `pbiServiceXmlaStyleLive` | v4 | Hybrid format for Pro perspectives |

**Important**: The `Cube` parameter (for perspectives) is NOT supported by `pbiServiceLive`. Attempting to use it causes Power BI Desktop to crash with a null reference error.

### Primary Data Classes

#### DataSourceConnection
Represents a single swappable connection discovered in a model.

```python
@dataclass
class DataSourceConnection:
    name: str                           # Connection identifier
    connection_type: ConnectionType     # Type of connection
    server: str                         # Server/endpoint URL
    database: str                       # Database/dataset name
    provider: Optional[str]             # OLE DB provider
    is_cloud: bool                      # Cloud vs local indicator
    connection_string: str              # Full connection string
    tom_reference: Any                  # TOM object for modification
    tom_reference_type: TomReferenceType
```

#### SwapTarget
Defines the target for a swap operation.

```python
@dataclass
class SwapTarget:
    target_type: str                    # "local" or "cloud"
    server: str                         # Target server
    database: str                       # Target database
    display_name: str                   # User-friendly name
    cloud_connection_type: CloudConnectionType
    perspective_name: Optional[str]     # For XMLA perspectives
    workspace_name: Optional[str]       # Cloud workspace
```

#### ConnectionMapping
Pairs a source connection with its target.

```python
@dataclass
class ConnectionMapping:
    source: DataSourceConnection        # Original connection
    target: Optional[SwapTarget]        # Swap destination
    status: SwapStatus                  # Current status
    original_connection_string: str     # For rollback
    error_message: Optional[str]        # Error details
```

### Swap Status Progression

```
PENDING -> MATCHED -> READY -> SWAPPING -> SUCCESS
                                       \-> ERROR
```

---

## Core Capabilities

### 1. Connection Detection

The `ConnectionDetector` class analyzes the Tabular Object Model (TOM) to discover all swappable connections.

**Detection Sources:**
- `Model.DataSources` collection - Standard datasources
- Partition expressions - DirectQuery in composite models
- Pure live connections - Models without DataSources collection

**Detected Information:**
- Connection type (Live, DirectQuery, Import)
- Server and database identifiers
- Cloud vs local classification
- Original connection strings

### 2. Smart Matching

The `LocalModelMatcher` automatically discovers open Power BI Desktop instances and suggests matches.

**Matching Algorithm:**
1. **Exact Match** - Names match exactly (case-insensitive)
2. **Contains Match** - One name contains the other
3. **Fuzzy Match** - SequenceMatcher with 60% threshold

**Discovery Methods:**
- Port file detection (fast path)
- Netstat port scanning (fallback)
- Window title parsing for friendly names

### 3. Cloud Integration

The `CloudWorkspaceBrowser` provides full Power BI Service integration.

**Features:**
- OAuth authentication via MSAL
- Persistent token cache (silent re-authentication)
- Workspace browsing (All/Favorites/Recent)
- Cross-workspace dataset search
- Manual XMLA endpoint entry
- Perspective selection (works in ALL workspace types: Pro, Premium, PPU, Fabric)
- Cloud-to-cloud swapping (switch between datasets, workspaces, or connection types)

**Authentication Flow:**
```
1. Check for cached token
2. Attempt silent token refresh
3. Interactive browser login (if needed)
4. Token persisted for future sessions
```

### 4. Thin Report Support

The `PbixModifier` handles reports without embedded models (thin reports).

**Supported Formats:**
- `.pbip` - Direct JSON file editing (no file lock while Desktop has it open) - **Recommended**
- `.pbix` - ZIP archive modification (requires closing Desktop first due to file lock)

**Key Operations:**
- Read current connection from file
- Swap to local model (server:port format)
- Restore original cloud connection
- Create timestamped backups

**Important:** Unlike composite models which can true hot-swap while open, ALL thin reports require closing and reopening the report after the swap to pick up changes. PBIP format is recommended because the file can be edited without closing Desktop first (no file lock), but the report must still be reloaded.

### 5. Preset System

The `PresetManager` handles configuration persistence.

**Preset Scopes:**

| Scope | Description | Use Case |
|-------|-------------|----------|
| GLOBAL | Single target, any model | "Always use Prod dataset" |
| MODEL | Full mapping snapshot | "Dev config for Sales Report" |

**Storage Locations:**

| Location | Path | Use Case |
|----------|------|----------|
| USER | `%APPDATA%/Analytic Endeavors/...` | Personal presets |
| PROJECT | `.pbi-hotswap/` folder | Team-shared presets |
| REPORT | Embedded in PBIP | Report-specific config |

### 6. Composite Model Support

Full support for models with multiple live connections:
- Detects all swappable connections
- Independent target assignment per connection
- Handles mixed DirectQuery/Import partitions

---

## Swap Workflows

### Standard TOM Swap (Model Open in Desktop)

```
+---------------------------------------------------------------+
|  1. DETECT                                                    |
|     ConnectionDetector scans TOM for swappable connections    |
+---------------------------------------------------------------+
                            |
                            v
+---------------------------------------------------------------+
|  2. SUGGEST                                                   |
|     LocalModelMatcher auto-matches by name similarity         |
+---------------------------------------------------------------+
                            |
                            v
+---------------------------------------------------------------+
|  3. CONFIGURE                                                 |
|     User selects targets for unmatched connections            |
|     - Browse cloud workspaces                                 |
|     - Select local models                                     |
|     - Apply presets                                           |
+---------------------------------------------------------------+
                            |
                            v
+---------------------------------------------------------------+
|  4. VALIDATE                                                  |
|     SchemaValidator checks target compatibility (warnings)    |
|     HealthChecker verifies target accessibility               |
+---------------------------------------------------------------+
                            |
                            v
+---------------------------------------------------------------+
|  5. SWAP                                                      |
|     ConnectionSwapper modifies TOM via SaveChanges()          |
|     RequestRefresh(Calculate) forces reconnection             |
+---------------------------------------------------------------+
                            |
                            v
+---------------------------------------------------------------+
|  6. VERIFY & LOG                                              |
|     Validate change persisted                                 |
|     Create history entry for rollback                         |
+---------------------------------------------------------------+
```

### Thin Report File Swap

```
+---------------------------------------------------------------+
|  1. DETECT                                                    |
|     Read Connections file from PBIX/PBIP                      |
|     Parse cloud connection details                            |
+---------------------------------------------------------------+
                            |
                            v
+---------------------------------------------------------------+
|  2. CACHE ORIGINAL                                            |
|     Store exact cloud connection format for restoration       |
+---------------------------------------------------------------+
                            |
                            v
+---------------------------------------------------------------+
|  3. BACKUP                                                    |
|     Create timestamped backup file                            |
|     Example: Report_backup_20241227_163045.pbix               |
+---------------------------------------------------------------+
                            |
                            v
+---------------------------------------------------------------+
|  4. CLOSE (PBIX only)                                         |
|     Close Power BI Desktop (file locked)                      |
|     PBIP can be modified while open                           |
+---------------------------------------------------------------+
                            |
                            v
+---------------------------------------------------------------+
|  5. MODIFY                                                    |
|     Update Connections file with local server:port            |
|     PBIX: Extract ZIP -> Modify -> Recompress                 |
|     PBIP: Direct JSON file edit                               |
+---------------------------------------------------------------+
                            |
                            v
+---------------------------------------------------------------+
|  6. REOPEN                                                    |
|     Launch Power BI Desktop with modified file                |
|     Report now connects to local model                        |
+---------------------------------------------------------------+
```

---

## Safety Mechanisms & Failsafes

### 1. Backup Systems

**PBIX/PBIP Backups:**
- Created automatically before any file modification
- Timestamped naming: `{filename}_backup_{YYYYMMDD_HHMMSS}.pbix`
- Stored alongside original file
- User option to disable via checkbox

**Backup Restoration:**
```
Original: Report.pbix
Backup:   Report_backup_20241227_163045.pbix

To restore: Delete modified file, rename backup
```

### 2. Cloud Connection Caching (Persistent)

When swapping a thin report from cloud to local, the original cloud connection is cached using a dual-layer system:

**Dual-Layer Cache:**
- **Memory Cache**: Fast access during active session
- **Disk Cache**: Persistent storage at `%APPDATA%/Analytic Endeavors/PBI Report Merger/hotswap_presets/connection_cache.json`

**Cached Format:**
```json
{
  "PbiServiceModelId": 16589041,
  "PbiModelVirtualServerName": "sobe_wowvirtualserver",
  "PbiModelDatabaseName": "f1a05c34-cb7e-4b08-bd82-fdf6d8412be1",
  "ConnectionType": "pbiServiceLive"
}
```

**Cache Behavior:**
- **Cloud-to-Local Swap**: Original connection cached to memory AND disk
- **Local-to-Cloud Swap**:
  - Tries memory cache first, then disk cache
  - If cached: restores EXACT original format with all fields
  - If no cache: generates generic connection WITH warning message
- **App Restart**: Disk cache automatically loaded into memory on startup

### 3. Schema Fingerprinting & Preset Validation

Each cloud connection has a schema fingerprint for change detection.

**Fingerprint Fields:**
- `PbiServiceModelId`
- `PbiModelVirtualServerName`
- `PbiModelDatabaseName`
- `ConnectionType`

**Preset Schema Validation:**
When applying a preset, the schema fingerprint is compared against the current file:

1. **Fingerprints match**: Preset applied normally
2. **Mismatch detected**: Warning dialog appears with options:
   - **Update Preset**: Update preset with current schema, then apply
   - **Apply Anyway**: Apply preset's stored connection (may have issues)
   - **Cancel**: Abort the operation

### 4. APPLY TARGET (Original Connection Restoration)

The APPLY TARGET feature allows quick restoration to the original connection state after swapping.

**How It Works:**
1. When a model is scanned, the starting configuration is automatically saved
2. After swapping to a different target, click APPLY TARGET to restore
3. Ready mappings are auto-selected for immediate swap execution

**Use Cases:**
- Quickly revert after testing with a different data source
- Round-trip development: local -> cloud -> local
- Recover from accidental target selection

### 5. Swap History & Rollback

Every swap operation is logged with full details:

```python
@dataclass
class SwapHistoryEntry:
    timestamp: datetime
    run_id: str                    # Groups batch operations
    source_name: str
    original_connection: str       # For rollback
    new_connection: str
    status: str
    model_file: str
```

**Rollback Capability:**
- Single connection rollback
- Batch rollback by run_id
- History persisted across sessions

### 6. Schema Validation

The `SchemaValidator` compares source and target model structures:

**Checks Performed:**
- Missing tables in target
- Missing columns in target
- Measure count differences
- Relationship compatibility

**Behavior:**
- Warnings only (does not block swaps)
- User informed of potential issues
- Logs detailed findings

### 7. Health Checks

The `HealthChecker` monitors target accessibility:

**Socket-Based Checks:**
- TCP connection test to target server:port
- 5-second timeout per check
- Background thread for non-blocking operation

**Health Status:**
```python
class HealthStatus(Enum):
    UNKNOWN = "unknown"
    CHECKING = "checking"
    HEALTHY = "healthy"
    UNHEALTHY = "unhealthy"
    ERROR = "error"
```

### 8. TOM Verification

After each swap, the tool verifies the change persisted:

```python
def _validate_swap_persisted(self, mapping):
    # Re-read connection from TOM
    # Compare with expected new value
    # Log discrepancy if found
```

### 9. Rollback on Error

If a swap fails mid-operation:
- Original connection string preserved in mapping
- Automatic restoration attempt
- User notified with recovery options

---

## Preset & Configuration System

### Preset JSON Schema (v2.0)

```json
{
  "version": "2.0",
  "global_presets": {
    "Production": {
      "name": "Production",
      "scope": "global",
      "storage_type": "user",
      "description": "Production Power BI Service",
      "created": "2024-12-27T16:30:00",
      "mappings": [
        {
          "target_type": "cloud",
          "server": "powerbi://api.powerbi.com/v1.0/myorg/Production",
          "database": "Sales Dataset",
          "display_name": "Sales Dataset (Production)",
          "cloud_connection_type": "pbiServiceLive",
          "workspace_name": "Production"
        }
      ]
    }
  },
  "model_presets": {
    "abc123def456": {
      "Dev Environment": {
        "name": "Dev Environment",
        "scope": "model",
        "model_hash": "abc123def456",
        "mappings": [...]
      }
    }
  },
  "last_configs": {
    "abc123def456": {
      "mappings": [...],
      "workspace_name": "Development",
      "friendly_name": "Sales Report"
    }
  },
  "settings": {
    "auto_backup": true,
    "auto_reconnect": true
  }
}
```

### Storage Locations

| Type | Path | Git-Tracked |
|------|------|-------------|
| USER | `%APPDATA%/Analytic Endeavors/PBI Report Merger/hotswap_presets/` | No |
| PROJECT | `.pbi-hotswap/presets/` | Optional |
| REPORT | `{report}.SemanticModel/.hotswap/` | Yes (with PBIP) |

---

## Cloud Browser Features

### Authentication

**MSAL Configuration:**
```python
CLIENT_ID = "a672d62c-fc7b-4e81-a576-e60dc46e951d"  # Power BI Public Client
AUTHORITY = "https://login.microsoftonline.com/common"
SCOPES = ["https://analysis.windows.net/powerbi/api/.default"]
```

**Token Cache:**
- Location: `%APPDATA%/Analytic Endeavors/PBI Report Merger/msal_token_cache.bin`
- Encrypted using MSAL's built-in protection
- Enables silent re-authentication

### Workspace Browsing

**Views:**
- **All** - All accessible workspaces (default view)
- **Favorites** - User-marked favorites (persisted locally)
- **Recent** - Recently accessed workspaces/datasets

**Default Behavior:**
- "(All Workspaces)" is selected by default when opening the browser
- All semantic models from all workspaces displayed initially
- Selecting a workspace filters to that workspace's models only

### Workspace Capacity Detection

The cloud browser automatically detects workspace capacity type and adjusts available options:

**Capacity Types:**
| Type | XMLA Support | Indicator |
|------|--------------|-----------|
| Premium | Yes | Filled diamond |
| PPU (Per-User) | Yes | Outlined diamond |
| Fabric | Yes | Filled diamond |
| Pro | No | (none) |

**XMLA Option Behavior:**
- Premium/PPU/Fabric: Both "Semantic Model" and "XMLA Endpoint" options available
- Pro workspace: "XMLA Endpoint" option disabled with tooltip explaining the limitation

### Cloud Connection Type Selector

When selecting a cloud target, users can choose the connection method:

**Semantic Model (pbiServiceLive)**
- Default for standard connections without perspectives
- Works with Pro, Premium, PPU, and Fabric workspaces
- Simpler connection format, lower overhead
- Does NOT support the `Cube` parameter for perspectives

**XMLA Endpoint (analysisServicesDatabaseLive)**
- Required for Premium/PPU/Fabric workspaces when using perspectives
- Full XMLA protocol access
- Supports perspective selection via `Cube` parameter

---

## Technical Requirements

### Required Dependencies

| Package | Purpose |
|---------|---------|
| `pythonnet` | .NET CLR integration for TOM access |
| `msal` | Microsoft Authentication Library |
| `requests` | Power BI REST API calls |

### Optional Dependencies

| Package | Purpose | Fallback |
|---------|---------|----------|
| `pyautogui` | Keyboard automation (Ctrl+S) | Manual save prompt |
| `cairosvg` | SVG icon rendering | Native Tkinter checkboxes |
| `Pillow` | Image manipulation | Reduced icon quality |

### Windows API Usage

**User32.dll:**
- `EnumWindows` - Window enumeration
- `GetWindowThreadProcessId` - Process identification
- `PostMessage` - Window messaging (WM_CLOSE)

**Kernel32.dll:**
- `OpenProcess` - Process access
- `QueryFullProcessImageName` - File path retrieval

### Power BI Desktop Requirements

- Power BI Desktop installed (Store or standalone)
- TOM libraries available in installation
- Composite models: Must be open in Desktop for TOM-based swaps
- Thin reports: Require PBIP format (PBIP can be modified while open)
- Cloud connections: Work with Pro, Premium, PPU, and Fabric workspaces
- Perspectives: Supported in ALL workspace types (including Pro)

---

## Connection String Formats

### Local Model
```
Provider=MSOLAP;Data Source=localhost:{port};
```

### Cloud - PBI Semantic Model
```
Provider=MSOLAP;Data Source=pbiazure://api.powerbi.com;
Initial Catalog={dataset-guid};
```

### Cloud - XMLA Endpoint
```
Provider=MSOLAP;Data Source=powerbi://api.powerbi.com/v1.0/myorg/{workspace};
Initial Catalog={dataset-name};
```

### With Perspective (XMLA - Premium/PPU/Fabric)
```
Provider=MSOLAP;Data Source=powerbi://api.powerbi.com/v1.0/myorg/{workspace};
Initial Catalog={dataset-name};
Cube={perspective-name};
```

### With Perspective (Pro Workspace)
Uses `pbiServiceXmlaStyleLive` connection type with XMLA-style data source:
```json
{
  "Version": 4,
  "Connections": [{
    "Name": "EntityDataSource",
    "ConnectionString": "Data Source=powerbi://api.powerbi.com/v1.0/myorg/{workspace};Initial Catalog={dataset-guid};Cube={perspective-name}",
    "ConnectionType": "pbiServiceXmlaStyleLive",
    "PbiServiceModelId": null,
    "PbiModelVirtualServerName": "sobe_wowvirtualserver",
    "PbiModelDatabaseName": "{dataset-guid}"
  }]
}
```

---

## Error Handling

### Common Errors

| Error | Cause | Resolution |
|-------|-------|------------|
| "Connection timed out" | Port not accessible | Ensure Power BI Desktop is open |
| "Not authenticated" | Token expired | Re-authenticate via cloud browser |
| "File is locked" | PBIX open in Desktop | Close file or use PBIP format |
| "Premium capacity required" | XMLA not enabled | Use PBI Semantic Model connector |
| "Schema mismatch" | Target has different structure | Review validation warnings |

### Recovery Procedures

**Failed TOM Swap:**
1. Check Power BI Desktop is still responsive
2. Use Rollback button to restore original
3. If unresponsive, close without saving

**Failed File Swap:**
1. Locate backup file (timestamped)
2. Delete modified file
3. Rename backup to original name

**Authentication Issues:**
1. Clear token cache (delete msal_token_cache.bin)
2. Re-authenticate via cloud browser
3. Check Azure AD permissions

---

## Best Practices

### Development Workflow

1. **Start Local**: Develop with local model copy
2. **Save Preset**: Create "Local Dev" preset
3. **Test Cloud**: Swap to cloud for integration testing
4. **Save Preset**: Create "Cloud Test" preset
5. **Quick Switch**: Use presets for rapid environment changes

### Team Collaboration

1. **Project Presets**: Store in `.pbi-hotswap/` folder
2. **Git Integration**: Include in version control (optional)
3. **Naming Convention**: `{Environment}_{Purpose}` (e.g., "Prod_ReadOnly")

### Thin Report Development

1. **PBIP Required**: File-based modification requires PBIP format
2. **PBIP Advantage**: File can be edited while Desktop has it open (no lock)
3. **PBIX Limitation**: Must close Desktop before modifying PBIX files (file locked)
4. **Reload Required**: All thin reports require close/reopen after swap to pick up changes
5. **Enable Backups**: Keep auto-backup enabled
6. **Cache Cloud**: Let tool cache cloud connection before local swap
7. **Easy Restore**: Use "Swap to Cloud" for seamless restoration

### Composite vs Thin Report Workflows

| Model Type | When to Use | Key Benefit |
|------------|-------------|-------------|
| Composite Model | Development with local data copies | True hot-swap (no restart needed) |
| Thin Report (PBIP) | Cloud-only reports needing local dev | File editable while open (but reload required) |
| Thin Report (PBIX) | Legacy files (convert to PBIP if possible) | Requires closing Desktop first |

---

## Troubleshooting

### Model Not Detected

**Symptoms:** Dropdown shows no models or "No swappable connections"

**Checks:**
1. Is Power BI Desktop running?
2. Is a model file open (not just Desktop)?
3. Does the model have live connections or DirectQuery?
4. Try Refresh button to re-scan

### Cloud Authentication Fails

**Symptoms:** "Authentication failed" or redirect loop

**Checks:**
1. Check internet connectivity
2. Verify Azure AD account has Power BI access
3. Clear token cache and retry
4. Check for proxy/firewall blocking

### Swap Completes but Model Unchanged

**Symptoms:** Status shows "Success" but data unchanged

**Checks:**
1. Refresh visuals manually (may need Ctrl+Shift+F5)
2. Check Data view for updated tables
3. Verify target model has expected data
4. Review log for "Swap persisted: False" warnings

### File Locked Error

**Symptoms:** "Cannot modify file - in use"

**Resolution:**
1. For PBIX: Close Power BI Desktop first
2. For PBIP: Should work while open - check for other locks
3. Check for OneDrive/sync conflicts
4. Try closing and reopening the file

---

## Version History

### v2.2.0 (Current)
- **APPLY TARGET Button**: Renamed from "LAST CONFIG" - restores original target connections after swapping
- **Cloud Browser Dialog Improvements**: Default selection of "(All Workspaces)", optimized panel widths
- **XMLA Access Detection**: Automatic detection of Pro workspaces, visual feedback when XMLA unavailable
- **UI Polish**: SVG-based radio buttons, consistent tooltip styling

### v2.1.0
- Perspective support for cloud connections (XMLA and Pro workspaces)
- `pbiServiceXmlaStyleLive` connection type for Pro workspace perspectives
- Cloud connection type selector (Semantic Model vs XMLA Endpoint)
- Automatic connection type selection based on perspective and workspace type

### v2.0.0
- Persistent cloud connection cache (survives app restart)
- Schema fingerprinting for change detection
- Preset schema validation with warning dialog
- Update/Apply Anyway/Cancel options for schema mismatches

### v1.0.0
- TOM-based connection swapping
- Thin report file modification
- Cloud workspace browser with OAuth
- Local model auto-matching
- Dual-scope preset system
- Backup and rollback capabilities
- Schema validation warnings
- Health monitoring

---

**Built by Reid Havens of Analytic Endeavors**
**Website**: https://www.analyticendeavors.com
**Email**: reid@analyticendeavors.com
