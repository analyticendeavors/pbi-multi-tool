# Sensitivity Scanner - Technical Guide

**Tool Version:** 1.2.0  
**Last Updated:** November 17, 2025  
**Built by:** Reid Havens of Analytic Endeavors

---

## Table of Contents

1. [Overview](#overview)
2. [Architecture](#architecture)
3. [Core Components](#core-components)
4. [Pattern Detection System](#pattern-detection-system)
5. [Risk Scoring Engine](#risk-scoring-engine)
6. [Finding Deduplication](#finding-deduplication)
7. [TMDL Scanning](#tmdl-scanning)
8. [User Interface](#user-interface)
9. [Usage Guide](#usage-guide)
10. [Rule Manager](#rule-manager)
11. [Pattern Configuration](#pattern-configuration)
12. [Extending the Tool](#extending-the-tool)
13. [Troubleshooting](#troubleshooting)

---

## Overview

### What is the Sensitivity Scanner?

The **Sensitivity Scanner** is a static analysis tool that scans Power BI semantic model TMDL files for potentially sensitive content. It uses pattern matching to detect:

- **Personal Identifiable Information (PII)**: SSN, email, phone numbers
- **Financial Data**: Credit cards, bank accounts, salary information
- **Credentials**: Passwords, API keys, connection strings
- **Infrastructure Details**: Server names, IP addresses, file paths
- **Security Expressions**: Row-Level Security (RLS) definitions
- **Healthcare Data**: Medical record numbers, patient identifiers

### Key Features

✅ **Pattern-Based Detection** - Uses regex patterns to identify sensitive data  
✅ **Risk Assessment** - Categorizes findings as HIGH/MEDIUM/LOW risk  
✅ **Context-Aware Scoring** - Adjusts risk based on surrounding context  
✅ **Intelligent Deduplication** - Combines duplicate findings from same location  
✅ **Power BI Recommendations** - Provides actionable Power BI-specific guidance  
✅ **Rule Manager** - Visual pattern editor with Simple and Advanced modes ⭐ *Enhanced in v1.2.0*  
✅ **Custom Date Patterns** - User-friendly date format conversion (dd/mm/yyyy) ⭐ *New in v1.2.0*  
✅ **Pattern Testing** - Real-time regex validation with test input  
✅ **Export Reports** - Generates detailed findings reports  
✅ **Static Analysis** - Analyzes model structure, not actual data  
✅ **Flexible Scan Modes** - Full scan or targeted category scanning

### What It Does NOT Do

❌ Does not query or access actual data values  
❌ Does not modify PBIP files  
❌ Does not send data to external servers  
❌ Does not replace proper security practices  
❌ Does not guarantee 100% detection (pattern-based)  
❌ Does not perform runtime analysis or data profiling

---

## Architecture

### Component Overview

```
Sensitivity Scanner
├── Tool Class (sensitivity_scanner_tool.py)
│   └── BaseTool implementation
│
├── UI Layer (ui/sensitivity_scanner_tab.py)
│   ├── Combined file input + scan options
│   ├── Results tree view (expandable)
│   ├── Finding details panel (with border)
│   ├── Scan log panel (side-by-side)
│   └── Export functionality
│
└── Logic Layer (logic/)
    ├── models.py              # Data structures
    ├── pattern_detector.py    # Pattern matching engine
    ├── risk_scorer.py         # Risk assessment + deduplication
    └── tmdl_scanner.py        # TMDL file scanner
```

### Data Flow

```
User selects PBIP file
    ↓
TmdlScanner finds TMDL files
    ↓
PatternDetector scans content
    ↓
RiskScorer assesses findings
    ↓
Deduplication removes duplicates
    ↓
Results displayed in UI
    ↓
User exports report (optional)
```

---

## Core Components

### 1. Data Models (models.py)

Defines the data structures used throughout the tool:

#### PatternMatch
```python
@dataclass
class PatternMatch:
    pattern_id: str           # e.g., "ssn_us"
    pattern_name: str         # e.g., "US Social Security Number"
    matched_text: str         # The actual text that matched
    start_pos: int           # Character position in file
    end_pos: int
    line_number: int         # Line number
    context_before: str      # Text before match
    context_after: str       # Text after match
```

#### Finding
```python
@dataclass
class Finding:
    risk_level: RiskLevel              # HIGH/MEDIUM/LOW
    pattern_match: PatternMatch
    file_path: str                     # Path to TMDL file
    file_type: str                     # "Table TMDL", "RLS Role", etc.
    location_description: str          # Human-readable location
    categories: List[FindingCategory]  # [PII, FINANCIAL, etc.]
    confidence_score: float            # 0.0-1.0
    description: str                   # Why this matters
    recommendation: str                # What to do about it
    context_risk_adjustment: float     # +/- risk adjustment
```

#### ScanResult
```python
@dataclass
class ScanResult:
    scan_id: str
    scan_time: datetime
    source_path: str                   # PBIP file path
    findings: List[Finding]            # All findings (deduplicated)
    high_risk_count: int
    medium_risk_count: int
    low_risk_count: int
    total_files_scanned: int
    scan_duration_seconds: float
```

### 2. Pattern Detector (pattern_detector.py)

The pattern detection engine loads patterns from JSON and performs regex matching.

#### Key Methods

```python
class PatternDetector:
    def __init__(self, patterns_file: Optional[Path] = None):
        """Load patterns from sensitivity_patterns.json"""
    
    def scan_text(self, text: str, risk_levels: List[RiskLevel]) -> List[PatternMatch]:
        """Scan text for sensitive patterns"""
    
    def check_context_keywords(self, text: str) -> Dict[str, int]:
        """Check for risk-adjusting keywords in context"""
    
    def get_pattern_by_id(self, pattern_id: str) -> Optional[PatternDefinition]:
        """Retrieve pattern definition by ID"""
```

#### Pattern Loading

Patterns are loaded from `src/data/sensitivity_patterns.json`:

```json
{
  "patterns": {
    "high_risk": [...],
    "medium_risk": [...],
    "low_risk": [...]
  },
  "whitelist": {
    "patterns": [...]
  },
  "context_keywords": {
    "high_risk_context": [...],
    "low_risk_context": [...]
  }
}
```

### 3. Risk Scorer (risk_scorer.py)

Converts pattern matches into findings with risk assessment and deduplication.

#### Confidence Scoring

Confidence is calculated based on:
- Pattern type (high-confidence patterns like SSN get +0.3)
- Match length (longer API keys get +0.1)
- Context keywords (high-risk context adds +0.15)
- Whitelist patterns (low-risk context subtracts -0.3)

```python
def _calculate_confidence(self, pattern_match, pattern_def, full_context) -> float:
    confidence = 0.5  # Base
    
    # Adjust based on pattern type
    if pattern_def.id in ['ssn_us', 'credit_card', 'email_address']:
        confidence += 0.3
    
    # Check context keywords
    context_keywords = self.pattern_detector.check_context_keywords(full_context)
    if context_keywords['high_risk'] > 0:
        confidence += min(0.15, context_keywords['high_risk'] * 0.05)
    
    return max(0.0, min(1.0, confidence))
```

#### Risk Level Adjustment

Risk levels can be adjusted based on context and confidence:

```python
def _adjust_risk_level(self, base_risk, context_adjustment, confidence) -> RiskLevel:
    # Lower risk if low-risk context and low confidence
    if context_adjustment < -0.3 and confidence < 0.5:
        return lower_risk_level(base_risk)
    
    # Raise risk if high-risk context and high confidence
    elif context_adjustment > 0.2 and confidence > 0.7:
        return higher_risk_level(base_risk)
    
    return base_risk
```

#### Deduplication Logic (NEW in v1.2.0)

Risk scorer now includes intelligent deduplication to combine multiple findings from the same location:

```python
def _deduplicate_findings(self, findings: List[Finding]) -> List[Finding]:
    """
    Deduplicate findings from the same location.
    
    Strategy:
    - Group by file_path + line_number
    - For connection strings: Combine multiple credential findings
    - For other patterns: Keep highest risk finding per line
    """
    grouped = defaultdict(list)
    for finding in findings:
        key = f"{finding.file_path}:{finding.pattern_match.line_number}"
        grouped[key].append(finding)
    
    deduplicated = []
    for location_findings in grouped.values():
        if len(location_findings) == 1:
            deduplicated.append(location_findings[0])
        else:
            # Check for connection string credentials
            cred_patterns = {'connection_string_creds', 'password', 'api_key'}
            pattern_ids = {f.pattern_match.pattern_id for f in location_findings}
            
            if pattern_ids.intersection(cred_patterns):
                # Combine credential findings
                combined = self._combine_credential_findings(location_findings)
                deduplicated.append(combined)
            else:
                # Keep highest risk
                highest_risk = max(location_findings, 
                                 key=lambda f: risk_level_numeric(f.risk_level))
                deduplicated.append(highest_risk)
    
    return deduplicated
```

**Benefits of Deduplication:**
- Reduces noise in scan results
- Combines related findings (e.g., User ID + Password on same line)
- Clearer, more actionable reports
- Example: Instead of 3 separate findings for a connection string with User ID, Password, and Server, you get 1 combined finding

#### Recommendations

Risk scorer provides Power BI-specific recommendations:

- **Pattern-Specific**: Detailed guidance for specific patterns (e.g., "api_key")
- **Category-Based**: General guidance for categories (e.g., "pii", "credentials")
- **Risk-Based**: Fallback recommendations by risk level

Example recommendation for SSN:
```
CRITICAL: Remove Social Security Number.
Power BI Best Practice: If SSN data is required, implement
Row-Level Security (RLS) to restrict access by user. Consider
masking (e.g., showing only last 4 digits) or removing from the
semantic model entirely. Use object-level security to hide sensitive
columns (set IsHidden = true in TMDL).
```

### 4. TMDL Scanner (tmdl_scanner.py)

Orchestrates the scanning process.

#### Scan Modes

```python
# Full scan - all TMDL files
scanner.scan_pbip(pbip_folder)

# Scan specific files
scanner.scan_specific_files(pbip_folder, ['Sales', 'model'])

# Scan by category
scanner.scan_by_category(pbip_folder, 'tables')      # Tables only
scanner.scan_by_category(pbip_folder, 'roles')       # RLS roles only
scanner.scan_by_category(pbip_folder, 'expressions') # Power Query only
```

#### File Type Detection

Scanner determines file types for better context:

| File Pattern | Type | Description |
|--------------|------|-------------|
| `role_*` | RLS Role | Row-Level Security role definition |
| `expression_*` | Power Query Expression | Named Power Query expression |
| `model` | Model Definition | Model-level settings |
| `database` | Database Definition | Database-level settings |
| `relationships` | Relationships | Relationship definitions |
| Other | Table TMDL | Table, measures, columns |

---

## Finding Deduplication

### Why Deduplication?

Connection strings and complex expressions often trigger multiple patterns on the same line:

**Before Deduplication:**
```
Finding #8: Connection String with Credentials (Password=Secret123!)
Finding #9: Connection String with Credentials (User ID=admin)
Finding #10: Connection String with Credentials (Server=PRODDB01)
```

**After Deduplication:**
```
Finding #8: Connection string with multiple embedded credentials: 
           User ID, Password, Server
```

### Deduplication Strategy

1. **Group by Location**: Findings are grouped by `file_path:line_number`
2. **Single Finding**: If only one finding at a location, keep it
3. **Credential Patterns**: If multiple credential-related patterns, combine them
4. **Other Patterns**: Keep the highest risk finding

### Combined Finding Description

When combining credential findings:
```python
if len(credential_types) > 1:
    description = (
        f"Connection string with multiple embedded credentials: "
        f"{', '.join(credential_types)}"
    )
```

This provides a single, comprehensive finding instead of cluttering the report with duplicates.

---

## User Interface

### Compact Layout (NEW in v1.2.0)

The UI has been redesigned for optimal space usage:

```
┌─────────────────────────────────────────────────────────┐
│ 📁 SENSITIVITY SCANNER SETUP                            │
│ PBIP File: [____________] [Browse] [SCAN]              │
│ Scan Mode: ○ Full  ○ Tables  ○ RLS  ○ Expressions     │
├─────────────────────────────────────────────────────────┤
│ 📊 SCAN RESULTS (60% of vertical space)                │
│ ┌─────────────────────────────────────────────────────┐ │
│ │ 🔴 HIGH RISK (25)                                   │ │
│ │   ├─ Credit Card Number: 3333-4444...              │ │
│ │   ├─ Connection String: User ID, Password          │ │
│ │ 🟡 MEDIUM RISK (42)                                 │ │
│ │ 🟢 LOW RISK (15)                                    │ │
│ └─────────────────────────────────────────────────────┘ │
│ [EXPORT REPORT]                                         │
├──────────────────────┬──────────────────────────────────┤
│ 📝 FINDING DETAILS   │ 📋 SCAN LOG                     │
│ (40% vertical space) │ (40% vertical space)            │
│ [Details with        │ [Log messages with              │
│  black border]       │  black border]                  │
│                      │ [Export Log] [Clear Log]        │
└──────────────────────┴──────────────────────────────────┘
│ [RESET] [HELP]                                          │
└─────────────────────────────────────────────────────────┘
```

**Key UI Improvements:**
- ✅ Combined file input and scan options into one compact section
- ✅ Narrower scan button (12 characters wide)
- ✅ Horizontal radio buttons save vertical space
- ✅ Results and Details/Log sections use 60:40 split
- ✅ Finding Details has matching black border
- ✅ Side-by-side layout for Details and Log
- ✅ Clean radio buttons without gray background
- ✅ Responsive expanding sections

### Results Tree Columns

| Column | Description |
|--------|-------------|
| Finding | Pattern name and matched text |
| Risk | Risk level (HIGH/MEDIUM/LOW) |
| File Type | Type of TMDL file |
| Location | Human-readable location |
| Pattern | Pattern that matched |
| Confidence | Confidence score percentage |

### Finding Details Panel

Shows detailed information for selected finding:
- Risk level and confidence (color-coded)
- Pattern name and description
- File location
- Matched text
- Context (text before/after match)
- Why it matters (description)
- Recommended action (Power BI specific)

**Features:**
- Black border matching scan log aesthetic
- Expands vertically with window
- Syntax-highlighted risk levels
- Scrollable for long content

---

## Usage Guide

### Basic Workflow

1. **Select PBIP File**
   - Click Browse button
   - Choose a `.pbip` file
   - Tool automatically finds companion `.SemanticModel` folder

2. **Choose Scan Mode**
   - Full Scan: All TMDL files (default)
   - Tables Only: Just table definitions
   - RLS Only: Security roles
   - Expressions Only: Power Query expressions

3. **Run Scan**
   - Click "SCAN" button
   - Progress shown in progress bar
   - Log messages display scan activity
   - Results appear in tree view

4. **Review Findings**
   - Click on findings to see details
   - Expand HIGH/MEDIUM/LOW risk groups
   - Note risk levels and recommendations
   - Plan remediation actions

5. **Export Report**
   - Click "EXPORT REPORT"
   - Save detailed text report
   - Share with team/compliance
   - Document findings for audit trail

### Interpreting Results

#### High Risk Findings 🔴
- **Action**: Remove or protect immediately
- **Examples**: SSN, passwords, API keys, credit cards
- **Recommendation**: Use RLS, parameters, or remove entirely
- **Typical Confidence**: 70-95%

#### Medium Risk Findings 🟡
- **Action**: Review and apply appropriate security
- **Examples**: Salaries, server names, file paths, usernames
- **Recommendation**: Implement RLS, hide columns, use parameters
- **Typical Confidence**: 50-80%

#### Low Risk Findings 🟢
- **Action**: Verify appropriateness for context
- **Examples**: Name columns, department references, generic IDs
- **Recommendation**: Consider if suitable for audience
- **Typical Confidence**: 40-75%

### Common Scenarios

#### Scenario 1: Pre-Publication Audit
```
1. Scan completed model before sharing
2. Review all HIGH risk findings
3. Remove/protect sensitive data
4. Re-scan to verify fixes
5. Export final report for documentation
```

#### Scenario 2: Template Sanitization
```
1. Scan template model
2. Remove all real data references
3. Replace with generic examples
4. Update RLS for template users
5. Verify no sensitive patterns remain
```

#### Scenario 3: Compliance Review
```
1. Scan production model
2. Export detailed report
3. Review with compliance team
4. Document approved exceptions
5. Implement required protections
```

#### Scenario 4: Credential Detection
```
1. Run full scan
2. Filter to HIGH risk findings
3. Check for connection strings
4. Move credentials to parameters
5. Configure via gateway settings
```

---

## Rule Manager

### Overview

The **Rule Manager** is a visual pattern editor that allows users to create, modify, and test detection patterns without editing JSON files directly. It provides two modes: **Simple Mode** for common patterns and **Advanced Mode** for custom regex patterns.

### Accessing the Rule Manager

1. Launch the Sensitivity Scanner tool
2. Click the **"Manage Rules"** button (top-right of main interface)
3. Rule Manager window opens (1035px wide, optimized layout)

### Simple Mode

**Purpose:** Create patterns using user-friendly templates and converters.

#### Pattern Types

**1. Select from Template**

Pre-built patterns ready to use:
- ✅ **Email Address** - Standard email format validation
- ✅ **Phone Number (US)** - US phone number formats
- ✅ **Credit Card Number** - Credit card validation
- ✅ **IP Address (IPv4)** - IPv4 address detection
- ✅ **URL** - Web URL matching
- ✅ **Date (MM/DD/YYYY)** - US date format
- ✅ **Date (DD/MM/YYYY)** - European date format
- ✅ **Date (YYYY-MM-DD)** - ISO date format

**Usage:**
```
1. Select "Select from Template"
2. Choose template from dropdown
3. Pattern Name and Regex auto-populate
4. Click "Add Pattern"
```

**2. Custom Date Pattern** ⭐ *New in v1.2.0*

User-friendly date format converter:

**Supported Placeholders:**
- `dd` - Day (01-31)
- `mm` - Month (01-12)
- `yyyy` - 4-digit year (1900-2099)
- `yy` - 2-digit year (00-99)

**Supported Separators:**
- `-` (dash), `/` (slash), `.` (period), `_` (underscore), ` ` (space)

**Examples:**

| Input Format      | Generated Regex Pattern                                                    |
|-------------------|----------------------------------------------------------------------------|
| `dd/mm/yyyy`      | `\\b(?:0?[1-9]|[12][0-9]|3[01])/(?:0?[1-9]|1[0-2])/(?:19|20)\\d{2}\\b`    |
| `mm-dd-yyyy`      | `\\b(?:0?[1-9]|1[0-2])-(?:0?[1-9]|[12][0-9]|3[01])-(?:19|20)\\d{2}\\b`   |
| `yyyy.mm.dd`      | `\\b(?:19|20)\\d{2}\\.(?:0?[1-9]|1[0-2])\\.(?:0?[1-9]|[12][0-9]|3[01])\\b` |
| `dd-mm-yy`        | `\\b(?:0?[1-9]|[12][0-9]|3[01])-(?:0?[1-9]|1[0-2])-\\d{2}\\b`             |

**Usage:**
```
1. Select "Custom Date Pattern"
2. Type date format: "dd-mm-yyyy"
3. Pattern converts automatically
4. Test with sample dates
5. Click "Add Pattern"
```

**Key Features:**
- ✅ Leading zero optional for dd and mm (matches both "01" and "1")
- ✅ Boundary detection (`\\b`) prevents partial matches
- ✅ Smart validation for day (01-31) and month (01-12) ranges
- ✅ 4-digit years limited to 1900-2099 for realistic dates
- ✅ Case-insensitive input (DD/dd/Dd all work)

**3. Search Text**

Simple literal text search (case-insensitive):

**Usage:**
```
1. Select "Search Text"
2. Enter text to find: "confidential"
3. Pattern Name auto-generates
4. Click "Add Pattern"
```

**Behavior:**
- Searches for exact text (with word boundaries)
- Case-insensitive matching
- Generates pattern name like "Text: confidential"

#### Testing Patterns

**Real-time Pattern Testing:**
```
1. Enter pattern (template, custom date, or search text)
2. Type test input in "Test Input" field
3. Click "Test Pattern" button
4. Results show:
   ✅ "Pattern matches!" (green)
   ❌ "No match found" (red)
```

**Examples:**
```
Pattern: dd/mm/yyyy
Test: "Born on 15/08/1985"
Result: ✅ Pattern matches!

Pattern: Email Address template
Test: "Contact: john@example.com"
Result: ✅ Pattern matches!

Pattern: Search "SSN"
Test: "Employee SSN: 123-45-6789"
Result: ✅ Pattern matches!
```

### Advanced Mode

**Purpose:** Full control over regex patterns with custom configurations.

#### Fields

**1. Pattern Name** *(required)*
- Descriptive name for the pattern
- Appears in scan results
- Example: "Employee ID Format"

**2. Regex Pattern** *(required)*
- Full regex pattern with escaping
- Full-width field for complex patterns
- Example: `\\bEMP-\\d{6}\\b`

**3. Category** *(required)*
- Risk level: HIGH, MEDIUM, or LOW
- Determines finding priority
- Used in result filtering

**4. Example Value** *(optional)*
- Sample value for documentation
- Helps users understand pattern
- Example: "EMP-123456"

**5. Whitelist Patterns** *(optional)*
- Patterns to exclude from results
- Comma-separated list
- Example: `test, example, demo`

**6. Power BI Recommendation** *(optional)*
- Custom guidance for this finding type
- Appears in Finding Details panel
- Power BI-specific remediation steps

#### Advanced Pattern Testing

**Full-width Test Input:**
```
1. Enter complex regex in "Regex Pattern" field
2. Type test string in "Test Input" (full-width field)
3. Click "Test Pattern"
4. Validate before adding pattern
```

**Example - Testing Complex Pattern:**
```
Pattern: (?i)\\b(?:EMP|EMPLOYEE)[-_]?\\d{4,6}\\b
Test: "Records: EMP-12345, EMPLOYEE_987654"
Result: ✅ Pattern matches! (finds both IDs)
```

### Pattern Management

#### Current Patterns List

**Layout:**
- Fixed 425px width panel (left side)
- Displays all currently loaded patterns
- Format: `[CATEGORY] Pattern Name`
- Example: `[HIGH] Credit Card Number`

**Features:**
- ✅ **Selection** - Click to select pattern
- ✅ **Editing** - Double-click to edit selected pattern
- ✅ **Deletion** - Select + click "Delete Selected" button
- ✅ **Scrolling** - Mouse wheel scrolls patterns list only

#### Adding Patterns

**Workflow:**
```
1. Configure pattern (Simple or Advanced mode)
2. Test pattern with sample input
3. Click "Add Pattern" button
4. Pattern appears in Current Patterns list
5. Pattern saved to sensitivity_patterns.json
6. Available immediately for scanning
```

#### Editing Patterns

**Workflow:**
```
1. Select pattern in Current Patterns list
2. Click "Edit Selected" button (or double-click)
3. Pattern details load into form
4. Modify fields as needed
5. Click "Update Pattern"
6. Changes saved to sensitivity_patterns.json
```

#### Deleting Patterns

**Workflow:**
```
1. Select pattern in Current Patterns list
2. Click "Delete Selected" button
3. Confirmation dialog appears
4. Confirm deletion
5. Pattern removed from sensitivity_patterns.json
```

**Note:** Built-in patterns cannot be deleted (only custom patterns).

### Window Layout & UX

#### Optimized Dimensions

**Window Size:**
- Width: 1035px (optimized for no horizontal scroll)
- Height: 700px (sufficient vertical space)
- Resizable: Yes, but optimized for default size

**Panel Widths:**
- **Current Patterns:** 425px fixed width
- **Pattern Configuration:** Remaining space (~580px)
- **Splitter:** 30px between panels

#### Smart Scrolling

**Behavior:**
- Mouse wheel only scrolls the panel under cursor
- Current Patterns list: Independent scrolling
- Pattern Configuration form: Independent scrolling
- Prevents accidental scroll conflicts

**Implementation:**
```python
def _on_mouse_wheel(event, scrollable_frame):
    """Only scroll the frame under the cursor"""
    widget = event.widget
    if self._is_descendant(widget, scrollable_frame):
        canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        return "break"  # Prevent event propagation
```

#### Layout Stability

**No Width Jumping:**
- Pattern configuration fields maintain consistent width
- Switching pattern types doesn't resize panels
- Full-width fields (regex, test input) match form width
- Helper text adapts without layout shift

### Context-Aware Help

#### Helper Text Adaptation

**Select from Template:**
```
"Choose a pre-built pattern template. 
The pattern will be automatically generated."
```

**Custom Date Pattern:**
```
"Enter a date format using dd (day), mm (month), yyyy (4-digit year), or yy (2-digit year).
Example: dd/mm/yyyy or mm-dd-yyyy
Supported separators: - / . _ (space)"
```

**Search Text:**
```
"Enter text to search for (case-insensitive word match).
Example: 'confidential' or 'SSN'"
```

### Pattern Validation

#### Validation Rules

**Pattern Name:**
- ❌ Empty strings
- ❌ Whitespace-only names
- ✅ Must be descriptive

**Regex Pattern:**
- ❌ Invalid regex syntax
- ❌ Empty patterns
- ✅ Must compile successfully
- ✅ Tested before saving

**Category:**
- ✅ Must be valid FindingCategory enum value
- ✅ Options: HIGH, MEDIUM, LOW, PII, CREDENTIALS, etc.

**Test Input (optional):**
- Used for validation
- Not saved to pattern file
- Helps verify pattern correctness

#### Error Handling

**Invalid Regex:**
```
Error: Invalid regex pattern
Details: Unbalanced parenthesis at position 15
```

**Duplicate Pattern:**
```
Error: Pattern already exists
Suggestion: Use "Edit Selected" to modify existing pattern
```

**Invalid Category:**
```
Error: Invalid category 'CRITICAL'
Valid options: HIGH, MEDIUM, LOW, PII, CREDENTIALS, FINANCIAL, INFRASTRUCTURE, HEALTHCARE, COMPENSATION, PERSONAL, ORGANIZATIONAL, IDENTITY
```

### Best Practices

#### Pattern Design

1. **Use Word Boundaries** (`\\b`)
   - Prevents partial matches
   - Example: `\\bSSN\\b` matches "SSN" but not "ASSIGNMENT"

2. **Test Thoroughly**
   - Test with positive matches (should match)
   - Test with negative matches (should NOT match)
   - Use real-world examples

3. **Use Appropriate Risk Levels**
   - HIGH: PII, credentials, financial data
   - MEDIUM: Infrastructure, usernames, internal IDs
   - LOW: Organizational data, generic names

4. **Add Power BI Recommendations**
   - Provide actionable guidance
   - Reference Power BI features (RLS, parameters, gateway)
   - Include step-by-step remediation

#### Performance Optimization

1. **Avoid Greedy Quantifiers**
   - Use `.*?` instead of `.*` when possible
   - Prevents excessive backtracking

2. **Use Specific Patterns**
   - `\\d{4}` instead of `\\d+` for 4-digit numbers
   - More efficient matching

3. **Whitelist Aggressively**
   - Reduce false positives
   - Example: Whitelist "test", "example", "demo" for emails

### Troubleshooting

#### Common Issues

**Pattern doesn't match expected text**
- ✅ Use Test Pattern feature to validate
- ✅ Check for escaping issues (`\\` vs `\`)
- ✅ Verify word boundaries if using `\\b`

**Too many false positives**
- ✅ Add whitelist patterns
- ✅ Make regex more specific
- ✅ Use context-aware patterns

**Pattern causes slow scans**
- ✅ Avoid nested quantifiers like `(.*)*`
- ✅ Use atomic groups where possible
- ✅ Test with large TMDL files

**Custom date pattern not working**
- ✅ Check format uses supported placeholders (dd, mm, yyyy, yy)
- ✅ Verify separator is supported (- / . _ space)
- ✅ Ensure consistent casing (all lowercase or all uppercase)

---

## Pattern Configuration

### Adding New Patterns

To add a new pattern to `sensitivity_patterns.json`:

```json
{
  "id": "my_new_pattern",
  "name": "My Pattern Name",
  "pattern": "\\bpattern_regex\\b",
  "description": "What this pattern detects",
  "categories": ["category1", "category2"],
  "examples": ["example1", "example2"]
}
```

**Guidelines:**
- Use unique, descriptive IDs
- Provide clear pattern names
- Write accurate regex patterns
- Test with examples
- Choose appropriate categories
- Add to correct risk level section

### Pattern Categories

Available categories:
- `pii` - Personally Identifiable Information
- `financial` - Financial data
- `credentials` - Passwords, keys, tokens
- `government_id` - SSN, passport, etc.
- `contact` - Email, phone
- `infrastructure` - Servers, IPs, paths
- `security` - Security expressions, RLS
- `connection_string` - Database connections
- `dax` - DAX expressions
- `column_name` - Column naming patterns
- `organizational` - Dept, division, etc.
- `business` - Business entity references
- `file_path` - File system paths
- `healthcare` - Medical records, patient IDs

### Regex Best Practices

**DO:**
- Use word boundaries (`\b`) for whole-word matching
- Escape special characters (`\.`, `\(`, `\)`)
- Test patterns with examples
- Keep patterns specific but flexible
- Consider international formats

**DON'T:**
- Make patterns too broad (too many false positives)
- Make patterns too narrow (miss valid matches)
- Use greedy quantifiers without limits
- Forget to handle case sensitivity
- Assume only US formats (unless pattern-specific)

---

## Extending the Tool

### Adding New Risk Scorer Recommendations

In `risk_scorer.py`, add pattern-specific recommendations:

```python
def _get_pattern_specific_recommendations(self) -> dict:
    return {
        'my_pattern_id': (
            "CRITICAL: [What to do immediately]. "
            "Power BI Best Practice: [Specific Power BI guidance]. "
            "[Implementation details with examples]."
        ),
        # ... other patterns
    }
```

### Adding New Scan Modes

In `tmdl_scanner.py`, add new category filtering:

```python
def scan_by_category(self, pbip_folder: str, category: str):
    # ... existing code ...
    
    elif category == 'my_new_category':
        files_to_scan = {
            name: path for name, path in tmdl_files.items()
            if my_filter_condition(name)
        }
```

### Adding New UI Features

In `sensitivity_scanner_tab.py`, extend the UI:

```python
def _create_my_new_section(self):
    """Create a new UI section."""
    frame = ttk.LabelFrame(
        self.frame,
        text="🆕 MY NEW SECTION",
        style='Section.TLabelframe',
        padding="15"
    )
    frame.grid(row=X, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
    
    # Add components...
```

---

## Troubleshooting

### Common Issues

#### Issue: "PBIP Folder Not Found"
**Cause**: `.pbip` file missing companion `.SemanticModel` folder  
**Solution**: Ensure PBIP uses enhanced metadata format (PBIR)

#### Issue: "No TMDL files found"
**Cause**: PBIP folder missing `definition/` directory  
**Solution**: Verify PBIP format - must have TMDL structure

#### Issue: No findings but expected some
**Cause**: Patterns may not match your data format  
**Solution**: Check pattern regex, add custom patterns if needed

#### Issue: Too many duplicate findings
**Cause**: Multiple patterns matching same text  
**Solution**: Deduplication should handle this automatically (v1.2.0+)

#### Issue: Scan takes too long
**Cause**: Large model with many TMDL files  
**Solution**: Use category-based scanning to focus on specific areas

#### Issue: "Pattern 'travel' is not a valid FindingCategory"
**Cause**: Pattern file has category not in FindingCategory enum  
**Solution**: Update patterns.json or add category to models.py

### Performance Tips

- Use category scans for large models (100+ tables)
- Review patterns for overly complex regex
- Close other applications during scan
- Scan development models, not production (faster)
- Consider splitting very large models

### Debugging

Enable debug logging:
```python
import logging
logging.basicConfig(level=logging.DEBUG)
```

Check log output for:
- Pattern compilation errors
- File reading issues
- Match filtering details
- Confidence calculation steps
- Deduplication decisions

---

## Appendix

### File Structure
```
sensitivity_scanner/
├── __init__.py
├── sensitivity_scanner_tool.py
├── logic/
│   ├── __init__.py
│   ├── models.py
│   ├── pattern_detector.py
│   ├── risk_scorer.py           # Now includes deduplication
│   └── tmdl_scanner.py
├── ui/
│   ├── __init__.py
│   └── sensitivity_scanner_tab.py  # Compact responsive layout
└── TECHNICAL_GUIDE.md
```

### Dependencies
- Python 3.8+
- tkinter (standard library)
- pathlib (standard library)
- json (standard library)
- re (standard library)
- collections.defaultdict (for deduplication)
- Core PBIPReader (shared)
- Core UI Base classes

### Version History
- **1.0.0** (2024-11-14): Initial release
  - Pattern-based detection
  - Risk scoring
  - Basic UI
  - Export functionality
  
- **1.2.0** (2024-11-17): Major UI and functionality improvements
  - ✨ Intelligent finding deduplication
  - 🎨 Compact responsive UI layout
  - ✅ Side-by-side Details and Log panels
  - ✅ Clean radio buttons without background
  - ✅ Improved space allocation (60:40 split)
  - ✅ Finding Details panel with matching border
  - 🐛 Fixed PBIP folder detection for modern format
  - 🐛 Fixed export dialog parameter error

---

**Built by Reid Havens of Analytic Endeavors**  
**Website**: https://www.analyticendeavors.com  
**Email**: reid@analyticendeavors.com
