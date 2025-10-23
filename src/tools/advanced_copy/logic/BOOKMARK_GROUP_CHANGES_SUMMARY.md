# Bookmark Group vs Individual Visual Detection - Changes Summary

## Date: October 22, 2025
## Files Modified: `advanced_copy_bookmark_analyzer.py`

---

## ✅ Changes Made

### 1. Updated `extract_bookmark_visuals()` Method

**Old Signature:**
```python
def extract_bookmark_visuals(self, bookmark: Dict) -> List[str]:
```

**New Signature:**
```python
def extract_bookmark_visuals(self, bookmark: Dict, page_dir: Path) -> Dict[str, List[str]]:
```

**New Return Structure:**
```python
{
    'group_ids': [list of group container IDs],
    'individual_visual_ids': [list of individual visual IDs],
    'all_visual_ids': [combined list for backwards compatibility]
}
```

**Detection Logic:**
- Scans `visualContainerGroups` in bookmark's `explorationState` to identify group toggles
- For each visual in `targetVisualNames`, reads the visual's JSON file
- Categorizes based on presence of `visualGroup` property:
  - **Has `visualGroup`** → Group container
  - **No `visualGroup`** → Individual visual

---

### 2. Added `get_all_group_members()` Helper Method

**Purpose:** Get current members of a group at copy-time (not from bookmark's snapshot)

**Signature:**
```python
def get_all_group_members(self, page_dir: Path, group_id: str) -> List[str]:
```

**Logic:**
- Scans all visuals in page's `visuals/` directory
- Returns visual IDs where `parentGroupName` matches the `group_id`
- Provides dynamic, up-to-date group membership

---

## 🎯 How This Solves The Problem

### Before (Problem):
- Bookmarks store `targetVisualNames` as a snapshot at creation time
- If group gains new visuals later → Not copied (breaks bookmark)
- Couldn't distinguish group toggles from individual selections

### After (Solution):
**For Group Toggles:**
- Bookmark is identified as targeting a group (via `visualContainerGroups`)
- Call `get_all_group_members(page_dir, group_id)` at copy-time
- Copy ALL current members, regardless of what was in `targetVisualNames`
- ✅ New visuals added to group → Automatically included
- ✅ Visuals removed from group → Safely ignored

**For Individual Visuals:**
- Bookmark is identified as targeting specific visuals (not in `visualContainerGroups`)
- Copy only those specific visual IDs from `targetVisualNames`
- ✅ Precise control maintained

---

## 📋 Usage Example

```python
# Initialize analyzer
analyzer = BookmarkAnalyzer(logger_callback=log)

# Extract and categorize visuals
bookmark_info = analyzer.extract_bookmark_visuals(bookmark_data, page_dir)

# Process group toggles (copy ALL current members)
for group_id in bookmark_info['group_ids']:
    current_members = analyzer.get_all_group_members(page_dir, group_id)
    for member_id in current_members:
        copy_visual(member_id)  # Copy each current group member

# Process individual visuals (copy only specific ones)
for visual_id in bookmark_info['individual_visual_ids']:
    copy_visual(visual_id)  # Copy only this specific visual
```

---

## 🔒 Backwards Compatibility

- `all_visual_ids` field provides combined list (same as old behavior)
- Methods are currently unused in codebase → Zero risk of breaking existing code
- Safe fallbacks if visual files can't be read (treats as individual visual)

---

## 🧪 Testing Scenarios

1. **Group toggle bookmark** - Add visual to group after bookmark creation
2. **Individual visual bookmark** - Verify only that visual is identified
3. **Mixed bookmark** - Group toggle + individual visuals in same bookmark
4. **Removed visuals** - Visual removed from group (no error)
5. **Missing visual files** - Graceful fallback to individual visual

---

## 📊 Detection Flow

```
Bookmark JSON
    ↓
Check visualContainerGroups
    ↓
    ├─ Has group ID? → GROUP TOGGLE
    │   ↓
    │   Call get_all_group_members(page_dir, group_id)
    │   → Returns current members
    │   → Copy all current members
    │
    └─ No group ID? → INDIVIDUAL VISUAL
        ↓
        Use targetVisualNames directly
        → Copy only specified visual IDs
```

---

## ✨ Benefits

✅ Groups stay current - New visuals automatically included  
✅ No breaking changes - Unused methods, safe to update  
✅ Robust - Handles removed visuals without errors  
✅ Smart detection - Distinguishes groups vs individuals  
✅ Backwards compatible - `all_visual_ids` field available  

---

*Built by Reid Havens of Analytic Endeavors*
