# Changelog

All notable changes to AE Power BI Multi-Tool Suite are documented here.

---

## [v2.0.0] - January 2026 - UI Redesign & Cleanup Edition

### Added
- **Dark/Light Mode** - Full theme support across all tools
- **Sidebar Navigation** - Redesigned for intuitive tool selection
- **Duplicate Image Detection** - MD5 hash detection with config dialog to select keeper
- **Unused Image Cleanup** - Find unreferenced images including orphan files
- **Saved Scripts Cleanup** - Remove DAX queries and TMDL scripts from semantic model
- **Multi-Diagram Optimization** - Optimize any saved diagram layout, not just default

### Changed
- **UI Consistent Theming** - Unified color system via ThemeManager
- **Layout Optimizer Scoring** - Relationship-aware 100-point scoring system with size-adjusted metrics

---

## [v1.3.1] - Cross-Page Bookmark Edition

### Added
- **Cross-Page Bookmark Mode** - Configure bookmarks to work across multiple pages without duplication
- **Cross-page UI indicator** - Link icon and page count display for cross-page bookmarks
- **Hover tooltips** - Shows which pages a cross-page bookmark spans
- **Bookmark navigator copying** - Navigators properly copied in cross-page mode

### Fixed
- Cross-page bookmarks now work on target pages (visualContainerGroups)
- Bookmarks no longer disappear after cross-page copy (comma-separated activeSection)

---

## [v1.3.0] - Cross-Page Bookmark Edition (Initial)

### Added
- Initial cross-page bookmark functionality

---

## [v1.2.0] - Sensitivity Scanner Edition

### Added
- **Sensitivity Scanner Tool** - Comprehensive TMDL scanning for sensitive data
- **Pattern-based detection** - 50+ patterns for PII, credentials, financial data, infrastructure
- **Rule Manager** - Visual pattern editor with Simple and Advanced modes
- **Custom date patterns** - User-friendly date format converter (dd/mm/yyyy to regex)
- **Pattern testing** - Real-time regex validation with test input
- **Pre-built templates** - Email, phone, credit card, IP, URLs, and 3 date formats

---

## [v1.1.1] - Advanced Copy Enhancement Edition

### Added
- **Enhanced page copying** - Multi-page support within/across PBIPs
- **Bookmark management** - Automatic bookmark and bookmark group duplication
- **Bookmark + Popup Copier** - Copy bookmarks with associated popup visuals
- **Action reassignment** - All items with actions pointing to copied pages get reassigned

---

## [v1.0.0] - Enhanced Report Cleanup Edition

### Added
- **Enhanced Table Column Widths** - Fit to Totals defaults with intelligent width calculation
- **Matrix optimization** - Hierarchy-aware spacing and improved compression

---

## [v0.0.0] - Beta Version

### Added
- **Multi-tool suite** with plugin architecture
- **Tool Manager** with automatic discovery
- **Report Merger** with conflict resolution and theme management
- **Advanced Page Copy** with dependency tracking
- **Layout Optimizer** with middle-out positioning
- **Report Cleanup** for optimization
- **Table Column Widths** tool
- Security-enhanced architecture throughout
- Professional UI with context-sensitive help

---

## Links

- **GitHub**: [analyticendeavors/pbi-multi-tool](https://github.com/analyticendeavors/pbi-multi-tool)
- **Website**: [analyticendeavors.com](https://www.analyticendeavors.com)
- **Support**: support@analyticendeavors.com
