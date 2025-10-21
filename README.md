# AE Power BI Multi-Tool Suite

**Professional Power BI toolkit with plugin architecture for comprehensive report management**

![Version](https://img.shields.io/badge/version-1.1.1-blue)
![Python](https://img.shields.io/badge/python-3.8+-green)
![License](https://img.shields.io/badge/license-MIT-green)
![Platform](https://img.shields.io/badge/platform-Windows-lightgrey)
![Format](https://img.shields.io/badge/format-PBIP%20Only-orange)

---

## 🚀 What is AE Power BI Multi-Tool?

The **AE Power BI Multi-Tool** is a professional desktop application suite that provides multiple specialized tools for Power BI report management and optimization. Built with a modern plugin architecture, it automatically discovers and loads tools for comprehensive PBIP file manipulation.

**Built by [Reid Havens](https://linkedin.com/in/reid-havens) of [Analytic Endeavors](https://www.analyticendeavors.com)**

---

## 🛠️ Multi-Tool Suite Features

### 📊 **Report Merger Tool**
- **Intelligent report consolidation** with automatic conflict resolution
- **Theme management** and preservation of all visualizations
- **Smart bookmark and measure merging**
- **Professional validation and audit logging**

### 🎯 **Advanced Copy Tool** ⭐ *Enhanced in v1.1.1*
- **Multi-page copying** - Copy pages 1 or more times within a single PBIP or between PBIPs
- **Bookmark management** - Automatic bookmark and bookmark group duplication
- **Action reassignment** - All items with actions pointing to copied pages get reassigned to new bookmarks
- **Popup visual copying** - NEW! Copy bookmarks with their associated popup visuals
- **Cross-report support** - Works seamlessly within reports or between different PBIPs
- **Complete preservation** - Maintains all visual properties, themes, and interactions

### 📐 **PBIP Layout Optimizer**
- **Middle-out layout optimization** for model diagrams
- **Relationship-aware table positioning** with family grouping
- **Automatic table categorization** (fact, dimension, bridge, reference)
- **Chain-aware alignment** for optimal visual flow
- **Dimension optimization** with collision detection

### 🧹 **Report Cleanup Tool**
- **Comprehensive scanning** - Analyzes themes, custom visuals, bookmarks, and visual filters for usage
- **Theme cleanup** - Identifies and removes unused themes from BaseThemes and RegisteredResources
- **Custom visual management** - Distinguishes between AppSource visuals, build pane visuals, and hidden visuals
- **Smart usage detection** - Scans all pages and visuals to determine what's actually being used
- **Bookmark analysis** - Finds unused bookmarks and validates page references
- **Visual filter hiding** - Option to hide all visual-level filters across the report
- **Safe removal process** - Creates automatic backups and updates report.json references properly

### 📊 **Table Column Widths Tool** 
- **Visual scanning and analysis** - Discovers all Table (tableEx) and Matrix (pivotTable) visuals in PBIP reports
- **Field categorization** - Automatically identifies categorical fields vs measures from visual configurations
- **Width preset options** - Narrow, Medium, Wide, Auto-fit, Fit to Totals, and Custom width settings
- **Auto-fit calculations** - Font-aware width calculations based on actual text length and visual font settings
- **Fit to Totals optimization** - Specialized sizing for measures to prevent total/subtotal value wrapping
- **Enhanced matrix handling** - Hierarchy-aware spacing for Compact, Outline, and Tabular matrix layouts
- **Content-type detection** - Intelligent recognition of dates, currency, hierarchy levels, and data patterns
- **Selective application** - Choose individual visuals, multiple visuals, or apply to all tables/matrices

### 🔧 **Plugin Architecture**
- **Automatic tool discovery** and registration
- **Modular component system** with composition patterns
- **Extensible framework** for custom tools
- **Dynamic tab management** with context-sensitive help

---

## ✨ Enhanced Features & Intelligence

### 🎯 **Advanced Copy Intelligence** ⭐ *New in v1.1.1*
- **Multi-instance support** - Create multiple copies in one operation
- **Smart action mapping** - Automatically updates all action references to copied bookmarks
- **Popup preservation** - Maintains complex popup visual configurations across copies
- **Cross-PBIP intelligence** - Seamlessly handles dependencies when copying between reports

### 📊 **Advanced Matrix Handling**
- **Hierarchy-aware compression** (parent levels keep full width)
- **Layout-specific optimizations** (Compact vs Outline matrices)
- **Enhanced minimum widths** based on content and visual type
- **Currency-specific adjustments** for better readability
- **Date format intelligence** with complex pattern recognition

### 🎯 **Modern Plugin System**
- Automatic tool discovery from `/tools` directory
- Dynamic tab creation and management
- Component-based architecture for maximum reusability
- Enhanced error handling with comprehensive logging

### 🔒 **Enterprise-Grade Security**
- Professional security architecture
- Comprehensive audit logging
- Secure file handling with validation
- Thread-safe operations

### 🖥️ **Professional Interface**
- Clean, tabbed interface with tool-specific UIs
- Context-sensitive help system
- Real-time progress tracking
- Responsive design with dynamic height adjustment

### ⚡ **Performance Optimized**
- Efficient background processing
- Memory-conscious operations
- Optimized for large PBIP files
- Professional error recovery

---

## 📋 Requirements

### ⚠️ **Critical: PBIP Format Required**
This tool suite **ONLY** works with **PBIP files** in the enhanced report format (PBIR). 

**To enable PBIR format in Power BI Desktop:**
1. Go to **File** → **Options and settings** → **Options**
2. Select **Preview features**
3. Enable **"Store reports using enhanced metadata format"**
4. Restart Power BI Desktop
5. Save your reports - they will now be in .pbip format with a `.Report` folder

### 🖥️ **System Requirements**
- **Windows 10/11** (64-bit recommended)
- **Python 3.8+** (for running from source)
- **Power BI Desktop** (for creating PBIR files)
- **4GB RAM** minimum, 8GB recommended for large reports

---

## 🚀 Installation & Quick Start

1. **Download** `AE Power BI Multi-Tool.exe` from the [Releases page](../../releases)
2. **Run directly** - no installation required
3. **Launch** and select your tool from the tabbed interface

---

## 📖 Complete Tool Usage Guide

### 🏠 **Main Interface**
- **Tabbed design** with each tool in its own tab
- **Context-sensitive help** (❓ button adapts to current tool)
- **Professional branding** with company website access
- **Dynamic window sizing** optimized for each tool

### 📊 **Report Merger Tool**
1. **Select** Report A (.pbip file)
2. **Select** Report B (.pbip file)  
3. **Analyze** reports for conflicts and validation
4. **Choose** preferred theme if conflicts detected
5. **Execute** merge with real-time progress tracking
6. **Open** your consolidated report in Power BI Desktop

### 🎯 **Advanced Copy Tool** ⭐ *Enhanced in v1.1.1*

#### **Page Copying Mode**
1. **Choose** source report (.pbip file)
2. **Choose** destination report (.pbip file or same for within-PBIP copying)
3. **Select** specific pages to copy (single or multiple)
4. **Specify** number of copies (1 or more)
5. **Configure** copy options:
   - Include bookmarks and bookmark groups
   - Automatic action reassignment to copied bookmarks
   - Theme preservation
6. **Execute** copy operation with dependency tracking
7. **Review** copied pages with all bookmarks and actions properly reassigned

#### **Bookmark + Popup Copier Mode** 🆕
1. **Switch** to Bookmark/Popup Copier mode
2. **Choose** source report (.pbip file)
3. **Choose** destination report (.pbip file or same for within-PBIP copying)
4. **Select** bookmarks with associated popup visuals
5. **Execute** copy operation
6. **Review** copied bookmarks with complete popup visual preservation

**Use Cases:**
- Duplicate standardized pages within a report for different data views
- Copy popup templates from master reports to client reports
- Create multiple instances of interactive pages with unique bookmarks
- Replicate complex bookmark/popup combinations across reports

### 📐 **Layout Optimizer Tool**
1. **Select** PBIP file to optimize
2. **Choose** optimization strategy (Middle-Out recommended)
3. **Configure** canvas dimensions if needed
4. **Execute** optimization for improved model diagram layout
5. **Review** optimized table positioning and relationships

### 🧹 **Report Cleanup Tool**
1. **Select** report (.pbip file) to clean
2. **Choose** cleanup operations (unused measures, columns, etc.)
3. **Review** cleanup recommendations
4. **Execute** cleanup with backup creation
5. **Verify** optimized report performance

### 📊 **Table Column Widths Tool** ⭐ *Enhanced Edition*
1. **Select** PBIP file containing tables/matrices
2. **Scan visuals** to discover all table and matrix elements
3. **Set width preferences** (smart defaults):
   - **Categorical columns**: Auto-fit (intelligent width calculation)
   - **Measure columns**: Fit to Totals (prevents wrapping, default)
   - **Custom options**: Narrow/Medium/Wide presets or exact pixel values
4. **Select visuals** to update (tables, matrices, or both)
5. **Preview changes** to see calculated widths
6. **Apply changes** for professional column standardization

---

## 🏗️ Architecture Overview

### **Plugin-Based Design**
```
📁 AE Multi-Tool Structure:
   ├── 🏠 main.py (Tool Manager & Main Interface)
   ├── 🧠 core/ (Shared Components)
   │   ├── tool_manager.py (Auto-discovery)
   │   ├── enhanced_base_tool.py (Tool framework)
   │   ├── ui_base.py (UI components)
   │   └── constants.py (Shared constants)
   └── 🛠️ tools/ (Plugin Tools)
       ├── 📊 report_merger/
       ├── 🎯 advanced_copy/ (Enhanced in v1.1.1)
       ├── 📐 pbip_layout_optimizer/
       ├── 🧹 report_cleanup/
       └── 📊 column_width/
```

### **Component Composition**
- **ValidationComponent**: Input validation and file checking
- **FileInputComponent**: Path handling and file operations  
- **ThreadingComponent**: Background processing with proper closures
- **ProgressComponent**: User feedback and progress indication
- **BookmarkMapper**: Advanced bookmark tracking and action reassignment (New!)
- **PopupHandler**: Popup visual dependency management (New!)
- **IntelligentSizing**: Advanced width calculations with content-type detection

---

## ❓ Frequently Asked Questions

### **Q: What's new in v1.1.1?**
**A:** Version 1.1.1 significantly enhances the Advanced Copy Tool (formerly "Advanced Page Copy"). It now supports multi-page copying within or between PBIPs, automatic bookmark and action reassignment, and introduces a brand new Bookmark + Popup Copier for copying bookmarks with their associated popup visuals.

### **Q: What's the difference between page copying and bookmark/popup copying?**
**A:** Page copying duplicates entire report pages with all their content, bookmarks, and actions. Bookmark/popup copying specifically targets bookmarks that have associated popup visuals, perfect for replicating interactive elements across reports without copying entire pages.

### **Q: Can I copy pages multiple times in one operation?**
**A:** Yes! In v1.1.1, you can specify how many copies you want (1 or more), and the tool will create them all with properly reassigned bookmarks and actions.

### **Q: What's the difference between this and individual tools?**
**A:** This is a unified suite with shared components, consistent UI, and plugin architecture. All tools share security features, logging, and base functionality.

### **Q: Can I use just one tool from the suite?**
**A:** Yes! Each tool operates independently within the tabbed interface. Just use the tab you need.

### **Q: What file formats are supported?**
**A:** Only PBIP files in enhanced report format (PBIR). Traditional .pbix files are NOT supported by any tool in the suite.

### **Q: How do I get the best results from Table Column Widths?**
**A:** Use the default settings (Auto-fit for categorical, Fit to Totals for measures) for the most intelligent width calculations with optimal total display. These defaults prevent text wrapping in totals while providing smart sizing for field headers.

---

## 🛠️ Troubleshooting

### **Common Issues**

**"File format not supported"**
- Ensure you're using .pbip files (not .pbix)
- Enable PBIR format in Power BI Desktop settings
- Verify .Report folders exist alongside .pbip files

**"No visuals found" in Table Column Widths**
- Tool only works with Table and Matrix visuals
- Ensure your report contains tableEx or pivotTable visual types
- Check that visuals have proper field configurations

**"Bookmark references broken" after copying**
- Ensure you're using v1.1.1 or later with enhanced action reassignment
- Check that all bookmarks were properly selected during copy operation
- Verify source report bookmarks are not corrupted

**"No tools discovered"**
- Check that `/tools` directory exists and contains tool modules
- Verify Python import paths are correct
- Run dependency checks to verify installations

---

## 🔒 Security & Privacy

### **Data Protection**
- **Completely offline** - no data transmission
- **No telemetry or tracking** - privacy-first design
- **Comprehensive audit logging** for enterprise compliance
- **Security-enhanced architecture** throughout all tools

### **Professional Standards**
- Code signing available for enterprise deployment
- Regular security reviews and updates
- Professional error handling and recovery
- Secure file operations with validation

---

## 📞 Support & Community

### **Getting Help**
- 📚 **Documentation**: Check our [Wiki](../../wiki) for detailed tool guides
- 📋 **Release Notes**: See [VERSION_1.1.1_RELEASE_NOTES.md](VERSION_1.1.1_RELEASE_NOTES.md) for latest updates
- 🐛 **Bug Reports**: Use [Issues](../../issues) to report problems
- 💡 **Feature Requests**: Submit via [Issues](../../issues) with enhancement label
- 🌐 **Professional Support**: [Analytic Endeavors](https://www.analyticendeavors.com)
- 📧 **Direct Contact**: support@analyticendeavors.com

### **Contributing**
We welcome contributions to the tool suite! Please see our [Contributing Guide](CONTRIBUTING.md) for details on:
- Adding new tools to the plugin system
- Enhancing existing components
- Improving the AI intelligence systems
- Documentation updates

---

## ⚖️ Legal & Disclaimers

### **Important Notices**
- **Independent tool suite** - not officially supported by Microsoft
- **Use at your own risk** - always test thoroughly in development environments
- **Keep backups** of original reports before performing operations
- **PBIR format required** - traditional Power BI files not supported
- **Microsoft Power BI** is a trademark of Microsoft Corporation

### **License**
This software is released under the MIT License. See [LICENSE](LICENSE.txt) for full terms and conditions.

---

## 🏢 About Analytic Endeavors

**AE Power BI Multi-Tool** is developed by [Analytic Endeavors](https://www.analyticendeavors.com), a consulting firm specializing in business intelligence and advanced Power BI solutions.

**Founded by Reid Havens & Steve Campbell**, we create professional tools and provide expert consulting services for organizations looking to maximize their Power BI investments.

### **Our Services**
- **Power BI consulting** and custom development
- **Advanced analytics** and data architecture
- **Custom tool development** for enterprise workflows
- **Training and workshops** for Power BI optimization
- **Enterprise BI strategy** and implementation

### **Why Choose Our Tools?**
- **Professional quality** with enterprise-grade architecture
- **AI-enhanced intelligence** for smarter automation
- **Real-world tested** in production environments
- **Continuous improvement** based on client feedback
- **Expert support** from certified Power BI professionals

---

## 🔄 Version History

### **v1.1.1 - Advanced Copy Enhancement Edition** ⭐ *Current*
- ✅ **Enhanced page copying** - Multi-page support within/across PBIPs
- ✅ **Bookmark management** - Automatic bookmark and bookmark group duplication
- ✅ **Action reassignment** - All action references automatically updated to copied bookmarks
- ✅ **NEW: Bookmark + Popup Copier** - Copy bookmarks with associated popup visuals
- ✅ **Tool renamed** - "Advanced Page Copy" → "Advanced Copy"
- 🐛 **Bug fixes** - Improved bookmark reference handling and action mapping

### **v1.0.0 - Enhanced Report Cleanup Edition**
- ✅ **Enhanced Table Column Widths** - Fit to Totals defaults with intelligent width calculation
- ✅ **Matrix optimization** - Hierarchy-aware spacing and improved compression
- ✅ **Content detection** - Smart recognition of dates, currency, hierarchy levels
- ✅ **Complex date support** - Enhanced handling for "EOW: 01/07/24" formats
- ✅ **Professional UI** - Streamlined configuration with intelligent defaults

### **v0.0.0 - Beta Version**
- ✅ **Multi-tool suite** with plugin architecture
- ✅ **Tool Manager** with automatic discovery
- ✅ **Layout Optimizer** with middle-out positioning
- ✅ **Advanced Page Copy** with dependency tracking
- ✅ **Enhanced Report Merger** with improved conflict resolution
- ✅ **Report Cleanup** tool for optimization
- ✅ **Security-enhanced** architecture throughout
- ✅ **Professional UI** with context-sensitive help

---

## 🌟 Show Your Support

If you find this tool suite useful:
- ⭐ **Star this repository** to help others discover it
- 🐦 **Share** with your Power BI community
- 💼 **Connect** with us on [LinkedIn](https://linkedin.com/company/analytic-endeavors)
- 🌐 **Visit** our website: [analyticendeavors.com](https://www.analyticendeavors.com)
- 📧 **Subscribe** to our newsletter for Power BI tips and tool updates

---

**Made with ❤️ and 🧠 AI by [Reid Havens](https://www.analyticendeavors.com) for the Power BI professional community**

*Empowering data professionals with AI-enhanced, enterprise-grade tools for Power BI excellence*
