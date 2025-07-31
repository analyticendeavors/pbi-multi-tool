# AE Power BI Multi-Tool Suite

**Professional Power BI toolkit with plugin architecture for comprehensive report management**

![Version](https://img.shields.io/badge/version-2.0.0-blue)
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

### 🎯 **Advanced Page Copy Tool**
- **Cross-report page copying** with dependency tracking
- **Visual element preservation** including custom themes
- **Bookmark and navigation maintenance**
- **Selective content transfer capabilities**

### 📐 **PBIP Layout Optimizer**
- **Middle-out layout optimization** for model diagrams
- **Relationship-aware table positioning** with family grouping
- **Automatic table categorization** (fact, dimension, bridge, reference)
- **Chain-aware alignment** for optimal visual flow
- **Dimension optimization** with collision detection

### 🔧 **Plugin Architecture**
- **Automatic tool discovery** and registration
- **Modular component system** with composition patterns
- **Extensible framework** for custom tools
- **Dynamic tab management** with context-sensitive help

---

## ✨ Key Architecture Features

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
3. Enable **\"Store reports using enhanced metadata format\"**
4. Restart Power BI Desktop
5. Save your reports - they will now be in .pbip format with a `.Report` folder

### 🖥️ **System Requirements**
- **Windows 10/11** (64-bit recommended)
- **Python 3.8+** (for running from source)
- **Power BI Desktop** (for creating PBIR files)
- **4GB RAM** minimum, 8GB recommended for large reports

---

## 🚀 Installation & Quick Start

### Option 1: Standalone Executable (Recommended)
1. **Download** `AE Power BI Multi-Tool.exe` from the [Releases page](../../releases)
2. **Run directly** - no installation required
3. **Launch** and select your tool from the tabbed interface

### Option 2: Run from Source
1. **Clone** this repository
2. **Install** Python 3.8+ from [python.org](https://python.org/downloads)
3. **Navigate** to the `builds` directory
4. **Run** `build_ae_pbi_multi_tool.bat` to build your own executable
5. **Or run directly** from source using `src/run_ae_multi_tool.bat`

---

## 📖 Tool Usage Guide

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

### 🎯 **Advanced Page Copy Tool**
1. **Choose** source report (.pbip file)
2. **Choose** destination report (.pbip file)
3. **Select** specific pages to copy
4. **Configure** copy options (themes, bookmarks, etc.)
5. **Execute** copy operation with dependency tracking

### 📐 **Layout Optimizer Tool**
1. **Select** PBIP file to optimize
2. **Choose** optimization strategy (Middle-Out recommended)
3. **Configure** canvas dimensions if needed
4. **Execute** optimization for improved model diagram layout
5. **Review** optimized table positioning and relationships

---

## 🏗️ Architecture Overview

### **Plugin-Based Design**
```
📁 AE Multi-Tool Structure:
   ├── 🏠 main.py (Tool Manager & Main Interface)
   ├── 🧠 core/ (Shared Components)
   │   ├── tool_manager.py (Auto-discovery)
   │   ├── enhanced_base_tool.py (Tool framework)
   │   └── composition/ (Component system)
   └── 🛠️ tools/ (Plugin Tools)
       ├── report_merger/
       ├── page_copy/
       └── pbip_layout_optimizer/
```

### **Component Composition**
- **ValidationComponent**: Input validation and file checking
- **FileInputComponent**: Path handling and file operations  
- **ThreadingComponent**: Background processing with proper closures
- **ProgressComponent**: User feedback and progress indication

---

## 🔧 Development & Extension

### **Adding New Tools**
1. Create new directory under `/tools`
2. Implement tool class inheriting from `EnhancedBaseExternalTool`
3. Tool Manager automatically discovers and registers
4. Add UI components using composition framework

### **Custom Components**
```python
from core.composition import ToolComponent

class MyComponent(ToolComponent):
    def initialize(self) -> bool:
        # Custom logic here
        self.mark_initialized()
        return True
```

### **Plugin Architecture Benefits**
- **Modular development** - each tool is independent
- **Easy testing** - components can be tested in isolation
- **Extensible design** - add tools without modifying core
- **Professional structure** - enterprise-grade codebase organization

---

## ❓ Frequently Asked Questions

### **Q: What's the difference between this and individual tools?**
**A:** This is a unified suite with shared components, consistent UI, and plugin architecture. All tools share security features, logging, and base functionality.

### **Q: Can I use just one tool from the suite?**
**A:** Yes! Each tool operates independently within the tabbed interface. Just use the tab you need.

### **Q: How do I add custom tools?**
**A:** Create a new tool class in the `/tools` directory following the plugin pattern. The Tool Manager will automatically discover it.

### **Q: What file formats are supported?**
**A:** Only PBIP files in enhanced report format (PBIR). Traditional .pbix files are NOT supported by any tool in the suite.

### **Q: Will this integrate with Power BI Desktop?**
**A:** Yes, the suite can be configured as an External Tool in Power BI Desktop for seamless workflow integration.

---

## 🛠️ Troubleshooting

### **Common Issues**

**\"File format not supported\"**
- Ensure you're using .pbip files (not .pbix)
- Enable PBIR format in Power BI Desktop settings
- Verify .Report folders exist alongside .pbip files

**\"No tools discovered\"**
- Check that `/tools` directory exists and contains tool modules
- Verify Python import paths are correct
- Run `test_imports.py` to verify dependencies

**\"Tool initialization failed\"**
- Check the application logs for specific error details
- Verify all required Python packages are installed
- Ensure file permissions allow tool operation

**Security warnings**
- The executable is built with security-enhanced architecture
- Some antivirus software may flag new applications
- All operations are logged for audit compliance

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
- 🐛 **Bug Reports**: Use [Issues](../../issues) to report problems
- 💡 **Feature Requests**: Submit via [Issues](../../issues) with enhancement label
- 🌐 **Professional Support**: [Analytic Endeavors](https://www.analyticendeavors.com)
- 📧 **Direct Contact**: support@analyticendeavors.com

### **Contributing**
We welcome contributions to the tool suite! Please see our [Contributing Guide](CONTRIBUTING.md) for details on:
- Adding new tools to the plugin system
- Enhancing existing components
- Improving the architecture
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
This software is released under the MIT License. See [LICENSE](LICENSE) for full terms and conditions.

---

## 🏢 About Analytic Endeavors

**AE Power BI Multi-Tool** is developed by [Analytic Endeavors](https://www.analyticendeavors.com), a consulting firm specializing in business intelligence and advanced Power BI solutions.

**Founded by Reid Havens**, we create professional tools and provide expert consulting services for organizations looking to maximize their Power BI investments.

### **Our Services**
- **Power BI consulting** and custom development
- **Advanced analytics** and data architecture
- **Custom tool development** for enterprise workflows
- **Training and workshops** for Power BI optimization
- **Enterprise BI strategy** and implementation

### **Why Choose Our Tools?**
- **Professional quality** with enterprise-grade architecture
- **Real-world tested** in production environments
- **Continuous improvement** based on client feedback
- **Expert support** from certified Power BI professionals

---

## 🔄 Version History

### **v2.0.0 - Plugin Architecture Edition**
- ✅ **Multi-tool suite** with plugin architecture
- ✅ **Tool Manager** with automatic discovery
- ✅ **Layout Optimizer** with middle-out positioning
- ✅ **Advanced Page Copy** with dependency tracking
- ✅ **Enhanced Report Merger** with improved conflict resolution
- ✅ **Composition framework** for modular development
- ✅ **Security-enhanced** architecture throughout
- ✅ **Professional UI** with context-sensitive help

### **v1.0.0 - Foundation Release**
- ✅ Initial report merger functionality
- ✅ PBIR format support
- ✅ Basic UI framework
- ✅ Core security features

---

## 🎯 Roadmap & Future Tools

### **Planned Tools**
- **📈 Report Analytics Tool** - Comprehensive report analysis and metrics
- **🔄 Model Optimizer** - Data model performance optimization
- **📋 Documentation Generator** - Automatic report documentation
- **🎨 Theme Manager** - Advanced theme management and customization
- **🔍 Content Auditor** - Report content validation and compliance checking

### **Architecture Enhancements**
- **Configuration-driven** tool assembly
- **Plugin marketplace** integration
- **Advanced logging** and monitoring
- **API integration** for enterprise workflows

---

## 🌟 Show Your Support

If you find this tool suite useful:
- ⭐ **Star this repository** to help others discover it
- 🐦 **Share** with your Power BI community
- 💼 **Connect** with us on [LinkedIn](https://linkedin.com/company/analytic-endeavors)
- 🌐 **Visit** our website: [analyticendeavors.com](https://www.analyticendeavors.com)
- 📧 **Subscribe** to our newsletter for Power BI tips and tool updates

---

**Made with ❤️ by [Reid Havens](https://www.analyticendeavors.com) for the Power BI professional community**

*Empowering data professionals with enterprise-grade tools for Power BI excellence*"