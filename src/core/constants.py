"""
Power BI Report Merger - Application Constants
Built by Reid Havens of Analytic Endeavors
"""

class AppConstants:
    """Essential application constants - simplified and optimized."""
    
    # =============================================================================
    # CORE APPLICATION INFO
    # =============================================================================
    
    # Company & Branding
    COMPANY_NAME = "Analytic Endeavors"
    COMPANY_FOUNDER = "Reid Havens"
    COMPANY_WEBSITE = "https://www.analyticendeavors.com"
    
    # Application
    APP_NAME = "Power BI Report Merger"
    APP_VERSION = "v1.0 Enhanced"
    WINDOW_TITLE = f"{COMPANY_NAME} - {APP_NAME}"
    WINDOW_SIZE = "1000x1040"
    MIN_WINDOW_SIZE = (900, 700)
    
    # =============================================================================
    # UI STYLING - SIMPLIFIED PROFESSIONAL THEME
    # =============================================================================
    
    COLORS = {
        # Primary colors
        'primary': '#066c7c',
        'secondary': '#0891a5',
        'accent': '#10b9d1',
        
        # Status colors
        'success': '#059669',
        'warning': '#d97706',
        'error': '#dc2626',
        'info': '#2563eb',
        
        # Interface colors
        'background': '#f8fafc',
        'surface': '#ffffff',
        'border': '#e2e8f0',
        'text_primary': '#1e293b',
        'text_secondary': '#64748b',
        'text_muted': '#94a3b8'
    }
    
    # =============================================================================
    # UI TEXT CONTENT
    # =============================================================================
    
    # Header text
    BRAND_TEXT = f"📊 {COMPANY_NAME.upper()}"
    MAIN_TITLE = APP_NAME
    TAGLINE = "Professional-grade tool for intelligent PBIP report consolidation"
    BUILT_BY_TEXT = f"Built by {COMPANY_FOUNDER} of {COMPANY_NAME}"
    
    # Quick start steps (embedded in UI)
    QUICK_START_STEPS = [
        "1. Navigate to your .pbip file in File Explorer",
        "2. Right-click the .pbip file and select 'Copy as path'", 
        "3. Paste (Ctrl+V) into the path field above",
        "4. Path quotes will be automatically cleaned",
        "5. Repeat for the second report file",
        "6. Click 'Analyze Reports' to begin"
    ]
    
    # =============================================================================
    # TECHNICAL CONFIGURATION
    # =============================================================================
    
    # Power BI schema URLs (only the ones actually used)
    SCHEMA_URLS = {
        'platform': "https://developer.microsoft.com/json-schemas/fabric/item/platformMetadata/1.0.0/schema.json",
        'pbip': "https://developer.microsoft.com/json-schemas/fabric/pbip/pbipProperties/1.0.0/schema.json",
        'bookmarks': "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/bookmarksMetadata/1.0.0/schema.json",
        'pages': "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/pagesMetadata/1.0.0/schema.json",
        'report_extension': "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/reportExtension/1.0.0/schema.json"
    }
    
    # File settings
    SUPPORTED_EXTENSIONS = ['.pbip']
    MAX_CACHE_SIZE = 100


# Export main constants class
__all__ = ['AppConstants']
