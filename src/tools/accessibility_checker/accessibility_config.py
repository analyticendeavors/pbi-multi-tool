"""
Accessibility Checker - Configuration Management
Built by Reid Havens of Analytic Endeavors

Configuration for accessibility check rules, severity filtering, and contrast levels.
Settings are persisted to JSON for user preference retention.
"""

import json
import os
from pathlib import Path
from dataclasses import dataclass, field, asdict
from typing import Dict, Optional

# Ownership fingerprint
_AE_FP = "QWNjZXNzQ29uZmlnOkFFLTIwMjQ="


# Contrast level thresholds per WCAG standards
CONTRAST_THRESHOLDS = {
    "AA": {"normal": 4.5, "large": 3.0},      # WCAG 2.1 Level AA
    "AAA": {"normal": 7.0, "large": 4.5},     # WCAG 2.1 Level AAA
    "AA_large": {"normal": 3.0, "large": 3.0} # Treats all text as large (relaxed)
}


def get_config_path() -> Path:
    """Get the path to the config file in AppData"""
    if os.name == 'nt':  # Windows
        appdata = os.environ.get('APPDATA', os.path.expanduser('~'))
        config_dir = Path(appdata) / 'AE-Multi-Tool'
    else:  # Linux/Mac
        config_dir = Path.home() / '.config' / 'ae-multi-tool'

    config_dir.mkdir(parents=True, exist_ok=True)
    return config_dir / 'accessibility_config.json'


@dataclass
class AccessibilityCheckConfig:
    """Configuration for accessibility checks"""

    # Which checks to run (all enabled by default)
    enabled_checks: Dict[str, bool] = field(default_factory=lambda: {
        "tab_order": True,
        "alt_text": True,
        "color_contrast": True,
        "page_title": True,
        "visual_title": True,
        "data_labels": True,
        "bookmark_name": True,
        "hidden_page": True,
    })

    # Severity filtering for display (show issues at this level and above)
    # Values: "error" (Critical only), "warning" (Critical + Should Fix), "info" (All)
    min_severity: str = "info"

    # Contrast level requirement
    # Values: "AA" (4.5:1), "AAA" (7:1), "AA_large" (3:1 - relaxed)
    contrast_level: str = "AA"

    # Whether to report AAA failures as "Review" issues
    flag_aaa_failures: bool = True

    # Whether to report AA failures as "Review" issues (when using AA_large threshold)
    flag_aa_failures: bool = False

    def save(self, path: Optional[Path] = None) -> None:
        """Save config to JSON file"""
        if path is None:
            path = get_config_path()

        try:
            with open(path, 'w', encoding='utf-8') as f:
                json.dump(asdict(self), f, indent=2)
        except Exception as e:
            print(f"Warning: Could not save accessibility config: {e}")

    @classmethod
    def load(cls, path: Optional[Path] = None) -> 'AccessibilityCheckConfig':
        """Load config from JSON file, returns defaults if not found"""
        if path is None:
            path = get_config_path()

        try:
            if path.exists():
                with open(path, 'r', encoding='utf-8') as f:
                    data = json.load(f)

                # Handle potential missing keys from older config files
                config = cls()
                if 'enabled_checks' in data:
                    config.enabled_checks.update(data['enabled_checks'])
                if 'min_severity' in data:
                    config.min_severity = data['min_severity']
                if 'contrast_level' in data:
                    config.contrast_level = data['contrast_level']
                if 'flag_aaa_failures' in data:
                    config.flag_aaa_failures = data['flag_aaa_failures']
                if 'flag_aa_failures' in data:
                    config.flag_aa_failures = data['flag_aa_failures']

                return config
        except Exception as e:
            print(f"Warning: Could not load accessibility config: {e}")

        return cls()  # Return defaults

    def get_contrast_threshold(self, text_size: str = "normal") -> float:
        """Get the contrast threshold based on current settings

        Args:
            text_size: "normal" or "large"

        Returns:
            The required contrast ratio
        """
        thresholds = CONTRAST_THRESHOLDS.get(self.contrast_level, CONTRAST_THRESHOLDS["AA"])
        return thresholds.get(text_size, thresholds["normal"])

    def is_check_enabled(self, check_type: str) -> bool:
        """Check if a specific check type is enabled"""
        return self.enabled_checks.get(check_type, True)

    def should_show_severity(self, severity: str) -> bool:
        """Determine if an issue should be shown based on severity filter

        Args:
            severity: "error", "warning", or "info"

        Returns:
            True if the issue should be displayed
        """
        severity_order = {"error": 3, "warning": 2, "info": 1}
        min_level = severity_order.get(self.min_severity, 1)
        issue_level = severity_order.get(severity, 1)
        return issue_level >= min_level


# Singleton instance for easy access
_config_instance: Optional[AccessibilityCheckConfig] = None


def get_config() -> AccessibilityCheckConfig:
    """Get the global config instance (loads from file on first access)"""
    global _config_instance
    if _config_instance is None:
        _config_instance = AccessibilityCheckConfig.load()
    return _config_instance


def save_config() -> None:
    """Save the current config to file"""
    global _config_instance
    if _config_instance is not None:
        _config_instance.save()


def reset_config() -> AccessibilityCheckConfig:
    """Reset config to defaults and save"""
    global _config_instance
    _config_instance = AccessibilityCheckConfig()
    _config_instance.save()
    return _config_instance


__all__ = [
    'AccessibilityCheckConfig',
    'CONTRAST_THRESHOLDS',
    'get_config',
    'save_config',
    'reset_config',
    'get_config_path',
]
