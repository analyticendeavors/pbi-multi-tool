"""
Registration Manager - Handles user registration and license verification
Built by Reid Havens of Analytic Endeavors
"""

import json
import hashlib
import requests
from pathlib import Path
from datetime import datetime
from typing import Optional, Dict, Any


class RegistrationManager:
    """
    Manages user registration with Google Apps Script backend.
    Uses singleton pattern matching existing architecture (ThemeManager, RecentFilesManager).
    """

    _instance: Optional['RegistrationManager'] = None

    # Google Apps Script endpoint - UPDATE THIS after deploying your script
    REGISTRATION_API_URL = "https://script.google.com/macros/s/AKfycbzWUCPF19WC-39HAzESVy2sFHoA9kodwlue9zMZZBzkbkX_mg9EeNMkwHU8KOYbDSSonQ/exec"

    # Request timeout in seconds
    REQUEST_TIMEOUT = 15

    def __init__(self):
        self._registration_path = (
            Path.home() / "AppData" / "Local" / "AnalyticEndeavors" / "registration.json"
        )
        self._registration_data: Optional[Dict[str, Any]] = None
        self._load_registration()

    @classmethod
    def get_instance(cls) -> 'RegistrationManager':
        """Get singleton instance"""
        if cls._instance is None:
            cls._instance = RegistrationManager()
        return cls._instance

    @classmethod
    def reset_instance(cls):
        """Reset singleton instance (for testing)"""
        cls._instance = None

    def _load_registration(self):
        """Load registration data from disk"""
        try:
            if self._registration_path.exists():
                with open(self._registration_path, 'r', encoding='utf-8') as f:
                    self._registration_data = json.load(f)
        except (json.JSONDecodeError, IOError):
            self._registration_data = None

    def _save_registration(self, name: str, email: str, app_version: str = ""):
        """Save registration marker to disk"""
        self._registration_path.parent.mkdir(parents=True, exist_ok=True)

        self._registration_data = {
            "registered": True,
            "email_hash": hashlib.sha256(email.lower().encode()).hexdigest(),
            "name": name,
            "registration_date": datetime.now().isoformat(),
            "app_version": app_version
        }

        # Atomic write pattern (temp file + rename)
        temp_path = self._registration_path.with_suffix('.tmp')
        try:
            with open(temp_path, 'w', encoding='utf-8') as f:
                json.dump(self._registration_data, f, indent=2)
            temp_path.replace(self._registration_path)
        except IOError:
            if temp_path.exists():
                temp_path.unlink()
            raise

    def is_registered(self) -> bool:
        """Check if user is registered (local marker exists and is valid)"""
        if self._registration_data is None:
            return False
        return self._registration_data.get("registered", False) is True

    def get_registration_info(self) -> Optional[Dict[str, Any]]:
        """Return stored registration data"""
        return self._registration_data

    def register_user(
        self,
        name: str,
        email: str,
        company: str,
        job_title: str,
        app_version: str = ""
    ) -> tuple[bool, str]:
        """
        Register a new user via Google Apps Script.

        Returns:
            tuple: (success: bool, message: str)
        """
        # Normalize email: lowercase and strip whitespace for consistent storage/lookup
        normalized_email = email.strip().lower()
        name = name.strip()
        company = company.strip()
        job_title = job_title.strip()

        try:
            response = requests.post(
                self.REGISTRATION_API_URL,
                json={
                    "name": name,
                    "email": normalized_email,
                    "company": company,
                    "job_title": job_title,
                    "version": app_version
                },
                timeout=self.REQUEST_TIMEOUT,
                headers={"Content-Type": "application/json"}
            )

            if response.status_code == 200:
                result = response.json()
                if result.get("success"):
                    self._save_registration(name, normalized_email, app_version)
                    return True, "Registration successful"
                else:
                    return False, result.get("error", "Registration failed")
            else:
                return False, f"Server error: {response.status_code}"

        except requests.exceptions.Timeout:
            return False, "Connection timed out. Please check your internet connection."
        except requests.exceptions.ConnectionError:
            return False, "Unable to connect. Please check your internet connection."
        except requests.exceptions.RequestException as e:
            return False, f"Network error: {str(e)}"
        except json.JSONDecodeError:
            # Apps Script returned success but not valid JSON - treat as success
            self._save_registration(name, normalized_email, app_version)
            return True, "Registration successful"
        except IOError as e:
            return False, f"Failed to save registration: {str(e)}"

    def verify_existing_email(self, email: str) -> tuple[bool, str]:
        """
        Check if an email is already registered (for users on new machines).

        Returns:
            tuple: (found: bool, message: str)
        """
        # Normalize email: lowercase and strip whitespace for consistent matching
        normalized_email = email.strip().lower()

        try:
            response = requests.get(
                self.REGISTRATION_API_URL,
                params={"email": normalized_email},
                timeout=self.REQUEST_TIMEOUT
            )

            if response.status_code == 200:
                result = response.json()
                if result.get("registered"):
                    # Email found - save local marker with normalized email
                    self._save_registration("", normalized_email, "")
                    return True, "Email verified successfully"
                else:
                    return False, "Email not found. Please register with the form."
            else:
                return False, f"Server error: {response.status_code}"

        except requests.exceptions.Timeout:
            return False, "Connection timed out. Please check your internet connection."
        except requests.exceptions.ConnectionError:
            return False, "Unable to connect. Please check your internet connection."
        except requests.exceptions.RequestException as e:
            return False, f"Network error: {str(e)}"
        except json.JSONDecodeError:
            return False, "Invalid server response"

    def clear_registration(self):
        """Clear registration (for testing/support purposes)"""
        if self._registration_path.exists():
            self._registration_path.unlink()
        self._registration_data = None
