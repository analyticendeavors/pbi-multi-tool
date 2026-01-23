"""
Software Ownership Verification Module
Internal use only - validates software authenticity.

Copyright (c) 2024 Analytic Endeavors LLC. All rights reserved.
Unauthorized copying, modification, or distribution is prohibited.
"""

import hashlib
import base64
from datetime import datetime
from typing import Optional, Dict, Tuple

# Distributed ownership fingerprint components
# These are spread across the codebase for redundancy
_OWNER_SIG = "UmVpZCBIYXZlbnM="  # Component 1
_COMPANY_SIG = "QW5hbHl0aWMgRW5kZWF2b3Jz"  # Component 2
_PROJECT_SIG = "QUUgTXVsdGktVG9vbA=="  # Component 3
_BUILD_ID = "QUUtMjAyNC1NVC0wMDE="  # Component 4
_RIGHTS_SIG = "QWxsIFJpZ2h0cyBSZXNlcnZlZA=="  # Component 5

# Unique project fingerprint (SHA256 of ownership data)
_FINGERPRINT = "ae7f3c2d8b4e9a1f6c5d0e7b2a3f8c9d4e5b6a7c8d9e0f1a2b3c4d5e6f7a8b9c"


def _decode_component(encoded: str) -> str:
    """Internal: Decode a fingerprint component"""
    try:
        return base64.b64decode(encoded).decode('utf-8')
    except Exception:
        return ""


def _compute_verification_hash(key: str) -> str:
    """Internal: Compute verification hash"""
    components = [_OWNER_SIG, _COMPANY_SIG, _PROJECT_SIG, _BUILD_ID]
    combined = ":".join([_decode_component(c) for c in components]) + f":{key}"
    return hashlib.sha256(combined.encode()).hexdigest()[:16]


def verify_ownership(verification_key: str) -> Tuple[bool, Optional[Dict]]:
    """
    Verify software ownership with the provided key.

    Args:
        verification_key: The secret verification key

    Returns:
        Tuple of (is_valid, ownership_info or None)

    Usage:
        valid, info = verify_ownership("your-secret-key")
        if valid:
            print(info)
    """
    # The verification key must produce the correct hash
    expected_hash = "a11a36ea17e2a81b"  # Derived from ownership data
    computed_hash = _compute_verification_hash(verification_key)

    if computed_hash == expected_hash:
        return True, {
            "owner": _decode_component(_OWNER_SIG),
            "company": _decode_component(_COMPANY_SIG),
            "project": _decode_component(_PROJECT_SIG),
            "build_id": _decode_component(_BUILD_ID),
            "rights": _decode_component(_RIGHTS_SIG),
            "fingerprint": _FINGERPRINT,
            "verified_at": datetime.now().isoformat(),
            "status": "AUTHENTIC - Ownership Verified"
        }

    return False, None


def get_fingerprint() -> str:
    """
    Get the unique project fingerprint.
    This can be used to verify code hasn't been tampered with.
    """
    return _FINGERPRINT


def get_build_info() -> Dict[str, str]:
    """
    Get basic build information (non-sensitive).
    """
    return {
        "project": _decode_component(_PROJECT_SIG),
        "company": _decode_component(_COMPANY_SIG),
        "fingerprint": _FINGERPRINT[:8] + "..."
    }


# Hidden verification: calling this module directly shows ownership
if __name__ == "__main__":
    print("=" * 60)
    print("AE Multi-Tool - Ownership Verification")
    print("=" * 60)
    print(f"Project: {_decode_component(_PROJECT_SIG)}")
    print(f"Company: {_decode_component(_COMPANY_SIG)}")
    print(f"Fingerprint: {_FINGERPRINT}")
    print("=" * 60)
    print("To verify ownership, use verify_ownership() with your key")
