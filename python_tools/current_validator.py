"""
============================================================
File:
    current_validator.py
Project:
    VOBES Migration Tool
Authors:
    Lin, Hua
    M365 Copilot
Purpose:
    Current field validation.
Original VBA Replacement:
    CurrentChk()

Description:
    Validates current values entered
    in I1-I5 fields.

Rules:
    Allowed:
        0-9
        .
        ,

    Not Allowed:
        Letters
        Symbols
        Multiple decimal separators

============================================================
"""

class CurrentValidator:
    @staticmethod
    def validate_current(value):
        if value is None:
            return []
        
        text = str(value).strip()
        
        if not text:
            return []
        
        decimal_count = 0
        
        for ch in text:
            if ch in ".,":
                decimal_count += 1
                if decimal_count > 1:
                    return [
                        "Multiple decimal separators"
                    ]
            elif not ch.isdigit():
                return [
                    f"Invalid character '{ch}'"
                ]
        return []