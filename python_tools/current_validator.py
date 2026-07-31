"""
============================================================
File:
    current_validator.py

Project:
    VOBES Migration Tool

Purpose:
    Validation logic for current-related and pinchar rules.

Description:
    Provides validation helpers for numeric current values and
    for pin characterisation rules used by the migration tool.

Last Updated:
    2026-07-31

============================================================
"""
from pinchar_rules import PINCHAR_RULES
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

    @staticmethod
    def validate_pinchar(pin):

        errors = []

        pin_type = str(
            pin.get(
                "pin_type",
                ""
            )
        )

        rules = PINCHAR_RULES.get(
            pin_type,
            {}
        )

        #
        # required fields
        #

        for field in rules.get(
            "required_fields",
            []
        ):

            value = pin.get(field)

            if value in (
                None,
                ""
            ):

                errors.append(
                    f"{field.upper()} required"
                )

        return errors
    
    @staticmethod
    def validate_formula_rules(pin):

        errors = []

        characteristic = str(
            pin.get(
                "characteristic",
                ""
            )
        )

        rules = PINCHAR_RULES.get(
            characteristic,
            {}
        )

        formulas = rules.get(
            "formula_rules",
            {}
        )

        for target, rule in formulas.items():
            
            source = rule["source"]

            multiplier = rule["multiplier"]

            try:            
                source_value = float(pin.get(source))

                target_value = float(pin.get(target))
                
            except (
                TypeError,
                ValueError
            ):
                continue
            
            expected = (
                source_value
                * multiplier
            )
            
            tolerance = 0.01

            if abs(expected - target_value) > tolerance:

                errors.append(
                    f"{target.upper()} "
                    f"expected {expected:.3f} "
                    f"actual {target_value:.3f}"
                )

        return errors
