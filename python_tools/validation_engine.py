"""
============================================================
File:
    validation_engine.py

Project:
    VOBES Migration Tool

Purpose:
    Validation framework for workbook pin data.

Description:
    Provides reusable validation helpers for required fields
    and business-rule checks used by the migration workflow.

Last Updated:
    2026-07-31

============================================================
"""
from current_validator import CurrentValidator

class ValidationEngine:

    @staticmethod
    def validate_pin(pin):

        errors = []

        if not pin.get("function"):
            errors.append(
                "Function missing"
            )

        if not pin.get("direction"):
            errors.append(
                "Direction missing"
            )

        if not pin.get("voltage"):
            errors.append(
                "Voltage missing"
            )

        if not pin.get("utilization"):
            errors.append(
                "Utilization missing"
            )

        if not pin.get("description"):
            errors.append(
                "Description missing"
            )
        
        return errors
    
    @staticmethod
    def validate_cl_consistency(
            pin,
            cl_info):

        errors = []
        if not cl_info:
            errors.append(
                "Unknown Function"
            )
            return errors
        #
        # Direction
        #
        expected = str(
            cl_info.get(
                "direction",
                ""
            )
        ).strip()
        actual = str(
            pin.get(
                "direction",
                ""
            )
        ).strip()
        if (
            expected
            and
            actual
            and
            expected != actual
        ):
            errors.append(
                f"Direction mismatch "
                f"(Expected={expected}, "
                f"Actual={actual})"
            )

        expected = str(
            cl_info.get(
                "utilization",
                ""
            )
        ).strip()
        actual = str(
            pin.get(
                "utilization",
                ""
            )
        ).strip()
        if (
            expected
            and
            actual
            and
            expected != actual
        ):
            errors.append(
                f"Utilization mismatch "
                f"(Expected={expected}, "
                f"Actual={actual})"
            )

        return errors
    
    @staticmethod
    def validate_database_values(
            pin,
            svc):
        errors = []
        #
        # Direction
        #
        direction = str(
            pin.get(
                "direction",
                ""
            )
        ).strip()
        if (
            direction
            and
            not svc.is_valid_direction(
                direction
            )
        ):
            errors.append(
                f"Invalid Direction: {direction}"
            )
        #
        # Voltage
        #
        voltage = str(
            pin.get(
                "voltage",
                ""
            )
        ).strip()
        if (
            voltage
            and
            not svc.is_valid_voltage(
                voltage
            )
        ):
            errors.append(
                f"Invalid Voltage: {voltage}"
            )
        
        return errors
    
    @staticmethod
    def validate_current_fields(
        pin):
        errors = []
        
        for field in [
            "i1",
            "i2",
            "i3",
            "i4",
            "i5",
            "i6",
            "t1",
            "t2"
        ]:
            value = pin.get(
                field
            )
            curr_errors = \
                CurrentValidator.validate_current(
                    value
                )
            for err in curr_errors:
                errors.append(
                    f"{field.upper()}: {err}"
                )
        errors.extend(
            CurrentValidator.validate_formula_rules(
                pin
            )
        )
        return errors
    
    
        