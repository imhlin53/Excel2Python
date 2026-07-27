"""
============================================================
File:
    validation_engine.py

Project:
    VOBES Migration Tool

Authors:
    Lin, Hua
    M365 Copilot

AI Contributor Information:
    Product:
        Microsoft 365 Copilot

    Model Family:
        GPT-5 Chat

    Contribution Areas:
        - Validation Architecture
        - VBA Migration Analysis
        - Python Development
        - Software Design
        - Documentation
        - Code Review

Purpose:
    Central validation framework for
    VOBES workbook data.

Description:
    Provides validation services used
    to verify the completeness and
    correctness of pin definitions.

    This module is intended to become
    the primary replacement for the
    original VBA validation logic.

Primary Class:
    ValidationEngine

Original VBA Replacements:
    CheckMe()

Current Responsibilities:

    - Required field checks
    - Function validation
    - Direction validation
    - Voltage validation
    - Utilization validation

Planned Responsibilities:

    Phase 2
        - CheckMe() migration

    Phase 3
        - CurrentChk() integration

    Future
        - Signal consistency checks
        - Pin type validation
        - Current limit validation
        - Workbook integrity checks

Architecture Role:

        GUI
          ↓
    ValidationEngine
          ↓
      Error List

Used By:
    vobes_gui.py
    vobes_pin_manager.py
    vobes_service.py

Dependencies:
    None

Status:
    Active Development

Creation Date:
    2026-07-21

Last Updated:
    2026-07-21

Maintainers:
    Lin, Hua
    M365 Copilot

Validation Philosophy:

    Validation logic should be
    centralized in this module.

    GUI code should only display
    validation results.

    Business rules should remain
    inside ValidationEngine.

Current Rule Set:

    Rule 1
        Function must exist.

    Rule 2
        Direction must exist.

    Rule 3
        Voltage must exist.

    Rule 4
        Utilization must exist.

Future Rule Set:

    Rule 5
        Function must exist in
        CL_InfoDB.

    Rule 6
        Pin type must be valid.

    Rule 7
        Direction must be
        compatible with function.

    Rule 8
        Voltage type must be
        compatible with function.

    Rule 9
        Description must be valid.

Change Log:

2026-07-21
    Lin, Hua / M365 Copilot

    Initial implementation

        - Created ValidationEngine
        - Added required field checks

    Future Work

        - Migrate CheckMe()
        - Add CurrentChk() support
        - Add CL_InfoDB validation
        - Add rule-based framework

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
        #
        # Voltage
        #
        # expected = str(
        #     cl_info.get(
        #         "voltage_type",
        #         ""
        #     )
        # ).strip()
        # actual = str(
        #     pin.get(
        #         "voltage",
        #         ""
        #     )
        # ).strip()
        # if (
        #     expected
        #     and
        #     actual
        #     and
        #     expected != actual
        # ):
        #     errors.append(
        #         f"Voltage mismatch "
        #         f"(Expected={expected}, "
        #         f"Actual={actual})"
        #     )
        #
        # Utilization
        #
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
            "i5"
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
        return errors
    
    
        