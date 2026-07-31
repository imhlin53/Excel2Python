"""
============================================================
File:
    pinchar_engine.py

Project:
    VOBES Migration Tool

Purpose:
    Pin characteristic logic for the migration workflow.

Description:
    Loads pin-characteristic mappings from the workbook and
    provides helpers for resolving sheet definitions, default
    values, and formula-based pinchar rules.

Last Updated:
    2026-07-31

============================================================
"""
from pinchar_rules import PINCHAR_RULES
class PincharEngine:

    def __init__(self, workbook):

        self.workbook = workbook

        self.pin_type_map = self._load_mapping()

    # -------------------------------------------------
    # Daten Sheet
    # -------------------------------------------------
    def _load_mapping(self):

        ws = self.workbook.wb["Daten"]

        mapping = {}

        row = 2

        while True:

            pin_type = ws.cell(row=row, column=1).value
            pinchar = ws.cell(row=row, column=2).value

            if pin_type is None and pinchar is None:
                break

            mapping[str(pin_type).strip()] = pinchar

            row += 1

        return mapping

    # -------------------------------------------------
    # Pin Type -> Pinchar Sheet
    # -------------------------------------------------
    def get_pinchar_sheet(self, pin_type):

        if not pin_type:
            return "Pinchar0"

        return self.pin_type_map.get(
            str(pin_type).strip(),
            "Pinchar0"
        )

    # -------------------------------------------------
    # Read Pinchar Definition
    # -------------------------------------------------
    def get_definition(self, pin_type):

        sheet_name = self.get_pinchar_sheet(pin_type)

        ws = self.workbook.wb[sheet_name]

        return {

            "sheet": sheet_name,

            "I1": ws["B2"].value,
            "I2": ws["C2"].value,
            "I3": ws["D2"].value,
            "I4": ws["E2"].value,
            "I5": ws["F2"].value,
            "I6": ws["G2"].value,

            "T1": ws["H2"].value,
            "T2": ws["I2"].value
        }

    # -------------------------------------------------
    # Component Characteristics
    # -------------------------------------------------
    def get_characteristics(self, pin_type):

        sheet_name = self.get_pinchar_sheet(
            pin_type
        )

        ws = self.workbook.wb[sheet_name]

        chars = []

        row = 5

        while True:

            value = ws.cell(
                row=row,
                column=1
            ).value

            if value is None:
                break

            chars.append(str(value))

            row += 1

        return chars

    # -------------------------------------------------
    # Defaults
    # -------------------------------------------------
    def apply_defaults(
            self,
            pin_type,
            characteristic,
            i1=None,
            i2=None):

        result = {}

        if pin_type == "Nicht verbunden":
            result["i1"] = 0

        #
        # Pinchar3
        #
        if pin_type == "Senke (Komponente)":

            if (
                characteristic == "Motor"
                and i2
            ):
                result["i3"] = float(i2) * 4
                result["t1"] = 15

            elif (
                characteristic == "Kapazität"
                and i2
            ):
                result["i3"] = float(i2) * 5
                result["t1"] = 2

            elif (
                characteristic == "Glühlampe"
                and i2
            ):
                result["i3"] = float(i2) * 3.5
                result["t1"] = 0.1

            elif (
                characteristic == "Spule"
                and i1
            ):
                result["i2"] = float(i1) * 4
                result["t1"] = 0.3

        return result
    
    def calculate_defaults(
            self,
            characteristic,
            values):

        result = {}

        rules = PINCHAR_RULES.get(
            characteristic,
            {}
        )

        #
        # Formula Rules
        #
        formulas = rules.get(
            "formula_rules",
            {}
        )

        for target, formula in formulas.items():
            source = formula["source"]
            multiplier = formula["multiplier"]
            source_value = values.get(source)

            if source_value in (None, ""):
                continue

            result[target] = (
                float(source_value)
                *
                multiplier
            )

        #
        # Fixed Defaults
        #
        defaults = rules.get(
            "default_values",
            {}
        )

        result.update(defaults)

        return result
    
    
    