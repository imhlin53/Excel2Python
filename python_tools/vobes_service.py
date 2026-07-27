"""
============================================================
File:
    vobes_service.py

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
        - VBA Migration Analysis
        - Python Development
        - Software Architecture
        - CL_InfoDB Integration
        - Validation Framework
        - Documentation
        - Code Review

Purpose:
    Service layer between GUI,
    workbook access, and CL_InfoDB.

Description:
    Provides the central business logic
    layer for the VOBES Migration Tool.

    Coordinates:
        GUI
        ↓
    VobesService
        ↓
    Workbook
        ↓
    CL_InfoDB

Responsibilities:

    Workbook Operations
        - Read pin data
        - Read slot data
        - Update workbook fields
        - Save workbook changes
    CL_InfoDB Operations
        - Function search
        - Clamp search
        - Voltage search
        - Utilization search
    Lookup Services
        - find_cl_info()
        - get_direction()
        - get_voltage_type()
    Validation Services
        - validate_pin()
        - validate_all_pins()
Primary Class:
    VobesService
Original VBA Replacements:
    Recherche()
    finde_CL_info()
    genPinDaten_actCell()
    Partial CheckMe()
Current Migration Status:
    Phase 1-A
        ✅ find_cl_info()
    Phase 1-B
        ✅ Search Integration
        ✅ Auto Fill
    Phase 1-C
        ✅ Description Generation
    Phase 2-A
        ✅ Basic CheckMe
    Phase 2-B
        ✅ CL_InfoDB Consistency
    Phase 2-C
        ✅ Database Validation
    Phase 2-D
        ✅ Check All Pins

    Phase 3
        ⏳ CurrentChk()

Dependencies:
    vobes_workbook.py
    vobes_db.py
    validation_engine.py

Used By:
    vobes_gui.py
    vobes_pin_manager.py

Status:
    Active Development

Creation Date:
    2026-07-16

Last Updated:
    2026-07-23

Maintainers:
    Lin, Hua
    M365 Copilot

Change Log:
    2026-07-23

        Lin, Hua / M365 Copilot

        Phase 2-D
            - Added validate_all_pins()
            - Added Validation Tab support
            - Added validation reporting

    2026-07-22
        Lin, Hua / M365 Copilot

        Phase 2-B
            - Added CL consistency validation
            - Added exact-match
                CL function lookup

    2026-07-21
        Lin, Hua / M365 Copilot

        Phase 1-A
            - Added find_cl_info()
        Phase 1-B
            - Added auto-fill:
                Voltage
                Direction
                Utilization
        Phase 1-C
            - Added description support

    2026-07-16
        Lin, Hua / M365 Copilot

        Initial implementation
============================================================
"""
from pathlib import Path

from vobes_workbook import VobesWorkbook
from vobes_db import VobesDB
from validation_engine import ValidationEngine

class VobesService:

    def __init__(
            self,
            workbook_file,
            csv_file):

        self.workbook = VobesWorkbook(
            workbook_file
        )

        self.db = VobesDB()

        self.db.load_csv(
            csv_file
        )

    # ==========================================
    # Workbook Summary
    # ==========================================

    def print_summary(self):

        print()

        print("=" * 60)
        print("VOBES SUMMARY")
        print("=" * 60)

        self.workbook.print_summary()

        print(
            "Database Sections:",
            len(
                self.db.get_section_names()
            )
        )

    # ==========================================
    # Search Functions
    # ==========================================

    def search_function(
            self,
            text):

        return self.db.search_function(
            text
        )

    def search_clamp(
            self,
            text):

        return self.db.search_clamp(
            text
        )

    def search_utilization(
            self,
            text):

        return self.db.search_utilization(
            text
        )

    def search_voltage(
            self,
            text):

        return self.db.search_voltage(
            text
        )
    
    #
    # NEW METHOD
    #
    def find_cl_info(
            self,
            function_name):
        # section = self.db.get_section("#CL_Functions###")
        # print("find_cl_info() input =", repr(function_name))
        
        function_name = str(
            function_name
        ).strip()
        
        rows = self.db.get_section("#CL_Functions###")
        function_name = str(
            function_name
        ).strip().lower()
        for row in rows:
            if not row:
                continue
            candidate = str(
                row[0]
            ).strip().lower()
            #
            # Exact match only
            #
            if candidate == function_name:
                # print()
                # print("Function = ", function_name)
                # print("Row = ", row)
                return {
                    "function":      row[0],
                    "english":       row[1] if len(row) > 1 else "",
                    "german":        row[2] if len(row) > 2 else "",
                    "voltage_type":  row[3] if len(row) > 3 else "",
                    "utilization":   row[4] if len(row) > 4 else "",
                    "direction":     row[5] if len(row) > 5 else ""
                }
        return None
                
        print()
        print("Function = ", function_name)
        print("Row = ", row)
        return {
            "function":      row[0],
            "english":       row[1] if len(row) > 1 else "",
            "german":        row[2] if len(row) > 2 else "",
            "voltage_type":  row[3] if len(row) > 3 else "",
            "utilization":   row[4] if len(row) > 4 else "",
            "direction":     row[5] if len(row) > 5 else ""
        }
        return None
    

    # ==========================================
    # Pin Operations
    # ==========================================

    def get_all_pins(self):

        pins = self.workbook.get_pins()
        print(pins[0]["pinchar"])
        return pins

    def get_pin(
            self,
            slot_name,
            pin_number):

        return self.workbook.get_pin_by_name(
            slot_name,
            pin_number
        )
        
    def get_voltage_type(
            self,
            function_name):
        info = self.find_cl_info(
            function_name
        )
        if not info:
            return ""
        
        return info[
            "voltage_type"
        ]
        
    def get_direction(
            self,
            function_name):
        info = self.find_cl_info(
            function_name
        )
        if not info:
            return ""
        
        return info[
            "direction"
        ]

    def update_pin_description(
            self,
            row,
            description):

        self.workbook.set_pin_description(
            row,
            description
        )

    def update_pin_type(
            self,
            row,
            pin_type):

        self.workbook.set_pin_type(
            row,
            pin_type
        )
        
    def update_pin_clamp(
            self,
            row,
            value):
        self.workbook.set_pin_clamp(
            row,
            value
        )
        
    def update_pin_function(
            self,
            row,
            value):
        self.workbook.set_pin_function(
            row,
            value
        )
        
    def update_pin_direction(
            self,
            row,
            value):
        self.workbook.set_pin_direction(
            row,
            value
        )
        
    def update_pin_voltage(
            self,
            row,
            value):
        self.workbook.set_pin_voltage(
            row,
            value
        )
        
    def update_pin_utilization(
            self,
            row,
            value):
        self.workbook.set_pin_utilization(
            row,
            value
        )

    # ==========================================
    # CheckMe Replacement
    # ==========================================

    def validate_pin(self, pin):

        errors = []
        #
        # Basic Checks
        #
        errors.extend(
            ValidationEngine.validate_pin(
                pin
            )
        )
        #
        # CL Checks
        #
        function_name = pin.get(
            "function",
            ""
        )
        if function_name:
            info = self.find_cl_info(
                function_name
            )
            errors.extend(
                ValidationEngine.validate_cl_consistency(
                    pin,
                    info
                )
            )
            errors.extend(
                ValidationEngine.validate_database_values(
                    pin,
                    self
                )
            )
        
        errors.extend(
            ValidationEngine
            .validate_current_fields(
                pin
            )
        )
        
        return errors

    def validate_all_pins(self):
        results = []
        #
        # We need to get all pins first
        #
        pins = self.get_all_pins()
        for pin in pins:
            errors = self.validate_pin(pin)
            results.append({
                "pin": pin,
                "errors": errors
            })
        return results

    # ==========================================
    # Save
    # ==========================================

    def save(self, filename=None):

        self.workbook.save(
            filename
        )
        
    def is_valid_direction(
            self,
            value):

        value = str(
            value
        ).strip()

        rows = self.db.search_direction(
            value
        )

        for row in rows:

            for item in row:

                if (
                    str(item).strip()
                    ==
                    value
                ):
                    return True

        return False
    
    def is_valid_voltage(
            self,
            value):
        value = str(
            value
        ).strip()
        rows = self.db.search_voltage(value)
        for row in rows:
            for item in row:
                if (
                    str(item).strip()
                    ==
                    value
                ):
                    return True
        return False
    
    def is_valid_utilization(
            self,
            value):
        value = str(value).strip()
        # print()
        # print("UTIL CHECK =", value)
        rows = self.db.search_utilization(
            value
        )
        # print("ROWS =", rows)
        
        for row in rows:
            for item in row:
                if (
                    str(item).strip()
                    ==
                    value
                ):
                    return True
        return False
