"""
============================================================
File:
    vobes_db.py

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

    Role:
        VBA Migration Support
        Python Development
        Software Architecture
        Documentation
        Code Review

Purpose:
    CL_InfoDB database access layer.

Description:
    Loads and manages the CL_InfoDB.csv
    database used by the VOBES Migration Tool.

    Provides:
        - Section loading
        - Database indexing
        - Generic searches
        - Function searches
        - Clamp searches
        - Utilization searches
        - Voltage searches
        - Signal searches

Primary Class:
    VobesDB

Original VBA Replacements:
    DB_Import.bas
    StarteImport_CSVdatei()
    untersuch()

Responsibilities:
    - Import CL_InfoDB.csv
    - Maintain section dictionary
    - Provide search services
    - Support CL information lookup
    - Support future CheckMe validation

Dependencies:
    pathlib

Used By:
    vobes_service.py
    validation_engine.py
    pin_description_generator.py

Status:
    Active Development

Creation Date:
    2026-07-16

Last Updated:
    2026-07-21

Change Log:

2026-07-21
    Lin, Hua / M365 Copilot

    Phase 1-A
        - Added support for
          find_cl_info()

        - Verified CL_Functions
          database searches

        - Integrated with
          VobesService

2026-07-16
    Lin, Hua / M365 Copilot

    Initial implementation

        - CSV loader
        - Section indexing
        - Generic search
        - Function search
        - Voltage search
        - Clamp search

============================================================
"""
from pathlib import Path

class VobesDB:
    """
    CL_InfoDB database manager.

    Responsibilities:
        - Load CL_InfoDB.csv
        - Organize data by section
        - Provide search APIs
        - Support CL information lookup

    Original VBA Replacements:
        DB_Import.bas
        StarteImport_CSVdatei()
        untersuch()
    """

    def __init__(self):

        self.sections = {}

    # ==================================================
    # Load CL_InfoDB.csv
    # ==================================================

    def load_csv(self, filename):
        """
        Load CL_InfoDB.csv.
        Args:
            filename (str): Path to CL_InfoDB.csv

        Creates:
            self.sections

        Example:
            #CL_Functions###
            #CL_Clamp###
            #CL_Utilization###

        Original VBA:
            StarteImport_CSVdatei()
        """
        current_section = None

        filename = Path(filename)

        with open(
            filename,
            "r",
            encoding="utf-8",
            errors="ignore"
        ) as f:

            for line in f:

                line = line.strip()

                if not line:
                    continue

                # Section marker

                if (
                    line.startswith("#")
                    and
                    line.endswith("###")
                ):

                    current_section = line

                    if current_section not in self.sections:

                        self.sections[current_section] = []

                    continue

                # Store row

                if current_section:

                    values = line.split(";")

                    self.sections[
                        current_section
                    ].append(values)

    # ==================================================
    # Section Access
    # ==================================================

    def get_section(self, section_name):

        return self.sections.get(
            section_name,
            []
        )

    def get_section_names(self):

        return sorted(
            self.sections.keys()
        )

    # ==================================================
    # Generic Search
    # ==================================================

    def search(
        self,
        section_name,
        text
    ):

        text = str(text).lower()

        results = []

        rows = self.get_section(
            section_name
        )

        for row in rows:

            joined = ";".join(row)

            if text in joined.lower():

                results.append(row)

        return results

    # ==================================================
    # CL Functions
    # ==================================================

    def search_function(self, text):

        return self.search(
            "#CL_Functions###",
            text
        )

    # ==================================================
    # CL Utilization
    # ==================================================

    def search_utilization(self, text):

        return self.search(
            "#CL_Utilization###",
            text
        )

    # ==================================================
    # CL Clamp
    # ==================================================

    def search_clamp(self, text):

        return self.search(
            "#CL_Clamp###",
            text
        )

    # ==================================================
    # Voltage
    # ==================================================

    def search_voltage(self, text):

        return self.search(
            "#CL_Voltage###",
            text
        )

    # ==================================================
    # Signal Abbreviation
    # ==================================================

    def search_signal_abbrev(self, text):

        return self.search(
            "#CL_SignalAbbreviation###",
            text
        )

    # ==================================================
    # Signal Characteristic
    # ==================================================

    def search_signal_char(self, text):

        return self.search(
            "#CL_SignalCharacteristic###",
            text
        )

    # ==================================================
    # Direction
    # ==================================================

    def search_direction(self, text):

        return self.search(
            "#CL_SignalDirection###",
            text
        )

    # ==================================================
    # BGER
    # ==================================================

    def search_bger(self, text):

        return self.search(
            "#CL_BGER_Pins###",
            text
        )

    # ==================================================
    # Summary
    # ==================================================

    def print_summary(self):

        print()

        print("=" * 60)
        print("CL_InfoDB Summary")
        print("=" * 60)

        for section in self.get_section_names():

            count = len(
                self.sections[section]
            )

            print(
                f"{section:<35} {count:>8}"
            )

        print()