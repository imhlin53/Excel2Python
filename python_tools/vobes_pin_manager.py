"""
File: vobes_pin_manager.py

Author:
    Lin, Hua
    M365 Copilot

Date:
    2026-07-16

Purpose:
    Pin management layer.

Replaces core VBA functionality:

    Generate_PinSlot_Rows()
    DelPins()
    AddPins2Con()
    ExchangePins()
    EndCheck_Pin()
"""

from vobes_service import VobesService


class VobesPinManager:

    def __init__(self, service):

        self.service = service

    # ===================================
    # Read Pins
    # ===================================

    def get_all_pins(self):

        return self.service.get_all_pins()

    def get_pin(self,
                slot_name,
                pin_number):

        return self.service.get_pin(
            slot_name,
            pin_number
        )

    # ===================================
    # Validation
    # ===================================

    def validate_pin(self, pin):
        errors = []
        if not pin.get("slot"):
            errors.append(
                "Missing slot"
            )
        if not pin.get("pin"):
            errors.append(
                "Missing pin"
            )
        if not pin.get("description"):
            errors.append(
                "Missing description"
            )
        if not pin.get("pin_type"):
            errors.append(
                "Missing pin type"
            )
        if not pin.get("voltage"):
            errors.append(
                "Missing voltage"
            )
        if not pin.get("direction"):
            errors.append(
                "Missing direction"
            )
        if not pin.get("utilization"):
            errors.append(
                "Missing utilization"
            )

        #
        # New checks
        #

        desc = str(
            pin.get(
                "description",
                ""
            )
        )
        if "#" not in desc:
            errors.append(
                "Description missing '#'"
            )
        return errors

    def validate_all_pins(self):

        results = []

        for pin in self.get_all_pins():

            results.append({

                "row":
                    pin["row"],

                "errors":
                    self.validate_pin(pin)
            })

        return results

    # ===================================
    # CheckMe replacement
    # ===================================

    def check_me(self):

        problems = []

        for item in self.validate_all_pins():

            if item["errors"]:

                problems.append(item)

        return problems

    # ===================================
    # Exchange Pins
    # ===================================

    def exchange_pin_descriptions(
            self,
            slot_a,
            pin_a,
            slot_b,
            pin_b):

        first = self.get_pin(
            slot_a,
            pin_a
        )

        second = self.get_pin(
            slot_b,
            pin_b
        )

        if first is None:

            raise ValueError(
                "First pin not found"
            )

        if second is None:

            raise ValueError(
                "Second pin not found"
            )

        tmp = first["description"]

        self.service.update_pin_description(
            first["row"],
            second["description"]
        )

        self.service.update_pin_description(
            second["row"],
            tmp
        )

    # ===================================
    # Pin Search
    # ===================================

    def find_pins_by_description(
            self,
            text):

        text = str(text).lower()

        matches = []

        for pin in self.get_all_pins():

            desc = str(
                pin["description"]
            ).lower()

            if text in desc:

                matches.append(pin)

        return matches

    # ===================================
    # Summary
    # ===================================

    def print_pin_summary(self):

        print()

        print("=" * 60)
        print("PIN SUMMARY")
        print("=" * 60)

        rows = self.get_all_pins()

        print(
            "Total Pins:",
            len(rows)
        )

        for p in rows:

            print(
                f"{p['slot']}{p['pin']}  "
                f"{p['description']}"
            )