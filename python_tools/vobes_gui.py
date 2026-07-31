"""
============================================================
File:
    vobes_gui.py

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
        - Software Architecture
        - Python Development
        - VBA Migration Analysis
        - GUI Enhancement
        - CL_InfoDB Integration
        - Documentation
        - Code Review

Purpose:
    Main graphical user interface for the
    VOBES Migration Tool.

Description:
    Provides a modern Python replacement
    for the original VBA UserForm-based
    workflow used in the VOBES workbook.

    The GUI allows users to:

        - View workbook information
        - Edit pins
        - Search CL_InfoDB
        - Generate descriptions
        - Perform validation
        - Save workbook changes

Primary Class:
    VobesGUI

Original VBA Replacements:
    UserForm1
    UserForm2
    Recherche()
    finde_CL_info()
    genPinDaten_actCell()

Current Migration Status:

    Phase 1-A:
        ✅ find_cl_info()

    Phase 1-B:
        ✅ Database Search
        ✅ Automatic CL lookup

    Phase 1-C:
        ✅ Description Generator
        ✅ Auto-fill:
              Voltage
              Direction
              Utilization

    Phase 2:
        ⏳ CheckMe() Validation

    Phase 3:
        ⏳ CurrentChk()

    Phase 4:
        ⏳ AddPins2Con()
        ⏳ DelPins()

Dependencies:
    tkinter
    vobes_service.py
    vobes_workbook.py
    vobes_pin_manager.py
    pin_description_generator.py
    validation_engine.py

Major Responsibilities:

    GUI Layer
        - Window creation
        - Widget creation
        - User interaction

    Pin Management
        - Pin selection
        - Pin editing
        - Pin saving

    CL Database Integration
        - Search operations
        - CL lookup
        - Auto-fill support

    Validation Support
        - Validation display
        - Error reporting

Status:
    Active Development

Creation Date:
    2026-07-16

Last Updated:
    2026-07-31

Maintainers:
    Lin, Hua
    M365 Copilot

Change Log:

2026-07-31
    Lin, Hua / M365 Copilot

    Updated module header documentation

2026-07-21
    Lin, Hua / M365 Copilot

    Phase 1-A

        - Implemented find_cl_info()

    Phase 1-B

        - Added automatic CL lookup
        - Added auto population of:

              Voltage
              Direction
              Utilization

    Phase 1-C

        - Added PinDescriptionGenerator
        - Added automatic description generation

2026-07-16
    Lin, Hua / M365 Copilot

        Initial migration framework

        - Notebook GUI
        - Pin tab
        - Search tab
        - Validation tab
        - Workbook integration

============================================================
"""
from tkinter import messagebox
from statistics import mode
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from pathlib import Path
from tkinter import filedialog
import csv
from tkinter import filedialog
from vobes_pin_manager import VobesPinManager
from vobes_service import VobesService
from pin_description_generator import PinDescriptionGenerator

class VobesGUI:
    def __init__(self, root):
        self.root = root
        self.root.title('VOBES Migration Tool V3')
        self.root.geometry('1400x800')

        source_dir = Path(__file__).resolve().parent.parent / 'source'
        xlsm = next(source_dir.glob('*.xlsm'))
        csv_file = source_dir / 'CL_InfoDB.csv'

        self.svc = VobesService(xlsm, csv_file)
        self.pm = VobesPinManager(self.svc)
        
        rows = self.svc.db.search_function(
            "ASV_1"
            )
        # print("Row Count = ", len(rows))
        # for row in rows:
        #     print(row)
        
        self.selected_pin = None
        self.pins = []

        self.create_widgets()
        self.load_pins()
        self.populate_combos()

    def create_widgets(self):
        # ==========================================
        # MENU
        # ==========================================
        menu_bar = tk.Menu(self.root)
        file_menu = tk.Menu(
            menu_bar,
            tearoff=0
        )
        file_menu.add_command(
            label="Save",
            command=self.save_pin
        )
        file_menu.add_command(
            label="Save As",
            command=self.save_as
        )
        file_menu.add_separator()
        file_menu.add_command(
            label="Exit",
            command=self.root.destroy
        )
        menu_bar.add_cascade(
            label="File",
            menu=file_menu
        )
        self.root.config(
            menu=menu_bar
        )
        # ==========================================
        # HEADER
        # ==========================================
        header = ttk.Frame(self.root)
        header.pack(
            fill="x",
            padx=10,
            pady=5
        )
        ttk.Label(
            header,
            text=f"Part Number: {self.svc.workbook.get_part_number()}"
        ).pack(
            side="left"
        )
        ttk.Label(
            header,
            text=f"Pins: {self.svc.workbook.get_pin_count()}"
        ).pack(
            side="left",
            padx=20
        )
        ttk.Label(
            header,
            text=f"Slots: {self.svc.workbook.get_slot_count()}"
        ).pack(
            side="left",
            padx=20
        )
        try:
            ttk.Label(
                header,
                text=f"Status: {self.svc.workbook.get_data_checked()}"
            ).pack(
                side="left",
                padx=20
            )
        except Exception:
            pass
        # ==========================================
        # NOTEBOOK
        # ==========================================
        self.notebook = ttk.Notebook(
            self.root
        )
        self.notebook.pack(
            fill="both",
            expand=True
        )
        self.create_pin_tab()
        self.create_search_tab()
        self.create_validation_tab()
        self.create_status_bar()    

    def create_pin_tab(self):
        self.tab_pins = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_pins, text='Pins')

        left = ttk.Frame(self.tab_pins)
        left.pack(side='left', fill='y', padx=5, pady=5)
        right = ttk.Frame(self.tab_pins)
        right.pack(side='left', fill='both', expand=True, padx=5, pady=5)

        self.pin_list = tk.Listbox(left, width=50)
        self.pin_list.pack(fill='y', expand=True)
        self.pin_list.bind('<<ListboxSelect>>', self.on_pin_selected)

        self.lbl_pin_info = ttk.Label(right, text='No pin selected')
        self.lbl_pin_info.pack(anchor='w')
        
        self.lbl_pin_stats = ttk.Label(right, text="")
        self.lbl_pin_stats.pack(anchor='w')
        
        # ---------------------------------------
        # Pinchar / Current Characteristics
        # ---------------------------------------
        self.current_frame = ttk.LabelFrame(
            right,
            text="Current Characteristics"
        )
        self.current_frame.pack(
            fill="x",
            pady=5
        )
        
        self.current_vars = {}
        self.current_labels = {}
        print("Creating Pinchar Labels")
        
        for field in [
            "I1",
            "I2",
            "I3",
            "I4",
            "I5",
            "I6",
            "T1",
            "T2"
        ]:
            row_frame = ttk.Frame(
                self.current_frame
            )
            row_frame.pack(
                fill="x"
            )
            lbl = ttk.Label(
                row_frame,
                width=35,
                text=f"{field}:"
            )
            lbl.pack(
                side="left"
            )
            var = tk.StringVar()
            ent = ttk.Entry(
                row_frame,
                textvariable=var,
                width=15
            )
            ent.pack(
                side="left",
                padx=5
            )
            self.current_labels[field] = lbl
            self.current_vars[field] = var

        self.desc_var = tk.StringVar()
        ttk.Label(right, text='Description').pack(anchor='w')
        ttk.Entry(right, textvariable=self.desc_var).pack(fill='x')

        self.type_var = tk.StringVar()
        ttk.Label(right, text='Pin Type').pack(anchor='w')
        ttk.Entry(right, textvariable=self.type_var).pack(fill='x')
        
        # ----------------------------------------
        # Clamp
        # ----------------------------------------

        ttk.Label(
            right,
            text="Clamp"
        ).pack(anchor="w")

        self.clamp_var = tk.StringVar()

        self.clamp_combo = ttk.Combobox(
            right,
            textvariable=self.clamp_var
        )

        self.clamp_combo.pack(fill="x")

        # ----------------------------------------
        # Function
        # ----------------------------------------

        ttk.Label(
            right,
            text="Function"
        ).pack(anchor="w")

        self.function_var = tk.StringVar()

        self.function_combo = ttk.Combobox(
            right,
            textvariable=self.function_var
        )

        self.function_combo.pack(fill="x")
        
        # ----------------------------------------
        # Characteristic
        # ----------------------------------------
        ttk.Label(
            right,
            text="Characteristic"
        ).pack(
            anchor="w"
        )
        
        self.characteristic_var = tk.StringVar()
        
        self.characteristic_combo = ttk.Combobox(
            right,
            textvariable=self.characteristic_var
        )
        
        self.characteristic_combo["values"] = (
            "Motor",
            "Kapazität",
            "Glühlampe",
            "Spule",
            "Ohmsch"
        )
        
        self.characteristic_combo.pack(
            fill="x"
        )

        # ----------------------------------------
        # Utilization
        # ----------------------------------------

        ttk.Label(
            right,
            text="Utilization"
        ).pack(anchor="w")

        self.util_var = tk.StringVar()

        self.util_combo = ttk.Combobox(
            right,
            textvariable=self.util_var
        )

        self.util_combo.pack(fill="x")

        # ----------------------------------------
        # Voltage
        # ----------------------------------------

        ttk.Label(
            right,
            text="Voltage"
        ).pack(anchor="w")

        self.volt_var = tk.StringVar()

        self.volt_combo = ttk.Combobox(
            right,
            textvariable=self.volt_var
        )

        self.volt_combo.pack(fill="x")
        
        # ----------------------------------------
        # Direction
        # ----------------------------------------
        ttk.Label(
            right,
            text="Direction"
        ).pack(
            anchor="w"
        )
        self.dir_var = tk.StringVar()
        self.dir_combo = ttk.Combobox(
            right,
            textvariable=self.dir_var
        )
        self.dir_combo.pack(
            fill="x"
        )
        
        # ----------------------------------------
        # Button
        # ----------------------------------------
        actions = ttk.LabelFrame(
            right,
            text="Actions"
        )
        actions.pack(
            fill="x",
            pady=5
        )
        
        ttk.Button(
            actions,
            text="Generate Description",
            command=self.generate_description
        ).grid(row=0,column=0,padx=2,pady=2)
        
        #----------------------------------------
        # Save Button
        #----------------------------------------
        ttk.Button(
            actions,
            text="Save Pin",
            command=self.save_pin
        ).grid(row=0,column=1,padx=2,pady=2)
        
        ttk.Button(
            actions,
            text="CheckMe",
            command=self.check_pin
        ).grid(row=0,column=2,padx=2,pady=2)
        
        # print("Creating Apply Pinchar Defaults button")
        
        ttk.Button(
            actions,
            text="Apply Pinchar Defaults",
            command=self.apply_pinchar_defaults
        ).grid(row=0,column=3,padx=2,pady=2)
        
        ttk.Button(
            actions,
            text="Recalculate All Pins",
            command=self.recalculate_all_pins
        ).grid(row=1,column=0,padx=2,pady=2)
        
        ttk.Button(
            actions,
            text="Validate All Pins",
            command=self.validate_all_pins
        ).grid(row=1,column=1,padx=2,pady=2)
        
        ttk.Button(
            actions,
            text="Export Validation Report",
            command=self.export_validation_report
        ).grid(row=1,column=2,padx=2,pady=2)
        
        ttk.Button(
            actions,
            text="Fill All Pins",
            command=self.fill_all_pins
        ).grid(row=1,column=3,padx=2,pady=2)
        
        print(
            self.current_labels.keys()
        )
    
    def check_pin(self):
        if not hasattr(
                self,
                "selected_pin"):
            messagebox.showerror(
                "Validation Error",
                "No pin selected"
            )
            return
        #
        # Use full workbook pin
        #
        pin = dict(
            self.selected_pin
        )
        
        pin["function"] = \
                self.function_var.get()
        pin["direction"] = \
                self.dir_var.get()
        pin["voltage"] = \
                self.volt_var.get()
        pin["utilization"] = \
                self.util_var.get()
        pin["description"] = \
                self.desc_var.get()
        pin["characteristic"] = \
                self.characteristic_var.get()

        errors = self.svc.validate_pin(pin)

        if errors:
            message = "\n".join(errors)
            messagebox.showerror("Validation Error", message)
        else:
            messagebox.showinfo("Validation", "PASS")

    def build_validation_report(
            self,
            results):

        lines = []

        pass_count = 0
        error_count = 0

        for item in results:

            pin = item["pin"]
            errors = item["errors"]

            pin_name = pin.get(
                "description",
                "UNKNOWN"
            )

            if errors:

                error_count += 1

                lines.append(
                    f"ERROR  {pin_name}"
                )

                for err in errors:

                    lines.append(
                        f"    {err}"
                    )

            else:

                pass_count += 1

                lines.append(
                    f"PASS   {pin_name}"
                )

        lines.append("")
        lines.append("-" * 40)
        lines.append(
            f"PASS : {pass_count}"
        )
        lines.append(
            f"ERROR: {error_count}"
        )

        return "\n".join(lines)

    def generate_description(self):
        desc = PinDescriptionGenerator.build(
            self.clamp_var.get(),
            self.function_var.get(),
            self.util_var.get()
        )
        self.desc_var.set(
            desc
        )
        
        self.status_var.set(
            "Description generated"
        )
    
    def get_first_column(
        self,
        section):

        data = self.svc.db.get_section(
            section
        )

        values = []

        for row in data:

            if row:

                values.append(row[0])

        values = sorted(
            list(set(values))
        )

        return values    

    def populate_combos(self):

        self.dir_combo["values"] = (
            "IN",
            "OUT",
            "BI"
        )
        try:
            self.clamp_combo["values"] = \
                self.get_first_column(
                    "#CL_Clamp###"
                )
        except:
            pass

        try:
            self.function_combo["values"] = \
                self.get_first_column(
                    "#CL_Functions###"
                )
        except:
            pass

        try:
            self.util_combo["values"] = \
                self.get_first_column(
                    "#CL_Utilization###"
                )
        except:
            pass

        try:
            self.volt_combo["values"] = \
                self.get_first_column(
                    "#CL_Voltage###"
                )
        except:
            pass
    
    def create_search_tab(self):

        self.tab_search = ttk.Frame(
            self.notebook
        )

        self.notebook.add(
            self.tab_search,
            text="Database Search"
        )

        # -------------------------
        # Search Controls
        # -------------------------

        top = ttk.Frame(
        self.tab_search
        )

        top.pack(
            fill="x",
            padx=10,
            pady=10
        )

        self.search_type = tk.StringVar(
            value="Function"
        )

        self.cbo_search = ttk.Combobox(
            top,
            state="readonly",
            width=20,
            textvariable=self.search_type
        )

        self.cbo_search["values"] = (
            "Function",
            "Clamp",
            "Utilization",
            "Voltage",
            "BGER"
        )

        self.cbo_search.pack(
            side="left"
        )

        self.search_var = tk.StringVar()

        self.search_entry = ttk.Entry(
            top,
            width=40,
            textvariable=self.search_var
        )

        self.search_entry.pack(
            side="left",
            padx=5
        )

        ttk.Button(
            top,
            text="Search",
            command=self.run_search
        ).pack(
            side="left"
        )

        # -------------------------
        # Results
        # -------------------------

        self.result_box = tk.Listbox(
            self.tab_search,
            font=("Consolas", 10)
        )

        self.result_box.pack(
            fill="both",
            expand=True,
            padx=10,
            pady=10
        )

        self.result_box.bind(
            "<Double-Button-1>",
            self.on_result_double_click
        )

    def run_search(self):

        text = self.search_var.get().strip()
        mode = self.search_type.get()

        self.result_box.delete(0, tk.END)

        if not text:
            self.status_var.set('Enter a search term first')
            return

        if mode == 'Function':
            results = self.svc.search_function(text)
        elif mode == 'Clamp':
            results = self.svc.search_clamp(text)
        elif mode == 'Utilization':
            results = self.svc.search_utilization(text)
        elif mode == 'Voltage':
            results = self.svc.search_voltage(text)
        elif mode == 'BGER':
            results = self.svc.db.search_bger(text)
        else:
            results = []
        print("=" * 50)
        print("Mode =", mode)
        print("Text =", text)
        print("Results Found =", len(results))
        print("=" * 50)
        if not results:
            self.result_box.insert(tk.END, 'No matches found')
            self.status_var.set('0 results')
            return

        for row in results[:200]:
            self.result_box.insert(tk.END, ' ; '.join(row))

        self.status_var.set(f'{len(results)} results')
        
    def run_validation(self):
        results = self.svc.validate_all_pins()
        
        report = self.build_validation_report(
            results
        )
        self.validation_box.delete(
            0,
            tk.END
        )
        
        for line in report.splitlines():
            self.validation_box.insert(
                tk.END,
                line
            )
        
    def on_result_double_click(
            self,
            event):
        sel = self.result_box.curselection()
        if not sel:
            return
        text = self.result_box.get(
            sel[0]
        )
        parts = text.split(";")
        if not parts:
            return
        value = parts[0].strip()
        mode = self.search_type.get()
        if mode == "Clamp":
            self.clamp_var.set(value)
        elif mode == "Function":
            self.function_var.set(value)
            info = self.svc.find_cl_info(value)
            if info:
                self.util_var.set(info["utilization"])
                self.volt_var.set(info["voltage_type"])
                self.dir_var.set(info["direction"])
                self.generate_description()
        elif mode == "Utilization":
            self.util_var.set(value)
        elif mode == "Voltage":
            self.volt_var.set(value)
        self.generate_description()
        self.notebook.select(
            self.tab_pins
        )
        self.status_var.set(
            f"{mode} selected: {value}"
        )
        
    def create_validation_tab(self):

        self.tab_validation = ttk.Frame(
            self.notebook
        )

        self.notebook.add(
            self.tab_validation,
            text="Validation"
        )

        button_frame = ttk.Frame(
            self.tab_validation
        )

        button_frame.pack(
            fill="x",
            padx=10,
            pady=10
        )

        ttk.Button(
            button_frame,
            text="Run Validation",
            # command=self.validate_pins
            command=self.run_validation
        ).pack(
            side="left"
        )

        self.validation_box = tk.Listbox(
            self.tab_validation,
            font=("Consolas", 10)
        )

        self.validation_box.pack(
            fill="both",
            expand=True,
            padx=10,
            pady=10
        )    

    def validate_pins(self):
        self.validation_box.delete(
            0,
            tk.END
        )
        issues = self.pm.check_me()
        if not issues:
            self.validation_box.insert(
                tk.END,
                "PASS - No validation issues found."
            )
            self.status_var.set(
                "Validation Passed"
            )
            return
        count = 0
        for item in issues:
            row = item["row"]
            errors = item["errors"]
            if errors:
                count += 1
                self.validation_box.insert(
                    tk.END,
                    f"Row {row}: {', '.join(errors)}"
                )
        self.status_var.set(
            f"{count} issue(s) found"
        )
    
    def create_status_bar(self):
        self.status_var = tk.StringVar(value='Ready')
        ttk.Label(self.root, textvariable=self.status_var, relief='sunken').pack(side='bottom', fill='x')

    def load_pins(self):
        self.pins = self.pm.get_all_pins() or []
        self.pin_list.delete(0, tk.END)
        for pin in self.pins:
            text = (
                f"{pin['slot']}{pin['pin']} | "
                f"{pin['description']:<20} |"
                f"{pin['pin_type']}"
            )
            self.pin_list.insert(tk.END, text)

    def on_pin_selected(self, event):
        selection = self.pin_list.curselection()
        if not selection or not self.pins:
            return
        pin = self.pins[selection[0]]
        self.selected_pin = pin
        
        self.svc.workbook.debug_pin_row(
            pin["row"]
        )
        
        # print()
        # print("=" * 50)
        # print("PINCHAR")
        # print(pin.get("pinchar"))
        # print("=" * 50)
        
        self.desc_var.set(
            pin.get("description", "")
        )
        self.type_var.set(
            pin.get("pin_type", "")
        )
        self.clamp_var.set(
            pin.get(
                "clamp",
                ""
            )
        )
        self.function_var.set(
            pin.get(
                "function",
                ""
            )
        )
        self.util_var.set(
            pin.get(
                "utilization",
                ""
            )
        )
        self.volt_var.set(
            pin.get(
                "voltage",
                ""
            )
        )
        self.dir_var.set(
            pin.get(
                "direction",
                ""
            )
        )
        self.characteristic_var.set(
            pin.get(
                "characteristic",
                ""
            )
        )

        self.lbl_pin_info.config(
            text=
            f"Row: {pin['row']}   "
            f"Slot: {pin['slot']}   "
            f"Pin: {pin['pin']}   "
            f"Checksum: {pin['checksum']}"
        )
        self.status_var.set(
            f"Selected {pin['slot']}{pin['pin']}"
        )
        self.lbl_pin_stats.config(
            text = 
            f"Description Length: {len(self.desc_var.get())}  "
            f"Pin Type Length: {len(self.type_var.get())}"
        )
        
        definition = pin.get(
            "pinchar",
            {}
        )
        
        for field in [
            "I1",
            "I2",
            "I3",
            "I4",
            "I5",
            "I6",
            "T1",
            "T2"
        ]:
            value = definition.get(
                field,
                ""
            )
            
            if field in self.current_labels:
                actual_value = pin.get(field.lower(), "")
                if actual_value is None:
                    actual_value = ""
                
                self.current_vars[field].set(
                    str(actual_value)
                )
                self.current_labels[field].config(
                    text=
                    f"{field}: {value}"
                )
        
        # desc = pin.get("description", "")
        # if "#" in desc:
        #     parts = desc.replace("*", "").split("#")
        #     if len(parts) >= 1:
        #         self.clamp_var.set(parts[0])
        #     if len(parts) >= 2:
        #         self.function_var.set(parts[1])
        #     if len(parts) >= 3:
        #         self.util_var.set(parts[2])
        #     if len(parts) >= 4:
        #         self.volt_var.set(parts[3])
    
    def save_pin(self):
        if self.selected_pin is None:
            return
        row = self.selected_pin["row"]
        self.svc.update_pin_description(
            row,
            self.desc_var.get()
        )
        self.svc.update_pin_type(
            row,
            self.type_var.get()
        )
        self.svc.update_pin_characteristic(
            row,
            self.characteristic_var.get()
        )
        self.svc.update_pin_clamp(
            row,
            self.clamp_var.get()
        )
        self.svc.update_pin_function(
            row,
            self.function_var.get()
        )
        self.svc.update_pin_utilization(
            row,
            self.util_var.get()
        )
        self.svc.update_pin_voltage(
            row,
            self.volt_var.get()
        )
        self.svc.update_pin_direction(
            row,
            self.dir_var.get()
        )
        self.svc.update_pin_i1(
            row,
            self.current_vars["I1"].get()
        )
        self.svc.update_pin_i2(
            row,
            self.current_vars["I2"].get()
        )
        self.svc.update_pin_i3(
            row,
            self.current_vars["I3"].get()
        )
        self.svc.update_pin_i4(
            row,
            self.current_vars["I4"].get()
        )
        self.svc.update_pin_i5(
            row,
            self.current_vars["I5"].get()
        )
        self.svc.update_pin_i6(
            row,
            self.current_vars["I6"].get()
        )
        self.svc.update_pin_t1(
            row,
            self.current_vars["T1"].get()
        )
        self.svc.update_pin_t2(
            row,
            self.current_vars["T2"].get()
        )
        print("Saving")
        print("Description:", self.desc_var.get())
        print("Type:", self.type_var.get())
        print("Clamp:", self.clamp_var.get())
        print("Function:", self.function_var.get())
        self.svc.save()
        self.status_var.set(
            f"Saved Pin {self.selected_pin['slot']}{self.selected_pin['pin']}"
        )
        self.load_pins()
        
    def save_as(self):
        filename = filedialog.asksaveasfilename(
            defaultextension=".xlsm",
            filetypes=[
                (
                    "Excel Macro Workbook",
                    "*.xlsm"
                )
            ]
        )
        if not filename:
            return
        try:
           self.svc.save(filename)
           self.status_var.set(
                f"Saved: {filename}"
            )
        except Exception as ex:
            self.status_var.set(
                str(ex)
            )
            
    def apply_pinchar_defaults(self):

        print()
        print("=" * 50)
        print("Apply Pinchar Defaults Clicked")
        print(
            "selected Pin =",
            self.selected_pin
        )
        print("=" * 50)
        if self.selected_pin is None:
            return

        characteristic = self.characteristic_var.get()

        values = {

            "i1": self.current_vars["I1"].get(),
            "i2": self.current_vars["I2"].get(),
            "i3": self.current_vars["I3"].get(),
            "i4": self.current_vars["I4"].get(),
            "i5": self.current_vars["I5"].get(),
            "i6": self.current_vars["I6"].get(),
            "t1": self.current_vars["T1"].get(),
            "t2": self.current_vars["T2"].get(),
        }

        defaults = \
            self.svc.workbook.pinchar.calculate_defaults(
                characteristic,
                values
            )

        print()
        print("PINCHAR DEFAULTS")
        print(defaults)

        for field, value in defaults.items():

            self.current_vars[
                field.upper()
            ].set(
                str(value)
            )
    
    def recalculate_all_pins(self):
        
        print()
        print("=" * 60)
        print("RECALCULATE ALL PINS")
        print("=" * 60)
        
        pins = self.svc.get_all_pins()
        
        count = 0
        
        for pin in pins:
            characteristic = str(
                pin.get(
                    "characteristic",
                    ""
                )
            ).strip()
            
            if not characteristic:
                continue
            
            values = {
                "i1": pin.get("i1"),
                "i2": pin.get("i2"),
                "i3": pin.get("i3"),
                "i4": pin.get("i4"),
                "i5": pin.get("i5"),
                "i6": pin.get("i6"),
                "t1": pin.get("t1"),
                "t2": pin.get("t2")
            }
            
            defaults = \
                self.svc.workbook.pinchar.calculate_defaults(
                    characteristic,
                    values
                )
            
            row = pin["row"]
            
            #
            # write results
            #
            if "i2" in defaults:
                self.svc.update_pin_i2(
                    row,
                    defaults["i2"]
                )
            
            if "i3" in defaults:
                self.svc.update_pin_i3(
                    row,
                    defaults["i3"]
                )
                
            if "t1" in defaults:
                self.svc.update_pin_t1(
                    row,
                    defaults["t1"]
                )
            
            if "t2" in defaults:
                self.svc.update_pin_t2(
                    row,
                    defaults["t2"]
                )
                
            count += 1
            
        print(
            f"Pins processed: {count}"
        )
        
        self.svc.save()
        
        messagebox.showinfo(
            "Pinchar",
            f"{count} pins recalculated and saved"
        )
    
    def validate_all_pins(self):
        
        results = self.svc.validate_all_pins()
        
        total = len(results)
        
        errors = 0
        
        for item in results:
            if item["errors"]:
                errors += 1
        
        passed = total - errors
        messagebox.showinfo(
            "Validation Results",
            f"Total Pins: {total}\n"
            f"PASS: {passed}\n"
            f"FAIL: {errors}"
        )
        
    def export_validation_report(self):
        filename = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[
                ("CSV Files", "*.csv")
            ]
        )
        
        if not filename:
            return
        
        results = self.svc.validate_all_pins()
        
        with open(
            filename,
            "w",
            newline="",
            encoding="utf-8"
        ) as f:
            writer = csv.writer(f)
            writer.writerow([
                "Slot",
                "Pin",
                "Description",
                "Status",
                "Errors"
            ])
            
            for item in results:
                pin = item["pin"]
                errors = item["errors"]
                writer.writerrom([
                    pin.get("slot"),
                    pin.get("Pin"),
                    pin.get(
                        "description",
                        ""
                    ),
                    (
                        "PASS"
                        if not errors
                        else "FAIL"
                    ),
                    " | ".join(errors)
                ])
    
    def fill_all_pins(self):
        from tkinter import messagebox
        pins = self.svc.get_all_pins()
        count = 0
        for pin in pins:
            function_name = str(
                pin.get(
                    "function",
                    ""
                )
            ).strip()
            
            if not function_name:
                continue
            
            info = self.svc.find_cl_info(
                function_name
            )
            
            if not info:
                continue
            
            row = pin["row"]
            
            #
            # CL Info
            #
            
            self.svc.update_pin_direction(
                row,
                info.get(
                    "direction",
                    ""
                )
            )
            
            self.svc.update_pin_voltage(
                row,
                info.get(
                    "voltage_type",
                    ""
                )
            )
            
            self.svc.update_pin_utilization(
                row,
                info.get(
                    "utilization",
                    ""
                )
            )
            
            count += 1
        self.svc.save()
        messagebox.showinfo(
            "Fill All Pins",
            f"{count} pins updated."
        )
        
if __name__ == '__main__':
    root = tk.Tk()
    app = VobesGUI(root)
    root.mainloop()
