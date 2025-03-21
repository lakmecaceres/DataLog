import sys
import os
import json
import pandas as pd
import pyperclip
import dateutil.parser
from datetime import datetime
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                             QLabel, QLineEdit, QComboBox, QPushButton, QScrollArea,
                             QMessageBox, QGridLayout, QGroupBox, QTabWidget, QFileDialog)
from PyQt6.QtCore import Qt, QTimer


class DataLogGUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Krienen Data Logger")
        self.init_constants()  # Initialize constants first
        self.black_fill = PatternFill(start_color='000000', fill_type='solid')
        self.bold_font = Font(bold=True)
        QTimer.singleShot(0, self.init_ui)

    def get_save_location(self):
        config_file = 'config.json'

        # Check if the config file exists and read the file location from it
        if os.path.exists(config_file):
            with open(config_file, 'r') as f:
                config = json.load(f)
                file_location = config.get('file_location')

                # Check if the saved file location is writable
                if os.access(os.path.dirname(file_location), os.W_OK):
                    return file_location

        # Prompt the user to select a save location
        file_name, _ = QFileDialog.getSaveFileName(
            self,
            "Save Excel File",
            "",  # Default directory, empty means last used
            "Excel Files (*.xlsx);;All Files (*)"
        )

        if file_name:
            # Save the file location to the config file
            with open(config_file, 'w') as f:
                json.dump({'file_location': file_name}, f)

        return file_name

    def on_submit(self):
        try:
            if not self.validate_inputs():
                return

            # Get or prompt for file location
            file_location = self.get_save_location()
            if not file_location:
                return  # User canceled the save dialog

            # Process the form data and update Excel
            self.process_form_data(file_location)

            # Adjust column widths
            workbook = load_workbook(file_location)
            worksheet = workbook.active

            for column in worksheet.columns:
                max_length = 0
                column_letter = get_column_letter(column[0].column)
                for cell in column:
                    try:
                        cell_value = str(cell.value)
                        if len(cell_value) > max_length:
                            max_length = len(cell_value)
                    except:
                        pass
                adjusted_width = (max_length + 2)
                worksheet.column_dimensions[column_letter].width = adjusted_width

            workbook.save(file_location)

            QMessageBox.information(
                self,
                "Success",
                f"Data successfully appended to {file_location}"
            )

            # Clear form fields after successful submission
            self.clear_form_fields()

        except Exception as e:
            QMessageBox.critical(
                self,
                "Error",
                f"An error occurred while processing the data:\n{str(e)}"
            )

    # Then find where you connect your save button and update it to use this method:
    def setup_buttons(self):  # or whatever method contains your button setup
        self.save_button = QPushButton("Save")
        self.save_button.clicked.connect(self.save_data)  # Connect to the new save_data method

    def delayed_init(self):
        self.init_ui()

    def init_constants(self):
        if getattr(sys, 'frozen', False):
            self.script_dir = os.path.dirname(sys.executable)
        else:
            self.script_dir = os.path.dirname(os.path.abspath(__file__))

        self.COUNTER_FILE = os.path.join(self.script_dir, 'sample_name_counter.json')
        self.workbook_path = os.path.join(self.script_dir, 'datalog.xlsx')

        self.name_to_code = {
            "Croissant": "CJ23.56.002",
            "Nutmeg": "CJ23.56.003",
            "Jellybean": "CJ24.56.001",
            "Rambo": "CJ24.56.004",
            "Morel": "CJ24.56.015"
        }

        self.tile_location_map = {
            "BRAINSTEM": "BS",
            "BS": "BS",
            "CORTEX": "CX",
            "CX": "CX",
            "CEREBELLUM": "CB",
            "CB": "CB"
        }

        self.load_counter_data()

    def init_ui(self):
        self.setGeometry(100, 100, 1000, 800)

        main_widget = QWidget()
        self.setCentralWidget(main_widget)

        scroll = QScrollArea()
        main_widget.layout = QVBoxLayout(main_widget)
        main_widget.layout.addWidget(scroll)

        content_widget = QWidget()
        scroll.setWidget(content_widget)
        scroll.setWidgetResizable(True)

        self.layout = QVBoxLayout(content_widget)

        # Create tab widget
        tab_widget = QTabWidget()
        self.layout.addWidget(tab_widget)

        # Create tabs with new names
        tissue_tab = QWidget()
        facs_tab = QWidget()
        library_tab = QWidget()
        indices_tab = QWidget()

        # Setup layouts for each tab
        self.setup_basic_tab(tissue_tab)
        self.setup_facs_tab(facs_tab)
        self.setup_library_tab(library_tab)
        self.setup_indices_tab(indices_tab)

        # Add tabs to widget with new names
        tab_widget.addTab(tissue_tab, "Tissue")
        tab_widget.addTab(facs_tab, "FACS")
        tab_widget.addTab(library_tab, "cDNA")
        tab_widget.addTab(indices_tab, "Libraries")

        # Add submit button
        self.submit_btn = QPushButton('Submit')
        self.submit_btn.clicked.connect(self.on_submit)
        self.layout.addWidget(self.submit_btn)

    def setup_basic_tab(self, tab):
        layout = QGridLayout()

        # Project selection (at top)
        self.project_input = QComboBox()
        self.project_input.addItems(["HMBA_CjAtlas_Subcortex", "Other"])
        self.project_input.currentTextChanged.connect(self.on_project_change)
        self.project_name_input = QLineEdit()
        self.project_name_input.setVisible(False)
        layout.addWidget(QLabel("Project:"), 0, 0)
        layout.addWidget(self.project_input, 0, 1)
        layout.addWidget(self.project_name_input, 0, 2)

        # Date input with validation
        self.date_input = QLineEdit()
        self.date_input.setPlaceholderText("YYMMDD or MM/DD/YY")
        layout.addWidget(QLabel("Experiment Date:"), 1, 0)
        layout.addWidget(self.date_input, 1, 1)

        # Marmoset name with dropdown
        self.marmoset_input = QComboBox()
        self.marmoset_input.addItems(self.name_to_code.keys())
        layout.addWidget(QLabel("Marmoset Name:"), 2, 0)
        layout.addWidget(self.marmoset_input, 2, 1)

        # Hemisphere selection
        self.hemisphere_input = QComboBox()
        self.hemisphere_input.addItems(["Left (LH)", "Right (RH)", "Both"])
        layout.addWidget(QLabel("Hemisphere:"), 3, 0)
        layout.addWidget(self.hemisphere_input, 3, 1)

        # Tile location multiselect (moved right after hemisphere)
        self.tile_location_input = QComboBox()
        self.tile_location_input.addItems(["BS", "CX", "CB"])
        self.tile_location_input.setEditable(True)
        layout.addWidget(QLabel("Tile Location:"), 4, 0)
        layout.addWidget(self.tile_location_input, 4, 1)

        # Slab and tile numbers
        self.slab_input = QLineEdit()
        self.slab_input.setPlaceholderText("Enter numeric value")
        layout.addWidget(QLabel("Slab Number:"), 5, 0)
        layout.addWidget(self.slab_input, 5, 1)

        self.tile_input = QLineEdit()
        self.tile_input.setPlaceholderText("Enter numeric value")
        layout.addWidget(QLabel("Tile Number:"), 6, 0)
        layout.addWidget(self.tile_input, 6, 1)

        tab.setLayout(layout)

    def setup_facs_tab(self, tab):
        layout = QGridLayout()

        # Sorter Initials
        self.sorter_initials_input = QLineEdit()
        self.sorter_initials_input.setPlaceholderText("Enter sorter's initials")
        layout.addWidget(QLabel("Sorter Initials:"), 0, 0)
        layout.addWidget(self.sorter_initials_input, 0, 1)

        # Sort method
        self.sort_method_input = QComboBox()
        self.sort_method_input.addItems(["pooled", "unsorted", "DAPI"])
        self.sort_method_input.currentTextChanged.connect(self.on_sort_method_change)
        layout.addWidget(QLabel("Sort Method:"), 1, 0)
        layout.addWidget(self.sort_method_input, 1, 1)

        # FACS population (moved before number of reactions)
        self.facs_population_input = QLineEdit()
        self.facs_population_input.setPlaceholderText("Format: XX/XX/XX (e.g., 70/20/10)")
        layout.addWidget(QLabel("FACS Population:"), 2, 0)
        layout.addWidget(self.facs_population_input, 2, 1)

        # Number of Reactions
        self.rxn_number_input = QLineEdit()
        self.rxn_number_input.setPlaceholderText("Enter number of reactions")
        layout.addWidget(QLabel("Number of Reactions:"), 3, 0)
        layout.addWidget(self.rxn_number_input, 3, 1)

        # Expected Recovery (moved from library tab)
        self.expected_recovery_input = QLineEdit()
        layout.addWidget(QLabel("Expected Recovery:"), 4, 0)
        layout.addWidget(self.expected_recovery_input, 4, 1)

        # Nuclei Concentration (moved from library tab)
        self.nuclei_concentration_input = QLineEdit()
        layout.addWidget(QLabel("Nuclei Concentration:"), 5, 0)
        layout.addWidget(self.nuclei_concentration_input, 5, 1)

        # Nuclei Volume (moved from library tab)
        self.nuclei_volume_input = QLineEdit()
        layout.addWidget(QLabel("Nuclei Volume (µL):"), 6, 0)
        layout.addWidget(self.nuclei_volume_input, 6, 1)

        tab.setLayout(layout)

    def setup_library_tab(self, tab):
        layout = QGridLayout()

        # Library dates
        self.cdna_amp_date_input = QLineEdit()
        self.cdna_amp_date_input.setPlaceholderText("YYMMDD or MM/DD/YY")
        layout.addWidget(QLabel("cDNA Amplification Date:"), 0, 0)
        layout.addWidget(self.cdna_amp_date_input, 0, 1)

        self.atac_prep_date_input = QLineEdit()
        self.atac_prep_date_input.setPlaceholderText("YYMMDD or MM/DD/YY")
        layout.addWidget(QLabel("ATAC Library Prep Date:"), 1, 0)
        layout.addWidget(self.atac_prep_date_input, 1, 1)

        self.rna_prep_date_input = QLineEdit()
        self.rna_prep_date_input.setPlaceholderText("YYMMDD or MM/DD/YY")
        layout.addWidget(QLabel("cDNA Library Prep Date:"), 2, 0)
        layout.addWidget(self.rna_prep_date_input, 2, 1)

        # PCR cycles
        self.cdna_pcr_cycles_input = QLineEdit()
        self.cdna_pcr_cycles_input.setPlaceholderText("Comma-separated values")
        layout.addWidget(QLabel("cDNA PCR Cycles:"), 3, 0)
        layout.addWidget(self.cdna_pcr_cycles_input, 3, 1)

        # cDNA metrics (reordered)
        self.cdna_concentration_input = QLineEdit()
        self.cdna_concentration_input.setPlaceholderText("Comma-separated values (ng/µL)")
        layout.addWidget(QLabel("cDNA Concentration:"), 4, 0)
        layout.addWidget(self.cdna_concentration_input, 4, 1)

        self.percent_cdna_400bp_input = QLineEdit()
        self.percent_cdna_400bp_input.setPlaceholderText("Comma-separated values")
        layout.addWidget(QLabel("Percent cDNA > 400bp:"), 5, 0)
        layout.addWidget(self.percent_cdna_400bp_input, 5, 1)

        tab.setLayout(layout)

    def on_sort_method_change(self, value):
        # Update FACS population field based on sort method
        if value.lower() == "pooled":
            self.facs_population_input.setEnabled(True)
            self.facs_population_input.setPlaceholderText("Format: XX/XX/XX (e.g., 70/20/10)")
        elif value.lower() == "unsorted":
            self.facs_population_input.setEnabled(False)
            self.facs_population_input.setText("no_FACS")
        else:  # DAPI
            self.facs_population_input.setEnabled(False)
            self.facs_population_input.setText("DAPI")

    def on_project_change(self, value):
        self.project_name_input.setVisible(value == "Other")

    def load_counter_data(self):
        if os.path.exists(self.COUNTER_FILE):
            with open(self.COUNTER_FILE, 'r') as f:
                try:
                    self.counter_data = json.load(f)
                except json.JSONDecodeError:
                    self.counter_data = {}
        else:
            self.counter_data = {}

        self.counter_data.setdefault("next_counter", 90)
        self.counter_data.setdefault("date_info", {})
        self.counter_data.setdefault("amp_counter", {})

    def setup_indices_tab(self, tab):
        layout = QGridLayout()

        # ATAC fields grouped together
        self.library_cycles_atac_input = QLineEdit()
        self.library_cycles_atac_input.setPlaceholderText("Comma-separated values")
        layout.addWidget(QLabel("ATAC Library Cycles:"), 0, 0)
        layout.addWidget(self.library_cycles_atac_input, 0, 1)

        self.atac_indices_input = QLineEdit()
        self.atac_indices_input.setPlaceholderText("Comma-separated values (e.g., D4,E5,F6)")
        layout.addWidget(QLabel("ATAC Indices:"), 1, 0)
        layout.addWidget(self.atac_indices_input, 1, 1)

        self.atac_lib_concentration_input = QLineEdit()
        self.atac_lib_concentration_input.setPlaceholderText("Comma-separated values (ng/µL)")
        layout.addWidget(QLabel("ATAC Library Concentration:"), 2, 0)
        layout.addWidget(self.atac_lib_concentration_input, 2, 1)

        self.atac_sizes_input = QLineEdit()
        self.atac_sizes_input.setPlaceholderText("Comma-separated values")
        layout.addWidget(QLabel("ATAC Library Sizes (bp):"), 3, 0)
        layout.addWidget(self.atac_sizes_input, 3, 1)

        # cDNA fields grouped together
        self.library_cycles_rna_input = QLineEdit()
        self.library_cycles_rna_input.setPlaceholderText("Comma-separated values")
        layout.addWidget(QLabel("cDNA Library Cycles:"), 4, 0)
        layout.addWidget(self.library_cycles_rna_input, 4, 1)

        self.rna_indices_input = QLineEdit()
        self.rna_indices_input.setPlaceholderText("Comma-separated values (e.g., A1,B2,C3)")
        layout.addWidget(QLabel("cDNA Indices:"), 5, 0)
        layout.addWidget(self.rna_indices_input, 5, 1)

        self.rna_lib_concentration_input = QLineEdit()
        self.rna_lib_concentration_input.setPlaceholderText("Comma-separated values (ng/µL)")
        layout.addWidget(QLabel("cDNA Library Concentration:"), 6, 0)
        layout.addWidget(self.rna_lib_concentration_input, 6, 1)

        self.rna_sizes_input = QLineEdit()
        self.rna_sizes_input.setPlaceholderText("Comma-separated values")
        layout.addWidget(QLabel("cDNA Library Sizes (bp):"), 7, 0)
        layout.addWidget(self.rna_sizes_input, 7, 1)

        tab.setLayout(layout)

    def convert_index(self, index):
        index = index.strip().upper()
        if len(index) == 3:
            if index[0].isdigit() and index[1].isdigit() and index[2].isalpha():
                return f"{index[2]}{index[0]}{index[1]}"
            elif index[0].isalpha() and index[1].isdigit() and index[2].isdigit():
                return index
        elif len(index) == 2:
            if index[0].isdigit() and index[1].isalpha():
                return f"{index[1]}0{index[0]}"
            elif index[0].isalpha() and index[1].isdigit():
                return f"{index[0]}0{index[1]}"
        return None

    def pad_index(self, index):
        if len(index) == 2 and index[0].isalpha() and index[1].isdigit():
            return f"{index[0]}0{index[1]}"
        return index

    def convert_date(self, exp_date):
        clean_date = "".join(c for c in exp_date if c.isdigit())
        if len(clean_date) == 6:
            try:
                datetime.strptime(clean_date, '%y%m%d')
                return clean_date
            except ValueError:
                pass
        try:
            parsed_date = dateutil.parser.parse(exp_date)
            return parsed_date.strftime('%y%m%d')
        except ValueError:
            return None

    def validate_inputs(self):
        # Basic validation
        current_date = self.convert_date(self.date_input.text())
        if not current_date:
            QMessageBox.warning(self, "Validation Error", "Please enter a valid date.")
            return False

        try:
            rxn_number = int(self.rxn_number_input.text())
            if rxn_number <= 0:
                raise ValueError
        except ValueError:
            QMessageBox.warning(self, "Validation Error", "Please enter a valid number of reactions.")
            return False

        # Validate numerical inputs
        try:
            slab_int = int(self.slab_input.text())
            tile_int = int(self.tile_input.text())
        except ValueError:
            QMessageBox.warning(self, "Validation Error", "Slab and tile numbers must be numeric values.")
            return False

        # Validate FACS population for pooled samples
        if self.sort_method_input.currentText().lower() == "pooled":
            proportions = self.facs_population_input.text().strip()
            if "/" not in proportions:
                QMessageBox.warning(self, "Validation Error",
                                    "Please enter FACS population proportions in format XX/XX/XX")
                return False
            try:
                proportions_list = [int(p) for p in proportions.split("/")]
                if len(proportions_list) != 3 or sum(proportions_list) != 100:
                    raise ValueError
            except ValueError:
                QMessageBox.warning(self, "Validation Error", "FACS proportions must be three numbers that sum to 100")
                return False

        for field, field_name in [
            (self.percent_cdna_400bp_input, "Percent cDNA > 400bp"),
            (self.cdna_concentration_input, "cDNA concentration"),
            (self.rna_lib_concentration_input, "RNA library concentration"),
            (self.atac_lib_concentration_input, "ATAC library concentration")
        ]:
            values = field.text().strip().split(',')
            try:
                values = [float(x.strip()) for x in values]
                if len(values) != rxn_number:
                    QMessageBox.warning(self, "Validation Error",
                                        f"{field_name} must have {rxn_number} comma-separated values")
                    return False
            except ValueError:
                QMessageBox.warning(self, "Validation Error",
                                    f"{field_name} values must be numbers")
                return False

        # Validate comma-separated inputs match reaction number
        fields_to_validate = [
            (self.cdna_pcr_cycles_input, "cDNA PCR cycles"),
            (self.rna_indices_input, "RNA indices"),
            (self.atac_indices_input, "ATAC indices"),
            (self.rna_sizes_input, "RNA library sizes"),
            (self.atac_sizes_input, "ATAC library sizes"),
            (self.library_cycles_rna_input, "RNA library cycles"),
            (self.library_cycles_atac_input, "ATAC library cycles")
        ]

        for field, field_name in fields_to_validate:
            values = field.text().strip().split(',')
            if len(values) != rxn_number:
                QMessageBox.warning(self, "Validation Error",
                                    f"{field_name} must have {rxn_number} comma-separated values")
                return False

        return True

    def initialize_excel(self):
        wb = Workbook()
        ws = wb.active
        ws.title = "HMBA"
        headers = ['krienen_lab_identifier', 'seq_portal', 'elab_link', 'experiment_start_date',
                   'mit_name', 'donor_name', 'tissue_name', 'tissue_name_old',
                   'dissociated_cell_sample_name', 'facs_population_plan', 'cell_prep_type',
                   'study', 'enriched_cell_sample_container_name', 'expc_cell_capture',
                   'port_well', 'enriched_cell_sample_name', 'enriched_cell_sample_quantity_count',
                   'barcoded_cell_sample_name', 'library_method', 'cDNA_amplification_method',
                   'cDNA_amplification_date', 'amplified_cdna_name', 'cDNA_pcr_cycles',
                   'rna_amplification_pass_fail', 'percent_cdna_longer_than_400bp',
                   'cdna_amplified_quantity_ng', 'cDNA_library_input_ng', 'library_creation_date',
                   'library_prep_set', 'library_name', 'tapestation_avg_size_bp',
                   'library_num_cycles', 'lib_quantification_ng', 'library_prep_pass_fail',
                   'r1_index', 'r2_index', 'ATAC_index']

        ws.append(headers)

        # Apply Arial 10 font to headers
        header_font = Font(name="Arial", size=10, bold=True)
        for col_num, header in enumerate(headers, start=1):
            ws.cell(row=1, column=col_num).font = header_font

        return wb

    def process_form_data(self, file_location):
        # Load or create workbook
        if os.path.exists(file_location):
            workbook = load_workbook(file_location)
        else:
            workbook = self.initialize_excel()

        worksheet = workbook.active
        current_row = worksheet.max_row
        if current_row == 1 and not any(cell.value for cell in worksheet[1]):
            current_row = 1
        else:
            current_row += 1

        # Get form values
        current_date = self.convert_date(self.date_input.text())
        mit_name_input = self.marmoset_input.currentText()
        mit_name = "cj" + mit_name_input
        donor_name = self.name_to_code[mit_name_input]  # Define donor_name here

        # Process slab and hemisphere
        slab = self.slab_input.text().strip()
        hemisphere = self.hemisphere_input.currentText().split()[0].upper()
        if hemisphere == "RIGHT":
            slab = str(int(slab) + 40).zfill(2)
        elif hemisphere == "BOTH":
            slab = str(int(slab) + 90).zfill(2)
        else:
            slab = slab.zfill(2)

        tile = str(int(self.tile_input.text())).zfill(2)

        # Process tile location
        tile_location_abbr = self.tile_location_input.currentText()

        # Sort method and FACS population
        sort_method = self.sort_method_input.currentText()
        sort_method = sort_method.upper() if sort_method.lower() == "dapi" else sort_method

        if sort_method.lower() == "pooled":
            facs_population = self.facs_population_input.text()
        elif sort_method.lower() == "unsorted":
            facs_population = "no_FACS"
        else:
            facs_population = "DAPI"

        # Get reaction number and update counters
        rxn_number = int(self.rxn_number_input.text())

        # Update date_info
        if current_date not in self.counter_data["date_info"]:
            self.counter_data["date_info"][current_date] = {
                "total_reactions": 0,
                "batches": []
            }

        date_info = self.counter_data["date_info"]
        date_entry = date_info[current_date]
        existing_total = date_entry["total_reactions"]

        # Calculate batch information
        total_reactions_after = existing_total + rxn_number
        batches_before = (existing_total + 7) // 8
        batches_after = (total_reactions_after + 7) // 8
        new_batches_needed = batches_after - batches_before

        new_p_numbers = [self.counter_data["next_counter"] + i for i in range(new_batches_needed)]
        self.counter_data["next_counter"] += new_batches_needed

        all_batches = date_entry["batches"].copy()
        all_batches.extend({"p_number": p, "count": 0} for p in new_p_numbers)

        # Calculate port wells
        port_wells = []
        for x in range(rxn_number):
            global_idx = existing_total + x + 1
            batch_idx = (global_idx - 1) // 8
            p_number = all_batches[batch_idx]["p_number"]
            port_well = (global_idx - 1) % 8 + 1
            port_wells.append((p_number, port_well))

        # Update counters
        date_entry["total_reactions"] = total_reactions_after
        date_entry["batches"] = all_batches

        # Process indices
        atac_indices = [self.convert_index(index) for index in self.atac_indices_input.text().split(",")]
        atac_indices = [self.pad_index(index) for index in atac_indices]

        rna_indices = [self.convert_index(index) for index in self.rna_indices_input.text().split(",")]
        rna_indices = [self.pad_index(index) for index in rna_indices]

        # Initialize common values
        seq_portal = "no"
        elab_link = pyperclip.paste()
        tissue_name = f"{donor_name}.{tile_location_abbr}.{slab}.{tile}"
        dissociated_cell_sample_name = f'{current_date}_{tissue_name}.Multiome'
        cell_prep_type = "nuclei"

        sorting_status = "PS" if sort_method.lower() in ["pooled", "dapi"] else "PN"
        sorter_initials = self.sorter_initials_input.text().strip().upper()
        enriched_cell_sample_container_name = f"MPXM_{current_date}_{sorting_status}_{sorter_initials}"

        # Get study name
        study = "HMBA_CjAtlas_Subcortex" if self.project_input.currentText() == "HMBA_CjAtlas_Subcortex" else self.project_name_input.text()

        # Process the data for each reaction and modality
        dup_index_counter = {}
        headers = [cell.value for cell in worksheet[1]]

        for x in range(rxn_number):
            p_number, port_well = port_wells[x]
            barcoded_cell_sample_name = f'P{str(p_number).zfill(4)}_{port_well}'

            for modality in ["RNA", "ATAC"]:
                self.write_modality_data(
                    worksheet, current_row, modality, x,
                    current_date, mit_name, slab, tile, sort_method,
                    port_well, barcoded_cell_sample_name,
                    sorting_status, sorter_initials,
                    tissue_name, dissociated_cell_sample_name,
                    enriched_cell_sample_container_name,
                    study, seq_portal, elab_link,
                    facs_population, cell_prep_type,
                    rna_indices, atac_indices,
                    headers, dup_index_counter,
                    donor_name  # Add donor_name here
                )
                current_row += 1

        # Save workbook and counter data
        workbook.save(file_location)
        with open(self.COUNTER_FILE, 'w') as f:
            json.dump(self.counter_data, f, indent=4)

    def write_modality_data(self, worksheet, current_row, modality, x, *args):
        (current_date, mit_name, slab, tile, sort_method,
         port_well, barcoded_cell_sample_name,
         sorting_status, sorter_initials,
         tissue_name, dissociated_cell_sample_name,
         enriched_cell_sample_container_name,
         study, seq_portal, elab_link,
         facs_population, cell_prep_type,
         rna_indices, atac_indices,
         headers, dup_index_counter, donor_name) = args

        krienen_lab_identifier = f'{current_date}_HMBA_{mit_name}_Slab{int(slab)}_Tile{int(tile)}_{sort_method}_{modality}{x + 1}'
        enriched_cell_sample_name = f'MPXM_{current_date}_{sorting_status}_{sorter_initials}_{port_well}'

        library_prep_date = (self.convert_date(self.rna_prep_date_input.text()) if modality == "RNA"
                             else self.convert_date(self.atac_prep_date_input.text()))

        if modality == "RNA":
            library_method = "10xMultiome-RSeq"
            library_type = "LPLCXR"
            library_index = rna_indices[x]

            # Calculate RNA-specific metrics
            cdna_concentration = float(self.cdna_concentration_input.text().split(',')[x])
            cdna_amplified_quantity = cdna_concentration * 40  # 40µL volume for cDNA
            cdna_library_input = cdna_amplified_quantity * 0.25  # 25% of amplified quantity
            percent_cdna_400bp = float(self.percent_cdna_400bp_input.text().split(',')[x])
            rna_concentration = float(self.rna_lib_concentration_input.text().split(',')[x])
            lib_quant = rna_concentration * 35  # Fixed 35µL volume for RNA library

            cdna_pcr_cycles = int(self.cdna_pcr_cycles_input.text().split(',')[x])
            rna_size = int(self.rna_sizes_input.text().split(',')[x])
            library_cycles = int(self.library_cycles_rna_input.text().split(',')[x])
        else:  # ATAC
            library_method = "10xMultiome-ASeq"
            library_type = "LPLCXA"
            library_index = atac_indices[x]

            # Calculate ATAC-specific metrics
            atac_concentration = float(self.atac_lib_concentration_input.text().split(',')[x])
            lib_quant = atac_concentration * 20  # Fixed 20µL volume for ATAC library

            atac_size = int(self.atac_sizes_input.text().split(',')[x])
            library_cycles = int(self.library_cycles_atac_input.text().split(',')[x])

        # Update library prep set counter
        key = (library_type, library_prep_date, library_index)
        dup_index_counter[key] = dup_index_counter.get(key, 0) + 1
        library_prep_set = f"{library_type}_{library_prep_date}_{dup_index_counter[key]}"
        library_name = f"{library_prep_set}_{library_index}"

        # Calculate common metrics
        expected_cell_capture = int(self.expected_recovery_input.text())
        concentration = float(self.nuclei_concentration_input.text().replace(",", ""))
        volume = float(self.nuclei_volume_input.text())
        enriched_cell_sample_quantity_count = round(concentration * volume)

        # Prepare row data
        row_data = [
            krienen_lab_identifier,  # Column 1
            seq_portal,
            elab_link,
            current_date,
            mit_name,
            donor_name,
            tissue_name,
            None,  # tissue_name_old (will be filled black)
            dissociated_cell_sample_name,
            facs_population,
            cell_prep_type,
            study,
            enriched_cell_sample_container_name,
            expected_cell_capture,
            port_well,
            enriched_cell_sample_name,
            enriched_cell_sample_quantity_count,
            barcoded_cell_sample_name,
            library_method,
            "10xMultiome-RSeq" if modality == "RNA" else None,
            self.convert_date(self.cdna_amp_date_input.text()) if modality == "RNA" else None,
            None,  # amplified_cdna_name (filled conditionally)
            cdna_pcr_cycles if modality == "RNA" else None,
            "Pass" if modality == "RNA" else None,
            percent_cdna_400bp if modality == "RNA" else None,
            cdna_amplified_quantity if modality == "RNA" else None,
            cdna_library_input if modality == "RNA" else None,
            library_prep_date,
            library_prep_set,
            library_name,
            rna_size if modality == "RNA" else atac_size,
            library_cycles,
            lib_quant,
            "Pass",
            f"SI-TT-{rna_indices[x]}_i7" if modality == "RNA" else None,
            f"SI-TT-{rna_indices[x]}_b(i5)" if modality == "RNA" else None,
            f"SI-NA-{atac_indices[x]}" if modality == "ATAC" else None
        ]

        # Handle amplified_cdna_name for RNA
        if modality == "RNA":
            if current_date not in self.counter_data["amp_counter"]:
                self.counter_data["amp_counter"][current_date] = 0
            reaction_count = self.counter_data["amp_counter"][current_date]
            letter = chr(65 + (reaction_count % 8))  # A-H (65 is ASCII for 'A')
            batch_num_for_amp = (reaction_count // 8) + 1  # Increment batch number every 8 reactions
            cdna_amp_date = self.convert_date(self.cdna_amp_date_input.text())
            row_data[21] = f"APLCXR_{cdna_amp_date}_{batch_num_for_amp}_{letter}"
            self.counter_data["amp_counter"][current_date] += 1

        # Write to Excel
        for col_num, value in enumerate(row_data, 1):
            cell = worksheet.cell(row=current_row, column=col_num, value=value)
            # Apply black fill for ATAC empty cells
            if modality == "ATAC" and value is None:
                cell.fill = self.black_fill

        # Apply black fill to tissue_name_old
        tissue_old_col = headers.index('tissue_name_old') + 1
        worksheet.cell(row=current_row, column=tissue_old_col).fill = self.black_fill

    def on_submit(self):
        try:
            if not self.validate_inputs():
                return

            # Ask user where to save the file
            file_location = self.get_save_location()
            if not file_location:
                return  # User canceled the save dialog

            # Process the form data and update Excel
            self.process_form_data(file_location)

            # Adjust column widths
            workbook = load_workbook(file_location)
            worksheet = workbook.active

            for column in worksheet.columns:
                max_length = 0
                column_letter = get_column_letter(column[0].column)
                for cell in column:
                    try:
                        cell_value = str(cell.value)
                        if len(cell_value) > max_length:
                            max_length = len(cell_value)
                    except:
                        pass
                adjusted_width = (max_length + 2)
                worksheet.column_dimensions[column_letter].width = adjusted_width

            workbook.save(file_location)

            QMessageBox.information(
                self,
                "Success",
                f"Data successfully appended to {file_location}"
            )

            # Clear form fields after successful submission
            self.clear_form_fields()

        except Exception as e:
            QMessageBox.critical(
                self,
                "Error",
                f"An error occurred while processing the data:\n{str(e)}"
            )

    def clear_form_fields(self):
        """Clear all form fields after successful submission"""
        # Clear basic info
        self.date_input.clear()
        self.marmoset_input.setCurrentIndex(0)
        self.slab_input.clear()
        self.tile_input.clear()
        self.hemisphere_input.setCurrentIndex(0)

        # Clear sample info
        self.tile_location_input.setCurrentIndex(0)
        self.sort_method_input.setCurrentIndex(0)
        self.rxn_number_input.clear()
        self.facs_population_input.clear()
        self.project_input.setCurrentIndex(0)
        self.project_name_input.clear()

        # Clear cDNA metrics
        self.percent_cdna_400bp_input.clear()
        self.cdna_concentration_input.clear()
        self.rna_lib_concentration_input.clear()
        self.atac_lib_concentration_input.clear()

        # Clear library info
        self.cdna_amp_date_input.clear()
        self.atac_prep_date_input.clear()
        self.rna_prep_date_input.clear()
        self.cdna_pcr_cycles_input.clear()
        self.expected_recovery_input.clear()
        self.nuclei_concentration_input.clear()
        self.nuclei_volume_input.clear()

        # Clear indices tab
        self.rna_indices_input.clear()
        self.atac_indices_input.clear()
        self.rna_sizes_input.clear()
        self.atac_sizes_input.clear()
        self.library_cycles_rna_input.clear()
        self.library_cycles_atac_input.clear()

def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    gui = DataLogGUI()
    gui.show()
    sys.exit(app.exec())

if __name__ == '__main__':
    main()