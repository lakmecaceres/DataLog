import xlsxwriter
import pyperclip
import dateutil.parser

# Make the file and a worksheet
workbook = xlsxwriter.Workbook('datalog.xlsx')
worksheet = workbook.add_worksheet('hmba')
bold = workbook.add_format({'bold': True})  # Define bold formatting
black_fill = workbook.add_format({'bg_color': 'black'})  # Define black fill formatting
yellow_fill = workbook.add_format({'bg_color': '#FFFF00'})  # Define yellow fill formatting for duplicates


# Date formatter
def convert(exp_date):
    try:
        parsed_date = dateutil.parser.parse(exp_date)  # Detect date format
        return parsed_date.strftime('%y%m%d')  # Convert to YYMMDD
    except ValueError:
        print('Invalid date format. Please try again.')
        return None


# Dictionary to map names to codes
name_to_code = {
    "Croissant": "CJ23.56.002",
    "Nutmeg": "CJ23.56.003",
    "Jellybean": "CJ24.56.004",
    "Rambo": "CJ24.56.005"
}

# Mapping of full names to abbreviations (case-insensitive)
tile_location_map = {
    "BRAINSTEM": "BS",
    "BS": "BS",
    "CORTEX": "CX",
    "CX": "CX",
    "CEREBELLUM": "CB",
    "CB": "CB"
}

# Date prompt
while True:
    date_input = input('Input the date of the experiment: ')
    date = convert(date_input)
    if date:
        break

# Marmoset name input with validation
while True:
    mit_name = input("Input the name of the marmoset: ").strip().title()  # Convert input to title case
    if mit_name in name_to_code:
        donor_name = name_to_code[mit_name]  # Get the corresponding donor code
        break
    else:
        print("Invalid name. Please enter one of: Croissant, Nutmeg, Jellybean, Rambo.")

# Ensure slab and tile are zero-padded to two digits
slab = input("Input the slab number: ")
tile = input("Input the tile number: ").zfill(2)  # Zero-pad to 2 digits

# Prompt the user for hemisphere information
while True:
    hemisphere = input(
        "Did the tile come from the left hemisphere (LH), right hemisphere (RH), or both? ").strip().lower()
    if hemisphere in ["left", "lh", "right", "rh", "both"]:
        break
    else:
        print("Invalid input. Please enter left/LH, right/RH, or both.")

# Normalize hemisphere input
if hemisphere in ["left", "lh"]:
    hemisphere = "LH"
elif hemisphere in ["right", "rh"]:
    hemisphere = "RH"
else:
    hemisphere = "BOTH"

# Adjust the slab number based on hemisphere
if hemisphere == "RH":
    slab = str(int(slab) + 40).zfill(2)  # Add 40 and zero-pad to 2 digits
elif hemisphere == "BOTH":
    slab = str(int(slab) + 90).zfill(2)  # Add 90 and zero-pad to 2 digits
else:
    slab = slab.zfill(2)  # Zero-pad to 2 digits (no adjustment for LH)

# Add a new prompt for the tile location
while True:
    tile_location_input = input(
        "Is the tile from the Brainstem (BS), Cortex (CX), and/or Cerebellum (CB)?").strip().upper()  # Convert to uppercase

    # Split the input into a list of locations
    tile_locations = []
    for part in tile_location_input.replace(" and ", ",").split(","):
        part = part.strip()
        if part in tile_location_map:
            tile_locations.append(tile_location_map[part])
        elif part in ["BS", "CX", "CB"]:  # Allow direct abbreviations
            tile_locations.append(part)

    if tile_locations:
        tile_location_abbr = "-".join(tile_locations)  # Join locations with dashes
        break
    else:
        print("Invalid input. Please enter Brainstem/BS, Cortex/CX, or Cerebellum/CB, separated by commas or 'and'.")

# Prompt for sort method only once
while True:
    sort_method = input("Input the sort method (pooled/unsorted/DAPI?): ").strip()
    if sort_method.lower() in ["pooled", "unsorted", "dapi"]:
        break
    print("Invalid sort method. Please enter pooled, unsorted, or DAPI.")

# Convert "dapi" to "DAPI" if the user enters it in lowercase
if sort_method.lower() == "dapi":
    sort_method = "DAPI"  # Force uppercase

rxn_number = int(input("Input the number of reactions you ran: "))
seq_portal = "no"  # Automatically set seq_portal status to "no"
elab_link = pyperclip.paste()  # Automatically copies link from clipboard
tissue_name = f"{donor_name}.{tile_location_abbr}.{slab}.{tile}"
dissociated_cell_sample_name = f'{date}_{tissue_name}.Multiome'
cell_prep_type = "nuclei"
sorter_initials = input("Enter the sorter's first and last initials: ").strip().upper()

# Determine facs_population_plan based on sort_method
if sort_method == "pooled":
    while True:
        proportions = input("Enter the proportions of NeuN+/Dneg/Olig2+ (e.g., 70/20/10 or 100/0/0): ").strip()
        if "/" in proportions:
            proportions = proportions.split("/")
            if len(proportions) == 3:
                try:
                    proportions = [int(p) for p in proportions]
                    if sum(proportions) == 100:
                        facs_population = "/".join(map(str, proportions))
                        break
                    else:
                        print("Invalid input. The proportions must add up to 100.")
                except ValueError:
                    print("Invalid input. Please enter numbers for the proportions.")
            else:
                print("Invalid input. Please enter three proportions separated by slashes (e.g., 70/20/10).")
        else:
            print("Invalid input. Please use slashes to separate the proportions (e.g., 70/20/10).")
elif sort_method == "unsorted":
    facs_population = "no_FACS"
elif sort_method == "DAPI":
    facs_population = "DAPI"
else:
    print("Invalid sort method. Please enter pooled, unsorted, or DAPI.")
    exit()

# Ask if the sample is for the HMBA Subcortex project
is_hmba_subcortex = input("Is the sample for the HMBA Subcortex project? (yes/no): ").strip().lower()
if is_hmba_subcortex in ["yes", "y"]:
    study = "HMBA_CjAtlas_Subcortex"
else:
    study = input("Enter the project name: ").strip()

if sort_method in ["pooled", "DAPI"]:
    sorting_status = "PS"
elif sort_method == "unsorted":
    sorting_status = "PN"
else:
    print("Invalid sort method. Please enter pooled, unsorted, or DAPI.")
    exit()

enriched_cell_sample_container_name = f"MPXM_{date}_{sorting_status}_{sorter_initials}"

expected_cell_capture = int(input("What is the expected recovery?: "))

# Ask user for concentration and volume to calculate enriched_cell_sample_quantity_count
concentration = float(input("Enter the concentration of nuclei/cells (cells/µL): "))
volume = float(input("Enter the volume used (µL): "))
enriched_cell_sample_quantity_count = concentration * volume

# Track the last used PXXXX number for each experiment date (for library_name batch; no longer used for library_name)
barcode_counter = {}
starting_p_number = 83
if date in barcode_counter:
    p_number = barcode_counter[date]
else:
    p_number = starting_p_number
    barcode_counter[date] = p_number

# Prompt for the cDNA amplification date
while True:
    cdna_amplification_date_input = input('Input the cDNA amplification date: ')
    cdna_amplification_date = convert(cdna_amplification_date_input)
    if cdna_amplification_date:
        break

# Track reactions for amplified_cdna_name
date_reaction_counter = {}
rna_amplification_pass_fail = "Pass"

# Prompt user for comma-separated values for cDNA amplification data
while True:
    cdna_pcr_cycles_list = input("Enter the PCR cycles for each reaction (comma-separated): ").split(',')
    if len(cdna_pcr_cycles_list) == rxn_number:
        break
    print(f"Please enter {rxn_number} values.")

while True:
    cdna_input = input("Enter the percent of cDNA >400bp for each reaction (comma-separated): ")
    percent_cdna_long_400bp_list = cdna_input.split(',')
    try:
        percent_cdna_long_400bp_list = [round(float(x.strip())) for x in percent_cdna_long_400bp_list]
        if len(percent_cdna_long_400bp_list) == rxn_number:
            break
    except ValueError:
        pass
    print(f"Please enter {rxn_number} numeric values.")

while True:
    cdna_concentration_list = input(
        "Enter the concentration of amplified cDNA (ng/uL) for each reaction (comma-separated): ").split(',')
    try:
        cdna_concentration_list = [float(x.strip()) for x in cdna_concentration_list]
        if len(cdna_concentration_list) == rxn_number:
            break
    except ValueError:
        pass
    print(f"Please enter {rxn_number} numeric values.")

# Calculate total cDNA quantity in ng (each value multiplied by 40µL)
cdna_amplified_quantity_ng_list = [conc * 40 for conc in cdna_concentration_list]

# Ask user for ATAC and RNA library prep dates
while True:
    atac_library_prep_date_input = input("Enter the ATAC library preparation date: ")
    atac_library_prep_date = convert(atac_library_prep_date_input)
    if atac_library_prep_date:
        break

while True:
    rna_library_prep_date_input = input("Enter the cDNA (RNA) library preparation date: ")
    rna_library_prep_date = convert(rna_library_prep_date_input)
    if rna_library_prep_date:
        break


# Function to convert index to letter-number format (e.g., 1A -> A1)
def convert_index(index):
    index = index.strip().upper()
    if len(index) == 2:
        if index[0].isdigit() and index[1].isalpha():
            return f"{index[1]}{index[0]}"
        elif index[0].isalpha() and index[1].isdigit():
            return index
    return None


# Prompt user for comma-separated indices for ATAC and RNA
while True:
    atac_indices_input = input("Enter the ATAC indices (comma-separated, e.g., A1,2B,C3): ").strip().upper()
    atac_indices = [convert_index(index) for index in atac_indices_input.split(",")]
    if all(atac_indices) and len(atac_indices) == rxn_number:
        break
    print(f"Please enter {rxn_number} valid ATAC indices (e.g., A1, 2B, C3).")

while True:
    rna_indices_input = input("Enter the RNA indices (comma-separated, e.g., D4,5E,F6): ").strip().upper()
    rna_indices = [convert_index(index) for index in rna_indices_input.split(",")]
    if all(rna_indices) and len(rna_indices) == rxn_number:
        break
    print(f"Please enter {rxn_number} valid RNA indices (e.g., D4, 5E, F6).")


# Pad indices to 3 characters (e.g., A1 -> A01)
def pad_index(index):
    if len(index) == 2 and index[0].isalpha() and index[1].isdigit():
        return f"{index[0]}0{index[1]}"
    return index


atac_indices = [pad_index(index) for index in atac_indices]
rna_indices = [pad_index(index) for index in rna_indices]

# Prompt for Tapestation average size (RNA)
while True:
    rna_sizes_input = input(
        f"Enter the Tapestation average size (bp) for RNA libraries (comma-separated, {rxn_number} values): ").strip()
    rna_sizes = rna_sizes_input.split(',')
    try:
        rna_sizes = [int(size.strip()) for size in rna_sizes]
        if len(rna_sizes) == rxn_number:
            break
    except ValueError:
        pass
    print(f"Please enter {rxn_number} integer values separated by commas.")

# Prompt for Tapestation average size (ATAC)
while True:
    atac_sizes_input = input(
        f"Enter the Tapestation average size (bp) for ATAC libraries (comma-separated, {rxn_number} values): ").strip()
    atac_sizes = atac_sizes_input.split(',')
    try:
        atac_sizes = [int(size.strip()) for size in atac_sizes]
        if len(atac_sizes) == rxn_number:
            break
    except ValueError:
        pass
    print(f"Please enter {rxn_number} integer values separated by commas.")

# --- New Code: Prompt for library_num_cycles ---
# For RNA libraries:
while True:
    library_num_cycles_rna_input = input(
        f"Enter the library_num_cycles for RNA libraries (comma-separated, {rxn_number} values): ").strip()
    try:
        library_num_cycles_rna = [int(x.strip()) for x in library_num_cycles_rna_input.split(',')]
        if len(library_num_cycles_rna) == rxn_number:
            break
    except ValueError:
        pass
    print(f"Please enter {rxn_number} integer values separated by commas.")

# For ATAC libraries:
while True:
    library_num_cycles_atac_input = input(
        f"Enter the library_num_cycles for ATAC libraries (comma-separated, {rxn_number} values): ").strip()
    try:
        library_num_cycles_atac = [int(x.strip()) for x in library_num_cycles_atac_input.split(',')]
        if len(library_num_cycles_atac) == rxn_number:
            break
    except ValueError:
        pass
    print(f"Please enter {rxn_number} integer values separated by commas.")

# --- New Code: Prompt for lib_quantification_ng (library concentrations in ng/uL) ---
# For RNA libraries:
while True:
    lib_quant_rna_input = input(
        f"Enter the RNA library concentrations (ng/uL) (comma-separated, {rxn_number} values): ").strip()
    try:
        lib_quant_rna = [float(x.strip()) for x in lib_quant_rna_input.split(',')]
        if len(lib_quant_rna) == rxn_number:
            break
    except ValueError:
        pass
    print(f"Please enter {rxn_number} numeric values separated by commas.")

# For ATAC libraries:
while True:
    lib_quant_atac_input = input(
        f"Enter the ATAC library concentrations (ng/uL) (comma-separated, {rxn_number} values): ").strip()
    try:
        lib_quant_atac = [float(x.strip()) for x in lib_quant_atac_input.split(',')]
        if len(lib_quant_atac) == rxn_number:
            break
    except ValueError:
        pass
    print(f"Please enter {rxn_number} numeric values separated by commas.")

# Define column headers with new ordering.
headers = [
    'krienen_lab_identifier',  # 0
    'seq_portal',  # 1
    'elab_link',  # 2
    'experiment_start_date',  # 3
    'mit_name',  # 4
    'donor_name',  # 5
    'tissue_name',  # 6
    'tissue_name_old',  # 7
    'dissociated_cell_sample_name',  # 8
    'facs_population_plan',  # 9
    'cell_prep_type',  # 10
    'study',  # 11
    'enriched_cell_sample_container_name',  # 12
    'expc_cell_capture',  # 13
    'port_well',  # 14
    'enriched_cell_sample_name',  # 15
    'enriched_cell_sample_quantity_count',  # 16
    'barcoded_cell_sample_name',  # 17
    'library_method',  # 18
    'cDNA_amplification_method',  # 19
    'cDNA_amplification_date',  # 20
    'amplified_cdna_name',  # 21
    'cDNA_pcr_cycles',  # 22
    'rna_amplification_pass_fail',  # 23
    'percent_cdna_longer_than_400bp',  # 24
    'cdna_amplified_quantity_ng',  # 25
    'cDNA_library_input_ng',  # 26
    'library_creation_date',  # 27
    'library_prep_set',  # 28
    'library_name',  # 29
    'tapestation_avg_size_bp',  # 30
    'library_num_cycles',  # 31
    'lib_quantification_ng',  # 32
    'library_prep_pass_fail',  # 33
    'r1_index',  # 34
    'r2_index',  # 35
    'ATAC_index',  # 36
    'library_pool_name'  # 37
]

# Write headers to the first row
for col_index, header in enumerate(headers):
    worksheet.write(0, col_index, header, bold)

# Track library indices to detect duplicates (for overall tracking, if needed)
library_index_tracker = {"LPLCXA": {}, "LPLCXR": {}}
# Dictionary to track duplicate indices per (library_type, library_prep_date, library_index)
dup_index_counter = {}

# Make duplicate rows for every reaction that was run for RNA and ATAC
row_index = 1
for x in range(rxn_number):
    port_well = x + 1  # Well number starting from 1
    for modality in ["RNA", "ATAC"]:
        krienen_lab_identifier = f'{date}_HMBA_cj{mit_name}_Slab{int(slab)}_Tile{int(tile)}_{sort_method}_{modality}{x + 1}'
        enriched_cell_sample_name = f'MPXM_{date}_{sorting_status}_{sorter_initials}_{port_well}'
        library_prep_date = rna_library_prep_date if modality == "RNA" else atac_library_prep_date
        barcoded_cell_sample_name = f'P{str(p_number).zfill(4)}_{port_well}'

        # Determine library method and set cDNA_amplification_method accordingly
        if modality == "RNA":
            library_method = "10xMultiome-RSeq"
            cDNA_amplification_method = library_method
            library_type = "LPLCXR"
            library_index = rna_indices[x]  # Use RNA index
        else:
            library_method = "10xMultiome-ASeq"
            library_type = "LPLCXA"
            library_index = atac_indices[x]  # Use ATAC index
            cDNA_amplification_method = None

        # Update barcode_counter for library_name batch (if needed elsewhere)
        if library_prep_date not in barcode_counter:
            barcode_counter[library_prep_date] = 1
        else:
            barcode_counter[library_prep_date] += 1

        # Compute library_prep_set based on duplicate index count
        library_prep_prefix = "LPLCXR_" if modality == "RNA" else "LPLCXA_"
        dup_key = (library_type, library_prep_date, library_index)
        if dup_index_counter.get(dup_key):
            dup_index_counter[dup_key] += 1
        else:
            dup_index_counter[dup_key] = 1
        library_prep_set = f"{library_prep_prefix}{library_prep_date}_{dup_index_counter[dup_key]}"

        # Generate library_name as library_prep_set appended with the index
        library_name = f"{library_prep_set}_{library_index}"

        # Write data to worksheet columns
        worksheet.write(row_index, 0, krienen_lab_identifier)
        worksheet.write(row_index, 1, seq_portal)
        worksheet.write(row_index, 2, elab_link)
        worksheet.write(row_index, 3, date)
        worksheet.write(row_index, 4, mit_name)
        worksheet.write(row_index, 5, donor_name)
        worksheet.write(row_index, 6, tissue_name)
        worksheet.write(row_index, 8, dissociated_cell_sample_name)
        worksheet.write(row_index, 9, facs_population)
        worksheet.write(row_index, 10, cell_prep_type)
        worksheet.write(row_index, 11, study)
        worksheet.write(row_index, 12, enriched_cell_sample_container_name)
        worksheet.write(row_index, 13, expected_cell_capture)
        worksheet.write(row_index, 14, port_well)
        worksheet.write(row_index, 15, enriched_cell_sample_name)
        worksheet.write(row_index, 16, enriched_cell_sample_quantity_count)
        worksheet.write(row_index, 17, barcoded_cell_sample_name)
        worksheet.write(row_index, 18, library_method)
        if modality == "RNA":
            worksheet.write(row_index, 19, cDNA_amplification_method)
        else:
            worksheet.write(row_index, 19, '', black_fill)
        worksheet.write(row_index, 20, cdna_amplification_date)
        if modality == "ATAC":
            worksheet.write(row_index, 20, '', black_fill)
        worksheet.write(row_index, 22, cdna_pcr_cycles_list[x] if modality == "RNA" else '',
                        black_fill if modality == "ATAC" else None)
        worksheet.write(row_index, 24, percent_cdna_long_400bp_list[x] if modality == "RNA" else '',
                        black_fill if modality == "ATAC" else None)
        worksheet.write(row_index, 25, cdna_amplified_quantity_ng_list[x] if modality == "RNA" else '',
                        black_fill if modality == "ATAC" else None)
        # Write cDNA_library_input_ng (25% of cdna amplified quantity) into column 26
        if modality == "RNA":
            cdna_library_input_ng = cdna_amplified_quantity_ng_list[x] * 0.25
            worksheet.write(row_index, 26, cdna_library_input_ng)
        else:
            worksheet.write(row_index, 26, '', black_fill)
        # Write shifted columns according to header order:
        worksheet.write(row_index, 27, library_prep_date)  # library_creation_date
        worksheet.write(row_index, 28, library_prep_set)  # library_prep_set
        worksheet.write(row_index, 29, library_name)  # library_name
        # Write Tapestation average size (column 30)
        if modality == "RNA":
            worksheet.write(row_index, 30, rna_sizes[x])
        else:
            worksheet.write(row_index, 30, atac_sizes[x])
        # Write library_num_cycles (column 31)
        if modality == "RNA":
            worksheet.write(row_index, 31, library_num_cycles_rna[x])
        else:
            worksheet.write(row_index, 31, library_num_cycles_atac[x])
        # Write lib_quantification_ng (column 32)
        # Multiply RNA concentrations by 35 uL and ATAC concentrations by 20 uL
        if modality == "RNA":
            worksheet.write(row_index, 32, lib_quant_rna[x] * 35)
        else:
            worksheet.write(row_index, 32, lib_quant_atac[x] * 20)
        # Automatically fill library_prep_pass_fail (column 33) with "Pass"
        worksheet.write(row_index, 33, "Pass")

        # --- New Code: Write r1_index and r2_index ---
        # These indices are only for RNA (cDNA) rows.
        if modality == "RNA":
            r1_val = f"SI-TT-{rna_indices[x]}_i7"
            r2_val = f"SI-TT-{rna_indices[x]}_b(i5)"
            worksheet.write(row_index, 34, r1_val)
            worksheet.write(row_index, 35, r2_val)
            worksheet.write(row_index, 36, '', black_fill)  # ATAC_index left blank for RNA rows
        else:
            worksheet.write(row_index, 34, '', black_fill)
            worksheet.write(row_index, 35, '', black_fill)
            # For ATAC rows, fill ATAC_index with the corresponding index
            worksheet.write(row_index, 36, atac_indices[x])

        # Generate amplified_cdna_name for RNA rows
        if modality == "RNA":
            if date not in date_reaction_counter:
                date_reaction_counter[date] = 0
            reaction_count = date_reaction_counter[date]
            letter = chr(65 + (reaction_count % 8))  # A-H
            batch_num_for_amp = (reaction_count // 8) + 1
            amplified_cdna_name = f"APLCXR_{cdna_amplification_date}_{batch_num_for_amp}_{letter}"
            worksheet.write(row_index, 21, amplified_cdna_name)
            date_reaction_counter[date] += 1
        else:
            worksheet.write(row_index, 21, '', black_fill)

        if modality == "RNA":
            worksheet.write(row_index, 23, rna_amplification_pass_fail)
        else:
            worksheet.write(row_index, 23, '', black_fill)

        row_index += 1

# Apply black fill to the tissue_name_old column for all rows
for x in range(rxn_number * 2):  # Multiply by 2 for RNA and ATAC rows
    worksheet.write(x + 1, 7, '', black_fill)

worksheet.autofit()
workbook.close()