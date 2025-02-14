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
    hemisphere = input("Did the tile come from the left hemisphere (LH), right hemisphere (RH), or both? ").strip().lower()
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
    tile_location_input = input("Is the tile from the Brainstem (BS), Cortex (CX), and/or Cerebellum (CB)?").strip().upper()  # Convert to uppercase

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
elab_link = clipboard_content = pyperclip.paste()  # Automatically copies link from clipboard
tissue_name = f"{donor_name}.{tile_location_abbr}.{slab}.{tile}"
dissociated_cell_sample_name = f'{date}_{tissue_name}.Multiome'
cell_prep_type = "nuclei"
sorter_initials = input("Enter the sorter's first and last initials: ").strip().upper()

# Determine facs_population_plan based on sort_method
if sort_method == "pooled":
    while True:
        proportions = input("Enter the proportions of NeuN+/Dneg/Olig2+ (e.g., 70/20/10 or 100/0/0): ").strip()

        # Normalize input
        if "/" in proportions:
            proportions = proportions.split("/")
            if len(proportions) == 3:
                try:
                    # Convert proportions to integers
                    proportions = [int(p) for p in proportions]
                    # Check if the proportions add up to 100
                    if sum(proportions) == 100:
                        facs_population = "/".join(map(str, proportions))  # Join proportions with slashes
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
elif sort_method == "DAPI":  # Ensure "DAPI" is uppercase
    facs_population = "DAPI"  # Set to "DAPI" instead of "DAPI_sorted"
else:
    print("Invalid sort method. Please enter pooled, unsorted, or DAPI.")
    exit()

# Ask the user if the sample is for the HMBA Subcortex project
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

# Track the last used PXXXX number for each experiment date
barcode_counter = {}
# Set the starting PXXXX number to 83 (P0083)
starting_p_number = 83
# Get the current PXXXX number for the experiment date
if date in barcode_counter:
    p_number = barcode_counter[date]
else:
    # Start with P0083 for a new date
    p_number = starting_p_number
    barcode_counter[date] = p_number

# Prompt for the cDNA amplification date
while True:
    cdna_amplification_date_input = input('Input the cDNA amplification date: ')
    cdna_amplification_date = convert(cdna_amplification_date_input)
    if cdna_amplification_date:
        break

# Track the reactions and batch number for amplified_cdna_name
date_reaction_counter = {}  # Dictionary to track number of reactions per date

rna_amplification_pass_fail = "Pass"

# Prompt user for comma-separated values for cDNA amplification data
while True:
    cdna_pcr_cycles_list = input("Enter the PCR cycles for each reaction (comma-separated): ").split(',')
    if len(cdna_pcr_cycles_list) == rxn_number:
        break
    print(f"Please enter {rxn_number} values.")

while True:
    percent_cdna_long_400bp_list = input("Enter the percent of cDNA >400bp for each reaction (comma-separated): ").split(',')
    try:
        percent_cdna_long_400bp_list = [round(float(x.strip())) for x in percent_cdna_long_400bp_list]
        if len(percent_cdna_long_400bp_list) == rxn_number:
            break
    except ValueError:
        pass
    print(f"Please enter {rxn_number} numeric values.")

while True:
    cdna_concentration_list = input("Enter the concentration of amplified cDNA (ng/uL) for each reaction (comma-separated): ").split(',')
    try:
        cdna_concentration_list = [float(x.strip()) for x in cdna_concentration_list]
        if len(cdna_concentration_list) == rxn_number:
            break
    except ValueError:
        pass
    print(f"Please enter {rxn_number} numeric values.")

# Calculate total cDNA quantity in ng (Multiply each value by 40µL)
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

# Function to convert index to letter-number format (e.g., 1A -> A1, A1 -> A1)
def convert_index(index):
    index = index.strip().upper()
    if len(index) == 2:
        if index[0].isdigit() and index[1].isalpha():  # Format: 1A
            return f"{index[1]}{index[0]}"
        elif index[0].isalpha() and index[1].isdigit():  # Format: A1
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

# Define column headers
headers = [
    'krienen_lab_identifier', 'seq_portal', 'elab_link', 'experiment_start_date', 'mit_name',
    'donor_name', 'tissue_name', 'tissue_name_old', 'dissociated_cell_sample_name', 'facs_population_plan',
    'cell_prep_type', 'study', 'enriched_cell_sample_container_name', 'expc_cell_capture', 'port_well',
    'enriched_cell_sample_name', 'enriched_cell_sample_quantity_count', 'barcoded_cell_sample_name', 'library_method',
    'cDNA_amplification_method', 'cDNA_amplification_date', 'amplified_cdna_name', 'cDNA_pcr_cycles',
    'rna_amplification_pass_fail', 'percent_cdna_longer_than_400bp', 'cdna_amplified_quantity_ng',
    'library_creation_date', 'library_prep_set', 'library_name', 'cDNA_library_input_ng', 'tapestation_avg_size_bp',
    'library_num_cycles', 'lib_quantification_ng', 'library_prep_pass_fail', 'r1_index', 'r2_index',
    'ATAC_index', 'library_pool_name'
]

# Write headers to the first row
for col_index, header in enumerate(headers):
    worksheet.write(0, col_index, header, bold)

# Track library indices to detect duplicates
library_index_tracker = {"LPLCXA": {}, "LPLCXR": {}}  # Track indices for each library type

# Make duplicate rows for every reaction that was run for RNA and ATAC
row_index = 1
for x in range(rxn_number):
    port_well = x + 1  # Assign well numbers starting from 1
    for modality in ["RNA", "ATAC"]:
        krienen_lab_identifier = f'{date}_HMBA_cj{mit_name}_Slab{int(slab)}_Tile{int(tile)}_{sort_method}_{modality}{x + 1}'
        enriched_cell_sample_name = f'MPXM_{date}_{sorting_status}_{sorter_initials}_{port_well}'

        library_prep_date = rna_library_prep_date if modality == "RNA" else atac_library_prep_date
        library_prep_set = f"LPLCXR_{library_prep_date}_1" if modality == "RNA" else f"LPLCXA_{library_prep_date}_1"

        # Generate barcoded_cell_sample_name
        barcoded_cell_sample_name = f'P{str(p_number).zfill(4)}_{port_well}'

        # Determine the library method based on modality
        if modality == "RNA":
            library_method = "10xMultiome-RSeq"
            cDNA_amplification_method = library_method
            library_type = "LPLCXR"
            library_index = rna_indices[x]  # Use RNA index
        else:  # ATAC
            library_method = "10xMultiome-ASeq"
            library_type = "LPLCXA"
            library_index = atac_indices[x]  # Use ATAC index

        # Ensure batch tracking for the entered date
        if library_prep_date not in barcode_counter:
            barcode_counter[library_prep_date] = 1  # Start at batch 1
        else:
            barcode_counter[library_prep_date] += 1  # Increment after each use

        # Calculate batch number correctly
        batch_number = ((barcode_counter[library_prep_date] - 1) // 8) + 1

        # Column AB: Library Prep Set (Prefix + Date + Batch Number)
        library_prep_prefix = "LPLCXR_" if modality == "RNA" else "LPLCXA_"
        library_prep_set = f"{library_prep_prefix}{library_prep_date}_{batch_number}"

        # Check for duplicate indices within the same library type
        if library_index in library_index_tracker[library_type]:
            library_index_tracker[library_type][library_index].append(row_index)
        else:
            library_index_tracker[library_type][library_index] = [row_index]

        # Generate library name
        library_name = f"{library_type}_{library_prep_date}_{batch_number}_{library_index}"

        # Write data to the worksheet
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
        worksheet.write(row_index, 19, cDNA_amplification_method)
        worksheet.write(row_index, 20, cdna_amplification_date)
        worksheet.write(row_index, 22, cdna_pcr_cycles_list[x] if modality == "RNA" else '', black_fill if modality == "ATAC" else None)
        worksheet.write(row_index, 24, percent_cdna_long_400bp_list[x] if modality == "RNA" else '', black_fill if modality == "ATAC" else None)
        worksheet.write(row_index, 25, cdna_amplified_quantity_ng_list[x] if modality == "RNA" else '', black_fill if modality == "ATAC" else None)
        worksheet.write(row_index, 26, library_prep_date)
        worksheet.write(row_index, 27, library_prep_set)
        worksheet.write(row_index, 28, library_name)  # Write library name

        # Black out cDNA amplification date for ATAC rows
        if modality == "ATAC":
            worksheet.write(row_index, 20, '', black_fill)

        # Insert cDNA library input ng (25% of cdna_amplified_quantity_ng for RNA rows, black out for ATAC rows)
        if modality == "RNA":
            cdna_library_input_ng = cdna_amplified_quantity_ng_list[x] * 0.25
            worksheet.write(row_index, 26, cdna_library_input_ng)  # Column AA (index 26)
        else:
            worksheet.write(row_index, 26, '', black_fill)  # Black out for ATAC rows

        # Generate amplified_cdna_name for RNA rows
        if modality == "RNA":
            # Track the reaction count for each date
            if date not in date_reaction_counter:
                date_reaction_counter[date] = 0

            # Determine the letter for the reaction (A-H, then restart at A)
            reaction_count = date_reaction_counter[date]
            letter = chr(65 + (reaction_count % 8))  # A-H (chr(65) = 'A')

            # Determine the batch number (increments after every 8 reactions)
            batch_number = (reaction_count // 8) + 1

            # Format the amplified cDNA name
            amplified_cdna_name = f"APLCXR_{cdna_amplification_date}_{batch_number}_{letter}"

            # Write the amplified cDNA name to the worksheet
            worksheet.write(row_index, 21, amplified_cdna_name)

            # Increment reaction count for this date
            date_reaction_counter[date] += 1
        else:  # ATAC Rows get black fill
            worksheet.write(row_index, 21, '', black_fill)

        if modality == "RNA":
            worksheet.write(row_index, 23, rna_amplification_pass_fail)
        else:  # ATAC Rows get black fill
            worksheet.write(row_index, 23, '', black_fill)

        row_index += 1

# Apply black fill to the tissue_name_old column for all rows
for x in range(rxn_number * 2):  # Multiply by 2 for RNA and ATAC rows
    worksheet.write(x + 1, 7, '', black_fill)

worksheet.autofit()
workbook.close()