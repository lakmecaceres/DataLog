import xlsxwriter
import pyperclip
import dateutil.parser

# Make the file and a worksheet
workbook = xlsxwriter.Workbook('datalog.xlsx')
worksheet = workbook.add_worksheet('hmba')
bold = workbook.add_format({'bold': True}) # Define bold formatting
black_fill = workbook.add_format({'bg_color': 'black'}) # Define black fill formatting

# Date formatter
def convert(exp_date):
    try:
        parsed_date = dateutil.parser.parse(exp_date) # Detect date format
        return parsed_date.strftime('%y%m%d') # Convert to YYMMDD
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
    tile_location_input = input("Is the tile from the Brainstem (BS), Cortex (CX), and/or Cerebellum (CB)?").strip().upper() # Convert to uppercase

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


sort_method = input("Input the sort method (pooled/unsorted/DAPI?): ")
rxn_number = int(input("Input the number of reactions you ran: "))
seq_portal = input('What is the seq portal status for this reaction (yes/no/done/reseq?): ')
elab_link = clipboard_content = pyperclip.paste() # Automatically copies link from clipboard
tissue_name = f"{donor_name}.{tile_location_abbr}.{slab}.{tile}"
dissociated_cell_sample_name = f'{date}_{tissue_name}.Multiome'

# Determine facs_population_plan based on sort_method
if sort_method.lower() == "pooled":
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
elif sort_method.lower() == "unsorted":
    facs_population = "no_FACS"
elif sort_method.lower() == "dapi":
    facs_population = "DAPI"  # Set to "DAPI" instead of "DAPI_sorted"
else:
    print("Invalid sort method. Please enter pooled, unsorted, or DAPI.")
    exit()

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

# Make duplicate rows for every reaction that was run for RNA and ATAC
row_index = 1
for x in range(rxn_number):
    for modality in ["RNA", "ATAC"]:
        krienen_lab_identifier = f'{date}_HMBA_cj{mit_name}_Slab{slab}_Tile{tile}_{sort_method}_{modality}{x + 1}'

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

        row_index += 1

# Apply black fill to the tissue_name_old column for all rows
for x in range(rxn_number * 2):  # Multiply by 2 for RNA and ATAC rows
    worksheet.write(x + 1, 7, '', black_fill)

worksheet.autofit()
workbook.close()