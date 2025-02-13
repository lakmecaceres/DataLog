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

# Make column headers and give their positions; also call "bold" formatting
worksheet.write('A1', 'krienen_lab_identifier', bold)
worksheet.write('B1', 'seq_portal', bold)
worksheet.write('C1', 'elab_link', bold)
worksheet.write('D1', 'experiment_start_date', bold)
worksheet.write('E1', 'mit_name', bold)
worksheet.write('F1', 'donor_name', bold)
worksheet.write('G1', 'tissue_name', bold)
worksheet.write('H1', 'tissue_name_old', bold)
worksheet.write('I1', 'dissociated_cell_sample_name', bold)
worksheet.write('J1', 'facs_population_plan', bold)
worksheet.write('K1', 'cell_prep_type', bold)
worksheet.write('L1', 'study', bold)
worksheet.write('M1', 'enriched_cell_sample_container_name', bold)
worksheet.write('N1', 'expc_cell_capture', bold)
worksheet.write('O1', 'port_well', bold)
worksheet.write('P1', 'enriched_cell_sample_name', bold)
worksheet.write('Q1', 'enriched_cell_sample_quantity_count', bold)
worksheet.write('R1', 'barcoded_cell_sample_name', bold)
worksheet.write('S1', 'library_method', bold)
worksheet.write('T1', 'cDNA_amplification_method', bold)
worksheet.write('U1', 'cDNA_amplification_date', bold)
worksheet.write('V1', 'amplified_cdna_name', bold)
worksheet.write('W1', 'cDNA_pcr_cycles', bold)
worksheet.write('X1', 'rna_amplification_pass_fail', bold)
worksheet.write('Y1', 'percent_cdna_longer_than_400bp', bold)
worksheet.write('Z1', 'cdna_amplified_quantity_ng', bold)
worksheet.write('AA1', 'library_creation_date', bold)
worksheet.write('AB1', 'library_prep_set', bold)
worksheet.write('AC1', 'library_name', bold)
worksheet.write('AD1', 'cDNA_library_input_ng', bold)
worksheet.write('AE1', 'tapestation_avg_size_bp', bold)
worksheet.write('AF1', 'library_num_cycles', bold)
worksheet.write('AG1', 'lib_quantification_ng', bold)
worksheet.write('AH1', 'library_prep_pass_fail', bold)
worksheet.write('AI1', 'r1_index', bold)
worksheet.write('AJ1', 'r2_index', bold)
worksheet.write('AK1', 'ATAC_index', bold)
worksheet.write('AL1', 'library_pool_name', bold)


# Make duplicate rows for every reaction that was run for RNA and ATAC
row_index = 1
for x in range(rxn_number):
    krienen_lab_identifier_rna = f'{date}_HMBA_cj{mit_name}_Slab{slab}_Tile{tile}_{sort_method}_RNA{x + 1}'
    worksheet.write(row_index, 0, krienen_lab_identifier_rna)
    worksheet.write(row_index, 1, seq_portal)
    worksheet.write(row_index, 2, elab_link)
    worksheet.write(row_index, 3, date)
    worksheet.write(row_index, 4, mit_name)
    worksheet.write(row_index, 5, donor_name)
    worksheet.write(row_index, 6, tissue_name)
    worksheet.write(row_index, 8, dissociated_cell_sample_name)
    worksheet.write(row_index,9, facs_population)
    row_index += 1

    krienen_lab_identifier_atac = f'{date}_HMBA_cj{mit_name}_Slab{slab}_Tile{tile}_{sort_method}_ATAC{x + 1}'
    worksheet.write(row_index, 0, krienen_lab_identifier_atac)
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
    worksheet.write(x + 1, 7, '', black_fill)  # Column H (index 7)

worksheet.autofit()
workbook.close()