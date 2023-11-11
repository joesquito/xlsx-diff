#!/usr/bin/env python3
import sys
import zipfile
import os
import tempfile
import re
import subprocess

from difflib import SequenceMatcher
from lxml import etree as ET

XMLNS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'

def build_shared_strings_map(shared_strings_xml_content):
    # Encode the content to bytes
    shared_strings_xml_bytes = shared_strings_xml_content.encode('utf-8')
    tree = ET.fromstring(shared_strings_xml_bytes)
    shared_strings = {}

    for index, si in enumerate(tree.findall(XMLNS + 'si')):
        t = si.find(XMLNS + 't')
        if t is not None:
            shared_strings[index] = t.text

    return shared_strings

def replace_shared_strings_in_sheet(sheet_xml_content, shared_strings):
    # Encode the content to bytes
    sheet_xml_bytes = sheet_xml_content.encode('utf-8')
    tree = ET.fromstring(sheet_xml_bytes)

    for c in tree.findall('.//' + XMLNS + 'c[@t="s"]'):  # cells with shared strings
        v = c.find(XMLNS + 'v')  # the <v> tag inside <c> holds the index
        if v is not None and v.text is not None:
            index = int(v.text)
            if index in shared_strings:
                v.text = shared_strings[index]

    return ET.tostring(tree, encoding='utf-8').decode('utf-8')

def replace_shared_strings(directory):
    shared_strings_path = os.path.join(directory, 'xl', 'sharedStrings.xml')
    with open(shared_strings_path, 'r', encoding='utf-8') as f:
        shared_strings_content = f.read()
        
    shared_strings = build_shared_strings_map(shared_strings_content)

    modified_sheets = {}
    for root, _, files in os.walk(directory):
        for file in files:
            if re.match(r'^sheet\d+\.xml$', file):
                sheet_path = os.path.join(root, file)
                with open(sheet_path, 'r', encoding='utf-8') as f:
                    sheet_content = f.read()
                modified_content = replace_shared_strings_in_sheet(sheet_content, shared_strings)
                modified_sheets[sheet_path] = modified_content

    return modified_sheets

def unpack_xlsx(xlsx_path, output_dir):
    with zipfile.ZipFile(xlsx_path, 'r') as zip_ref:
        zip_ref.extractall(output_dir)

def get_sheet_names(directory):
    """Extract the sheet names from the workbook.xml file."""
    workbook_path = os.path.join(directory, 'xl', 'workbook.xml')
    tree = ET.parse(workbook_path)
    root = tree.getroot()

    # Extract the sheet names and map them to the sheet IDs
    sheet_names = {}
    for sheet in root.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheet'):
        sheet_id = sheet.get('sheetId')
        sheet_name = sheet.get('name')
        sheet_names[sheet_id] = sheet_name

    return sheet_names

def get_next_excel_column_name(column_name):
    result = ""
    carry = 1
    for char in reversed(column_name):
        new_char_val = ord(char) - ord('A') + carry
        carry = new_char_val // 26
        result = chr(ord('A') + new_char_val % 26) + result
    if carry:
        result = 'A' + result
    return result

def get_letters(s):
    return ''.join([c for c in s if c.isalpha()])

def summarize_sheet_files(modified_sheets, sheet_names):
    summaries = {} # Dictionary to store summaries
    pattern = re.compile(r'([A-Z]+)\d+')

    for sheet_path, sheet_content in modified_sheets.items():
      # Extract the sheet ID from the file name and get the corresponding sheet name
        sheet_id = re.search(r'sheet(\d+)\.xml', sheet_path).group(1)
        sheet_name = sheet_names.get(sheet_id, '')
        xml_root = ET.fromstring(sheet_content.encode('utf-8'))

        elements_text = [] # String to store summary for this file
        location = []
        row_number = 0

        for row in xml_root.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}row'):
            row_number = int(row.get('r')) if row.get('r') is not None else row_number + 1
            last_cell_location = None
            last_column_letter = None
            # Iterate through cells
            for cell in row.findall('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c'):
                cell_location = cell.get('r')  # Get the cell's location
                if cell_location is not None:
                    last_cell_location = cell_location
                elif last_cell_location is not None:
                    # Increment the column letter and keep the row number
                    # last_column_letter = get_letters(cell_location)
                    last_column_letter = pattern.match(last_cell_location).group(1)
                    last_column_letter = get_next_excel_column_name(last_column_letter)
                    cell_location = last_column_letter + str(row_number)
                else:
                    last_column_letter = 'A'
                    cell_location = last_column_letter + str(row_number)
                    last_cell_location =  cell_location
                value = cell[0] if len(cell) > 0 else None
                if value is not None and value.text is not None:
                    text = value.text.replace('\n', '\\n').replace('\r', '\\r') + "\n"
                    elements_text.append(text)
                    location.append(cell_location)

        summaries[ sheet_name ] = [elements_text, location] # Store summary in dictionary

    return summaries

def diff_files(lines_from, lines_to, file_from_name, file_to_name):

    with tempfile.NamedTemporaryFile(mode='w+', delete=False) as file_from, tempfile.NamedTemporaryFile(mode='w+', delete=False) as file_to:
        file_from.write(''.join(lines_from))
        file_to.write(''.join(lines_to))
        file_from.flush()
        file_to.flush()

    command = ["git", "diff", "--no-index", file_from.name, file_to.name]
    result = subprocess.run(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)

    if result.returncode not in (0, 1):
        print(f"An error occurred while running git diff: {result.stderr}")
        return []

    diff_output = result.stdout.splitlines(True)[2:]

    if len(diff_output) > 1:
        diff_output[0] = '--- ' + file_from_name + '\n'
        diff_output[1] = '+++ ' + file_to_name + '\n'

    #you can remove the temporary files from here if they are no longer needed
    os.remove(file_from.name)
    os.remove(file_to.name)

    return diff_output

def compare_dirs(summaries_from, summaries_to, file_from, file_to, removed_sheets, renamed_sheets):
    differences = []

    for file_name in summaries_from.keys():

        index = next((i for i, item in enumerate(renamed_sheets) if item[0] == file_name), None)
        file_name_to = file_name
        # Skip the comparison for removed sheets
        if file_name in removed_sheets:
            continue
        elif index is not None:
            file_name_to = renamed_sheets[index][1]
        
        lines_from = summaries_from.get(file_name, [])[0]
        lines_to = summaries_to.get(file_name_to, [])[0]
        locations_from = summaries_from.get(file_name, [])[1]
        locations_to = summaries_to.get(file_name_to, [])[1]
        diff_list = diff_files(lines_from, lines_to, file_from_name=file_from + ' ' + file_name, file_to_name=file_to + ' ' + file_name_to)
        diff_list_final = list(diff_list)

        # Each item in 'line_changes' is a tuple of four integers: from_start, from_count, to_start, to_count
        line_changes = []
        for index, line in enumerate(diff_list):
            match = re.match(r'@@ -(?P<from_start>\d+)(,(?P<from_count>\d+))? \+(?P<to_start>\d+)(,(?P<to_count>\d+))? @@', line)
            if match:
                from_start = int(match.group('from_start'))
                from_count = int(match.group('from_count')) if match.group('from_count') else 1  # handle case where count is blank: @@ -1 +1,2 @@
                to_start = int(match.group('to_start'))
                to_count = int(match.group('to_count')) if match.group('to_count') else 1 
                line_changes.append((from_start, from_count, to_start, to_count, index)) 

        offset = 0
        for change in line_changes:
            from_start, from_count, to_start, to_count, index_start = change
            locations_from_subset = locations_from[from_start - 1:from_start - 1 + from_count]
            locations_to_subset = locations_to[to_start - 1:to_start - 1 + to_count]

            gray = '\033[38;5;232m'  #'\033[90m'
            reset_format = '\033[0m'

            i = 0
            j = 0
            idx = index_start + offset + 1
            while i < from_count or j < to_count:
                line = diff_list_final[idx]
                if line.startswith('-'):
                    diff_list_final.insert(idx, gray + "(-" + locations_from_subset[i] + ")" + reset_format + "\n")
                    i += 1
                elif line.startswith('+'):
                    diff_list_final.insert(idx, gray + "(+" + locations_to_subset[j] + ")" + reset_format + "\n")
                    j += 1
                else:
                    diff_list_final.insert(idx, gray + "(-" + locations_from_subset[i] + " +" + locations_to_subset[j] + ")" + reset_format + "\n")
                    i += 1
                    j += 1
                idx +=2
                offset += 1

        differences.extend(diff_list_final)

    return differences

def similarity(content_from, content_to):
    return SequenceMatcher(None, content_from, content_to).ratio()

def compare_sheet_names(sheet_names_from, sheet_names_to, summaries_from, summaries_to):
    added_sheets = []
    removed_sheets = []
    renamed_sheets = []

    # Check for removed or renamed sheets
    for name_from in sheet_names_from:
        content_from = summaries_from.get(name_from, [])[0]
        if name_from not in sheet_names_to:
            is_renamed = False
            for name_to in sheet_names_to:
                if name_to != name_from: # Ensure names are not the same
                    content_to = summaries_to.get(name_to, [])[0]
                    if similarity(content_from, content_to) > 0.5:
                        renamed_sheets.append((name_from, name_to))
                        is_renamed = True
                        break # renamed sheet identified when first match found, so there shouldn't be a situation where the original sheet is identified as being renamed to multiple new names.
            if not is_renamed:
                removed_sheets.append(name_from)

    # Check for added sheets (excluding those identified as renames)
    for name_to in sheet_names_to:
        if name_to not in sheet_names_from and name_to not in [rename[1] for rename in renamed_sheets]:
            added_sheets.append(name_to)

    return added_sheets, removed_sheets, renamed_sheets

def custom_diff(summaries_from, summaries_to, file_from, file_to):

      # Get the sheet names from the summaries
    sheet_names_from = list(summaries_from.keys())
    sheet_names_to = list(summaries_to.keys())

    # Compare sheet names to find added, removed, and renamed sheets
    added_sheets, removed_sheets, renamed_sheets = compare_sheet_names(
        sheet_names_from, sheet_names_to, summaries_from, summaries_to
    )

    # Print added sheets
    for sheet in added_sheets:
        print('\033[92m' + f"Sheet {sheet} has been added.\n", end='\033[0m')

    # Print removed sheets
    for sheet in removed_sheets:
        print('\033[91m' + f"Sheet {sheet} has been removed.\n", end='\033[0m')

    # Print renamed sheets
    for old_name, new_name in renamed_sheets:
        print('\033[94m' + f"Sheet {old_name} has been renamed to {new_name}.\n", end='\033[0m')


    differences = compare_dirs(summaries_from, summaries_to, file_from, file_to, removed_sheets, renamed_sheets)

    # Accumulate the colored lines
    colored_lines = ""
    for line in differences:
        if line.startswith('+'):
            colored_lines += '\033[92m' + line.replace('\\n', '\n ').replace('\\r', '\r') + '\033[0m'
        elif line.startswith('-'):
            colored_lines += '\033[91m' + line.replace('\\n', '\n ').replace('\\r', '\r') + '\033[0m'  # Red for removed lines
        elif line.startswith('@@'):
            colored_lines += '\033[94m' + line.replace('\\n', '\n ').replace('\\r', '\r') + '\033[0m'  # Blue for @@ lines
        else:
            colored_lines += line.replace('\\n', '\n ').replace('\\r', '\r')
    # Print the result
    print(colored_lines, end='')


if __name__ == "__main__":
    file_to = sys.argv[1]
    file_from = sys.argv[2]

    if not os.path.exists(file_to):
        print(f"File {file_to} has been deleted.")
        sys.exit(0)
    
    unpack_dir_from = tempfile.TemporaryDirectory( prefix= os.path.basename(file_from) + "_unpacked" )
    unpack_dir_to = tempfile.TemporaryDirectory( prefix= os.path.basename(file_to) + "_unpacked" )

    unpack_xlsx(file_from, unpack_dir_from.name)
    unpack_xlsx(file_to, unpack_dir_to.name)

    modified_sheet_from = replace_shared_strings(unpack_dir_from.name)
    modified_sheet_to = replace_shared_strings(unpack_dir_to.name)

    sheet_names_from = get_sheet_names(unpack_dir_from.name)
    sheet_names_to = get_sheet_names(unpack_dir_to.name)

    summaries_from = summarize_sheet_files(modified_sheet_from, sheet_names_from)
    summaries_to = summarize_sheet_files(modified_sheet_to, sheet_names_to)

    custom_diff(summaries_from, summaries_to, file_from, file_to)

    unpack_dir_from.cleanup()
    unpack_dir_to.cleanup()
