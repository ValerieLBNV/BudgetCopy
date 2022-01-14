import openpyxl as xl
import shutil
from openpyxl import Workbook
import os

INPUT_PATH = "C:\\Users\\vval\\Timex Group\\Sangion, Martina - IKA\\Budget\\2022\\Budget Tracker\\Flash files per entity\\100AMS_IntlKA_Weekly Flash Update.xlsx"
OUTPUT_PATH = "C:\\Users\\vval\\Timex Group\\Sangion, Martina - IKA\\Budget\\2022\\Budget Tracker\\Others\\Datasource Budget-Revenue tracker 2022.xlsx"
ARCHIVE_FOLDER_PATH = "C:\\Users\\vval\\Timex Group\\Sangion, Martina - IKA\\Budget\\2022\\Budget Tracker\\Flash files per entity\\Archive monthly versions\\"
ARCHIVE_FILE_NAME = os.path.basename(os.path.splitext(INPUT_PATH)[0])
SHEET_NAME = 'All'

dictionary = {
    "100AMS_IntlKA_Weekly Flash Update": '100 AMS',
    "220UK_IntlKA_Weekly Flash Update": '220 UK',
    "570 980 Vert_IntlKA_Weekly Flash Update": '570 980 Vert',
    "720TSH_IntlKA_Weekly Flash Update": '720 TSH',
    "788HK_IntlKA_Weekly Flash Update": '788 HK'
}

def archive(version):
    new_name = ARCHIVE_FILE_NAME + " " + version + ".xlsx"
    archive_file_path = shutil.copyfile(INPUT_PATH ,ARCHIVE_FOLDER_PATH + new_name)

    archive_file = xl.load_workbook(archive_file_path)
    consolidated_sheet = archive_file[SHEET_NAME]

    consolidated_sheet['A1'].value = int(version)
    archive_file.save(archive_file_path)

#TODO: Create a copy of output file before adding anything
def load_input_worksheet():
    input_workbook = xl.load_workbook(INPUT_PATH)
    input_worksheet = input_workbook[SHEET_NAME]
    return input_worksheet

def find_starting_cell(output_version_cell, output_worksheet, output_version_column, version):
    chosen_cell = ''
    for number in range(output_version_cell.row+1, output_worksheet.max_row+1): #row+1: excl.header, max+1 incl last max row
        chosen_cell = output_worksheet.cell(number, output_version_column)
        # case 1 where we found a cell with the same version
        if chosen_cell.value == int(version):
            return chosen_cell
    # case 2 where we did not find a cell with the same version, and we return the last cell in the column
    return output_worksheet.cell(chosen_cell.row+1, chosen_cell.column)

def copy_data(version):
    # loading from input excel
    input_workbook = xl.load_workbook(INPUT_PATH, data_only=True)
    input_worksheet = input_workbook[SHEET_NAME]

    # input_worksheet = load_input_worksheet()
    input_version_cell = get_version_cell(input_worksheet)
    input_version_cell_below = input_worksheet.cell(input_version_cell.row +1, input_version_cell.column) # eliminate header row

    max_cell = input_worksheet.cell(input_worksheet.max_row, input_worksheet.max_column)
    cell_range = input_worksheet[input_version_cell_below.coordinate :max_cell.coordinate]

    # find version cell in output to get the version column
    output_workbook = xl.load_workbook(OUTPUT_PATH)
    output_worksheet = output_workbook[SHEET_NAME]
    output_version_cell = get_version_cell(output_worksheet)
    output_version_column = output_version_cell.column

    # find location to paste to
    chosen_cell = find_starting_cell(output_version_cell, output_worksheet, output_version_column, version)

    # Paste cell range
    row_counter = 0
    for row in cell_range:
        cell_counter = 0
        for cell in row:
            if cell_counter == 0: # overwrite all value of version cells to indicated version
                output_worksheet.cell(chosen_cell.row + row_counter, chosen_cell.column).value = int(version)
            else: # paste the rest of cell range
                output_worksheet.cell(chosen_cell.row + row_counter, chosen_cell.column + cell_counter).value = cell.value
            cell_counter += 1
        row_counter += 1

    #  output_worksheet.cell(chosen_cell.row + row_counter, chosen_cell.column) <- specify coordinate of each cell

    # TODO: delete previous data of 1 measure
    output_workbook.save(OUTPUT_PATH)

def get_version_cell(sheet):
    for row in sheet:
        for cell in row:
            value = str(cell.value)
            if value.lower() == "version":
                return cell
    raise ValueError('There is no specified cell')
# create: make a copy of the output file before changing anything to prevent accidental overwrites


def choose_files(type, version):
    print("Here are the files we have: ")
    a = 1
    for entity in dictionary:
        print(entity +": " + str(a))
        a += 1
    choice = input("Which file do you want to work with: ")
    files_chosen = []
    for character in choice:
        if character.isdigit():
            files_chosen.append(character)

    type_choice = ""
    if type == "1":
        type_choice = "archive"
    else:
        type_choice = "copy"

    action = ""
    while action != "yes" and action != "no":
        print("Run " + type_choice + " version " + version + " for: " + str(files_chosen) + " proceed? yes/no")
        action = input()

    if action == "yes":
        return files_chosen
    else:
        choose_files(type, version)

def main():
    type = input("Do you want to archive file or copy data. Type 1 for archive, 2 for copy: ")
    version = input("Specify version number: ")

    if type == "1":
        chosen_files = choose_files(type, version)
        archive(version)

    elif type == "2":
        chosen_files = choose_files(type, version)
        copy_data(version)

    else:
        main()

#TODO: you already have an archive / output copy of this, do you want to overwrite again?

# main method
if __name__ == '__main__':
    main()
