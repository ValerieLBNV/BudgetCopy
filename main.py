import openpyxl as xl
import shutil
from openpyxl import Workbook
import os
from pathlib3x import Path

INPUT_PATH = "C:\\Users\\vval\\Timex Group\\Sangion, Martina - IKA\\Budget\\2022\\Budget Tracker\\Flash files per entity\\"
OUTPUT_PATH = "C:\\Users\\vval\\Timex Group\\Sangion, Martina - IKA\\Budget\\2022\\Budget Tracker\\Others\\Datasource Budget-Revenue tracker 2022.xlsx"
ARCHIVE_FOLDER_PATH = "C:\\Users\\vval\\Timex Group\\Sangion, Martina - IKA\\Budget\\2022\\Budget Tracker\\Flash files per entity\\Archive monthly versions\\"
# ARCHIVE_FILE_NAME = os.path.basename(os.path.splitext(INPUT_PATH)[0])
SHEET_NAME = 'All'

dictionary = {"100AMS_IntlKA_Weekly Flash Update": '100 AMS',
              "220UK_IntlKA_Weekly Flash Update": '220 UK',
              "570 980 Vert_IntlKA_Weekly Flash Update": '570 980 Vert',
              "720TSH_IntlKA_Weekly Flash Update": '720 TSH',
              "788HK_IntlKA_Weekly Flash Update": '788 HK'}

keys = list(dictionary.keys())
values = list(dictionary.values())

def get_version_cell(sheet):
    for row in sheet:
        for cell in row:
            value = str(cell.value)
            if value.lower() == "version":
                return cell
    raise ValueError('There is no specified cell')

def idk():

def find_starting_cell(output_version_cell, output_worksheet, output_version_column, version, overwrite):
    chosen_cell = ''
    overwrite = ""
    for number in range(output_version_cell.row + 1,
                        output_worksheet.max_row + 1):  # row+1: excl.header, max+1 incl last max row
        chosen_cell = output_worksheet.cell(number, output_version_column)
        # case 1 where we found a cell with the same version

        if chosen_cell.value == int(version):
            while overwrite.lower() != "yes" and overwrite.lower() != "no":
                overwrite = input("version" + version + "in" + output_worksheet + "already exists. Overwrite? Yes/no:")
            # TODO: here, you have found a cell that matches your version. That means that overwriting will happen.
            #  We will need to add another value to return. Just write return x, y
            #  REMEMBER - you will also need to save both of the values too (where you are calling the method from).
            #  You can save them by saying x, y = method()
            #  You can easily find where the method is called from by clicking wheel on your mouse and hovering over the name of the methdod

            # TODO: Now to put that variable to good use:
            #  Print out a question where you ask if you want to overwrite. Whatever the answer, return the value (together with the value that is being returned currently, do as described above)
            #  It might be useful to check if the user input makes sense (if its either yes or no). Use while loop as we did before

            return chosen_cell, overwrite
    # case 2 where we did not find a cell with the same version, and we return the last cell in the column
    overwrite ="yes"
    return output_worksheet.cell(chosen_cell.row + 1, chosen_cell.column)


def copy_data(version, chosen_files):
    # put entire code below into a for loop (VERY similar to archive method)
    # for loop should loop over the numbers that you chose, based on those numbers - take the value from list "keys" and also list "values" (get data from both of them from the same position -1)
    # from keys you will get the name of the file to read from and from values you will get the names of the sheets to paste to
    # change these values in the code below to these newly gotten values -> nice and dynamic.

    for entity_number in chosen_files:
        # loading from input excel
        input_file_name = keys[int(entity_number) - 1] + ".xlsx"
        input_workbook = xl.load_workbook(INPUT_PATH + input_file_name, data_only=True)
        input_worksheet = input_workbook[SHEET_NAME]

        input_version_cell = get_version_cell(input_worksheet)
        input_version_cell_below = input_worksheet.cell(input_version_cell.row + 1,
                                                        input_version_cell.column)  # eliminate header row

        max_cell = input_worksheet.cell(input_worksheet.max_row, input_worksheet.max_column)
        cell_range = input_worksheet[input_version_cell_below.coordinate:max_cell.coordinate]

        # find version cell in output to get the version column
        output_workbook = xl.load_workbook(OUTPUT_PATH)
        output_sheet_name = values[int(entity_number) - 1]
        output_worksheet = output_workbook[output_sheet_name]
        output_version_cell = get_version_cell(output_worksheet)
        output_version_column = output_version_cell.column

        # find location to paste to
        chosen_cell = find_starting_cell(output_version_cell, output_worksheet, output_version_column, version)
        # TODO: when you have returned both values, you need to make a check here. (if you are not yet returning two values, see TODO inside of the method find_starting_cell)
        # TODO: inside of the check (if statement), check if the returned input is no (no overwriting). If it is that, then you should close the file (you need to close it, otherwise it will lag ur pc)
        #   and execute (write to the code) - continue
        #   Continue will make the code skip the entire loop - which means that the entire file will be ignored and we will move to another one.
        #   No need to check for anything else, just let the code run.
        # Paste cell range
        row_counter = 0
        for row in cell_range:
            cell_counter = 0
            for cell in row:
                if cell_counter == 0:  # overwrite all value of version cells to indicated version
                    output_worksheet.cell(chosen_cell.row + row_counter, chosen_cell.column).value = int(version)
                else:  # paste the rest of cell range
                    output_worksheet.cell(chosen_cell.row + row_counter,
                                          chosen_cell.column + cell_counter).value = cell.value
                cell_counter += 1
            row_counter += 1

        #  output_worksheet.cell(chosen_cell.row + row_counter, chosen_cell.column) <- specify coordinate of each cell

        # TODO: delete previous data of measure YTG (or find how in DAX)
        output_workbook.save(OUTPUT_PATH)

def archive_action(input_file_name, version, new_file_name):
    archive_file_path = shutil.copyfile(INPUT_PATH + input_file_name + ".xlsx", ARCHIVE_FOLDER_PATH + new_file_name)

    archive_file = xl.load_workbook(archive_file_path)
    consolidated_sheet = archive_file[SHEET_NAME]

    consolidated_sheet['A1'].value = int(version)
    archive_file.save(archive_file_path)
    print("Archived " + new_file_name)

def archive(version, chosen_files):
    for entity_number in chosen_files:
        input_file_name = keys[int(entity_number) - 1]
        new_file_name = input_file_name + " " + version + ".xlsx"
        if Path(ARCHIVE_FOLDER_PATH + new_file_name).is_file():
            overwrite = ""
            while overwrite.lower() != "yes" and overwrite.lower() != "no":
                overwrite = input(new_file_name + " already exists. Overwrite? Yes/no: ")
                if overwrite.lower() == "yes":
                    archive_action(input_file_name, version, new_file_name)
                elif overwrite.lower() == "no":
                    print('Skipping')
                    continue
        else:
            archive_action(input_file_name, version, new_file_name)

def choose_files(type, version):
    print("Here are the files we have: ")
    a = 1
    for entity in dictionary:
        print("\t" + entity + ": " + str(a))
        a += 1
    choice = input("Which file do you want to work with: ")
    files_chosen = []
    for character in choice:
        if character.isdigit():
            files_chosen.append(character)

    if type == "1":
        type_choice = "archive"
    else:
        type_choice = "copy"

    action = ""
    while action.lower() != "yes" and action.lower() != "no":
        print("Run " + type_choice + " version " + version + " for: " + str(files_chosen) + " proceed? yes/no")
        action = input()

    if action.lower() == "yes":
        return files_chosen
    else:
        choose_files(type, version)

def main():
    type = input("Do you want to archive file or copy data. Type 1 for archive, 2 for copy: ")
    version = input("Specify version number: ")

    if type == "1":
        chosen_files = choose_files(type, version)
        archive(version, chosen_files)

    elif type == "2":
        chosen_files = choose_files(type, version)
        copy_data(version, chosen_files)

    else:
        main()


# TODO: you already have an archive / output copy of this, do you want to overwrite again?

# main method
if __name__ == '__main__':
    main()
