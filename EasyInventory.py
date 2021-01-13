#!/usr/bin/python3
# Author: Anthony Garrett
#
# Small script that will read in inventory location data from one spreadsheet
# and transfer that data to another spreadsheet
#
import openpyxl

#
# Input file name is set to the default file name used by the system that
#   generates the initial spreadsheet to be parsed
#

INPUT_FILENAME = 'sqlexec.xlsx'
OUTPUT_FILENAME = 'completed.xlsx'

#
# Helper function to generate the dictionary that will hold the inventory count
# for the number of cases and pallets in the aisle in the base case with no
# special cases
#
# The value array is used as [pallet count, case count] and set to 0 for all
# as the default
#
# Input paramenters are the aisle number as an int and the ending location for
# the aisle
#

def generate_inner_dict(aisle_number, ending_number):
    """
    Returns the default generated dictionary with the given aisle number
    and ending aisle location number.
    """

    locations_data = {};

    for i in range(1, ending_number):
        locations_data[str(aisle_number) + '-' + str((100 + i)) + '-A'] = [0, 0]
        locations_data[str(aisle_number) + '-' + str((100 + i)) + '-B'] = [0, 0]
        locations_data[str(aisle_number) + '-' + str((100 + i)) + '-C'] = [0, 0]
        locations_data[str(aisle_number) + '-' + str((100 + i)) + '-D'] = [0, 0]

    for j in range(1, ending_number):
        locations_data[str(aisle_number) + '-' + str((200 + j)) + '-A'] = [0, 0]
        locations_data[str(aisle_number) + '-' + str((200 + j)) + '-B'] = [0, 0]
        locations_data[str(aisle_number) + '-' + str((200 + j)) + '-C'] = [0, 0]
        locations_data[str(aisle_number) + '-' + str((200 + j)) + '-D'] = [0, 0]

    return locations_data;


#
# Function to generate the dictionary that will hold the inventory count
# for the number of cases and pallets in the aisle.
#
# The value array is used as [pallet count, case count] and set to 0 for all
# as the default
#
# Special cases of uneven aisles and aisles that lack A levels on one side
# are dealt with in this function and base case aisles are handled in 
# generate_inner_dict()
#
# Input paramenters are the aisle number as an int
#
def create_dictionary(aisle_number):
    """
    Returns the default generated dictionary with the given aisle number including special cases
    """
    locations = {}

#
# Decides what to do with each aisle number based on how the aisle is organized
# in the warehouse. Each dictionary is created with the correct amount of valid
# aisle locations all set to [0, 0]
#
#
    #
    # Aisles 1-3 only have 56 locations on each side
    #
    if(aisle_number < 4):
        locations = generate_inner_dict(aisle_number, 57);

    #
    # Aisle 4 has only have 56 locations the 100 side and 62 locations on the
    # 200 side
    #
    elif (aisle_number == 4):

        for i in range(1, 57):
            locations[str(aisle_number) + '-' +
                          str((100 + i)) + '-A'] = [0, 0]
            locations[str(aisle_number) + '-' + str((100 + i)) + '-B'] = [0, 0]
            locations[str(aisle_number) + '-' + str((100 + i)) + '-C'] = [0, 0]
            locations[str(aisle_number) + '-' + str((100 + i)) + '-D'] = [0, 0]

        for j in range(1, 63):
            locations[str(aisle_number) + '-' +
                          str((200 + j)) + '-A'] = [0, 0]
            locations[str(aisle_number) + '-' + str((200 + j)) + '-B'] = [0, 0]
            locations[str(aisle_number) + '-' + str((200 + j)) + '-C'] = [0, 0]
            locations[str(aisle_number) + '-' + str((200 + j)) + '-D'] = [0, 0]

    #
    # Aisles 5 - 12 have base case locations with 62 on each side
    #
    elif(4 < aisle_number < 13):
        locations = generate_inner_dict(aisle_number, 62)
    
    #
    # Aisle 13 is missing the A levels on the 200 side
    #
    elif(aisle_number == 13):

        for i in range(1, 63):
            locations[str(aisle_number) + '-' +
                          str((100 + i)) + '-A'] = [0, 0]
            locations[str(aisle_number) + '-' + str((100 + i)) + '-B'] = [0, 0]
            locations[str(aisle_number) + '-' + str((100 + i)) + '-C'] = [0, 0]
            locations[str(aisle_number) + '-' + str((100 + i)) + '-D'] = [0, 0]

        for j in range(1, 63):
            locations[str(aisle_number) + '-' + str((200 + j)) + '-B'] = [0, 0]
            locations[str(aisle_number) + '-' + str((200 + j)) + '-C'] = [0, 0]
            locations[str(aisle_number) + '-' + str((200 + j)) + '-D'] = [0, 0]

    #
    # Aisle 14 is missing the A levels on the 100 side
    #
    elif(aisle_number == 14):

        for i in range(1, 63):
            locations[str(aisle_number) + '-' + str((100 + i)) + '-B'] = [0, 0]
            locations[str(aisle_number) + '-' + str((100 + i)) + '-C'] = [0, 0]
            locations[str(aisle_number) + '-' + str((100 + i)) + '-D'] = [0, 0]

        for j in range(1, 63):
            locations[str(aisle_number) + '-' +
                          str((200 + j)) + '-A'] = [0, 0]
            locations[str(aisle_number) + '-' + str((200 + j)) + '-B'] = [0, 0]
            locations[str(aisle_number) + '-' + str((200 + j)) + '-C'] = [0, 0]
            locations[str(aisle_number) + '-' + str((200 + j)) + '-D'] = [0, 0]

    #
    # Aisles 15 - 26 have base case locations with 62 on each side
    #
    elif(aisle_number <= 26):
        locations = generate_inner_dict(aisle_number, 63)

    return locations

def main():

    wb = openpyxl.load_workbook(INPUT_FILENAME)

    sheet = wb.active

    location_codes = sheet['A']
    actual_quantity = sheet['N']

    a = location_codes[0].value.split("-")

    data_core = create_dictionary(int(a[0]))

    for count,code in enumerate(location_codes):
        data_core[code.value.strip()][0] += 1
        data_core[code.value.strip()][1] += actual_quantity[count].value

    out_book = openpyxl.Workbook()

    out_sheet = out_book.active

    count = [1, 1, 1, 1]

    for key, value in data_core.items():
        find_level = key.split("-")
        level = find_level[2]

        if (str(level) == 'A'):
            out_sheet['A' + str(count[0])] = key
            out_sheet['B' + str(count[0])] = value[0]
            out_sheet['C' + str(count[0])] = value[1]
            count[0] += 1

        elif (str(level) == 'B'):
            out_sheet['E' + str(count[1])] = key
            out_sheet['F' + str(count[1])] = value[0]
            out_sheet['G' + str(count[1])] = value[1]
            count[1] += 1

        elif (str(level) == 'C'):
            out_sheet['I' + str(count[2])] = key
            out_sheet['J' + str(count[2])] = value[0]
            out_sheet['K' + str(count[2])] = value[1]
            count[2] +=1

        elif (str(level) == 'D'):
            out_sheet['M' + str(count[3])] = key
            out_sheet['N' + str(count[3])] = value[0]
            out_sheet['O' + str(count[3])] = value[1]
            count[3] += 1


    out_book.save(OUTPUT_FILENAME)



if __name__ == "__main__":
    main()