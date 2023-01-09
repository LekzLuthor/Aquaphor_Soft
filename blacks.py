import os
import pprint

import openpyxl


def load_database():
    files_name = os.listdir("sours/data/")
    for f_index, f in enumerate(files_name):

        print()

        excel_file = openpyxl.open(f'sours/data/{f}', read_only=True)
        sheet = excel_file.active

        start_line_ind = 0
        while sheet[f'B{start_line_ind}'].value != "№ п/п":
            start_line_ind += 1
        start_line_ind += 3

        end_line_ind = start_line_ind + 1
        while sheet[f'B{end_line_ind}'].value is not None:
            end_line_ind += 1

        equipment = []
        for ind in range(start_line_ind, end_line_ind + 1):
            line = [i.value for i in sheet[f'A{ind}':f'L{ind}'][0]]
            equipment.append(line)


load_database()
