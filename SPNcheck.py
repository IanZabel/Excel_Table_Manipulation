import openpyxl as openpy
import os


class SPNcheck:

    # This method gathers the user input and determines the intended panel to analyze and output data for.
    def panel_select(self):

        # As needed, this dictionary is populated with the panels that are supported.
        panel_dictionary = {
            'cons': 'Consolidated Panel',
            'rl': 'Reading Light Panel'
        }

        # Take the above dictionary and present it in a clear fashion
        for i, j in panel_dictionary.items():
            print(i, '=', j)

        print('\nUse the key above for reference. Case sensitive.')
        print('Type "exit" to close the program.')
        user_selection = input('    Which Panel?: ')
        user_selection = user_selection.lower()

        # Remind the user what they asked for
        if user_selection.rstrip() in panel_dictionary:
            print('You selected', panel_dictionary[user_selection], '\n')
            return True, user_selection, panel_dictionary

        # Exit when the ask to
        elif user_selection.rstrip() is 'exit':
            print('Closing.')
            return False, user_selection

        # Tell them to try again
        else:
            print('Not supported.')
            return False, user_selection, panel_dictionary

    # This method calculates the center of mass and overall weight for each configuration present.
    def excel_parse(panel_selection):

        panel_selection_string = str(panel_selection)

        # Displays currently supported panels and their corresponding sheet names
        sheet_dictionary = {
            'cons': ['Cons Components', 'Cons CG', 'Cons Dash', 'Cons SPN Key'],
            'rl': ['RL Components', 'RL CG', 'RL Dash', 'RL SPN Key']
        }

        # Point to the location of the ScriptPullTable and ScriptPlaceTable documents.
        os.chdir(r'P:/DIRECTORY_PATH')
        print('Working Directory:', os.getcwd())

        print('Accessing source workbook...')
        workbook1 = openpy.load_workbook(filename='ScriptPullTable.xlsx', data_only=True)
        print('Accessed.')

        if panel_selection_string in sheet_dictionary:
            cg_sheet = workbook1[sheet_dictionary[panel_selection_string][1]]
            config_sheet = workbook1[sheet_dictionary[panel_selection_string][2]]
        else:
            return False

        item, dash, mx, my, mz, mass, spn = [], [], [], [], [], [], []
        group_position = 1
        column = 0
        group_quantity = 5

        # Population and calculation of the moments, masses, and categories
        while group_position <= group_quantity:
            for i in range(0, group_quantity):
                column = ((group_position-1)*group_quantity)+i+1

            for row in range(3, cg_sheet.max_row+1):
                if cg_sheet.cell(row=row, column=column).value is not None:

                    # item: category for determining context of mass and CG values
                    # mass: calculated total mass of the configuration
                    item.append(cg_sheet.cell(row=row, column=(group_quantity*group_position - 4)).value)
                    mass.append(cg_sheet.cell(row=row, column=(group_quantity*group_position)).value)

                    # Calculation and population of moment values
                    mx.append(mass[row]*cg_sheet.cell(row=row, column=(group_quantity*group_position - 3)).value)
                    my.append(mass[row]*cg_sheet.cell(row=row, column=(group_quantity*group_position - 2)).value)
                    mz.append(mass[row]*cg_sheet.cell(row=row, column=(group_quantity*group_position - 1)).value)

                else:
                    continue

            group_position += 1

        # Population of the dash number and SPN lists
        for row in range(1, config_sheet.max_row + 1):
            dash.append(config_sheet.cell(row=row, column=1).value)
            spn.append(config_sheet.cell(row=row, column=2).value)



