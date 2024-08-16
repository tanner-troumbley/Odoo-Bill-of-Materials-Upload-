import pdb

import pandas as pd
import openpyxl
from tkinter import * 
from tkinter import filedialog
import numpy as np

class PandasMagic:
    def choose_file(self):
        gui = Tk()
        gui.withdraw()

        file_path = filedialog.askopenfilename(multiple=True)

        return file_path

    def main(self):
        file_tuple = self.choose_file()
        # Loops through the list to be able to select multiple files.
        for filename in file_tuple:
            # Checks what type of file it is and then read's it.
            try:
                if 'xls' in filename:
                    data = pd.read_excel(filename)
                    index = filename.find('.xls')
                    NewFile = filename[:index] + '_edited.xlsx'
                    self.panadas_edit_magic(data, NewFile)

                elif 'csv' in filename:
                    data = pd.read_csv(filename)
                    index = filename.find('.csv')
                    NewFile = filename[:index] + '_edited.xlsx'
                    self.panadas_edit_magic(data, NewFile)

                else:
                    print(filename)
                    return print('File Not Supported')

            except PermissionError:
                print("Close ", filename, " before using it.")
                return
        print("Edits Finished")
    def panadas_edit_magic(self, data, filename):
        data = data.astype({'categ_id': str})
        data = data.astype({'TYPE': str})
        data = data.astype({'Description': str})
        data = data.astype({'TYPE': str})
        data.loc[data['Description'].isnull(), 'EPS Part Number'] = np.nan
        data.loc[data['Description'].str.isspace(), 'EPS Part Number'] = np.nan
        data = data.astype({'EPS Part Number': str})
        data['EPS Part Number'] = data['EPS Part Number'].str.upper()

        # Put a blank value into categ_id to be able to work with it.
        data.loc[data['EPS Part Number'].notnull(), 'categ_id'] = "FILL IN MANUALLY"
        # Type Logic
        data.loc[data['EPS Part Number'].notnull(), 'TYPE'] = "Storable Product"

        data.loc[data['EPS Part Number'].str.contains('NAN', case=False), 'categ_id'] = ''
        data.loc[data['EPS Part Number'].str.contains('NAN', case=False), 'TYPE'] = ''
        data.loc[data['EPS Part Number'].str.contains('NAN', case=False), 'EPS Part Number'] = ''

        data.loc[data['EPS Part Number'].str.contains('EWA_H', case=False), 'categ_id'] = 'Custom / Manufactured Parts'
        data.loc[data['EPS Part Number'].str.contains('Harness', case=False), 'categ_id'] = 'Custom / Manufactured Parts'

        data.loc[data['EPS Part Number'].str.contains('A0', case=False), 'categ_id'] = 'Custom / Assemblies'
        data.loc[data['EPS Part Number'].str.contains('ASM', case=False), 'categ_id'] = 'Custom / Assemblies'

        data.loc[data['EPS Part Number'].str.contains('P0', case=False), 'categ_id'] = 'Standard / Machined'
        data.loc[data['EPS Part Number'].str.contains('IP', case=False), 'categ_id'] = 'Standard / Machined'
        data.loc[data['EPS Part Number'].str.contains('PRT', case=False), 'categ_id'] = 'Standard / Machined'

        data.loc[data['EPS Part Number'].str.contains('ELC', case=False), 'categ_id'] = 'Standard / Electronics'
        data.loc[data['EPS Part Number'].str.contains('CON', case=False), 'categ_id'] = 'Standard / Consumables'
        data.loc[data['EPS Part Number'].str.contains('HRD', case=False), 'categ_id'] = 'Standard / Hardware'
        data.loc[data['EPS Part Number'].str.contains('CELL', case=False), 'categ_id'] = 'Standard / Cells'
        data.loc[data['EPS Part Number'].str.contains('CCA', case=False), 'categ_id'] = 'Standard / CCA'
        data.loc[data['EPS Part Number'].str.contains('BRD', case=False), 'categ_id'] = 'Standard / Board Components'

        #Removes a type for any time that dosen't fit in the categories.
        data.loc[data['categ_id'].str.contains('FILL IN MANUALLY'), 'TYPE'] = None

        data.to_excel(filename)

        print(filename + ' Created.')
        return


