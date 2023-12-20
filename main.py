import pandas as pd
import customtkinter
from tkinter import filedialog
import threading
# to check size of file and paths
import os
from collections import Counter
import math


class App(customtkinter.CTk):
    def __init__(self):
        super().__init__()
        customtkinter.set_appearance_mode("dark")
        self.title("Split to files")
        self.minsize(1000, 350)

        # Create grid 2x2
        self.grid_rowconfigure(2, weight=1)
        # noinspection PyTypeChecker
        self.grid_columnconfigure((0, 1, 2, 3, 4), weight=1)

        # Choose excel file button
        self.button = customtkinter.CTkButton(master=self, command=self.get_file, text="Choose excel file")
        self.button.grid(row=0, column=0, padx=10, pady=10, sticky='nsew')

        # Download function button
        self.start_download = customtkinter.CTkButton(master=self, text="Generate CSV files", command=self.run_download)
        self.start_download.grid(row=1, column=0, padx=10, pady=10, sticky='nsew')

        # Label with chosen Path
        global label
        label = customtkinter.CTkLabel(master=self, text='Directory')
        label.grid(row=2, columnspan=3, padx=20, pady=10, sticky='w')

        # Open chosen excel file
        self.open_file = customtkinter.CTkButton(master=self, text="Open chosen file", command=self.open_file)
        self.open_file.grid(row=3, column=0, padx=10, pady=10, sticky='nsew')

        # Open folder with downloaded files
        self.open_dl_folder = customtkinter.CTkButton(master=self, text="Open download folder", command=self.open_dl_folder)
        self.open_dl_folder.grid(row=4, column=0, padx=10, pady=10, sticky='nsew')

        # TextBox to show what's happening
        global textbox
        textbox = customtkinter.CTkTextbox(master=self)
        textbox.grid(row=0, column=3, rowspan=5, columnspan=2, padx=10, pady=10, sticky='nsew')

    @staticmethod
    def run_download():
        task = threading.Thread(target=split_to_files)
        task.start()
        return
    
    @staticmethod
    def open_dl_folder():
        try:
            os.startfile(save_path)
        except NameError:
            textbox.insert('end', 'No folder to be opened!\n')
            
    @staticmethod
    def open_file():
        try:
            os.startfile(workbook_path)
        except NameError:
            textbox.insert('end', 'No file chosen to be opened!\n')

    @staticmethod
    def get_file():
        global workbook_path
        workbook_path = filedialog.askopenfilename(initialdir=os.getcwd(), title="Select workbook with links",
                                                   filetypes=(('Excel with Macro', '*.xlsm'),('Excel files', '*.xlsx'), ('All files', '*.*')))
        label.configure(text=workbook_path)

# Definiujemy funkcję, która przyjmuje argument jako dataframe
def remove_duplicates_and_save(df,max_data, column_prefix, path_prefix, save_path):
    result = {}
    for index, row in df.iterrows():
        ELNR = row['ELNR']
        Filename = path_prefix + row['Filename']
        description = row['Description']
        if ELNR in result:
            result[ELNR].append((Filename, description))
        else:
            result[ELNR] = [(Filename, description)]

    # Obliczamy liczbę plików do zapisania
    num_files = math.ceil(sum(len(v) for v in result.values()) / max_data)

    # Zapisujemy dane do plików Excela
    start_index = 0
    for i in range(num_files):
        end_index = start_index
        while sum(len(v) for v in list(result.values())[start_index:end_index]) < max_data and end_index < len(result):
            end_index += 1
        values_slice = list(result.values())[start_index:end_index]
        if not values_slice:
            break
        # Tworzymy nowy dataframe dla każdego pliku
        new_df = pd.DataFrame()
        new_df['ELNR'] = list(result.keys())[start_index:end_index]
        max_links = max(len(v) for v in list(result.values())[start_index:end_index])
        for j in range(1, max_links + 1):
            column_name_file = f'{column_prefix}_{j}_FILE'
            column_name_description = f'{column_prefix}_{j}_DESCRIPTION'
            new_df[column_name_file] = [v[j-1][0] if len(v) >= j else None for v in list(result.values())[start_index:end_index]]
            new_df[column_name_description] = [v[j-1][1] if len(v) >= j else None for v in list(result.values())[start_index:end_index]]

        new_df.to_csv(f'{save_path}/{column_prefix}{i+1}.csv',sep=';', index=False, encoding='utf-8')
        start_index = end_index


def split_to_files():
    global save_path
    save_path = filedialog.askdirectory(initialdir=os.getcwd(), title="Select where to generate CSV files")
    dialog = customtkinter.CTkInputDialog(text='default 1000', title='Input how many entries per file.')
    try:
        max_data = int(dialog.get_input())
    except ValueError:
        textbox.insert('end', 'Not a valid number. Using default value (1000)\n')
        max_data = 1000
    with pd.ExcelFile(workbook_path) as xlsx:
        CARD = pd.read_excel(xlsx,0, converters={'ELNR':str})
        CERTIFICATE = pd.read_excel(xlsx,1, converters={'ELNR':str})
        DIALUX = pd.read_excel(xlsx,2, converters={'ELNR':str})
        REACH = pd.read_excel(xlsx,3, converters={'ELNR':str})
        ROHS = pd.read_excel(xlsx,4, converters={'ELNR':str})
        TECHDOC = pd.read_excel(xlsx,5, converters={'ELNR':str})
        IMAGE = pd.read_excel(xlsx,6, converters={'ELNR':str})
    array = [CARD, CERTIFICATE, DIALUX, REACH, ROHS, TECHDOC, IMAGE]
    # GET SHEET NAMES IN ORDER
    names = xlsx.sheet_names
    prefixes = ['/PDF/CARD/', '/PDF/CERTIFICATE/', '/PDF/DIALUX/', '/PDF/REACH/', '/PDF/ROHS/', '/PDF/TECHDOC/', 'photo\\']
    print("Names ", names)

    name_indicator = 0
    for entry in array:
        if not entry.empty:
            print("im inside not ", names[name_indicator], "data ", array[name_indicator])
            remove_duplicates_and_save(array[name_indicator], max_data, names[name_indicator], prefixes[name_indicator], save_path)
            name_indicator += 1
        else:
            name_indicator += 1
    textbox.insert('end', 'Files generated\n')

if __name__ == "__main__":
    app = App()
    app.mainloop()
