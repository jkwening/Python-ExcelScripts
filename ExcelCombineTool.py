import tkFileDialog
import os
import pandas as pd
import xlrd as xl
import Tkinter as tk


class ExcelScript:
    """
    Scripts allows user to combine multiple sheets or Excel files into one
    Excel file; leaving the original intact.
    """
    def __init__(self):
        root = tk.Tk()
        root.withdraw()
        # get request from user what they would like to do
        self.prompt()

    def prompt(self):
        """
        Entry point of the script that prompts the user for action to perform
        """
        # get request from user what they would like to do
        print "\nMAIN SCREEN"
        print "\nWhat would you like to do: "
        print "0: Exit"
        print "1: Combine multiple sheets into one Excel file?"
        print "2: Combine multiple Excel files into one?"
        print
        action = raw_input("Enter 0, 1, or 2:")

        # perform appropriate method
        if int(action) == 0:
            exit()
        elif int(action) == 1:
            self.append_sheets()
        elif int(action) == 2:
            self.append_files()
        else:
            print "Not a valid option!"
            self.prompt()

    def append_sheets(self, final_data=None):
        """
        Process for combining multiple sheets in one or more files into one Excel file
        :param final_data: the data frame consisting of combined sheets
        """
        # request Excel file from user
        print "\nAPPEND SHEETS TOOL"
        print "Find file open dialogue box; maybe behind current screen!"
        file_name = tkFileDialog.askopenfilename(title="SELECT EXCEL FILE WITH MULTIPLE SHEETS!",
                                                 filetypes=[("Excel", "*.xlsx")])

        # go back to main screen if user selects "cancel"
        if file_name == "":
            self.prompt()

        # get sheet names
        sheet_names = xl.open_workbook(file_name).sheet_names()

        # combine sheets into one data frame
        f_name = os.path.basename(file_name)
        print "\nCOMBINING SHEETS..."
        if final_data is None:  # initialize data frame if it doesn't exist yet
            final_data = pd.DataFrame()
        for name in sheet_names:  # iterate through each sheet and add to data frame
            print "Appending sheet: " + name + "..."
            df = pd.read_excel(file_name, name)
            # add new coln for form name
            df["Sheet Name"] = name
            df["File Name"] = f_name
            final_data = final_data.append(df, ignore_index=True)

        print "DONE!"

        # see if user wants to combine more sheets
        action = -1
        while action == -1:
            print "\nDo you want to combine more sheets from another file?"
            print "0: NO"
            print "1: YES"
            action = raw_input("ENTER 0 or 1:")

        if int(action) == 1:
            self.append_sheets(final_data)
        else:
            self.save_to_file(final_data)
            self.prompt()

    def append_files(self, final_data=None):
        """
        Process for combining multiple Excel file into one Excel file
        :param final_data: the data frame to append to if one already exists
        """
        # request Excel file from user
        print "\nAPPEND FILES TOOL"
        print "Find file open dialogue box; maybe behind current screen!"
        file_name = tkFileDialog.askopenfilenames(title="SELECT EXCEL FILE TO COMBINE!",
                                                  filetypes=[("Excel", "*.xlsx")])  # tuple of paths

        # go back to main screen if user selects "cancel"
        if file_name == "":
            self.prompt()

        # start processing files and notify user
        print "\nCOMBINING FILES..."
        if final_data is None:  # initialize data frame if doesn't already exist
            final_data = pd.DataFrame()

        # iterate through each file selected and append first sheet to data frame
        for name in file_name:
            f_name = os.path.basename(name)
            print "Appending file: " + f_name + "..."
            df = pd.read_excel(name)
            df["File Name"] = f_name
            final_data = final_data.append(df, ignore_index=True)

        print "DONE!"

        # see if user wants to combine more files
        action = -1
        while action == -1:
            print "\nDo you want to include more files?"
            print "0: NO"
            print "1: YES"
            action = raw_input("ENTER 0 or 1:")

        if int(action) == 1:
            self.append_files(final_data)
        else:
            self.save_to_file(final_data)
            self.prompt()

        # ask for next action to perform
        self.prompt()

    def save_to_file(self, all_data):
        """
        Saves the data frame as an excel file
        :param all_data: data frame that should be saved
        """
        print "Find file save dialogue box; maybe behind current screen!"
        file_path = tkFileDialog.asksaveasfilename(title="SAVE FILE", filetypes=[("Excel", "*.xlsx")])
        file_path += ".xlsx"
        try:
            all_data.to_excel(file_path, index=False)
        except ValueError:
            print "\nSave Canceled...Sending back to MAIN SCREEN"
            self.prompt()

        print "\nSaved Successfully at: " + file_path
        return


app = ExcelScript()
