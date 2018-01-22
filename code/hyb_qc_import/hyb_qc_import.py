"""
v1.0 - AB 2018/01/18

Requirements:
    ODBC connection to Moka
    Python 2.7
    pyodbc

usage: feature_extraction.py [-h] -i INPUT_FOLDER

Import QC data into Moka from Array feature extraction files

optional arguments:
  -h, --help            show this help message and exit
  -i INPUT_FOLDER, --input_folder INPUT_FOLDER
                        folder containing feature extraction files to be
                        imported
"""

import argparse
import os
import shutil
import Tkinter
import tkMessageBox
import pyodbc

def check_directory(dir_path):
    """
    Validates that filepath provided as command line argument exists. Raises argparse.ArgumentTypeError if it doesn't.
    """
    # If directory exists, return the path
    if os.path.isdir(dir_path):
        return dir_path
    # If directory doesn't exist, raise argparse.ArgumentTypeError
    else:
        raise argparse.ArgumentTypeError('{dir_path} does not exist'.format(dir_path=dir_path))

def process_arguments():
    """
    Uses argparse module to define and handle command line input arguments and help menu
    """
    # Create ArgumentParser object. Description message will be displayed as part of help message if script is run with -h flag
    parser = argparse.ArgumentParser(description='Import QC data into Moka from Array feature extraction files')
    # Define the arguments that will be taken.
    # This script takes one required argument, a directory containing fe files. Execute the check_directory function to validate that the supplied folder exists 
    parser.add_argument('-i', '--input_folder', required=True, type=check_directory, help='folder containing feature extraction files to be imported')
    # Return the arguments
    return parser.parse_args()

class FEParser(object):
    """
    Parses all fe_files in a directory and inserts the QC records into Moka. Takes a directory as argument.
    """
    def __init__(self, directory):
        # Store user supplied directory
        self.directory = directory
        # List to store fe_filepaths
        self.fe_filepaths = []
        # List to store QC metric dictionaries for each fe_file
        self.all_QC_metrics = []
        # Create connection to Moka
        self.cnxn = pyodbc.connect('DRIVER={SQL Server}; SERVER=GSTTV-MOKA; DATABASE=mokadata;', autocommit=True)
        # Create cursor for executing queries
        self.cursor = self.cnxn.cursor()

    def get_fe_filepaths(self):
        """
        Loops through folder, identifies fe_files and adds to list.
        """
        # Loop through file names (not paths) in the directory
        for fe_file in os.listdir(self.directory):
            # fe files should have .txt extension and should have 9 fields in the filename, spearated by underscores. 
            # If file meets these criteria, add full filepath to self.fe_filepaths list
            if fe_file.endswith('.txt') and len(fe_file.split('_')) == 9:
                # Store full path, use os.path.join to concatenate directory with filename
                self.fe_filepaths.append(os.path.join(self.directory, fe_file))

    def get_subarray_number(self, group, id):
        """
        Returns subarray number (1-8) from the 2 number code used in filename
        """
        # Subarrays split into two groups (1 & 2), with 4 subarrays in each group (1 - 4)
        # e.g. subarray 1 = group 1, id 1; subarray 8 = group 2, id 4 
        if group == 2:
            # If the group is 2, add 4 to the id to get the subarray number and return it
            return 4 + id
        elif group == 1:
            # Else if the the group is 1, subarray number is same as id, so return it
            return id

    def parse_files(self):
        """
        Extracts QC data from fe_files and stores in dictionary 
        """
        # Call self.get_fe_filepaths method to generate a list of filepaths for all fe files in the directory
        self.get_fe_filepaths()
        # Loop through the fe filepaths in self.fe_filepaths list
        for fe_filepath in self.fe_filepaths:
            # Create an empty dictionary to hold QC data for this file
            fe_file_dict = {}
            # Parse file to extract required information
            with open(fe_filepath, 'r') as fe_file:
                # Files are large so read line by line using enumerate
                for line_num, line in enumerate(fe_file):
                    # enumerate returns 0-based line numbers, so need to subtract 1 from 
                    if line_num == 1:
                        # Capture the headers from line 2
                        headers_line2 = line.split('\t')
                    elif line_num == 2:
                        # Capture barcode/subbarray field, (which contains the barcode and subarray number in format: barcode_group_id 
                        # e.g. subarray 8 for barcode 12345678 would be: 12345678_2_4)
                        # Use line 2 headers to find correct index for 'FeatureExtractor_Barcode' field
                        # Split on underscores to separate barcode, subarray group and subarray id
                        barcode_subarray = line.split('\t')[headers_line2.index('FeatureExtractor_Barcode')].split('_')
                        # Capture barcode and subarray id
                        barcode = long(barcode_subarray[0])
                        subarray = self.get_subarray_number(group=int(barcode_subarray[-2]), id=int(barcode_subarray[-1]))
                        # Store barcode and subarray number in dictionary
                        fe_file_dict['barcode'] = barcode
                        fe_file_dict['subarray'] = subarray
                    elif line_num == 5:
                        # Capture the headers from line 6
                        headers_line6 = line.split('\t')
                    elif line_num == 6:
                        # Capture the data for each of the QC fields. Use headers from line 6 row to get correct indexes.
                        # Store values in fe_file_dict. NU% and DLRS fields are rounded to 2dp, all other fields are rounded to integers
                        # Move to next line containing the QC data and split on tabs
                        data = line.split('\t')
                        fe_file_dict['nu'] = round(float(data[headers_line6.index('Metric_AnyColorPrcntFeatNonUnifOL')]), 2)
                        fe_file_dict['dlrs'] = round(float(data[headers_line6.index('Metric_DerivativeLR_Spread')]), 2)
                        fe_file_dict['cy3_s2n'] = int(round(float(data[headers_line6.index('Metric_g_Signal2Noise')])))
                        fe_file_dict['cy3_si'] = int(round(float(data[headers_line6.index('Metric_g_SignalIntensity')])))
                        fe_file_dict['cy5_s2n'] = int(round(float(data[headers_line6.index('Metric_r_Signal2Noise')])))
                        fe_file_dict['cy5_si'] = int(round(float(data[headers_line6.index('Metric_r_SignalIntensity')])))
                        # Add the dictionary to the all_QC_metrics list
                        self.all_QC_metrics.append(fe_file_dict)
                        # No need to read rest of file so break out of for loop
                        break

    def update_moka(self):
        """
        Updates Moka with QC data extracted from fe files.
        """
        # For each set of data in the all_QC_metrics list, use the barcode and subarray to pull out the DNALabellingID from Moka
        for fe_file_dict in self.all_QC_metrics:
            # SQL to get the DNALabellingID 
            dna_labelling_id_sql = (
                "SELECT ArrayLabelling.DNALabellingID FROM ArrayLabelling INNER JOIN Arrays ON ArrayLabelling.ArrayID = Arrays.ArrayID "
                "WHERE Arrays.ArrayBarCode = '{barcode}' and ArrayLabelling.Subarray = '{subarray}';".format(**fe_file_dict)
                )
            # Execute the SQL and store results in the fe_file_dict dictionary
            fe_file_dict['DNALabellingID'] = self.cursor.execute(dna_labelling_id_sql).fetchone().DNALabellingID
        # For each set of data in the all_QC_metrics list, update the ArrayLabelling table with the QC metrics from fe_file
        for fe_file_dict in self.all_QC_metrics:
            # SQL to update the ArrayLabelling table with QC data
            update_sql = (
                "UPDATE ArrayLabelling SET DSLR = {dlrs}, PercentNonUnifFeat = {nu}, g_Signal2Noise = {cy3_s2n}, g_SignalIntensity = {cy3_si}, r_Signal2Noise = {cy5_s2n}, r_SignalIntensity = {cy5_si} "
                "WHERE DNALabellingID = {DNALabellingID};".format(**fe_file_dict)
            )
            # Execute the SQL to update ArrayLabelling table with the QC metrics from fe file
            self.cursor.execute(update_sql)
    
    def move_to_subfolder(self):
        """
        Moves imported fe_files to a subfolder
        """
        # Build path for destination folder
        dest_folder = '{parent_directory}\imported_to_moka'.format(parent_directory = self.directory)
        # Create folder to move imported files to, if it doesnt already exist
        if not os.path.isdir(dest_folder):
            os.mkdir(dest_folder)
        # Move all files in fe_filepaths list to the subfolder
        # os.path.basename gets the filename from the full source filepath.
        [shutil.move(fe_filepath, os.path.join(dest_folder, os.path.basename(fe_filepath))) for fe_filepath in self.fe_filepaths]
    
    def dispay_complete_message(self):
        """
        Displays messagebox telling user the script is complete and states the number of files that were imported.
        """
        # Remove the root window so that only the messagebox is displayed
        Tkinter.Tk().withdraw()
        # Display the messagebox
        tkMessageBox.showinfo('Complete', 'QC data from {num_files} feature extraction files imported into Moka'.format(num_files=len(self.fe_filepaths)))

def main():
    # Create FEParser obeject, passing in the user supplied directory from command line
    f = FEParser(process_arguments().input_folder)
    # Extract required information from fe files in the directory
    f.parse_files()
    # Add the extracted QC data into Moka 
    f.update_moka()
    # Move fe_files into subdirectory to prevent them being reimported
    f.move_to_subfolder()
    # Display a message stating 
    f.dispay_complete_message()

if __name__ == '__main__':
    main()
