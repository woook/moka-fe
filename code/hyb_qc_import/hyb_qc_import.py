"""
v1.2 - AB 2019/09/11

Requirements:
    ODBC connection to Moka
    Python 2.7
    pyodbc

usage: hyb_qc_import.py [-h] -i INPUT_FOLDER --hyb HYB_RUN_NUMBER

Import QC data into Moka from Array feature extraction files

optional arguments:
  -h, --help            show this help message and exit
  -i INPUT_FOLDER, --input_folder INPUT_FOLDER
                        folder containing feature extraction files to be
                        imported
  --hyb HYB_RUN_NUMBER  the hyb run number for the batch to be imported
"""

import argparse
import os
import Tkinter
import tkMessageBox
import pyodbc
from ConfigParser import ConfigParser

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
    # The first required argument is the directory containing the fe files. Execute the check_directory function to validate that the supplied folder exists 
    parser.add_argument('-i', '--input_folder', required=True, type=check_directory, help='folder containing feature extraction files to be imported')
    # The second required argument is the Hyb run number. This will be used to identify which fe files to import data from.
    parser.add_argument('--hyb', dest='hyb_run_number', required=True, help='the hyb run number for the batch to be imported')
    # Return the arguments
    return parser.parse_args()

class FEParser(object):
    """
    Parses all fe_files in a directory and inserts the QC records into Moka. Takes a directory and hyb ID as arguments.
    """
    def __init__(self, moka_server, moka_db_name, directory, hyb_run):
        # Store user supplied directory
        self.directory = directory
        # Store hyb run number
        self.hyb_run = hyb_run
        # Subarray lookup dictionary. Convert 1 digit subarray number to 2 digit representation used in fe filename.
        self.subarray_lookup = {
            1: '1_1', 
            2: '1_2', 
            3: '1_3', 
            4: '1_4', 
            5: '2_1', 
            6: '2_2', 
            7: '2_3', 
            8: '2_4'
            }
        # Create the template filepath for fe files
        self.fe_filepath_template = os.path.join(directory, '{barcode}_S01_Guys121919_CGH_1100_Jul11_2_{subarray}.txt')
        # List to store dictionaries containing info from Moka for every hyb in the hyb run
        self.moka_details = []
        # List to store QC metric dictionaries for each fe_file found for the run
        self.all_QC_metrics = []
        # Create connection to Moka
        self.cnxn = pyodbc.connect('DRIVER={{SQL Server}}; SERVER={moka_server}; DATABASE={moka_db_name};'.format(
                    moka_server=moka_server,
                    moka_db_name=moka_db_name
                ),
            autocommit=True
            )
        # Create cursor for executing queries
        self.cursor = self.cnxn.cursor()

    def moka_lookup(self):
        """
        Retrieves records from Moka for user supplied HybID
        """
        # SQL to retrieve DNALabellingID, Barcode and Subarray number for each pair in run
        moka_select_sql = (
            "SELECT ArrayLabelling.DNALabellingID, ArrayLabelling.Subarray, Arrays.ArrayBarCode FROM ArrayLabelling INNER JOIN Arrays ON ArrayLabelling.ArrayID = Arrays.ArrayID "
            "WHERE ArrayLabelling.ArrayRunNumber = '{hyb_run}'".format(hyb_run=self.hyb_run)
            )
        # Execute query and store results
        records = self.cursor.execute(moka_select_sql).fetchall()
        # Add query results to dictionary and store dictionary in all_QC_metrics list
        for record in records:
            self.moka_details.append({
                'DNALabellingID': int(record.DNALabellingID),
                'subarray': int(record.Subarray),
                'barcode': record.ArrayBarCode
            })

    def get_fe_filepaths(self):
        """
        Loops through moka_details dictionary, constructs expected filepath and, if it exists, adds filepath to dictionary.
        """
        # Loop through each hyb in the moka_details dictionary 
        for hyb in self.moka_details:
            # Construct the expected filepath by adding the barcode and subarray to the fe_filepath_template attribute.
            # The subarray must be converted to the 2 digit representation using the subarray_lookup dictionary
            expected_filepath = self.fe_filepath_template.format(barcode=hyb['barcode'], subarray=self.subarray_lookup[hyb['subarray']])
            # Use basename to extract the filename from the filepath
            filename = os.path.basename(expected_filepath)
            # If the file exists, store the filepath in the dictionary and add the dictionary to all_QC_metrics list
            if os.path.isfile(expected_filepath):
                hyb['filepath'] = expected_filepath
                self.all_QC_metrics.append(hyb)


    def confirm_proceed(self):
        """
        Get user confirmation to proceed. Returns True if they say yes, False if they say no.
        """
        # Remove the root window so that only the messagebox is displayed
        Tkinter.Tk().withdraw()
        # Display the messagebox. Returns True if user clicks Yes (to proceed), returns False if user clicks No.
        return tkMessageBox.askyesno(
            'Import Array QC Metrics', '{num_files} feature extraction files found for run {hyb_run}.\n\nDo you want to import QC metrics?'.format(
                hyb_run=self.hyb_run,
                num_files=len(self.all_QC_metrics)
                )
            )


    def parse_files(self):
        """
        Extracts QC data from fe_files and stores in dictionary 
        """
        # Loop through the hybs in all_QC_metrics dictionary
        for hyb in self.all_QC_metrics:
            # Parse file to extract required information
            with open(hyb['filepath'], 'r') as fe_file:
                # Files are large so read line by line using enumerate
                for line_num, line in enumerate(fe_file):
                    # enumerate returns 0-based line numbers, so need to subtract 1 from actual line number in file
                    if line_num == 1:
                        # Capture the headers from line 2
                        headers_line2 = line.split('\t')
                    elif line_num == 2:
                        # Capture barcode/subarray field, which contains the barcode and 2 digit representation of subarray separated by underscores 
                        # e.g. subarray 8 for barcode 12345678 would be: 12345678_2_4)
                        # Use line 2 headers to find correct index for 'FeatureExtractor_Barcode' field
                        # Split on underscores to separate barcode, subarray group and subarray id
                        barcode_subarray = line.split('\t')[headers_line2.index('FeatureExtractor_Barcode')].split('_')
                        # Capture barcode and subarray id
                        barcode = barcode_subarray[0]
                        subarray = barcode_subarray[1] + '_' + barcode_subarray[2]
                        # Check that the barcode and subarray within the file match those in the filename.
                        # The subarray stored int the dictionary will be in single digit format, so use subarray_lookup dict to convert to 2 digit representation 
                        assert barcode == hyb['barcode'] and subarray == self.subarray_lookup[hyb['subarray']], 'Barcode/subarray in file does not match filename'
                    elif line_num == 5:
                        # Capture the headers from line 6
                        headers_line6 = line.split('\t')
                    elif line_num == 6:
                        # Capture the data for each of the QC fields. Use headers from line 6 row to get correct indexes.
                        # Store values in hyb. NU% and DLRS fields are rounded to 2dp, all other fields are rounded to integers
                        # Move to next line containing the QC data and split on tabs
                        data = line.split('\t')
                        hyb['nu'] = round(float(data[headers_line6.index('Metric_AnyColorPrcntFeatNonUnifOL')]), 2)
                        hyb['dlrs'] = round(float(data[headers_line6.index('Metric_DerivativeLR_Spread')]), 2)
                        hyb['cy3_s2n'] = int(round(float(data[headers_line6.index('Metric_g_Signal2Noise')])))
                        hyb['cy3_si'] = int(round(float(data[headers_line6.index('Metric_g_SignalIntensity')])))
                        hyb['cy5_s2n'] = int(round(float(data[headers_line6.index('Metric_r_Signal2Noise')])))
                        hyb['cy5_si'] = int(round(float(data[headers_line6.index('Metric_r_SignalIntensity')])))
                        # No need to read rest of file so break out of for loop
                        break

    def update_moka(self):
        """
        Updates Moka with QC data extracted from fe files.
        """
        # For each set of data in the all_QC_metrics list, update the ArrayLabelling table with the QC metrics from fe_file
        for hyb in self.all_QC_metrics:
            # SQL to update the ArrayLabelling table with QC data. Populate the fields using the hyb dictionary
            update_sql = (
                "UPDATE ArrayLabelling SET DSLR = {dlrs}, PercentNonUnifFeat = {nu}, g_Signal2Noise = {cy3_s2n}, g_SignalIntensity = {cy3_si}, r_Signal2Noise = {cy5_s2n}, r_SignalIntensity = {cy5_si} "
                "WHERE DNALabellingID = {DNALabellingID};".format(**hyb)
            )
            # Execute the SQL to update ArrayLabelling table with the QC metrics from fe file
            self.cursor.execute(update_sql)
    
    def dispay_complete_message(self):
        """
        Displays messagebox telling user the script is complete.
        """
        # Remove the root window so that only the messagebox is displayed
        Tkinter.Tk().withdraw()
        # Display the messagebox
        tkMessageBox.showinfo('Complete', 'Import complete')

def main():
    # Get command line arguments
    args = process_arguments()
    # Read config file
    config = ConfigParser()
    config.read(os.path.join(os.path.dirname(os.path.realpath(__file__)), "config.ini"))
    # Create FEParser object
    f = FEParser(config.get("MOKA", "SERVER"), config.get("MOKA", "DATABASE"), args.input_folder, args.hyb_run_number)
    # Find each hyb in Moka for the run and store details in dictionary
    f.moka_lookup()
    # Check if the files exist and store the filepaths in dictionary
    f.get_fe_filepaths()
    # Display files found and not found to user, and ask if they're happy to proceed. Function will return True if they click Yes.
    if f.confirm_proceed():
        # Extract required information from fe files in the directory
        f.parse_files()
        # Add the extracted QC data into Moka 
        f.update_moka()
        # Display a complete message stating number of files imported 
        f.dispay_complete_message()

if __name__ == '__main__':
    main()
