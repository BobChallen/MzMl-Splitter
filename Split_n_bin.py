import os
import math
import sys
from PyQt6 import uic
from PyQt6.QtGui import *
from PyQt6.QtCore import pyqtSlot, Qt, pyqtSignal
from PyQt6.QtWidgets import *
from PyQt6.uic import loadUi
from PySide6 import QtCore, QtGui
from Splitting_UI_modv2 import Ui_MainWindow
from pyopenms import *
from pyopenms import ChromatogramExtractorAlgorithm, ChromatogramExtractor, OSChromatogram
import numpy as np
import pandas as pd
from spectrum_binner import bin_spectra
from glob import glob, iglob
from openpyxl import load_workbook, worksheet
import pandas as pd
from pythoms.scripttime import ScriptTime

class WorkerThread(QtCore.QThread):
    finished = QtCore.Signal()
    progress_update = QtCore.Signal(int)

    def __init__(self, per_sample, parent=None):
        super().__init__()
        self.scans_to_sum = per_sample
        self.parent = parent
        

    def run(self):
        self.parent.split_file(self.scans_to_sum)
        self.progress_update.emit(100)
        self.finished.emit()

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        # setup mainwindow (load UI from .UI file)
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.ui.selectStartScanSpinBox.editingFinished.connect(self.start_scan)
        self.ui.doubleSpinBox.editingFinished.connect(self.sum_scans)
        self.scans_to_sum = 10
        self.starting_scan = 1
        self.ui.pushButton.clicked.connect(self.file_finder)
        self.ui.output_button.clicked.connect(self.output_finder)
        self.ui.go_button.clicked.connect(self.scans_per_timepoint) 
        self.ui.substrate_spinbox.editingFinished.connect(self.substrate_mz)
        self.ui.product_spinbox.editingFinished.connect(self.product_mz)    
        self.ui.IS_spinbox.editingFinished.connect(self.is_mz)
        self.ui.IS_checkBox.stateChanged.connect(self.is_normalisation)        
        self.show()

    def file_finder(self):
        #User input for input file location - must be mzml currently
        file_dialog_result = QFileDialog.getOpenFileName()
        selected_file_path = file_dialog_result[0]

        if selected_file_path:
            self.ui.selectFileLineEdit.setText(selected_file_path)

    def output_finder(self):
        #user input for output file location - recommend a new folder for each processing job due to the number of files created.
        #the number of files created can be reduced in the split_file function, but not recommended as the split mzml files are useful
        self.selected_directory = QFileDialog.getExistingDirectory()

        if self.selected_directory:
            self.ui.selectOutputDirectoryLineEdit.setText(
                self.selected_directory)

    def start_scan(self):
        #user input for the starting scan from the GUI.
        #might be nice to have an image window pop-up with the chromatogram for the user to see the outline without having to go into masslynx
        self.starting_scan = self.ui.selectStartScanSpinBox.value()
        print("Processing from scan",self.starting_scan)

    def sum_scans(self):
        #takes the user input from the UI as to how many scans to sum per timepoint, and sends the value to the worker thread
        self.scans_to_sum = self.ui.doubleSpinBox.value()
        print("Will sum", self.scans_to_sum, "scans for each timepoint")

    def scans_per_timepoint(self):        
        #sends the value and progress bar management to the worker thread
        self.worker_thread = WorkerThread(per_sample=self.scans_to_sum, parent=self)
        self.worker_thread.progress_update.connect(self.update_progress_bar)
        self.worker_thread.finished.connect(self.worker_finished)
        self.worker_thread.start()
  
    def worker_finished(self):
            #Prints completion method once the worker thread has finished
            print("Processing complete!") 
            self.ui.progressBar.setRange(0,100)      

    def substrate_mz(self):
        self.SUBSTRATE_MZ = self.ui.substrate_spinbox.value()
        print("Substrate m/z:", self.SUBSTRATE_MZ)
    
    def product_mz(self):
        self.PRODUCT_MZ = self.ui.product_spinbox.value()
        print("Product m/z:", self.PRODUCT_MZ)
            
    def is_mz(self):
        self.IS_MZ = self.ui.IS_spinbox.value()
        print("Internal Standard m/z:", self.IS_MZ)

    def is_normalisation(self):
        print("test")

    def update_progress_bar(self, value):
        self.ui.progressBar.setValue(value)
    
    def extract_intensity(self, spectrum, mz_value):
        #Peak pick for closest peak to desired m/z value
        closest_peak = min(spectrum, key=lambda x: abs(x.getMZ() - mz_value))
        #return intensity of the closest peak
        return closest_peak.getIntensity()

    def split_file(self, scans_to_sum, bin_width=0.5):
        # Get the input file path from the GUI
        input_file_path = self.ui.selectFileLineEdit.text()

        #create a list to store all extracted intensities
        all_df_list = []
        # Specify the output directory for the split files
        output_directory = self.selected_directory
        os.makedirs(output_directory, exist_ok=True)

        # Load the input file into an MSExperiment object using the PyOpenMS library
        exp = MSExperiment()
        MzMLFile().load(input_file_path, exp)

        # Iterate over spectra in the input file
        offset_value = 0
        current_spectrum_index = self.starting_scan - offset_value
        current_output_file_index = 1
        value = 0
        total_files = len(exp.getSpectra()) // int(scans_to_sum)
        value = self.ui.progressBar.value()
        while current_spectrum_index < len(exp.getSpectra()):
            # Create a new MSExperiment object for the specified number of scans
            new_exp = MSExperiment()

            # Specify the range of spectra to include in the current file
            start_spectrum_index = current_spectrum_index
            end_spectrum_index = min(
                current_spectrum_index + int(scans_to_sum), len(exp.getSpectra()))
            print("Start Spectrum Index:", start_spectrum_index)
            print("End Spectrum Index:", end_spectrum_index)

            #create a list with user defined m/z values to look for
            #add an internal standard if the 'IS used' box on the GUI is checked
            if self.ui.IS_checkBox.isChecked:
                mz_values_of_interest = [self.SUBSTRATE_MZ, self.PRODUCT_MZ, self.IS_MZ]
            else:
                mz_values_of_interest = [self.SUBSTRATE_MZ, self.PRODUCT_MZ]

            #create a dataframe to store the intensity values
            df_list = []

            for i in range(start_spectrum_index, end_spectrum_index):
                # Get the current spectrum
                current_spectrum = exp.getSpectra()[i]

                'Extract the intensity values for the desired peaks'
                intensity_values = []
                for mz_value in mz_values_of_interest:
                    intensity = self.extract_intensity(current_spectrum, mz_value)
                    intensity_values.append(intensity)
                
                #Extract scan information
                scan_number = i +1
                scan_df = pd.DataFrame([{'Scan': scan_number, **dict(zip(mz_values_of_interest, intensity_values))}])
                df_list.append(scan_df)
                # Add the current spectrum to the new MSExperiment
                new_exp.addSpectrum(current_spectrum)

                # Create a new MSChromatogram for the current spectrum
                new_chromatogram = MSChromatogram()

                # Get the total intensity for the specified range of scans
                total_intensity = sum(peak.getIntensity()
                                      for peak in current_spectrum)

                # Create a ChromatogramPeak with total intensity
                chromatogram_peak = ChromatogramPeak()
                chromatogram_peak.setRT(i)  # Using the scan number as RT
                chromatogram_peak.setIntensity(total_intensity)

                # Add the ChromatogramPeak to the new MSChromatogram
                new_chromatogram.push_back(chromatogram_peak)

                # Add the new MSChromatogram to the new MSExperiment
                new_exp.addChromatogram(new_chromatogram)

            # Store the new MSExperiment object in a new output file with scan range and index
            first_spectrum_num = start_spectrum_index + 1
            last_spectrum_num = end_spectrum_index
            output_file_name = os.path.join(output_directory, (os.path.basename(
                input_file_path) + f"scan_{first_spectrum_num}_to_{last_spectrum_num}_timepoint_{current_output_file_index}.mzML"))
            MzMLFile().store(output_file_name, new_exp)
            value = int(((current_output_file_index/total_files)*100)//2)
            self.worker_thread.progress_update.emit(value)
           
           # Increment the indices
            current_spectrum_index = end_spectrum_index
            current_output_file_index += 1
           
            #append extracted m/z intensity list to master list
            all_df_list.append(df_list)
        
        #concat all intensity lists and export to .xlsx file
        intensity_df = pd.concat([pd.concat(df_list) for df_list in all_df_list], ignore_index=True)

        #normalise the intensity of the product and substrate by the intensity of the internal standard
        intensity_df['NORM_SUBSTRATE_MZ'] = (intensity_df[self.SUBSTRATE_MZ]/intensity_df[self.IS_MZ])
        intensity_df['NORM_PRODUCT_MZ'] = (intensity_df[self.PRODUCT_MZ]/intensity_df[self.IS_MZ])
        
        #calculate the conversion % of substrate to product
        intensity_df['Conversion%'] = (intensity_df['NORM_PRODUCT_MZ']/(intensity_df['NORM_SUBSTRATE_MZ']+intensity_df['NORM_PRODUCT_MZ'])*100)

        #create a new dataframe with averaged data to smooth the data
        averaged_df_list = []
        averaged_df = pd.DataFrame(columns=intensity_df.columns)
        

        # #iterate over averaged data to calculate mean 
        grouped_df = intensity_df.groupby(np.arange(len(intensity_df))//self.scans_to_sum).mean()
        averaged_df_list.append(grouped_df)

        #concat the averaged list into one dataframe
        averaged_df = pd.concat(averaged_df_list, ignore_index=True)
        averaged_df['Conversion% Standard Deviation'] = intensity_df.groupby(np.arange(len(intensity_df)) // self.scans_to_sum)['Conversion%'].std()
        
        #rename the columns for averaged and extracted intenisites csv's to make the ouput easier to read
        averaged_df2 = averaged_df.rename({
            self.SUBSTRATE_MZ: 'Substrate Intensity', 
            self.PRODUCT_MZ: 'Product Intensity', 
            'NORM_SUBSTRATE_MZ': 'Normalised Substrate Intensity', 
            'NORM_PRODUCT_MZ':'Normalised Product Intensity',
            self.IS_MZ:'Internal Standard Intensity',
            }, 
            axis='columns')
        intensity_df2 = averaged_df.rename({
            self.SUBSTRATE_MZ: 'Substrate Intensity', 
            self.PRODUCT_MZ: 'Product Intensity', 
            'NORM_SUBSTRATE_MZ': 'Normalised Substrate Intensity', 
            'NORM_PRODUCT_MZ':'Normalised Product Intensity',
            self.IS_MZ:'Internal Standard Intensity',
            }, 
            axis='columns')
        averaged_filepath = os.path.join(output_directory, "averaged intensities.csv")
        averaged_df2.to_csv(averaged_filepath, index=False)

       
        #determine output directory for the extracted csv (same as user selected output path)
        extracted_filepath = os.path.join(output_directory, "extracted intensities.csv")
       
        #create new .csv file` `
        intensity_df2.to_csv(extracted_filepath, index=False)
       
        #run through list of created mzml files and run the bin_spectra script on them
        for filename in os.listdir(output_directory):
            if filename.endswith(".mzML"):
                file_path = os.path.join(output_directory, filename)
                bin_spectra(file_path)
                value = int(((current_output_file_index/total_files*100)//2))
                self.worker_thread.progress_update.emit(value)

        # Sort the list alphabetically so that the final output is in ascending order
        excel_files = glob(os.path.join(output_directory, "*.xlsx"))
        for file in excel_files:
            if '~$' in file:
                continue #bug fix for temp files causing the script to error out due to being unable to access the temp file
            else:
                print("Normalising summed intensity data of", file)
               
                # load to pandas data frame
                workbook = pd.read_excel(file)
                
                # normalise second column (counts) by dividing value by max value
                df_max_scaled = workbook.copy()
               
                # df_max_scaled['counts'] = df_max_scaled['counts'] / df_max_scaled['counts'].max()
                # rename counts column to filename
                df_max_scaled.rename(columns={'counts': file}, inplace=True)
               
                # save dataframe back to excel
                df_max_scaled.to_excel(file + "normalised.xlsx", index=False)
       
        # Load only excel files with normalised data into variable
        norm_list = glob(output_directory + "/*normalised.xlsx")
       
        # Tranpose the data in each excel file (row 1 = m/z, row 2 = intensity)
        st = ScriptTime()
        st.printstart()
        for file in norm_list:
            print("Transposing data of", file)
           
            # load workbook to dataframe
            df = pd.read_excel(file, index_col=0)
           
            # transpose data
            df_transposed = df.transpose()
            
            # save the updated workbook to a new file
            df_transposed.to_excel(file + "transposed_output.xlsx")
       
        # load only excel files with transposed data into variable
        merged_list = glob(output_directory + "/*transposed_output.xlsx")
        df = pd.DataFrame()

        print("Merging data to single document")
        df = pd.concat(pd.read_excel(file) for file in merged_list)
        df['sort_column'] = df[df.columns[0]].str.extract(
            r'(\d+)\.mzML\.xlsx').astype(int)
        # Sort the DataFrame based on the new column
        df_sorted = df.sort_values(by='sort_column')
        # Drop the temporary sort column
        df_sorted = df_sorted.drop(columns=['sort_column'])
        # merge together all created transposed files
        df.head()
        # save merged file
        df_sorted.to_csv(output_directory + '/merged_bins.csv', index=False)
        # gather all now unnecessary file paths into a list
        print("Deleting unnecessary files")
        # delete all unnecessary files (i.e intermediary excel files, mzML files)
        # alter lines as necessary to customise which files to keep
        files_to_delete = glob(output_directory + "/*.mzml.xlsx") + glob(output_directory + "/*.mzml.xlsxnormalised.xlsx") + \
            glob(output_directory + "/*.mzml.xlsxnormalised.xlsxtransposed_output.xlsx") + \
            glob(output_directory + "/*.mzML.gz")

        # iterate through all useless excel files + mzML files, so only the merged binned data is left
        iterator = iter(files_to_delete)
        try:
            while True:
                element = next(iterator)
                os.remove(element)
        except StopIteration:
            pass
        self.worker_thread.progress_update.emit(100)
        print("Done!")
        st.printend()

        print(f"Total output files created: {current_output_file_index - 1}")
        self.worker_thread.finished.emit()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    GUI = MainWindow()
    sys.exit(app.exec())
