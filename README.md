# MzMl-Splitter

Requirements:
-PyQT6
-Pyside6
-PyOpenMS
-Pythoms
-pandas
-numpy
-openpyxl


This script allows a user to split an mzML file into 'timepoints', with a user defined number of scans per timepoint.

A GUI is opened upon running the script asking for a file input and output file path, the data start scan, the number of scans to average per time point, and the m/z values for a product, substrate and IS.

The input is HIGHLY RECOMMENDED to be a pre-centroided mzML file.  

Continuum mzML files can be used, but only the merged_data.csv will produce appropriate results as the other two csv's take the intensity for a single m/z value.

By default, the files that will end up in the output folder are:

-the sliced mzML files

-merged_bins.csv (user defined scans summed, m/z 1 bin width, whole spectrum data)

-extracted_intensities.csv (intensity data for user defined m/z values, data for each individual scan)

-averaged_intensities.csv (intensity data and summary statistics for user defined m/z values, averaged over a user defined number of scans)
