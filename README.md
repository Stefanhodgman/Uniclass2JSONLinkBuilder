# 

TLDR; You download the NBS Uniclass tables in Excel format, Download and run this exe and point it at the downloaded and extracted Uniclass files and select Yes to skip the first two lines for all (becuase the files all have the first two lines with irrelevant data) then it will match the child codes with its parents and output JSON files for each file and one JSON with all data combined. You can then use this JSON to run the REST API for Asite.

-----------------------------------------------
This is a Python script that processes Excel files in a selected directory, extracts information from them, and saves it in JSON format. The script first creates a Tkinter root window to allow the user to select a directory using a file dialog. It then iterates over each file in the directory, checks if it is an Excel file, and processes it. For each Excel file, the script reads the data into a pandas DataFrame, extracts relevant information, and saves it to a dictionary in JSON format.

The script creates two output files for each Excel file: one file containing the processed data for that file, and one file containing the merged processed data for all files. The processed data is saved in a dictionary with parent codes as keys and child codes as values, separated by the "#" character. The script removes any trailing "#" characters from the child code lists. The merged processed data is obtained by merging the processed data for all files into a single dictionary.

The script also handles errors that may occur during the processing of the Excel files. If there is an error processing a file, the script adds the filename to a list of failed files and continues with the next file. If there are any failed files, the script shows a message box with the list of failed files.

Overall, the script is a simple and effective way to process Excel files and extract relevant information from them. It uses popular Python libraries like pandas and json to handle data and file formats, and the Tkinter library to create a simple user interface for selecting a directory.
 
