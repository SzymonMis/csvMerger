# Quick script for merging multiple csv into one excel file using https://pandas.pydata.org
# and https://pypi.org/project/XlsxWriter/

# Dependencies
# pip install pandas
# pip install XlsxWriter

import os
import pandas as pd

currentDirectory = os.path.dirname(os.path.abspath(__file__))

# Getting files relativl to merger.py localization folder
inputFiles = [file for file in os.listdir(currentDirectory) if file.endswith('.csv')]

# Output name for excel file
outputFile = 'mergedFile.xlsx' 

excelWriter = pd.ExcelWriter(outputFile, engine='xlsxwriter')

for inputFile in inputFiles:
    try:
        df = pd.read_csv(os.path.join(currentDirectory, inputFile))        
        sheet_name = os.path.splitext(inputFile)[0]        
        df.to_excel(excelWriter, sheet_name=sheet_name, index=False)

    except Exception as e:
        print(f"Error: '{inputFile}': {e}")

excelWriter.close()

print("Done.") 
