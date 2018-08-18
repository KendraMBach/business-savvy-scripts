import pandas as pd
import xlsxwriter
import xlrd
import os
import glob
import warnings
warnings.filterwarnings("always")

# filenames
import os
file_path = os.path.expanduser("~/Desktop/Test/")

# Create a Pandas Excel writer using XlsxWriter as the engine
writer = pd.ExcelWriter(os.path.expanduser("~/Desktop/Test/TestWorkbookName.xlsx"), engine='xlsxwriter')

# Attach each existing Excel document to the same workbook as separate worksheets
# Names of the files are retained as worksheet names
for path in glob.glob(os.path.expanduser("~/Desktop/Test/*.xls")):
    ds_name = str(os.path.basename(path))
    ds = pd.read_excel(path)
    ds.to_excel(writer, sheet_name=ds_name, index = False)

#Save and close writer to allow changes to new workbook to persist
writer.save()
writer.close()
