# Full documentation - https://github.com/aurghya-0/office-scripts/wiki/Unmerge-Sheet
import os
import pandas as pd

def unmerge_cells(directory):
    excel_files = []

    for file in os.listdir(directory):
        filename = os.fsdecode(file)
        if filename.endswith(".xlsx"):
            excel_files.append(filename)

    for file in excel_files:
        df = pd.read_excel(directory + file)
        df = df.fillna(method='ffill')
        df.to_excel(directory + "modified/" + file, index=False)


if __name__ == "__main__":
    excel_directory = "excel"
    unmerge_cells(excel_directory)
