# Full documentation - https://github.com/aurghya-0/office-scripts/wiki/Unmerge-Sheet
import os
import pandas as pd

def unmerge_cells(directory):
    """
    Unmerge cells in all Excel files within a specified directory.

    Parameters:
    ----------
    directory : str
        The path to the directory containing the Excel files.

    Functionality:
    --------------
    - Scans the specified directory for all `.xlsx` files.
    - For each Excel file:
        - Reads the content into a Pandas DataFrame.
        - Fills any merged cells with the forward-fill method to propagate values downwards.
        - Saves the modified DataFrame back to a new Excel file in a `modified/` subdirectory.

    Notes:
    ------
    - The function assumes that the `modified/` subdirectory exists in the specified directory.
      Ensure that this subdirectory is created beforehand, or modify the function to create it if
      it doesn't exist.

    Returns:
    --------
    None
    """
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
