# Full documentation - https://github.com/aurghya-0/office-scripts/wiki/Download-and-Rename
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
import pandas as pd


def authenticate_drive():
    """
    Authenticate with Google Drive using OAuth and return a GoogleDrive object.
    
    Returns:
    -------
    GoogleDrive
        A GoogleDrive object that can be used to interact with Google Drive.
    
    Notes:
    -----
    This function uses the pydrive library to authenticate with Google Drive.
    It uses the LocalWebserverAuth method to authenticate, which will open a web browser
    to authenticate the user.
    """
    gauth = GoogleAuth()
    gauth.LocalWebserverAuth()
    drive = GoogleDrive(gauth)
    return drive


def download_and_rename(index, file_id, new_name):
    """
    Download a file from Google Drive and rename it.
    
    Parameters:
    ----------
    index : int
        The index of the file being processed.
    file_id : str
        The ID of the file to download from Google Drive.
    new_name : str
        The new name to give the file.
    
    Notes:
    -----
    This function downloads the file from Google Drive using the pydrive library,
    and saves it to the local file system with the new name.
    
    Raises:
    ------
    Exception
        If an error occurs while downloading or renaming the file.
    """
    try:
        drive = authenticate_drive()
        file = drive.CreateFile({'id': file_id})
        file_name = file['title']
        print(f"{index + 1}. Currently formatting - {file_name}")
        file_extension = file_name.split('.')[-1]
        file.GetContentFile(f"CC_CA1/{new_name}.{file_extension}")
    except:
        print("Error occured! Trying next value")


def process_csv(csv_file, identifier="Roll", name="Name", link="Link", provide_name=False):
    """
    Process a CSV file to download and rename files from Google Drive from a Google Forms Response file.

    Parameters:
    ----------
    csv_file : str
        The path to the CSV file that contains the data.
        
    identifier : str, optional
        The name of the column in the CSV file that contains the unique identifier for each file.
        Default is "Roll".
        
    name : str, optional
        The name of the column in the CSV file that contains the name of the person or entity 
        associated with the file. Default is "Name".
        
    link : str, optional
        The name of the column in the CSV file that contains the Google Drive link to the file.
        Default is "Link".
        
    provide_name : bool, optional
        A flag that determines whether the `name` column value should be included in the new file name.
        Default is False.

    Functionality:
    --------------
    - Reads the specified CSV file into a Pandas DataFrame.
    - Iterates over each row of the DataFrame.
    - Generates a new file name based on the identifier and optionally the name.
    - Extracts the Google Drive file ID from the link column.
    - Calls the `download_and_rename` function to download and save the file with the new name.

    Returns:
    --------
    None
    """
    df = pd.read_csv(csv_file)
    for index, row in df.iterrows():
        if provide_name:
            new_name = str(row[identifier]) + " - " + row[name]
        else:
            new_name = str(row[identifier])
        file_link = row[link]
        file_id = file_link.split('=')[-1]
        download_and_rename(index, file_id, new_name)


if __name__ == "__main__":
    csv_file_path = "CC_CA1.csv"
    process_csv(csv_file_path)
