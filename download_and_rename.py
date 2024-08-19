# Full documentation - https://github.com/aurghya-0/office-scripts/wiki/Download-and-Rename
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
import pandas as pd


def authenticate_drive():
    gauth = GoogleAuth()
    gauth.LocalWebserverAuth()
    drive = GoogleDrive(gauth)
    return drive


def download_and_rename(index, file_id, new_name):
    try:
        file = drive.CreateFile({'id': file_id})
        file_name = file['title']
        print(f"{index + 1}. Currently formatting - {file_name}")
        file_extension = file_name.split('.')[-1]
        file.GetContentFile(f"CC_CA1/{new_name}.{file_extension}")
    except:
        print("Error occured! Trying next value")


def process_csv(csv_file):
    df = pd.read_csv(csv_file)
    for index, row in df.iterrows():
        new_name = str(row['Roll']) + " - " + row['Name']
        file_link = row['Link']
        file_id = file_link.split('=')[-1]
        download_and_rename(index, file_id, new_name)


if __name__ == "__main__":
    drive = authenticate_drive()
    csv_file_path = "CC_CA1.csv"
    process_csv(csv_file_path)
