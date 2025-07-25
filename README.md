# Simple SharePoint

This package is designed to be a simple code interface to manage files in a SharePoint account.

## How to Install

``pip install simple_sharepoint``

## How to Use

```py
import os
from simple_sharepoint import SharePoint


sp = SharePoint(
        user=os.getenv('USERNAME'),
        password=os.getenv('PASSWORD'),
        sharepoint_url="https://example.com/sites/YourSite",
        library_title="Main Folder"
    )

sp_path = 'Folder 1/Subfolder'
# Create a directory in SharePoint
sp.create_directory(sp_path, 'New Folder')

# Upload a file to SharePoint
sp.upload_file(r'/path/to/file.txt', sp_path)

# get file objects from a directory
files = sp.get_files(sp_path)
for file in files:
    sp.download_file(file)
    sp.move_file(file, sp_path + '/Archive')
    sp.delete_file(file)

# get folder by link
folder = sp.get_folder_by_link(r'/sites/YourSite/Shared Documents/' + sp_path)
files = sp.get_files(folder)
for file in files:
    print(f'Moving file to archive: {file.properties["Name"]}')



```
