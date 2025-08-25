# Simple SharePoint

This package is designed to be a simple code interface to manage files in a SharePoint account.

## How to Install

``pip install py_simple_sharepoint``

## How to Use

```py
import os
from py_simple_sharepoint import SharePointClient

sp = SharePointClient(
    tenant_id=os.getenv("SHAREPOINT_CLIENT_TENANT_ID"),
    client_id=os.getenv("SHAREPOINT_AZURE_APPLICATION_ID"),
    cert_path=r"Certificate.cer",
    key_path=r"Certificate.key",
    site_hostname="<sitename>.sharepoint.com",
    site_path='/sites/<location>',
    library_title='Documents'
)

files = sp.get_files(files_path)
sp.create_folder(folder_path='Path/to/New Folder')
sp.upload_file(local_path='file.txt', target_folder='Path/to/Upload Folder')
sp.download_file(file_path='Path/to/file.txt', download_dir='local/path/to/folder')
sp.move_file(file_path='path/to/file.txt', target_folder='other/path/to/file')
sp.rename_file(file_path='path/to/file.txt', new_name='path/to/new file name.txt')

```
