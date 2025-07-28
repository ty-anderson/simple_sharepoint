import os
import time
from pathlib import Path
from io import BytesIO
from collections.abc import Generator
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.files.file import File, Folder


class SharePoint:
    def __init__(self,
                 user: str,
                 password: str,
                 sharepoint_url: str,
                 library_title: str):
        """
        Create a SharePoint instance. This allows the user to more easily interact with files in SharePoint.

        :param user: email@email.com
        :param password: password
        :param sharepoint_url: "https://example.com/sites/FolderName"
        :param library_title: "Main Folder"
        """
        attempt_counter = 6
        while True:
            try:
                self.sharepoint_url = sharepoint_url
                self.ctx = ClientContext(self.sharepoint_url).with_credentials(
                    UserCredential(user, password))
                library = self.ctx.web.lists.get_by_title(library_title)
                self.root_folder = library.root_folder
                self.ctx.load(self.root_folder.folders)
                self.ctx.execute_query()
                print('Connected to SharePoint')
                break
            except Exception as e:
                print('Error connecting to SharePoint')
                time.sleep(1)
                attempt_counter -= 1
                if attempt_counter == 0:
                    raise Exception(f"Could not connect to SharePoint after 6 attempts. Error: {e}")

    def create_directory(self, existing_dir_path: str | Path, new_dir_name: str) -> Folder:
        """
        Create directory in SharePoint.

        :param existing_dir_path: Path or str, list of folder names to get to folder where new folder will be created
        :param new_dir_name: str, name of folder to create
        :return: Folder, target folder
        """
        if isinstance(existing_dir_path, str):
            existing_dir_path = Path(existing_dir_path)

        directory = self.get_directory(existing_dir_path)
        directory.folders.add(new_dir_name)
        self.ctx.execute_query()
        print(f'Created folder {new_dir_name} in {"/".join(existing_dir_path.parts)}')

        path_to_target = existing_dir_path / new_dir_name
        directory = self.get_directory(path_to_target)

        return directory

    def _get_or_create_subfolder(self, parent: Folder, name: str) -> Folder:
        """
        Return a subfolder called `name` under `parent`.
        If it doesn’t exist, create it.

        Needs additional testing but could replace create_directory and get_directory
        """
        # Load child folders just once
        self.ctx.load(parent.folders)
        self.ctx.execute_query()

        for f in parent.folders:  # check if it already exists
            if f.properties["Name"].lower() == name.lower():
                return f

        # Otherwise create it
        new_folder = parent.folders.add(name)
        self.ctx.execute_query()
        return new_folder

    def get_directory(self, path_to_target: str | Path):
        """
        Point context to target directory in SharePoint.

        :param path_to_target: list, list of folder names to get to target folder
        :return: Folder, target folder or None if not found
        """
        if isinstance(path_to_target, str):
            path_to_target = Path(path_to_target)

        starting_folder = self.root_folder
        found = False
        for folder_name in path_to_target.parts:
            folders = starting_folder.folders
            self.ctx.load(folders)
            self.ctx.execute_query()

            found = False
            for folder in folders:
                if folder.properties["Name"] == folder_name:
                    found = True
                    starting_folder = folder
                    break

        if found:
            return starting_folder
        else:
            raise Exception(
                f"Folder {'/'.join(path_to_target.parts)} not found. Your account may not have access to this folder.")

    def get_folder_by_link(self, folder_link: str):
        """
        Get a folder by its SharePoint link, from the root folder.

        Example:

            sp = SharePoint(sharepoint_url="https://base_url/sites/ShareFiles", library_title="Main Folder")
            folder = sp.get_folder_by_link("/sites/ShareFiles/Main Folder/Compliance Team/Target Dir")

        """
        if isinstance(folder_link, Path):
            folder_link = folder_link.as_posix()

        folder_ = self.ctx.web.get_folder_by_server_relative_url(folder_link)
        self.ctx.load(folder_)
        self.ctx.execute_query()
        print(f"Accessed folder: {folder_.properties['Name']}")
        return folder_

    def get_files(self, path_to_files: str | Path | Folder) -> Generator[File]:
        """
        Generator for files in a SharePoint folder.

        :param path_to_files: Path, path to target folder.

        Example:

            folder = sp.get_folder_by_link("/sites/ShareFiles/Main Folder")
            for f in get_folder_files(folder):
                print(f.properties['Name'])
        """
        if isinstance(path_to_files, str):
            path_to_files = Path(path_to_files)
        if isinstance(path_to_files, Path):
            directory = self.get_directory(path_to_files)
            files_ = directory.files
        else: # Folder
            files_ = path_to_files.files
        self.ctx.load(files_)
        self.ctx.execute_query()
        for f in files_:
            yield f

    def get_all_files(self, files_path: str | Path) -> list[File]:
        """
        Get a list of all file objects from a SharePoint directory.

        :param files_path: Path, path to target folder
        """
        if isinstance(files_path, str):
            files_path = Path(files_path)

        file_obj_list = []
        for file_ in self.get_files(files_path):
            file_obj_list.append(file_)

        return file_obj_list

    def upload_file(self, local_file: str | Path, path_to_target: str | Path) -> None:
        """
        Upload a file to SharePoint.

        :param local_file: str, local file path
        :param path_to_target: list, list of folder names to get to target folder
        """
        if isinstance(path_to_target, str):
            path_to_target = Path(path_to_target)

        print(f'Uploading "{local_file}" to "{"/".join(path_to_target.parts)}"')
        directory = self.get_directory(path_to_target)
        print(f'Selected directory {directory.resource_url} {directory.properties}')
        with open(local_file, 'rb') as file_input:
            file_content = file_input.read()
            directory.upload_file(os.path.basename(local_file), file_content)  # .execute_query()
            self.ctx.execute_query()
            print('File uploaded')

    @staticmethod
    def download_file(file_obj: File, dest_path: str | None = None, file_name: str | None = None) -> str:
        """
        Download a file from SharePoint.

        :param file_obj: File, SharePoint File object.
        :param dest_path: Str, directory to download to.
        :param file_name: Str, name of the file to download.

        Example:

            sp = SharePoint()
            files = sp.get_file_objects(["Audit Folder"])
            for file in files:
                sp.download_file(file, dest_path='Workday')

        """
        if dest_path is None:
            dest_path = os.getcwd()

        if file_name is None:
            dest_file = os.path.join(dest_path, file_obj.properties['Name'])
        else:
            dest_file = os.path.join(dest_path, file_name)

        with open(dest_file, "wb") as local_file:
            file_obj.download(local_file).execute_query()

        return dest_file

    def delete_file(self, file_obj: File):
        """
        Delete a file from SharePoint using its server-relative URL.

        Example:

            sp.delete_file(file_obj)

        """
        try:
            file = self.ctx.web.get_file_by_server_relative_url(file_obj.serverRelativeUrl)
            file.delete_object()
            self.ctx.execute_query()
            print(f"Deleted file: {file.properties['Name']}")
        except Exception as e:
            print(f"Failed to delete file. Error: {e}")

    def get_file_binary_contents(self, file_obj: File):
        response = file_obj.read()
        self.ctx.execute_query()
        return BytesIO(response)

    def move_file(self, file_obj: File, file_path: str | Path | Folder) -> None:
        """
        Move a file in SharePoint.

        :param: File, file object to move
        :param: Path, str with the file path

        Example:

            sp = SharePoint()
            files = sp.get_files(["Audit Folder"])
            for file in files:
                sp.move_file(file, '')
        """
        if isinstance(file_path, str):
            file_path = Path(file_path)

        file_path = self.get_directory(file_path)
        file_obj.moveto(file_path, 1)  # flag = 1 means overwrite if exists
        self.ctx.execute_query()

    def archive_files(self, file_obj_list: list[File], dir_path: str | Path, sub_folder_name: str) -> None:
        """
        Move processed files to archive folder in SharePoint. This will create a folder called 'Archive' in the
        dir_path variable (if doesn't already exist) and then create a subfolder named as the archive_folder variable.

        :param file_obj_list: Path, path to sharepoint file objects to archive.
        :param dir_path: Path, path to target folder.
        :param sub_folder_name: str, name of archive folder within 'Archive'

        Example:

            sp = SharePointContext()
            files_path = "Dir 1/Another Dir/Final Dir"
            file_obj_list = sp.get_file_objects(files_path)
            sp.archive_files(file_obj_list=file_obj_list, dir_path=files_path, archive_folder='Archive 2024-11-01')

        """
        if isinstance(dir_path, str):
            dir_path = Path(dir_path)

        if len(file_obj_list) > 0:
            self.create_directory(dir_path, 'Archive')
            new_dir = self.create_directory(dir_path / 'Archive', sub_folder_name)
            for file in file_obj_list:
                print(f'Moving file to archive: {file.properties["Name"]}')
                try:
                    self.move_file(file, new_dir)
                except Exception as e:
                    print(f'Error moving file to archive: {file.properties["Name"]}')
                    print(str(e))

    def upload_copy_delete(self, sp_file: File, target_folder: Folder):
        """
        For folders where you don't have access to the full parent directory, use this to upload a copy
        of the file to a different folder and then delete the original.

        Example:

            archive_folder = sp.get_folder_by_link(f"/sites/ShareFiles/Main Folder/Archive")
            sp.upload_copy_delete(file_obj, archive_folder)
        """
        file_name = sp_file.properties["Name"]
        file_url = sp_file.properties["ServerRelativeUrl"]

        # 1. Read contents from the source file
        file_content = File.open_binary(self.ctx, file_url).content

        # 2. Upload to the target
        target_folder.upload_file(file_name, file_content)
        self.ctx.execute_query()

        # 3. Delete the original
        sp_file.delete_object()
        self.ctx.execute_query()

        print(f"Moved (fallback): {file_name}")

