import os
import base64
import json
import uuid
import time
import requests
from cryptography.hazmat.primitives import hashes, serialization
from cryptography.hazmat.primitives.asymmetric import padding
from cryptography import x509
from pathlib import Path


class SharePointClient:
    def __init__(self, tenant_id, client_id, cert_path, key_path, site_hostname, site_path, library_title, key_password=None):
        self.tenant_id = tenant_id
        self.client_id = client_id
        self.cert_path = cert_path
        self.key_path = key_path
        self.site_hostname = site_hostname
        self.site_path = site_path
        self.library_title = library_title
        self.key_password = key_password

        self.private_key, self.certificate = self._load_key_and_cert()
        self.access_token = self._get_access_token()
        self.headers = {"Authorization": f"Bearer {self.access_token}", "Accept": "application/json"}

        # Resolve site + drive once and cache them
        self.site_id = self._resolve_site()
        self.drive_id = self._resolve_drive()

    # -------------------------
    # Internal helpers
    # -------------------------
    def _base64url_encode(self, data: bytes) -> str:
        return base64.urlsafe_b64encode(data).rstrip(b"=").decode("utf-8")

    def _load_key_and_cert(self):
        # Load private key
        with open(self.key_path, "rb") as f:
            private_key = serialization.load_pem_private_key(
                f.read(),
                password=(self.key_password.encode() if self.key_password else None),
            )
        # Load certificate (PEM or DER)
        with open(self.cert_path, "rb") as f:
            cert_data = f.read()
            try:
                certificate = x509.load_pem_x509_certificate(cert_data)
            except ValueError:
                certificate = x509.load_der_x509_certificate(cert_data)

        return private_key, certificate

    def _new_jwt_client_assertion(self):
        aud = f"https://login.microsoftonline.com/{self.tenant_id}/oauth2/v2.0/token"
        nbf = int(time.time()) - 60
        exp = int(time.time()) + 300
        jti = str(uuid.uuid4())

        header = {
            "alg": "RS256",
            "typ": "JWT",
            "x5t": self._base64url_encode(self.certificate.fingerprint(hashes.SHA1()))
        }
        payload = {
            "iss": self.client_id,
            "sub": self.client_id,
            "aud": aud,
            "nbf": nbf,
            "exp": exp,
            "jti": jti
        }

        header_b64 = self._base64url_encode(json.dumps(header, separators=(",", ":")).encode())
        payload_b64 = self._base64url_encode(json.dumps(payload, separators=(",", ":")).encode())
        unsigned = f"{header_b64}.{payload_b64}"

        signature = self.private_key.sign(
            unsigned.encode("ascii"),
            padding.PKCS1v15(),
            hashes.SHA256()
        )
        sig_b64 = self._base64url_encode(signature)

        return f"{unsigned}.{sig_b64}"

    def _get_access_token(self, scope="https://graph.microsoft.com/.default"):
        assertion = self._new_jwt_client_assertion()
        body = {
            "client_id": self.client_id,
            "scope": scope,
            "grant_type": "client_credentials",
            "client_assertion_type": "urn:ietf:params:oauth:client-assertion-type:jwt-bearer",
            "client_assertion": assertion
        }
        token_endpoint = f"https://login.microsoftonline.com/{self.tenant_id}/oauth2/v2.0/token"
        resp = requests.post(token_endpoint, data=body)
        resp.raise_for_status()
        return resp.json()["access_token"]

    def _resolve_site(self):
        site_url = f"https://graph.microsoft.com/v1.0/sites/{self.site_hostname}:{self.site_path}"
        site = requests.get(site_url, headers=self.headers).json()
        if "id" not in site:
            raise Exception(f"Failed to resolve site id for {self.site_hostname}{self.site_path}")
        return site["id"]

    def _resolve_drive(self):
        drives_url = f"https://graph.microsoft.com/v1.0/sites/{self.site_id}/drives"
        drives = requests.get(drives_url, headers=self.headers).json()
        drive = next((d for d in drives["value"] if d["name"] == self.library_title), None)
        if not drive:
            raise Exception(f"Drive (library) named '{self.library_title}' not found on site.")
        return drive["id"]

    # -------------------------
    # Public methods
    # -------------------------
    def list_folder(self, folder_name=""):
        """List files/folders inside the given folder (default = root)."""
        if not folder_name.strip():
            url = f"https://graph.microsoft.com/v1.0/drives/{self.drive_id}/root/children"
        else:
            encoded = requests.utils.quote(folder_name.strip("/"))
            folder_url = f"https://graph.microsoft.com/v1.0/drives/{self.drive_id}/root:/{encoded}"
            resp = requests.get(folder_url, headers=self.headers)
            if resp.status_code != 200:
                print(f"Folder '{folder_name}' not found. Listing root instead.")
                url = f"https://graph.microsoft.com/v1.0/drives/{self.drive_id}/root/children"
            else:
                folder_item = resp.json()
                url = f"https://graph.microsoft.com/v1.0/drives/{self.drive_id}/items/{folder_item['id']}/children"
        return requests.get(url, headers=self.headers).json()

    def print_folder(self, folder_name=""):
        """Pretty-print contents of a folder."""
        children = self.list_folder(folder_name)
        print("FILES:")
        for item in children.get("value", []):
            if "file" in item:
                print(f"{item['name']}\t{item['webUrl']}")
        print("\nFOLDERS:")
        for item in children.get("value", []):
            if "folder" in item:
                print(f"{item['name']}\t{item['webUrl']}")

    def create_folder(self, folder_path):
        """
        Create a folder (and its parent chain if needed) in SharePoint.

        folder_path: path relative to the library root (e.g. "HCM Audit/Archive/2025")
        """
        parts = folder_path.strip("/").split("/")
        current_path = ""

        created_folder = None
        for part in parts:
            current_path = f"{current_path}/{part}" if current_path else part
            encoded_path = requests.utils.quote(current_path)

            # Check if it already exists
            url = f"https://graph.microsoft.com/v1.0/drives/{self.drive_id}/root:/{encoded_path}"
            resp = requests.get(url, headers=self.headers)

            if resp.status_code == 200:
                # Folder already exists
                created_folder = resp.json()
                continue

            # Otherwise create it under its parent
            parent_path = "/".join(current_path.split("/")[:-1])
            if parent_path:
                parent_encoded = requests.utils.quote(parent_path)
                create_url = f"https://graph.microsoft.com/v1.0/drives/{self.drive_id}/root:/{parent_encoded}:/children"
            else:
                create_url = f"https://graph.microsoft.com/v1.0/drives/{self.drive_id}/root/children"

            body = {
                "name": part,
                "folder": {},
                "@microsoft.graph.conflictBehavior": "fail"
            }
            resp = requests.post(create_url, headers=self.headers, json=body)
            resp.raise_for_status()
            created_folder = resp.json()
            print(f"📁 Created folder: {current_path}")

        return created_folder

    def get_files(self, folder_name=""):
        """
        Return a list of file metadata objects in the given folder.
        Each object includes id, name, webUrl, size, lastModifiedDateTime, etc.
        """
        children = self.list_folder(folder_name)
        return [item for item in children.get("value", []) if "file" in item]

    def get_folders(self, folder_name=""):
        """
        Return a list of folder metadata objects in the given folder.
        """
        children = self.list_folder(folder_name)
        return [item for item in children.get("value", []) if "folder" in item]


    def download_files(self, folder_name="", download_dir="downloads"):
        """Download all files in a folder to `download_dir`."""
        os.makedirs(download_dir, exist_ok=True)
        children = self.list_folder(folder_name)
        for item in children.get("value", []):
            if "file" in item:
                url = item["@microsoft.graph.downloadUrl"]
                local_path = os.path.join(download_dir, item["name"])
                print(f"⬇ Downloading {item['name']} ...")
                resp = requests.get(url)
                resp.raise_for_status()
                with open(local_path, "wb") as f:
                    f.write(resp.content)
                print(f"✅ Saved to {local_path}")

    def download_file(self, file_path, download_dir="downloads", new_name=None):
        """
        Download a single file from SharePoint.

        file_path: path to the file relative to library root
                   (e.g. "HCM Audit/250711 - HCM Audit Findings.xlsx")
        download_dir: local folder to save into (default "downloads")
        new_name: optional new filename for saving locally
        """
        os.makedirs(download_dir, exist_ok=True)
        if isinstance(file_path, dict) and "webUrl" in file_path:
            file_path = file_path["webUrl"]

        # Resolve the file item
        encoded_path = requests.utils.quote(file_path.strip("/"))
        url = f"https://graph.microsoft.com/v1.0/drives/{self.drive_id}/root:/{encoded_path}"
        resp = requests.get(url, headers=self.headers)
        resp.raise_for_status()
        file_item = resp.json()

        if "@microsoft.graph.downloadUrl" not in file_item:
            raise Exception(f"Item '{file_path}' is not a file or has no download URL")

        # Get the direct download link
        download_url = file_item["@microsoft.graph.downloadUrl"]

        # Choose local filename
        local_filename = new_name if new_name else file_item["name"]
        local_path = os.path.join(download_dir, local_filename)

        # Download
        print(f"⬇ Downloading {file_item['name']} ...")
        file_resp = requests.get(download_url)
        file_resp.raise_for_status()

        with open(local_path, "wb") as f:
            f.write(file_resp.content)

        print(f"✅ Saved to {local_path}")
        return local_path


    def upload_file(self, local_path, target_folder="", chunk_size=5*1024*1024):
        """
        Upload a file to SharePoint.
        - If file <= 4MB, does simple PUT.
        - If file > 4MB, uses an upload session (chunked).
        """
        file_name = os.path.basename(local_path)
        folder_path = target_folder.strip("/")
        if folder_path:
            item_path = f"{folder_path}/{file_name}"
        else:
            item_path = file_name

        file_size = os.path.getsize(local_path)

        if file_size <= 4 * 1024 * 1024:  # small file
            url = f"https://graph.microsoft.com/v1.0/drives/{self.drive_id}/root:/{item_path}:/content"
            with open(local_path, "rb") as f:
                resp = requests.put(url, headers=self.headers, data=f)
            resp.raise_for_status()
            print(f"✅ Uploaded small file: {item_path}")
            return resp.json()

        # Large file: use upload session
        url = f"https://graph.microsoft.com/v1.0/drives/{self.drive_id}/root:/{item_path}:/createUploadSession"
        session = requests.post(url, headers=self.headers, json={"item": {"@microsoft.graph.conflictBehavior": "replace"}})
        session.raise_for_status()
        upload_url = session.json()["uploadUrl"]

        with open(local_path, "rb") as f:
            i = 0
            while True:
                chunk = f.read(chunk_size)
                if not chunk:
                    break
                start = i * chunk_size
                end = start + len(chunk) - 1
                headers = {
                    "Content-Length": str(len(chunk)),
                    "Content-Range": f"bytes {start}-{end}/{file_size}"
                }
                resp = requests.put(upload_url, headers=headers, data=chunk)
                resp.raise_for_status()
                i += 1

        print(f"✅ Uploaded large file: {item_path}")
        return resp.json()

    def move_file(self, file_path, target_folder):
        """
        Move a file to another folder in the same library.

        file_path: path to the file relative to the drive root (e.g. "HR/Payroll/report.xlsx")
        target_folder: target folder path relative to root (e.g. "HR/Archive")
        """
        # Get the file item first
        encoded_path = requests.utils.quote(file_path.strip("/"))
        url = f"https://graph.microsoft.com/v1.0/drives/{self.drive_id}/root:/{encoded_path}"
        resp = requests.get(url, headers=self.headers)
        resp.raise_for_status()
        file_item = resp.json()
        file_id = file_item["id"]

        # Resolve target folder
        encoded_target = requests.utils.quote(target_folder.strip("/"))
        folder_url = f"https://graph.microsoft.com/v1.0/drives/{self.drive_id}/root:/{encoded_target}"
        resp = requests.get(folder_url, headers=self.headers)
        resp.raise_for_status()
        folder_item = resp.json()
        folder_id = folder_item["id"]

        # Issue PATCH to move
        patch_url = f"https://graph.microsoft.com/v1.0/drives/{self.drive_id}/items/{file_id}"
        body = {
            "parentReference": {
                "id": folder_id
            }
        }
        move_resp = requests.patch(patch_url, headers=self.headers, json=body)
        move_resp.raise_for_status()

        print(f"✅ Moved '{file_path}' → '{target_folder}/'")
        return move_resp.json()

    def rename_file(self, file_path, new_name):
        """
        Rename a file in SharePoint.

        file_path: current path of the file relative to library root
                   (e.g. "HCM Audit/250711 - HCM Audit Findings.xlsx")
        new_name:  new filename (just the name, not a path, e.g. "Findings_2025.xlsx")
        """
        # Resolve the file item
        encoded_path = requests.utils.quote(file_path.strip("/"))
        url = f"https://graph.microsoft.com/v1.0/drives/{self.drive_id}/root:/{encoded_path}"
        resp = requests.get(url, headers=self.headers)
        resp.raise_for_status()
        file_item = resp.json()
        file_id = file_item["id"]

        # Rename via PATCH
        patch_url = f"https://graph.microsoft.com/v1.0/drives/{self.drive_id}/items/{file_id}"
        body = {"name": new_name}
        rename_resp = requests.patch(patch_url, headers=self.headers, json=body)
        rename_resp.raise_for_status()

        print(f"✅ Renamed '{file_path}' → '{new_name}'")
        return rename_resp.json()

    def delete_file(self, file_path):
        """
        Delete a file from SharePoint.

        file_path: path to the file relative to the library root
                   (e.g. "HR/Payroll/report.xlsx")
        """
        # Resolve the file item first
        encoded_path = requests.utils.quote(file_path.strip("/"))
        url = f"https://graph.microsoft.com/v1.0/drives/{self.drive_id}/root:/{encoded_path}"
        resp = requests.get(url, headers=self.headers)
        if resp.status_code == 404:
            raise FileNotFoundError(f"File '{file_path}' not found in SharePoint.")
        resp.raise_for_status()
        file_item = resp.json()
        file_id = file_item["id"]

        # Issue DELETE request
        delete_url = f"https://graph.microsoft.com/v1.0/drives/{self.drive_id}/items/{file_id}"
        del_resp = requests.delete(delete_url, headers=self.headers)

        if del_resp.status_code in (204, 200):
            print(f"🗑️ Deleted file: {file_path}")
            return True
        else:
            raise Exception(f"Failed to delete file '{file_path}': {del_resp.text}")

