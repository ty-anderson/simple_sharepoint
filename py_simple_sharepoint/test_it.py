import os
from pathlib import Path
from py_simple_sharepoint import SharePointClient
from dotenv import load_dotenv
load_dotenv()


root_path = Path(__file__).parent.parent.resolve()
sp = SharePointClient(
    tenant_id=os.getenv("SHAREPOINT_CLIENT_TENANT_ID"),
    client_id=os.getenv("SHAREPOINT_AZURE_APPLICATION_ID"),
    cert_path=root_path / r"PACS.SelfSigned.ETLSharepointAppReg.cer",
    key_path=root_path / r"PACS.SelfSigned.ETLSharepointAppReg.key",
    site_hostname="providencegroup.sharepoint.com",
    site_path="/sites/P_PACS",
    library_title='PACS'
)

files = sp.get_files('Data Team/Net Health/Archive')

for file in files:
    print(file['name'])
