import requests
import json
import msal
import configparser
import traceback
from typing import Optional, Dict, List, Any

class SharePointClient:
    def __init__(self, config_path: str = 'config.cfg'):
        """
        Initialize SharePoint client with configuration
        
        config.cfg should contain:
        [azure]
        clientId=your_client_id
        clientSecret=your_client_secret
        tenantId=your_tenant_id
        """
        self.config = configparser.ConfigParser()
        self.config.read(config_path)
        self.client_id = self.config.get('azure', 'clientId')
        self.client_secret = self.config.get('azure', 'clientSecret')
        self.authority = f"https://login.microsoftonline.com/{self.config.get('azure', 'tenantId')}"
        self.scope = ["https://graph.microsoft.com/.default"]
        self.access_token = None
        self.http_headers = None
        
    def authenticate(self) -> bool:
        """Authenticate with Microsoft Graph API"""
        try:
            app = msal.ConfidentialClientApplication(
                self.client_id,
                authority=self.authority,
                client_credential=self.client_secret,
            )
            
            # Try to get token from cache first
            result = app.acquire_token_silent(self.scope, account=None)
            
            if not result:
                print("Getting new token from Azure AD...")
                result = app.acquire_token_for_client(scopes=self.scope)
                
            if "access_token" in result:
                self.access_token = result['access_token']
                self.http_headers = {
                    'Authorization': f'Bearer {self.access_token}',
                    'Accept': 'application/json',
                    'Content-Type': 'application/json'
                }
                return True
            else:
                print(f"Error: {result.get('error')}")
                print(f"Description: {result.get('error_description')}")
                print(f"Correlation ID: {result.get('correlation_id')}")
                return False
        except Exception as e:
            print(f"Authentication error: {str(e)}")
            print(f"Detailed error: {traceback.format_exc()}")
            return False
    
    def get_site_info(self, site_path: str) -> Dict:
        """Get information about a SharePoint site"""
        try:
            graph_url = f'https://graph.microsoft.com/v1.0/sites/{site_path}'
            response = requests.get(graph_url, headers=self.http_headers)
            response_json = response.json()
            
            if response.status_code != 200:
                print(f"Error getting site info. Status code: {response.status_code}")
                print(f"Error details: {response_json}")
            
            return response_json
        except Exception as e:
            print(f"Error in get_site_info: {str(e)}")
            print(f"Detailed error: {traceback.format_exc()}")
            return {"error": str(e)}
    
    def get_lists(self, site_id: str) -> List[Dict]:
        """Get all lists in a SharePoint site"""
        try:
            graph_url = f'https://graph.microsoft.com/v1.0/sites/{site_id}/lists'
            response = requests.get(graph_url, headers=self.http_headers)
            return response.json().get('value', [])
        except Exception as e:
            print(f"Error in get_lists: {str(e)}")
            return []
    
    def upload_file(self, site_id: str, file_path: str, target_folder: str = "Shared Documents") -> Dict:
        """
        Upload a file to SharePoint document library
        
        Args:
            site_id (str): The SharePoint site ID
            file_path (str): Path to the local file to upload
            target_folder (str): Target folder in SharePoint (default is 'Shared Documents')
        
        Returns:
            Dict: Response from the API containing the uploaded file information
        """
        try:
            # Get the file name from the path
            file_name = file_path.split('/')[-1] if '/' in file_path else file_path.split('\\')[-1]
            
            # Read the file content
            with open(file_path, 'rb') as file_content:
                content = file_content.read()
            
            # Construct the upload URL
            upload_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{target_folder}/{file_name}:/content"
            
            # Upload the file
            response = requests.put(
                upload_url,
                headers={
                    'Authorization': f'Bearer {self.access_token}',
                    'Content-Type': 'application/octet-stream'
                },
                data=content
            )
            
            response_json = response.json()
            
            if response.status_code not in [200, 201]:
                print(f"Error uploading file. Status code: {response.status_code}")
                print(f"Error details: {response_json}")
            
            return response_json
            
        except Exception as e:
            print(f"Error in upload_file: {str(e)}")
            print(f"Detailed error: {traceback.format_exc()}")
            return {"error": str(e)}

def main():
    # Initialize the client
    client = SharePointClient('config.cfg')
    
    # Authenticate
    if not client.authenticate():
        print("Authentication failed!")
        return
    
    try:
        # Get site info using the full domain
        site_path = 'humcodetechnologies143.sharepoint.com:/sites/Sharepoint-Bedrock-Test'
        print(f"\nGetting site info for: {site_path}")
        
        site_info = client.get_site_info(site_path)
        print("\nSite response:", json.dumps(site_info, indent=2))
        
        if 'error' in site_info:
            print(f"Error getting site: {site_info['error']}")
            return
            
        site_id = site_info['id']
        print(f"\nSite found: {site_info.get('webUrl', 'No URL found')}")
        print(f"Site ID: {site_id}")
        
        # Create a test file
        test_file_path = "test.txt"
        print(f"\nCreating test file: {test_file_path}")
        with open(test_file_path, "w") as f:
            f.write("This is a test file for SharePoint upload")
        
        # Upload the test file
        print("\nUploading file...")
        result = client.upload_file(
            site_id=site_id,
            file_path=test_file_path,
            target_folder="Shared Documents"
        )
        
        if 'error' in result:
            print(f"Error uploading file: {result['error']}")
        else:
            print(f"File uploaded successfully!")
            print(f"File URL: {result.get('webUrl', 'No URL available')}")
            
    except Exception as e:
        print(f"\nAn error occurred in main: {str(e)}")
        print("Full error:", traceback.format_exc())

if __name__ == "__main__":
    main()