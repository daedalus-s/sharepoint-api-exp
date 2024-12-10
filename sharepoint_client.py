import requests
import json
import msal
import configparser
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
    
    def get_site_info(self, site_path: str = 'root') -> Dict:
        """Get information about a SharePoint site"""
        graph_url = f'https://graph.microsoft.com/v1.0/sites/{site_path}'
        response = requests.get(graph_url, headers=self.http_headers)
        return response.json()
    
    def get_lists(self, site_id: str) -> List[Dict]:
        """Get all lists in a SharePoint site"""
        graph_url = f'https://graph.microsoft.com/v1.0/sites/{site_id}/lists'
        response = requests.get(graph_url, headers=self.http_headers)
        return response.json().get('value', [])
    
    def get_list_items(self, site_id: str, list_id: str) -> List[Dict]:
        """Get items from a specific list"""
        graph_url = f'https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{list_id}/items'
        response = requests.get(graph_url, headers=self.http_headers)
        return response.json().get('value', [])
    
    def create_list_item(self, site_id: str, list_id: str, fields: Dict) -> Dict:
        """Create a new item in a list"""
        graph_url = f'https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{list_id}/items'
        body = {
            "fields": fields
        }
        response = requests.post(graph_url, headers=self.http_headers, json=body)
        return response.json()
    
    def update_list_item(self, site_id: str, list_id: str, item_id: str, fields: Dict) -> Dict:
        """Update an existing list item"""
        graph_url = f'https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{list_id}/items/{item_id}/fields'
        response = requests.patch(graph_url, headers=self.http_headers, json=fields)
        return response.json()
    
    def delete_list_item(self, site_id: str, list_id: str, item_id: str) -> bool:
        """Delete a list item"""
        graph_url = f'https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{list_id}/items/{item_id}'
        response = requests.delete(graph_url, headers=self.http_headers)
        return response.status_code == 204
    
    def get_list_by_name(self, site_id: str, list_name: str) -> Optional[Dict]:
        """Get a list by its display name"""
        lists = self.get_lists(site_id)
        for list_info in lists:
            if list_info['displayName'].lower() == list_name.lower():
                return list_info
        return None
    
    def get_site_by_name(self, site_name: str) -> Dict:
        """Get a site by its name"""
        graph_url = f'https://graph.microsoft.com/v1.0/sites/{site_name}'
        response = requests.get(graph_url, headers=self.http_headers)
        return response.json()
    
    # Add this new method to your SharePointClient class:

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
        
        return response.json()

    def get_subsites(self, site_id: str) -> List[Dict]:
        """Get all subsites of a site"""
        graph_url = f'https://graph.microsoft.com/v1.0/sites/{site_id}/sites'
        response = requests.get(graph_url, headers=self.http_headers)
        return response.json().get('value', [])

def main():
    # Initialize the client
    client = SharePointClient('config.cfg')
    
    # Authenticate
    if not client.authenticate():
        print("Authentication failed!")
        return
    
    try:
        # Get root site information
        site_info = client.get_site_info()
        print(f"\nSite URL: {site_info['webUrl']}")
        site_id = site_info['id']
        
        # Get lists
        print("\nFetching lists...")
        lists = client.get_lists(site_id)
        for list_info in lists:
            print(f"- {list_info['displayName']}")
        
        if lists:
            # Work with the first list as an example
            test_list = lists[0]
            print(f"\nWorking with list: {test_list['displayName']}")
            
            # Create a test item
            print("\nCreating new item...")
            new_item = client.create_list_item(site_id, test_list['id'], {
                "Title": "Test Item",
                "Description": "Created via API"
            })
            print(f"Created item with ID: {new_item.get('id')}")
            
            # Get all items
            print("\nFetching all items...")
            items = client.get_list_items(site_id, test_list['id'])
            for item in items:
                print(f"- {item['fields'].get('Title')}")
            
            # Get subsites if any
            print("\nFetching subsites...")
            subsites = client.get_subsites(site_id)
            for subsite in subsites:
                print(f"- {subsite['displayName']}: {subsite['webUrl']}")
            
    except Exception as e:
        print(f"An error occurred: {str(e)}")

if __name__ == "__main__":
    main()