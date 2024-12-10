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