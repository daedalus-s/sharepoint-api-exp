import requests
import json
import msal
import configparser
from typing import Dict, List, Any, Set
from dataclasses import dataclass
from concurrent.futures import ThreadPoolExecutor
import time

@dataclass
class GroupMember:
    id: str
    display_name: str
    user_principal_name: str
    email: str = None

@dataclass
class PermissionEntry:
    item_id: str
    item_type: str  # 'site', 'list', 'folder', 'file'
    item_path: str
    roles: List[Dict[str, Any]]
    inherited: bool
    parent_id: str = None
    group_members: Dict[str, List[GroupMember]] = None  # Map group id to list of members

class SharePointPermissionExtractor:
    def __init__(self, config_path: str = 'config.cfg'):
        """Initialize SharePoint permission extractor with configuration"""
        # Load configuration
        self.config = configparser.ConfigParser()
        self.config.read(config_path)
        self.client_id = self.config.get('azure', 'clientId')
        self.client_secret = self.config.get('azure', 'clientSecret')
        self.authority = f"https://login.microsoftonline.com/{self.config.get('azure', 'tenantId')}"
        self.scope = ["https://graph.microsoft.com/.default"]
        
        # Initialize authentication-related attributes
        self.access_token = None
        self.headers = None
        
        # Initialize extractor-related attributes
        self.permission_cache = {}
        self.batch_size = 20
        self.site_id = None

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
                self.headers = {
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
            return False

    def get_site_permissions(self, site_id: str) -> Dict:
        """Get permissions at site level"""
        url = f'https://graph.microsoft.com/v1.0/sites/{site_id}/permissions'
        response = requests.get(url, headers=self.headers)
        data = response.json()
        print(f"\nSite permissions structure: {json.dumps(data, indent=2)}")
        return data

    def get_site_groups(self) -> List[Dict]:
        """Get all SharePoint groups for the site"""
        url = f'https://graph.microsoft.com/v1.0/sites/{self.site_id}/groups'
        try:
            response = requests.get(url, headers=self.headers)
            data = response.json()
            print(f"Found site groups: {json.dumps(data, indent=2)}")
            return data.get('value', [])
        except Exception as e:
            print(f"Error getting site groups: {str(e)}")
            return []

    def get_drive_item_permissions(self, site_id: str, item_id: str) -> Dict:
        """Get permissions for a specific drive item"""
        url = f'https://graph.microsoft.com/v1.0/sites/{site_id}/drive/items/{item_id}/permissions'
        response = requests.get(url, headers=self.headers)
        return response.json()

    def get_list_item_permissions(self, site_id: str, list_id: str, item_id: str) -> Dict:
        """Get permissions for a specific list item"""
        url = f'https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{list_id}/items/{item_id}/driveItem/permissions'
        response = requests.get(url, headers=self.headers)
        return response.json()

    def get_all_subsites(self, site_id: str) -> List[Dict]:
        """Get all subsites recursively"""
        subsites = []
        url = f'https://graph.microsoft.com/v1.0/sites/{site_id}/sites'
        
        try:
            response = requests.get(url, headers=self.headers)
            data = response.json()
            
            if 'value' in data:
                subsites.extend(data['value'])
                
                # Recursively get subsites of each subsite
                for subsite in data['value']:
                    subsites.extend(self.get_all_subsites(subsite['id']))
                    
        except Exception as e:
            print(f"Error getting subsites: {str(e)}")
            
        return subsites

    def get_all_lists(self, site_id: str) -> List[Dict]:
        """Get all lists in a site"""
        url = f'https://graph.microsoft.com/v1.0/sites/{site_id}/lists'
        response = requests.get(url, headers=self.headers)
        return response.json().get('value', [])

    def get_list_items(self, site_id: str, list_id: str) -> List[Dict]:
        """Get all items in a list"""
        url = f'https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{list_id}/items?expand=fields'
        response = requests.get(url, headers=self.headers)
        return response.json().get('value', [])

    def get_group_members(self, group_id: str) -> List[GroupMember]:
        """Get all members of a specific group"""
        url = f'https://graph.microsoft.com/v1.0/groups/{group_id}/members'
        try:
            response = requests.get(url, headers=self.headers)
            if response.status_code == 404:
                return self.get_sharepoint_group_members(group_id)
            
            print(f"Group members response: {response.status_code}")
            print(f"Response content: {response.text}")
            
            data = response.json()
            members = []
            
            for member in data.get('value', []):
                members.append(GroupMember(
                    id=member.get('id'),
                    display_name=member.get('displayName'),
                    user_principal_name=member.get('userPrincipalName'),
                    email=member.get('mail')
                ))
            
            print(f"Found {len(members)} members for Azure AD group {group_id}")
            return members
        except Exception as e:
            print(f"Error getting group members for {group_id}: {str(e)}")
            return []

    def get_sharepoint_group_members(self, group_id: str) -> List[GroupMember]:
        """Get members of a SharePoint group"""
        url = f'https://graph.microsoft.com/v1.0/sites/{self.site_id}/siteGroups/{group_id}/users'
        try:
            response = requests.get(url, headers=self.headers)
            print(f"Group members response for group {group_id}: {response.status_code}")
            print(f"Response content: {response.text}")
            
            if response.status_code == 200:
                data = response.json()
                members = []
                
                for member in data.get('value', []):
                    members.append(GroupMember(
                        id=member.get('id'),
                        display_name=member.get('displayName'),
                        user_principal_name=member.get('userPrincipalName', ''),
                        email=member.get('email', member.get('mail', ''))
                    ))
                
                print(f"Found {len(members)} members for SharePoint group {group_id}")
                return members
            else:
                print(f"Failed to get group members. Status code: {response.status_code}")
                print(f"Response: {response.text}")
                return []
                
        except Exception as e:
            print(f"Error getting SharePoint group members for {group_id}: {str(e)}")
            return []

    def process_permission_roles(self, roles: List[Dict], item_id: str) -> Dict[str, List[GroupMember]]:
        """Process permission roles and get group members"""
        group_members = {}
        
        for role in roles:
            # Add debug logging
            print(f"Processing role: {role.get('roles', [])} for item: {item_id}")
            print(f"Role structure: {json.dumps(role, indent=2)}")
            
            # Check grantedToV2 for site groups
            if 'grantedToV2' in role:
                granted_to_v2 = role['grantedToV2']
                
                # Check for site groups
                if 'siteGroup' in granted_to_v2:
                    group_id = granted_to_v2['siteGroup'].get('id')
                    if group_id:
                        members = self.get_sharepoint_group_members(group_id)
                        if members:
                            group_members[group_id] = members
                
                # Check for Azure AD groups
                elif 'group' in granted_to_v2 and granted_to_v2['group'].get('@odata.type', '').endswith('sharePointIdentity'):
                    group_id = granted_to_v2['group'].get('id')
                    if group_id:
                        members = self.get_group_members(group_id)
                        if members:
                            group_members[group_id] = members

        return group_members

    def process_batch(self, items: List[Dict], site_id: str, list_id: str = None) -> List[PermissionEntry]:
        """Process a batch of items in parallel"""
        results = []
        with ThreadPoolExecutor(max_workers=self.batch_size) as executor:
            if list_id:
                futures = [executor.submit(self.get_list_item_permissions, site_id, list_id, item['id']) 
                          for item in items]
            else:
                futures = [executor.submit(self.get_drive_item_permissions, site_id, item['id']) 
                          for item in items]
            
            for item, future in zip(items, futures):
                try:
                    perm_data = future.result()
                    results.append(PermissionEntry(
                        item_id=item['id'],
                        item_type='file' if 'file' in item else 'folder',
                        item_path=item.get('webUrl', ''),
                        roles=perm_data.get('value', []),
                        inherited=any(p.get('inherited', False) for p in perm_data.get('value', [])),
                        parent_id=item.get('parentReference', {}).get('id')
                    ))
                except Exception as e:
                    print(f"Error processing item {item['id']}: {str(e)}")
        
        return results

    def extract_all_permissions(self, site_path: str) -> Dict[str, List[PermissionEntry]]:
        """
        Extract all permissions across the site hierarchy
        Args:
            site_path: The SharePoint site path (e.g., 'domain.sharepoint.com:/sites/site-name')
        Returns:
            Dictionary with different types of permissions
        """
        # Get site ID first
        url = f'https://graph.microsoft.com/v1.0/sites/{site_path}'
        response = requests.get(url, headers=self.headers)
        site_data = response.json()
        
        if 'id' not in site_data:
            raise ValueError(f"Could not find site ID for path: {site_path}")
            
        self.site_id = site_data['id']
        
        print("\nFetching site groups...")
        site_groups = self.get_site_groups()
        for group in site_groups:
            group_id = group.get('id')
            if group_id:
                members = self.get_sharepoint_group_members(group_id)
                if members:
                    print(f"Group {group.get('displayName')}: {len(members)} members")
        
        permissions_db = {
            'sites': [],
            'lists': [],
            'documents': [],
            'folders': []
        }

        # Get site permissions with group members
        site_perms = self.get_site_permissions(self.site_id)
        group_members = self.process_permission_roles(site_perms.get('value', []), self.site_id)
        
        permissions_db['sites'].append(PermissionEntry(
            item_id=self.site_id,
            item_type='site',
            item_path='/',
            roles=site_perms.get('value', []),
            inherited=False,
            group_members=group_members
        ))

        # Process subsites
        subsites = self.get_all_subsites(self.site_id)
        for subsite in subsites:
            subsite_perms = self.get_site_permissions(subsite['id'])
            group_members = self.process_permission_roles(subsite_perms.get('value', []), subsite['id'])
            
            permissions_db['sites'].append(PermissionEntry(
                item_id=subsite['id'],
                item_type='site',
                item_path=subsite.get('webUrl', ''),
                roles=subsite_perms.get('value', []),
                inherited=False,
                group_members=group_members
            ))

        # Process lists and their items
        lists = self.get_all_lists(self.site_id)
        for list_info in lists:
            list_perms = self.get_drive_item_permissions(self.site_id, list_info['id'])
            group_members = self.process_permission_roles(list_perms.get('value', []), list_info['id'])
            
            permissions_db['lists'].append(PermissionEntry(
                item_id=list_info['id'],
                item_type='list',
                item_path=list_info.get('webUrl', ''),
                roles=list_perms.get('value', []),
                inherited=True,
                group_members=group_members
            ))

            # Process list items
            items = self.get_list_items(self.site_id, list_info['id'])
            for i in range(0, len(items), self.batch_size):
                batch = items[i:i + self.batch_size]
                results = self.process_batch(batch, self.site_id, list_info['id'])
                
                for entry in results:
                    # Add group members for each entry
                    entry.group_members = self.process_permission_roles(entry.roles, entry.item_id)
                    
                    if entry.item_type == 'file':
                        permissions_db['documents'].append(entry)
                    else:
                        permissions_db['folders'].append(entry)

        return permissions_db

    def save_to_json(self, permissions_db: Dict[str, List[PermissionEntry]], output_file: str = 'sharepoint_permissions.json') -> None:
        """
        Save the permissions database to a JSON file
        """
        try:
            json_data = {}
            for category, entries in permissions_db.items():
                json_data[category] = [
                    {
                        'item_id': entry.item_id,
                        'item_type': entry.item_type,
                        'item_path': entry.item_path,
                        'roles': entry.roles,
                        'inherited': entry.inherited,
                        'parent_id': entry.parent_id,
                        'group_members': {
                            group_id: [
                                {
                                    'id': member.id,
                                    'display_name': member.display_name,
                                    'user_principal_name': member.user_principal_name,
                                    'email': member.email
                                }
                                for member in members
                            ]
                            for group_id, members in (entry.group_members or {}).items()
                        }
                    }
                    for entry in entries
                ]
            
            with open(output_file, 'w', encoding='utf-8') as f:
                json.dump(json_data, f, indent=4, ensure_ascii=False)
                
            print(f"\nPermissions successfully saved to {output_file}")
            
        except Exception as e:
            print(f"Error saving to JSON: {str(e)}")

    def load_from_json(self, json_file: str = 'sharepoint_permissions.json') -> Dict:
        """
        Load permissions from JSON file
        """
        try:
            with open(json_file, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            print(f"Error loading JSON: {str(e)}")
            return {}

def main():
    # Initialize the extractor
    extractor = SharePointPermissionExtractor('config.cfg')
    
    # Authenticate
    if not extractor.authenticate():
        print("Authentication failed!")
        return
    
    try:
        # Extract permissions for a specific site
        site_path = 'humcodetechnologies143.sharepoint.com:/sites/Sharepoint-Bedrock-Test'
        print(f"\nExtracting permissions for site: {site_path}")
        
        permissions_db = extractor.extract_all_permissions(site_path)
        
        # Print summary
        print("\nPermissions extracted successfully!")
        print("\nBreakdown:")
        for category in permissions_db:
            print(f"{category}: {len(permissions_db[category])} entries")
        
        # Save to JSON file
        extractor.save_to_json(permissions_db, 'sharepoint_permissions.json')
        
        # Example: Load and verify the saved data
        loaded_permissions = extractor.load_from_json('sharepoint_permissions.json')
        print("\nVerified saved data - Entries per category:")
        for category in loaded_permissions:
            print(f"{category}: {len(loaded_permissions[category])} entries")
            
    except Exception as e:
        print(f"\nAn error occurred: {str(e)}")

if __name__ == "__main__":
    main()