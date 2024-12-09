from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.client_credential import ClientCredential
import logging

class SharePointGraphClient:
    def __init__(self, site_url, client_id, client_secret):
        """Initialize SharePoint client using Graph API"""
        self.site_url = site_url
        self.client_id = client_id
        self.client_secret = client_secret
        self.logger = logging.getLogger(__name__)
        self._init_client()

    def _init_client(self):
        """Initialize the client context with credentials"""
        try:
            credentials = ClientCredential(self.client_id, self.client_secret)
            self.ctx = ClientContext(self.site_url).with_credentials(credentials)
            
        except Exception as e:
            self.logger.error(f"Failed to initialize client: {str(e)}")
            raise

    def test_connection(self):
        """Test connection and get basic site info"""
        try:
            web = self.ctx.web
            self.ctx.load(web)
            self.ctx.execute_query()
            self.logger.info(f"Successfully connected to: {web.properties['Title']}")
            return web.properties
        except Exception as e:
            self.logger.error(f"Connection test failed: {str(e)}")
            raise

    def get_all_lists(self):
        """Get all lists in the site"""
        try:
            lists = self.ctx.web.lists
            self.ctx.load(lists)
            self.ctx.execute_query()
            return [{'Title': lst.properties['Title'],
                    'ItemCount': lst.properties.get('ItemCount', 0),
                    'Id': lst.properties['Id']}
                   for lst in lists]
        except Exception as e:
            self.logger.error(f"Failed to get lists: {str(e)}")
            raise

    def get_list_items(self, list_title):
        """Get items from a specific list"""
        try:
            target_list = self.ctx.web.lists.get_by_title(list_title)
            items = target_list.items.get_all()
            return [item.properties for item in items]
        except Exception as e:
            self.logger.error(f"Failed to get items from list '{list_title}': {str(e)}")
            raise