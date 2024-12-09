from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File
import json

class SharePointClient:
    def __init__(self, site_url, client_id, client_secret):
        self.site_url = site_url
        self.client_id = client_id
        self.client_secret = client_secret
        self.ctx = self._get_client_context()

    def _get_client_context(self):
        auth = AuthenticationContext(self.site_url)
        auth.acquire_token_for_app(self.client_id, self.client_secret)
        return ClientContext(self.site_url, auth)

    def test_connection(self):
        """Test the connection by getting the web title"""
        web = self.ctx.web
        self.ctx.load(web)
        self.ctx.execute_query()
        return web.properties

    def get_all_lists(self):
        """Get all lists in the site"""
        lists = self.ctx.web.lists
        self.ctx.load(lists)
        self.ctx.execute_query()
        return [{'Title': lst.properties['Title'], 
                'ItemCount': lst.properties['ItemCount']} 
                for lst in lists]

    def get_list_items(self, list_title):
        """Get items from a specific list"""
        target_list = self.ctx.web.lists.get_by_title(list_title)
        items = target_list.items
        self.ctx.load(items)
        self.ctx.execute_query()
        return [item.properties for item in items]