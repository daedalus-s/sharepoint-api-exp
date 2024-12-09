from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.client_credential import ClientCredential
import logging

class SharePointClient:
    def __init__(self, site_url, client_id, client_secret):
        self.site_url = site_url
        self.client_id = client_id
        self.client_secret = client_secret
        self.logger = logging.getLogger(__name__)
        self._init_client()

    def _init_client(self):
        try:
            credentials = ClientCredential(self.client_id, self.client_secret)
            self.ctx = ClientContext(self.site_url).with_credentials(credentials)
        except Exception as e:
            self.logger.error(f"Failed to initialize client: {str(e)}")
            raise

    def test_connection(self):
        try:
            web = self.ctx.web
            self.ctx.load(web)
            self.ctx.execute_query()
            self.logger.info(f"Successfully connected to: {web.properties['Title']}")
            return web.properties
        except Exception as e:
            self.logger.error(f"Connection test failed: {str(e)}")
            raise