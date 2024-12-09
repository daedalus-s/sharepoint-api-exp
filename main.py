from sharepoint_client import SharePointClient
import logging

logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

SITE_URL = "https://humcodetechnologies143.sharepoint.com/sites/Sharepoint-Bedrock-Test"
CLIENT_ID = "4640e179-675e-4c20-8620-3a3f31fc2ea4"      # Just the ID, no realm
CLIENT_SECRET = "zDGyTwSshid+D8Tz3L6toW04A1RcMiQ7h/W9GOsdqeQ="     # The secret from appregnew.aspx

def main():
    try:
        logger.info("Initializing SharePoint client...")
        client = SharePointClient(SITE_URL, CLIENT_ID, CLIENT_SECRET)
        
        logger.info("Testing connection...")
        site_info = client.test_connection()
        
        if site_info:
            logger.info("Connection successful!")
            logger.info(f"Site Title: {site_info.get('Title', 'N/A')}")
                
    except Exception as e:
        logger.error(f"Error in main: {str(e)}", exc_info=True)

if __name__ == "__main__":
    main()