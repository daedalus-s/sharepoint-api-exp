import requests

site_url = "https://humcodetechnologies143.sharepoint.com/sites/Sharepoint-Bedrock-Test"
response = requests.get(f"{site_url}/_vti_bin/client.svc", 
                       headers={"Authorization": "Bearer"}, 
                       verify=True)
www_authenticate = response.headers['WWW-Authenticate']
realm = www_authenticate.split('Bearer realm="')[1].split('"')[0]
print(f"Your realm ID is: {realm}")