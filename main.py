import os
import requests
import sys
from datetime import datetime

# 配置
CLIENT_ID = os.environ.get('CLIENT_ID')
CLIENT_SECRET = os.environ.get('CLIENT_SECRET')
REFRESH_TOKEN = os.environ.get('REFRESH_TOKEN')
TENANT_ID = os.environ.get('TENANT_ID')
OUTPUT_FILE = os.environ.get('GITHUB_OUTPUT')

def main():
    if not REFRESH_TOKEN:
        print("Error: No REFRESH_TOKEN")
        sys.exit(1)

    url = f'https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token'
    data = {
        'client_id': CLIENT_ID,
        'client_secret': CLIENT_SECRET,
        'refresh_token': REFRESH_TOKEN,
        'grant_type': 'refresh_token',
        'scope': 'Files.ReadWrite.All Mail.Send Calendars.Read User.Read offline_access'
    }

    try:
        r = requests.post(url, data=data)
        if r.status_code != 200:
            print(f"Refresh failed: {r.text}")
            sys.exit(1)
        
        res = r.json()
        new_rt = res.get('refresh_token')
        
        # 核心：将新令牌写入 GitHub 规定的输出文件
        if new_rt and new_rt != REFRESH_TOKEN and OUTPUT_FILE:
            with open(OUTPUT_FILE, 'a') as f:
                f.write(f"NEW_REFRESH_TOKEN={new_rt}\n")
            print(">>> New Refresh Token detected and exported.")

        # 模拟一些业务调用保活
        headers = {'Authorization': f"Bearer {res['access_token']}"}
        requests.get('https://graph.microsoft.com/v1.0/me/drive/root', headers=headers)
        print(f"KeepAlive OK at {datetime.now()}")

    except Exception as e:
        print(f"Error: {e}")
        sys.exit(1)

if __name__ == '__main__':
    main()
