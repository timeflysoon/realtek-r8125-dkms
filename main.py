import os
import random
import string
import requests
import time
import sys
from datetime import datetime

# ================= 🛡️ 核心配置区域 =================
CLIENT_ID = os.environ.get('CLIENT_ID')
CLIENT_SECRET = os.environ.get('CLIENT_SECRET')
REFRESH_TOKEN = os.environ.get('REFRESH_TOKEN')
TENANT_ID = os.environ.get('TENANT_ID')
GITHUB_REPO = os.environ.get('GITHUB_REPO_NAME', 'Unknown-Repo')

# 获取 GitHub 传递的环境变量文件路径，用于更新 Secret
GITHUB_OUTPUT_FILE = os.environ.get('GITHUB_OUTPUT')

if not TENANT_ID:
    print("!! 致命错误: 缺少 TENANT_ID 环境变量")
    sys.exit(1)

TOKEN_URL = f'https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token'
GRAPH_URL = 'https://graph.microsoft.com/v1.0'
DATA_FOLDER = "/Data"
LOCK_FOLDER = "/Data/Lock"

# ================= 🔐 鉴权模块 (含令牌自动更新逻辑) =================
def get_access_token():
    print(f">>> [Auth] 正在尝试刷新令牌 (账号: {GITHUB_REPO})")
    data = {
        'client_id': CLIENT_ID,
        'client_secret': CLIENT_SECRET,
        'refresh_token': REFRESH_TOKEN,
        'grant_type': 'refresh_token',
        'scope': 'Files.ReadWrite.All Mail.Send Calendars.Read User.Read offline_access'
    }
    try:
        r = requests.post(TOKEN_URL, data=data, timeout=30)
        if r.status_code != 200:
            print(f"!! 令牌刷新失败: {r.text}")
            sys.exit(1)
        
        res_data = r.json()
        new_rt = res_data.get('refresh_token')
        
        # 【核心：检测并输出新令牌】
        if new_rt and new_rt != REFRESH_TOKEN:
            print(">>> [Auth] 检测到令牌已滑动更新，正在同步至 GitHub 输出缓冲区...")
            if GITHUB_OUTPUT_FILE:
                with open(GITHUB_OUTPUT_FILE, 'a') as f:
                    f.write(f"NEW_REFRESH_TOKEN={new_rt}\n")
            else:
                print(f"::set-output name=NEW_REFRESH_TOKEN::{new_rt}")
        
        return res_data['access_token']
    except Exception as e:
        print(f"!! 网络鉴权异常: {e}")
        sys.exit(1)

# ================= 🚧 抢占逻辑 =================
def try_lock(token, today_str, current_period):
    headers = {'Authorization': f'Bearer {token}', 'Content-Type': 'application/json'}
    lock_file = f"lock_{today_str}_{current_period}.json"
    lock_url = f"{GRAPH_URL}/me/drive/root:{LOCK_FOLDER}/{lock_file}:/content?@microsoft.graph.conflictBehavior=fail"
    lock_data = {"locked_by": GITHUB_REPO, "time": datetime.utcnow().strftime('%H:%M:%S')}
    try:
        r = requests.put(lock_url, headers=headers, json=lock_data, timeout=30)
        if r.status_code == 201:
            print(f"✅ [Lock] 抢占成功！本时段执行者: {GITHUB_REPO}")
            return True
        elif r.status_code == 409:
            print(f"ℹ️ [Lock] 抢占失败：此时间段已有其他账号在运行。")
            return False
        else:
            print(f"❌ [Lock] 错误: {r.status_code}。请确保网盘已手动创建 {LOCK_FOLDER} 文件夹。")
            return False
    except: return False

# ================= 🚀 业务模块 =================
def task_read_calendar(token):
    headers = {'Authorization': f'Bearer {token}'}
    try: requests.get(f'{GRAPH_URL}/me/events?$top=1', headers=headers, timeout=20)
    except: pass

def task_update_log(token):
    print("\n>>> [Task 2] 更新日志 (CSV)")
    headers = {'Authorization': f'Bearer {token}', 'Content-Type': 'application/json'}
    log_url = f'{GRAPH_URL}/me/drive/root:{DATA_FOLDER}/ActivityLog.csv:/content'
    try:
        old_content = "Time,Repo,Event"
        r = requests.get(log_url, headers=headers, timeout=20)
        if r.status_code == 200:
            old_content = r.text
            lines = old_content.splitlines()
            if len(lines) > 100: old_content = "\n".join(lines[:1] + lines[-99:])
        new_row = f"\n{datetime.now().strftime('%Y-%m-%d %H:%M:%S')},{GITHUB_REPO},KeepAlive_OK"
        requests.put(log_url, headers=headers, data=(old_content + new_row).encode('utf-8'), timeout=20)
        return old_content
    except: return ""

def task_send_mail(token, old_log, today_str):
    if f"{today_str},MAIL_SENT" in old_log: return
    headers = {'Authorization': f'Bearer {token}', 'Content-Type': 'application/json'}
    try:
        r_me = requests.get(f'{GRAPH_URL}/me', headers=headers, timeout=20)
        my_email = r_me.json().get('userPrincipalName')
        mail_data = {
            "message": {
                "subject": f"Office365 KeepAlive: {today_str}",
                "body": {"contentType": "Text", "content": f"运行正常。\n执行账号: {GITHUB_REPO}"},
                "toRecipients": [{"emailAddress": {"address": my_email}}]
            },
            "saveToSentItems": False
        }
        requests.post(f'{GRAPH_URL}/me/sendMail', headers=headers, json=mail_data, timeout=20)
    except: pass

def task_upload_large_file(token):
    headers = {'Authorization': f'Bearer {token}'}
    file_size = random.randint(1, 3) * 1024 * 1024
    file_name = f"Auto_{int(time.time())}.bin"
    try:
        session_url = f'{GRAPH_URL}/me/drive/root:{DATA_FOLDER}/{file_name}:/createUploadSession'
        r_session = requests.post(session_url, headers=headers, json={"item": {"@microsoft.graph.conflictBehavior": "rename"}}, timeout=20)
        upload_url = r_session.json()['uploadUrl']
        requests.put(upload_url, data=b'\0' * file_size, headers={'Content-Length': str(file_size), 'Content-Range': f'bytes 0-{file_size-1}/{file_size}'}, timeout=40)
    except: pass

    def cleanup(folder, suffix):
        try:
            url = f'{GRAPH_URL}/me/drive/root:{folder}:/children?$select=id,name,createdDateTime'
            items = [x for x in requests.get(url, headers=headers, timeout=20).json().get('value', []) if x['name'].endswith(suffix)]
            if len(items) >= 35:
                items.sort(key=lambda x: x['createdDateTime'])
                for item in items[:(len(items) - 3)]:
                    requests.delete(f'{GRAPH_URL}/me/drive/items/{item["id"]}', headers=headers, timeout=20)
        except: pass
    cleanup(DATA_FOLDER, ".bin")
    cleanup(LOCK_FOLDER, ".json")

def main():
    token = get_access_token()
    today_str = datetime.utcnow().strftime('%Y-%m-%d')
    current_period = datetime.utcnow().strftime('%H')
    if try_lock(token, today_str, current_period):
        task_read_calendar(token)
        old_log = task_update_log(token)
        task_send_mail(token, old_log, today_str)
        task_upload_large_file(token)
        print("\n>>> [Done] 任务完成。")
    else:
        sys.exit(0)

if __name__ == '__main__':
    main()
