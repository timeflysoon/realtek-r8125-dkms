import os
import random
import string
import requests
import time
from datetime import datetime

# ================= 配置区域 =================
CLIENT_ID = os.environ.get('Z_CLIENT_ID')
CLIENT_SECRET = os.environ.get('Z_CLIENT_SECRET')
REFRESH_TOKEN = os.environ.get('Z_REFRESH_TOKEN')

TOKEN_URL = 'https://login.microsoftonline.com/common/oauth2/v2.0/token'
GRAPH_URL = 'https://graph.microsoft.com/v1.0'
DATA_FOLDER = "/Data"

# ================= 核心：带有详细报错的 Token 获取函数 =================
def get_access_token():
    print(">>> [Auth] 正在刷新令牌...")
    
    # 检查 Secrets 是否读取成功
    if not CLIENT_ID or not CLIENT_SECRET or not REFRESH_TOKEN:
        print("!! [致命错误] GitHub Secrets 读取失败！")
        print(f"   Z_CLIENT_ID: {'有值' if CLIENT_ID else '空'}")
        print(f"   Z_CLIENT_SECRET: {'有值' if CLIENT_SECRET else '空'}")
        print(f"   Z_REFRESH_TOKEN: {'有值' if REFRESH_TOKEN else '空'}")
        exit(1)

    data = {
        'client_id': CLIENT_ID,
        'client_secret': CLIENT_SECRET,
        'refresh_token': REFRESH_TOKEN,
        'grant_type': 'refresh_token',
        'scope': 'Files.ReadWrite.All Mail.Send Calendars.Read User.Read offline_access'
    }
    
    # 发送请求
    r = requests.post(TOKEN_URL, data=data)
    
    # --- 调试核心：把微软的拒信打印出来 ---
    if r.status_code != 200:
        print("\n" + "="*40)
        print(f"!! [致命错误] 状态码: {r.status_code}")
        print(f"!! [微软原话]: {r.text}")  # <--- 这一行是解题的关键
        print("="*40 + "\n")
        
        if "invalid_client" in r.text:
            print(">> [诊断]: Client Secret 填错了！你可能填成了 Secret ID。")
            print(">> 请去 Azure 重新生成一个，务必复制那个【只显示一次的 Value】。")
        elif "invalid_grant" in r.text:
            print(">> [诊断]: Refresh Token 无效。可能是复制不全，或账号密码改过。")
        elif "unauthorized_client" in r.text:
            print(">> [诊断]: Client ID 错，或者 rclone 获取 token 时没有勾选 offline_access。")
            
        exit(1)
        
    return r.json()['access_token']

# ================= 其他辅助函数 (保持不变) =================
def get_me(headers):
    try:
        r = requests.get(f'{GRAPH_URL}/me', headers=headers)
        if r.status_code == 200: return r.json().get('userPrincipalName')
    except: pass
    return None

def random_string(length=8):
    return ''.join(random.choices(string.ascii_letters + string.digits, k=length))

def task_read_calendar(headers):
    print("\n>>> [Task 1] 读取日历")
    try:
        requests.get(f'{GRAPH_URL}/me/events?$top=1', headers=headers)
    except: pass

def task_update_log(headers):
    print("\n>>> [Task 2] 更新日志")
    try:
        filename = "ActivityLog.csv"
        new_row = f"\n{datetime.now().strftime('%Y-%m-%d %H:%M:%S')},AutoRun,OK,{random.randint(1000,9999)}"
        url = f'{GRAPH_URL}/me/drive/root:{DATA_FOLDER}/{filename}:/content'
        
        # 简单追加模式：先读后写
        r = requests.get(url, headers=headers)
        current = r.text if r.status_code == 200 else "Time,Status,ID"
        requests.put(url, headers=headers, data=current + new_row)
        print("   日志更新成功")
    except: print("   日志更新跳过")

def task_send_mail(headers, email):
    print("\n>>> [Task 3] 发送邮件")
    if not email: return
    try:
        data = {
            "message": {
                "subject": f"Report: {random_string(5)}",
                "body": {"contentType": "Text", "content": "Auto Run OK"},
                "toRecipients": [{"emailAddress": {"address": email}}]
            }
        }
        requests.post(f'{GRAPH_URL}/me/sendMail', headers=headers, json=data)
        print("   邮件发送成功")
    except: pass

def task_file_manage(headers):
    print("\n>>> [Task 4] 文件管理")
    try:
        # 上传 1MB 随机文件
        name = f"KeepAlive_{int(time.time())}.bin"
        url = f'{GRAPH_URL}/me/drive/root:{DATA_FOLDER}/{name}:/content'
        requests.put(url, headers=headers, data=b'\0'*1024*1024)
        print(f"   上传成功: {name}")
        
        # 清理
        r = requests.get(f'{GRAPH_URL}/me/drive/root:{DATA_FOLDER}:/children', headers=headers)
        if r.status_code == 200:
            items = [x for x in r.json().get('value',[]) if x['name'].endswith('.bin')]
            if len(items) > 25:
                items.sort(key=lambda x: x['createdDateTime'])
                for i in range(len(items) - 25):
                    requests.delete(f'{GRAPH_URL}/me/drive/items/{items[i]["id"]}', headers=headers)
                    print(f"   删除旧文件: {items[i]['name']}")
    except Exception as e: print(f"   文件操作异常: {e}")

def main():
    token = get_access_token()
    headers = {'Authorization': f'Bearer {token}', 'Content-Type': 'application/json'}
    me = get_me(headers)
    print(f"当前用户: {me}")
    
    task_read_calendar(headers)
    task_update_log(headers)
    task_send_mail(headers, me)
    task_file_manage(headers)
    print("\n>>> 完成")

if __name__ == '__main__':
    main()
