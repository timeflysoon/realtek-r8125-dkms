import os
import random
import requests
import time
import sys
from datetime import datetime, timezone

# ================= 🛡️ 核心配置区域 =================
CLIENT_ID = os.environ.get('CLIENT_ID')
CLIENT_SECRET = os.environ.get('CLIENT_SECRET')
REFRESH_TOKEN = os.environ.get('REFRESH_TOKEN')
TENANT_ID = os.environ.get('TENANT_ID')
GITHUB_REPO = os.environ.get('GITHUB_REPO_NAME', 'Unknown-Repo')

if not TENANT_ID:
    print("!! 致命错误: 缺少 TENANT_ID 环境变量")
    sys.exit(1)

TOKEN_URL = f'https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token'
GRAPH_URL = 'https://graph.microsoft.com/v1.0'
DATA_FOLDER = "/Data"
LOCK_FOLDER = "/Data/Lock"
REFRESH_TOKEN_FILE = f"{DATA_FOLDER}/RefreshToken.txt"

# ================= 🔐 鉴权模块（最终优化版） =================
def get_access_token(refresh_token):
    print(">>> [Auth] 正在刷新访问令牌...")
    data = {
        'client_id': CLIENT_ID,
        'client_secret': CLIENT_SECRET,
        'refresh_token': refresh_token,
        'grant_type': 'refresh_token',
        'scope': 'Files.ReadWrite.All Mail.Send Calendars.Read User.Read offline_access'
    }
    try:
        r = requests.post(TOKEN_URL, data=data, timeout=30)
        if r.status_code == 200:
            token_data = r.json()
            access_token = token_data['access_token']
            new_refresh_token = token_data.get('refresh_token')
            
            if new_refresh_token and new_refresh_token != refresh_token:
                print("✅ [Auth] 已获取新的 refresh_token")
            return access_token, new_refresh_token
        else:
            print(f"!! 令牌刷新失败: {r.status_code} {r.text[:400]}")
            if r.status_code == 400 and ("invalid_grant" in r.text or "expired" in r.text.lower()):
                print("⚠️  refresh_token 已失效！请使用 rclone 重新生成并更新 GitHub Secrets。")
            return None, None
    except Exception as e:
        print(f"!! 刷新令牌网络异常: {e}")
        return None, None


def save_new_refresh_token(token, new_refresh_token):
    if not new_refresh_token:
        return
    headers = {'Authorization': f'Bearer {token}', 'Content-Type': 'text/plain'}
    try:
        requests.put(
            f'{GRAPH_URL}/me/drive/root:{REFRESH_TOKEN_FILE}:/content',
            headers=headers,
            data=new_refresh_token.encode('utf-8'),
            timeout=15
        )
        print("✅ [Auth] 最新 refresh_token 已备份到 OneDrive")
    except Exception as e:
        print(f"⚠️ 保存 refresh_token 备份失败: {e}")


# ================= 🚧 抢占逻辑 =================
def try_lock(token, today_str, current_period):
    headers = {'Authorization': f'Bearer {token}', 'Content-Type': 'application/json'}
    lock_file = f"lock_{today_str}_{current_period}.json"
    lock_url = f"{GRAPH_URL}/me/drive/root:{LOCK_FOLDER}/{lock_file}:/content?@microsoft.graph.conflictBehavior=fail"
    
    lock_data = {"locked_by": GITHUB_REPO, "time": datetime.now(timezone.utc).strftime('%H:%M:%S')}
    
    try:
        r = requests.put(lock_url, headers=headers, json=lock_data, timeout=20)
        if r.status_code == 201:
            print(f"✅ [Lock] 抢占成功！执行者: {GITHUB_REPO}")
            return True
        elif r.status_code == 409:
            print(f"ℹ️ [Lock] 此时间段已被其他账号占用，跳过。")
            return False
        else:
            print(f"❌ [Lock] 错误 ({r.status_code})，请确保已创建 {LOCK_FOLDER} 文件夹")
            return False
    except Exception as e:
        print(f"⚠️ Lock 请求异常: {e}")
        return False


# ================= 🚀 业务模块 =================
def task_read_calendar(token):
    print("\n>>> [Task 1] 读取日历保活")
    try:
        requests.get(f'{GRAPH_URL}/me/events?$top=1', headers={'Authorization': f'Bearer {token}'}, timeout=15)
        print("    ✅ 操作成功")
    except:
        print("    ⚠️ 日历读取异常（可忽略）")


def task_update_log(token):
    print("\n>>> [Task 2] 更新活动日志")
    headers = {'Authorization': f'Bearer {token}', 'Content-Type': 'application/json'}
    log_url = f'{GRAPH_URL}/me/drive/root:{DATA_FOLDER}/ActivityLog.csv:/content'
    try:
        old_content = "Time,Repo,Event"
        r = requests.get(log_url, headers=headers, timeout=15)
        if r.status_code == 200:
            old_content = r.text
            lines = old_content.splitlines()
            if len(lines) > 100:
                old_content = "\n".join(lines[:1] + lines[-99:])
        
        new_row = f"\n{datetime.now().strftime('%Y-%m-%d %H:%M:%S')},{GITHUB_REPO},KeepAlive_OK"
        requests.put(log_url, headers=headers, data=(old_content + new_row).encode('utf-8'), timeout=15)
        print("    ✅ 日志更新成功")
        return old_content
    except Exception as e:
        print(f"    ⚠️ 日志更新异常: {e}")
        return ""


def task_send_mail(token, old_log, today_str):
    if f"{today_str},MAIL_SENT" in old_log:
        print("\n>>> [Task 3] 邮件跳过：今日已有账号发送。")
        return
   
    print("\n>>> [Task 3] 发送每日提醒邮件")
    headers = {'Authorization': f'Bearer {token}', 'Content-Type': 'application/json'}
    try:
        r_me = requests.get(f'{GRAPH_URL}/me', headers=headers, timeout=10)
        my_email = r_me.json().get('userPrincipalName')
        
        mail_data = {
            "message": {
                "subject": f"Office365 KeepAlive: {today_str}",
                "body": {"contentType": "Text", "content": f"系统运行正常。\n执行账号: {GITHUB_REPO}"},
                "toRecipients": [{"emailAddress": {"address": my_email}}]
            },
            "saveToSentItems": False
        }
        resp = requests.post(f'{GRAPH_URL}/me/sendMail', headers=headers, json=mail_data, timeout=20)
        if resp.status_code in [200, 202]:
            mark = f"\n{datetime.now().strftime('%Y-%m-%d %H:%M:%S')},{today_str},MAIL_SENT"
            requests.put(f'{GRAPH_URL}/me/drive/root:{DATA_FOLDER}/ActivityLog.csv:/content',
                        headers=headers, data=(old_log + mark).encode('utf-8'))
            print(f"    ✅ 邮件已发送至 {my_email}")
    except Exception as e:
        print(f"    ⚠️ 邮件发送异常: {e}")


def task_upload_large_file(token):
    print("\n>>> [Task 4] 模拟大文件操作 + 自动清理")
    headers = {'Authorization': f'Bearer {token}'}
    try:
        file_size = random.randint(1, 5) * 1024 * 1024
        file_name = f"Auto_{int(time.time())}.bin"
        
        session_url = f'{GRAPH_URL}/me/drive/root:{DATA_FOLDER}/{file_name}:/createUploadSession'
        r_session = requests.post(session_url, headers=headers, 
                                json={"item": {"@microsoft.graph.conflictBehavior": "rename"}}, timeout=15)
        upload_url = r_session.json()['uploadUrl']
        
        requests.put(upload_url, data=b'\0' * file_size,
                    headers={'Content-Length': str(file_size), 'Content-Range': f'bytes 0-{file_size-1}/{file_size}'}, 
                    timeout=60)
        print(f"    ✅ 上传完成: {file_name} ({file_size//(1024*1024)}MB)")
    except Exception as e:
        print(f"    ⚠️ 文件上传异常: {e}")

    # 自动清理
    def cleanup(folder, suffix):
        try:
            url = f'{GRAPH_URL}/me/drive/root:{folder}:/children?$select=id,name,createdDateTime'
            items = [x for x in requests.get(url, headers=headers, timeout=15).json().get('value', [])
                    if x['name'].endswith(suffix)]
            if len(items) >= 35:
                items.sort(key=lambda x: x['createdDateTime'])
                for item in items[:(len(items) - 3)]:
                    requests.delete(f'{GRAPH_URL}/me/drive/items/{item["id"]}', headers=headers, timeout=10)
                    print(f"        -> 清理: {item['name']}")
        except:
            pass
    cleanup(DATA_FOLDER, ".bin")
    cleanup(LOCK_FOLDER, ".json")


# ================= 🏁 主入口 =================
def main():
    if not REFRESH_TOKEN:
        print("!! 致命错误: GitHub Secrets 中缺少 REFRESH_TOKEN")
        sys.exit(1)

    token = None
    new_refresh = None

    # 使用 Secrets 中的 refresh_token
    token, new_refresh = get_access_token(REFRESH_TOKEN)

    if not token:
        print("!! 无法获取访问令牌，请检查 GitHub Secrets 中的 REFRESH_TOKEN 是否有效。")
        sys.exit(1)

    # 保存最新 refresh_token 作为备份
    if new_refresh:
        save_new_refresh_token(token, new_refresh)

    today_str = datetime.now(timezone.utc).strftime('%Y-%m-%d')
    current_period = datetime.now(timezone.utc).strftime('%H')

    if try_lock(token, today_str, current_period):
        task_read_calendar(token)
        old_log = task_update_log(token)
        task_send_mail(token, old_log, today_str)
        task_upload_large_file(token)
        print("\n>>> [Done] 本次任务圆满完成。")
    else:
        print("\n>>> [Skip] 任务已被其他账号执行，本次跳过。")

if __name__ == '__main__':
    main()
