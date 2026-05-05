# ================= 🔐 鉴权模块 (已优化) =================
def get_access_token():
    print(">>> [Auth] 刷新令牌...")
    data = {
        'client_id': CLIENT_ID,
        'client_secret': CLIENT_SECRET,
        'refresh_token': REFRESH_TOKEN,
        'grant_type': 'refresh_token',
        'scope': 'Files.ReadWrite.All Mail.Send Calendars.Read User.Read offline_access'
    }
    try:
        r = requests.post(TOKEN_URL, data=data)
        if r.status_code != 200:
            print(f"!! 令牌刷新失败: {r.text}")
            sys.exit(1)
        
        res_data = r.json()
        new_rt = res_data.get('refresh_token')
        
        # 【核心：如果微软返回了新 RT，通过环境变量输出给 GitHub Action】
        if new_rt and new_rt != REFRESH_TOKEN:
            print(f"::set-output name=NEW_REFRESH_TOKEN::{new_rt}")
            print(">>> [Auth] 检测到新刷新令牌，已准备同步。")
            
        return res_data['access_token']
    except Exception as e:
        print(f"!! 网络异常: {e}")
        sys.exit(1)
