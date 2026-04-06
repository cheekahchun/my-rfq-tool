import os
import json
import streamlit as st

# --- 路径配置 ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
USERS_FILE = os.path.join(BASE_DIR, "users.json")
CONFIG_PATH = os.path.join(BASE_DIR, "config.txt")

# --- 内部配置读取 ---
def _get_config():
    """读取配置，确保多用户能从 Settings 配置自己的登录账号"""
    cfg = {}
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                for line in f:
                    if "=" in line and not line.strip().startswith("#"):
                        k, v = line.split("=", 1)
                        cfg[k.strip()] = v.strip()
        except: pass
    return cfg

# --- 会话初始化 ---
def init_session():
    """初始化 Streamlit 会话状态，修复刚才的 AttributeError"""
    if 'logged_in' not in st.session_state:
        st.session_state['logged_in'] = False
    if 'user_info' not in st.session_state:
        st.session_state['user_info'] = None

# --- 用户持久化管理 ---
def load_users():
    """从 users.json 加载用户，如果文件不存在则创建一个默认 admin"""
    if os.path.exists(USERS_FILE):
        try:
            with open(USERS_FILE, "r") as f:
                return json.load(f)
        except: return []
    else:
        # 创建默认用户
        default_users = [
            {"username": "admin", "password": "admin123", "name": "System Admin", "role": "admin", "is_active": True}
        ]
        save_users(default_users)
        return default_users

def save_users(users):
    """保存用户到 json 文件"""
    with open(USERS_FILE, "w") as f:
        json.dump(users, f, indent=4)

# --- 核心业务逻辑 ---
def check_login(username, password):
    """验证登录逻辑：同时支持配置文件的 Master 密钥和 JSON 数据库"""
    # 1. 检查万能 Master 账号 (从 config.txt 加载)
    cfg = _get_config()
    master_user = cfg.get("APP_USERNAME", "admin")
    master_pw = cfg.get("APP_PASSWORD", "admin123")
    
    if username == master_user and password == master_pw:
        return True, {"username": username, "name": "Master Administrator", "role": "admin"}
    
    # 2. 检查普通数据库用户
    users = load_users()
    for u in users:
        if u['username'] == username and u['password'] == password:
            if u.get('is_active', True):
                return True, u
            else:
                return False, "🚫 账号已被禁用，请联系管理员"
    return False, "⚠️ 账号或密码错误"

def add_user(username, password, name, role):
    """添加新用户"""
    users = load_users()
    if any(u['username'] == username for u in users):
        return False, "❌ 用户名已存在"
    users.append({"username": username, "password": password, "name": name, "role": role, "is_active": True})
    save_users(users)
    return True, "✅ 用户创建成功"

def update_user_status(username, is_active):
    """启用/禁用用户"""
    users = load_users()
    for u in users:
        if u['username'] == username:
            u['is_active'] = is_active
            save_users(users)
            return True
    return False

def logout():
    """退出登录逻辑"""
    st.session_state['logged_in'] = False
    st.session_state['user_info'] = None
