import os
import streamlit as st

# 获取 config.txt 的路径
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(BASE_DIR, "config.txt")

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

def check_login(username, password):
    """验证逻辑：优先从 config.txt 里查，没有就用默认的 admin / admin123"""
    cfg = _get_config()
    saved_user = cfg.get("APP_USERNAME", "admin")
    saved_pw = cfg.get("APP_PASSWORD", "admin123")
    
    if username == saved_user and password == saved_pw:
        return True, {"name": "System Administrator", "role": "admin"}
    return False, "⚠️ 账号或密码错误 (Invalid credentials)"

def logout():
    """退出登录并清理会话"""
    if 'logged_in' in st.session_state:
        del st.session_state['logged_in']
    if 'user_info' in st.session_state:
        del st.session_state['user_info']
