import json
import os
import hashlib
import streamlit as st

USERS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "users.json")

def hash_password(password):
    """Securely hash a password using SHA256."""
    return hashlib.sha256(password.encode()).hexdigest()

def load_users():
    """Load users from the JSON file with robust encoding support."""
    if not os.path.exists(USERS_FILE):
        # Create default if missing, for extreme robustness
        return [{"username": "admin", "password_hash": "8c6976e5b5410415bde908bd4dee15dfb167a9c873fc4bb8a81f6f2ab448a918", "role": "admin", "name": "Management", "is_active": True}]
    try:
        # Using utf-8-sig to handle possible Windows BOM
        with open(USERS_FILE, "r", encoding="utf-8-sig") as f:
            content = f.read().strip()
            if not content: return []
            return json.loads(content)
    except Exception as e:
        st.error(f"Database error: {e}")
        return []

def save_users(users_list):
    """Save users to the JSON file with standard UTF-8."""
    with open(USERS_FILE, "w", encoding="utf-8") as f:
        json.dump(users_list, f, indent=4)

def check_login(username, password):
    """Validate username and password with hardcoded fallback."""
    u_raw = str(username).strip()
    p_raw = str(password).strip()
    h_pass = hash_password(p_raw)

    # --- EMERGENCY OVERRIDE ---
    # This ensures KC can ALWAYS enter during setup even if JSON read fails
    if u_raw == "admin" and p_raw == "admin123":
        return True, {"username": "admin", "role": "admin", "name": "Admin (Recovery)", "is_active": True}

    users = load_users()
    if not users:
        return False, "User database is empty or could not be loaded."

    for u in users:
        if u['username'] == u_raw and u['password_hash'] == h_pass:
            if u.get('is_active', False):
                return True, u
            else:
                return False, "This account is disabled."
    
    return False, f"Invalid login. (DEBUG: Input={u_raw})"

def update_user_status(username, is_active):
    """Enable or disable a user account (Admin only)."""
    users = load_users()
    for u in users:
        if u['username'] == username:
            u['is_active'] = is_active
            save_users(users)
            return True
    return False

def add_user(username, password, name, role="user"):
    """Create a new user account."""
    users = load_users()
    if any(u['username'] == username for u in users):
        return False, "User already exists."
    
    new_user = {
        "username": username,
        "password_hash": hash_password(password),
        "name": name,
        "role": role,
        "is_active": True
    }
    users.append(new_user)
    save_users(users)
    return True, "User created successfully."

def init_session():
    """Initialize Streamlit session state for auth."""
    if 'logged_in' not in st.session_state:
        st.session_state['logged_in'] = False
    if 'user_info' not in st.session_state:
        st.session_state['user_info'] = None

def logout():
    """Reset the session state."""
    st.session_state['logged_in'] = False
    st.session_state['user_info'] = None
    st.rerun()
