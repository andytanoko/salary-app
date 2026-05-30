"""
create_user.py — Add or update users in users.json
Run from your payslip_app folder:
    python create_user.py
"""
import os
import json
import bcrypt
import getpass

_BASE_DIR  = os.path.dirname(os.path.abspath(__file__))
USERS_FILE = os.path.join(_BASE_DIR, 'users.json')


def load_users():
    if not os.path.isfile(USERS_FILE):
        return {}
    with open(USERS_FILE, 'r') as f:
        return json.load(f)


def save_users(users):
    with open(USERS_FILE, 'w') as f:
        json.dump(users, f, indent=2)
    print(f"Saved to {USERS_FILE}")


def add_user():
    users    = load_users()
    username = input("Username: ").strip()
    if not username:
        print("Username cannot be empty.")
        return

    if username in users:
        confirm = input(f"User '{username}' already exists. Update password? (y/n): ").strip().lower()
        if confirm != 'y':
            print("Cancelled.")
            return

    while True:
        password = getpass.getpass("Password: ")
        confirm  = getpass.getpass("Confirm password: ")
        if password != confirm:
            print("Passwords do not match. Try again.")
        elif len(password) < 6:
            print("Password must be at least 6 characters.")
        else:
            break

    hashed = bcrypt.hashpw(password.encode('utf-8'), bcrypt.gensalt()).decode('utf-8')
    users[username] = {"password": hashed}
    save_users(users)
    print(f"User '{username}' saved successfully.")


def list_users():
    users = load_users()
    if not users:
        print("No users found.")
        return
    print("Existing users:")
    for u in users:
        print(f"  - {u}")


def delete_user():
    users = load_users()
    list_users()
    username = input("Username to delete: ").strip()
    if username not in users:
        print(f"User '{username}' not found.")
        return
    confirm = input(f"Delete '{username}'? (y/n): ").strip().lower()
    if confirm == 'y':
        del users[username]
        save_users(users)
        print(f"User '{username}' deleted.")


if __name__ == '__main__':
    print("=== Otopia Payslip — User Manager ===\n")
    print("1. Add / update user")
    print("2. List users")
    print("3. Delete user")
    choice = input("\nChoice (1/2/3): ").strip()
    if choice == '1':
        add_user()
    elif choice == '2':
        list_users()
    elif choice == '3':
        delete_user()
    else:
        print("Invalid choice.")
