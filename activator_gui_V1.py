import tkinter as tk
from tkinter import messagebox, simpledialog
import subprocess
import os

# Configurable paths
OFFICE_PATH = r"C:\Program Files\Microsoft Office\Office16"
OSPP = os.path.join(OFFICE_PATH, "OSPP.VBS")
WIN_KEY_FILE = "windows_keys.txt"
OFFICE_KEY_FILE = "office_keys.txt"
IID_FILE = "installation_id.txt"
LOG_FILE = "activation_log.txt"

# Global upgrade key to Pro
UPGRADE_KEY = "VNV2J-64BH6-G6Y6K-V7MVH-3J3GT"

def read_next_key(file):
    if not os.path.exists(file):
        return None, []
    with open(file, "r") as f:
        keys = f.readlines()
    if not keys:
        return None, []
    return keys[0].strip(), keys[1:]

def update_key_file(file, remaining_keys):
    with open(file, "w") as f:
        f.writelines(remaining_keys)

def log_key(action, key):
    with open(LOG_FILE, "a") as f:
        f.write(f"{action}: {key}\n")

def activate_windows():
    key, remaining = read_next_key(WIN_KEY_FILE)
    if not key:
        messagebox.showinfo("No Key", "No Windows keys left.")
        return
    try:
        subprocess.run(["slmgr.vbs", "/ipk", key], check=True)
        subprocess.run(["slmgr.vbs", "/ato"], check=True)
        update_key_file(WIN_KEY_FILE, remaining)
        log_key("Windows Activated", key)
        messagebox.showinfo("Success", f"✅ Windows activated with:\n{key}")
    except subprocess.CalledProcessError:
        messagebox.showerror("Error", "❌ Activation failed.")

def upgrade_windows_to_pro():
    try:
        subprocess.run(["slmgr.vbs", "/ipk", UPGRADE_KEY], check=True)
        messagebox.showinfo("Upgrade Triggered", "🪟 Windows will upgrade to Pro. A reboot may be required.")
        log_key("Windows Edition Upgrade Triggered", UPGRADE_KEY)
    except subprocess.CalledProcessError:
        messagebox.showerror("Error", "❌ Failed to upgrade to Windows Pro.")

def generate_office_iid():
    key, remaining = read_next_key(OFFICE_KEY_FILE)
    if not key:
        messagebox.showinfo("No Key", "No Office keys left.")
        return
    try:
        subprocess.run(["cscript", OSPP, f"/inpkey:{key}"], check=True)
        result = subprocess.run(["cscript", OSPP, "/dinstid"], capture_output=True, text=True)
        with open(IID_FILE, "w") as f:
            f.write(result.stdout)
        update_key_file(OFFICE_KEY_FILE, remaining)
        log_key("Office Key Set, IID Generated", key)
        messagebox.showinfo("IID Generated", "📥 Installation ID saved to installation_id.txt.\nSubmit to Microsoft to get CID.")
    except subprocess.CalledProcessError as e:
        messagebox.showerror("Error", f"Failed to generate IID.\n{e}")

def activate_office_with_cid():
    cid = simpledialog.askstring("CID Input", "Enter Confirmation ID (no dashes):")
    if not cid:
        return
    try:
        subprocess.run(["cscript", OSPP, f"/actcid:{cid}"], check=True)
        messagebox.showinfo("Activated", "✅ Office activated successfully.")
    except subprocess.CalledProcessError:
        messagebox.showerror("Error", "❌ Activation failed. CID might be incorrect.")

# GUI Setup
root = tk.Tk()
root.title("Windows + Office Activation Tool")

frame = tk.Frame(root, padx=20, pady=20)
frame.pack()

tk.Label(frame, text="Activation Tool", font=("Arial", 16)).pack(pady=10)

tk.Button(frame, text="🪟 Activate Windows", width=30, command=activate_windows).pack(pady=5)
tk.Button(frame, text="🔁 Upgrade Windows to Pro", width=30, command=upgrade_windows_to_pro).pack(pady=5)
tk.Button(frame, text="📥 Generate Office IID", width=30, command=generate_office_iid).pack(pady=5)
tk.Button(frame, text="🔑 Enter Office CID to Activate", width=30, command=activate_office_with_cid).pack(pady=5)

root.mainloop()
