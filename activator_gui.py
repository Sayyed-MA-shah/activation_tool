import requests
import tkinter as tk
from tkinter import messagebox, simpledialog
import subprocess
from tkinter import scrolledtext
import win32file
import string
import os
import re
import sys
import platform
import webbrowser
import pyperclip
import requests
import json
from tkinter import messagebox
from tkinter import messagebox, Toplevel, Listbox, Scrollbar, Button, Label
import ctypes, sys

def run_as_admin():
    if not ctypes.windll.shell32.IsUserAnAdmin():
        # Relaunch the script with admin privileges
        ctypes.windll.shell32.ShellExecuteW(
            None, "runas", sys.executable, ' '.join(sys.argv), None, 1
        )
        sys.exit()

run_as_admin()  # Run this early in your script


def log_error(message):
    with open("errorlog.txt", "a") as log:
        log.write(f"[CLIENT ERROR] {message}\n")

# Server settings
SERVER_URL = "http://192.168.0.72:5000"  # ✅ Base URL only
API_KEY = "secure123"

def fetch_key(key_type):
    try:
        response = requests.post(
            f"{SERVER_URL}/get-key",  # ✅ Correct endpoint
            json={"type": key_type},
            headers={"Authorization": f"Bearer {API_KEY}"}
        )

        if response.status_code == 200:
            key = response.json().get("key")
            if key:
                return key
            else:
                log_error(f"No key returned for type '{key_type}' from server.")
                return None
        else:
            log_error(f"Server error {response.status_code}: {response.text}")
            return None
    except Exception as e:
        log_error(f"Failed to connect to key server: {e}")
        return None




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
        # _______________________________________
        # WINDOW ACTIVATION
        # _______________________________________
def activate_windows():
    activation_key = fetch_key("windows")
    if not activation_key:
        messagebox.showerror("Error", "No activation key available.")
        return

    try:
        # Try using Sysnative to avoid redirection issues on 64-bit systems
        slmgr_path = os.path.join(os.environ.get('WINDIR', 'C:\\Windows'), "Sysnative", "slmgr.vbs")
        if not os.path.exists(slmgr_path):
            # Fallback to System32 if Sysnative not found
            slmgr_path = os.path.join(os.environ.get('WINDIR', 'C:\\Windows'), "System32", "slmgr.vbs")

        # Inject the key
        ipk_result = subprocess.run(
            ["cscript.exe", slmgr_path, "/ipk", activation_key],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            shell=True
        )

        if ipk_result.returncode != 0:
            messagebox.showerror("Key Error", f"❌ Failed to install key:\n{ipk_result.stderr}\n{ipk_result.stdout}")
            return

        # Activate
        ato_result = subprocess.run(
            ["cscript.exe", slmgr_path, "/ato"],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            shell=True
        )

        if ato_result.returncode != 0:
            messagebox.showerror("Activation Error", f"❌ Activation failed:\n{ato_result.stderr}\n{ato_result.stdout}")
            return

        messagebox.showinfo("Success", f"✅ Windows activated successfully:\n{ato_result.stdout.strip()}")

    except Exception as e:
        messagebox.showerror("Exception", f"❌ Activation exception:\n{e}")
# def activate_windows():
#     activation_key = fetch_key("windows")
#     if not activation_key:
#         messagebox.showerror("Error", "No activation key available.")
#         return

#     try:
#         slmgr_path = os.path.join(os.environ['SystemRoot'], "System32", "slmgr.vbs")
#         subprocess.run(["cscript", slmgr_path, "/ipk", activation_key], check=True)
#         subprocess.run(["cscript", slmgr_path, "/ato"], check=True)
#         messagebox.showinfo("Success", "Windows activated successfully.")
#     except subprocess.CalledProcessError as e:
#         log_error(f"Activation failed: {e}")
#         messagebox.showerror("Activation Failed", f"Failed to activate Windows:\n{e}")
# ____________________________
# WINDOR PRO
# ____________________________
def upgrade_windows_to_pro():
    try:
        subprocess.run(
            ["powershell", "changepk.exe", "/ProductKey", UPGRADE_KEY],
            check=True
        )
        messagebox.showinfo("Success", "Windows upgrade initiated. Your system may restart.")
    except subprocess.CalledProcessError as e:
        log_error(f"Windows Pro upgrade failed: {e}")
        messagebox.showerror("Upgrade Failed", f"Failed to upgrade to Windows Pro:\n{e}")

# ________________
# Office new code 7 / 21
# ________________
def fetch_and_set_office_key():
    try:
        response = requests.post(
            f"{SERVER_URL}/get-key",  # ✅ Correct endpoint
            headers={"Authorization": f"Bearer {API_KEY}"},
            json={"type": "office"}   # ✅ This must match the "office" key in key.json
        )

        print("Raw response:", response.text)  # Optional debug line

        if response.status_code != 200:
            messagebox.showerror("Server Error", f"❌ Server error:\n{response.text}")
            return

        key = response.json().get("key")
        if not key:
            messagebox.showerror("Key Error", "❌ No Office key received.")
            return

        subprocess.run(["cscript", OSPP, f"/inpkey:{key}"], check=True)
        messagebox.showinfo("Office Key Set", f"✅ Office key set:\n{key}")

    except Exception as e:
        messagebox.showerror("Error", f"❌ Failed to set Office key:\n{e}")

                            # ___________________________
                            # OFFICE IID GENERATION
                            # __________________________
def generate_office_iid():
    try:
        # Run the command to get the Installation ID
        result = subprocess.run(
            ["cscript", OSPP, "/dinstid"],
            capture_output=True,
            text=True,
            check=True
        )

        output = result.stdout
        print("IID Output:\n", output)

        # Save the Installation ID to a file
        with open("installation_id.txt", "w", encoding="utf-8") as f:
            f.write(output)

        messagebox.showinfo("IID Generated", "📥 Installation ID saved to installation_id.txt.\nSubmit this to Microsoft to get your CID.")
    except subprocess.CalledProcessError as e:
        messagebox.showerror("Error", f"❌ Failed to generate Installation ID:\n{e}")
                    # ___________________________
                    # Office IID COPYING
                    # ___________________________
def fast_office_iid_flow():
    try:
        result = subprocess.run(
            ["cscript", OSPP, "/dinstid"],
            capture_output=True,
            text=True,
            check=True
        )
        output = result.stdout

        # Save raw output
        with open("installation_id.txt", "w", encoding="utf-8") as f:
            f.write(output)

        # Extract just the IID digits (clean format)
        iid_match = re.search(r"Installation ID:\s+([\d\s]+)", output, re.IGNORECASE)
        if iid_match:
            iid = iid_match.group(1).strip().replace(" ", "")
            pyperclip.copy(iid)
            messagebox.showinfo("✅ IID Ready", "Installation ID copied to clipboard.\nOpening GetCID...")
        else:
            pyperclip.copy(output)
            messagebox.showwarning("⚠️ Output Copied", "Could not extract clean IID.\nRaw output copied instead.")

        # Open GetCID.info
        webbrowser.open("https://getcid.info")

    except Exception as e:
        messagebox.showerror("❌ Error", f"Failed to extract or copy IID:\n{e}")


                    # ___________________________
                    #   CONNECT TO WIFI
                    # ___________________________
def connect_to_wifi():
    ssid = "No Internet"
    password = "4sSsy34ux472547"

    # Create the network profile XML
    import subprocess
import os
from tkinter import messagebox

def connect_to_wifi():
    ssid = "No Internet"
    password = "4sSsy34ux472547"

    wifi_profile = f"""<?xml version="1.0"?>
    <WLANProfile xmlns="http://www.microsoft.com/networking/WLAN/profile/v1">
        <name>{ssid}</name>
        <SSIDConfig>
            <SSID>
                <name>{ssid}</name>
            </SSID>
        </SSIDConfig>
        <connectionType>ESS</connectionType>
        <connectionMode>auto</connectionMode>
        <MSM>
            <security>
                <authEncryption>
                    <authentication>WPA2PSK</authentication>
                    <encryption>AES</encryption>
                    <useOneX>false</useOneX>
                </authEncryption>
                <sharedKey>
                    <keyType>passPhrase</keyType>
                    <protected>false</protected>
                    <keyMaterial>{password}</keyMaterial>
                </sharedKey>
            </security>
        </MSM>
    </WLANProfile>
    """

    profile_path = "wifi_profile.xml"
    with open(profile_path, "w") as f:
        f.write(wifi_profile)

    try:
        result_add = subprocess.run(["netsh", "wlan", "add", "profile", f"filename={profile_path}"], check=True)
        result_connect = subprocess.run(["netsh", "wlan", "connect", f"name={ssid}"], check=True)

        if result_connect.returncode == 0:
            messagebox.showinfo("WiFi", f"✅ Connected to WiFi: {ssid}")
        else:
            messagebox.showerror("WiFi Error", f"❌ Failed to connect to WiFi.")

    except subprocess.CalledProcessError as e:
        messagebox.showerror("WiFi Error", f"❌ Wifi connection failed:\n{e}")
    finally:
        if os.path.exists(profile_path):
            os.remove(profile_path)



# ___________________________

#   LAPTOP MODEL DETECTOR
# ___________________________
def get_laptop_model():
    try:
        result = subprocess.run(
            ["wmic", "computersystem", "get", "model"],
            capture_output=True,
            text=True,
            check=True
        )
        lines = result.stdout.strip().splitlines()
        if len(lines) > 1:
            return lines[1].strip()
    except Exception as e:
        log_error(f"Failed to detect laptop model: {e}")
    return "Unknown Model"

def install_drivers_automatically():
    model = get_laptop_model()
    if model == "Unknown Model":
        messagebox.showerror("Model Error", "❌ Could not detect the laptop model.")
        return

    try:
        with open("drivers.json", "r") as f:
            drivers = json.load(f)

        if model not in drivers:
            messagebox.showwarning("Unsupported Model", f"No driver config found for: {model}")
            return

        driver_urls = drivers[model]
        for name, url in driver_urls.items():
            local_file = f"{name.lower()}_driver.exe"
            messagebox.showinfo("Downloading", f"📥 Downloading {name} driver for {model}")
            response = requests.get(url)
            with open(local_file, "wb") as f:
                f.write(response.content)

            subprocess.run([local_file, "/quiet", "/norestart"], check=False)

        messagebox.showinfo("Driver Installation", f"✅ All drivers installed for {model}")
    except Exception as e:
        log_error(f"Driver installation failed: {e}")
        messagebox.showerror("Error", f"❌ Failed to install drivers: {e}")
# ___________________________
# DRIVER USB DETECTION
# ___________________________
def pick_and_install_usb_driver():
    import os
    import string
    import subprocess
    import win32file
    from tkinter import Toplevel, Listbox, Button, END, messagebox

    relative_dir = os.path.join("Drivers", "Laptops", "Lenovo", "IdeaPad")
    usb_drive = None
    exe_files = []

    # Detect USB drive
    for letter in string.ascii_uppercase:
        drive = f"{letter}:\\"
        try:
            if win32file.GetDriveType(drive) == win32file.DRIVE_REMOVABLE:
                possible_path = os.path.join(drive, relative_dir)
                if os.path.exists(possible_path):
                    usb_drive = drive
                    # Collect all .exe files in that folder
                    exe_files = [f for f in os.listdir(possible_path) if f.lower().endswith(".exe")]
                    full_paths = [os.path.join(possible_path, f) for f in exe_files]
                    break
        except Exception:
            continue

    if not usb_drive or not exe_files:
        messagebox.showerror("Not Found", "❌ No .exe drivers found in USB at expected path.")
        return

    # Create a GUI to select which driver to install
    driver_window = Toplevel()
    driver_window.title("Select Driver to Install")

    listbox = Listbox(driver_window, height=10, width=50)
    for file in exe_files:
        listbox.insert(END, file)
    listbox.pack(padx=10, pady=10)

    def install_selected_driver():
        selection = listbox.curselection()
        if not selection:
            messagebox.showwarning("No Selection", "⚠️ Please select a driver.")
            return
        selected_file = exe_files[selection[0]]
        driver_path = os.path.join(usb_drive, relative_dir, selected_file)
        try:
            subprocess.run(driver_path, check=True)
            messagebox.showinfo("Success", f"✅ Driver installed: {selected_file}")
            driver_window.destroy()
        except subprocess.CalledProcessError:
            messagebox.showerror("Error", f"❌ Failed to install {selected_file}")

    install_button = Button(driver_window, text="Install", command=install_selected_driver)
    install_button.pack(pady=(0, 10))

    # ✅ Bind Enter key to trigger install
    driver_window.bind("<Return>", lambda event: install_button.invoke())




# ___________________________
# End Usb DRIVER DETECTION
# ___________________________

 # OFFICE VERSION
def check_office_activation():
    office_paths = [
        r"C:\Program Files\Microsoft Office\Office16\ospp.vbs",
        r"C:\Program Files (x86)\Microsoft Office\Office16\ospp.vbs",
        r"C:\Program Files\Microsoft Office\Office15\ospp.vbs",
        r"C:\Program Files (x86)\Microsoft Office\Office15\ospp.vbs",
    ]

    for path in office_paths:
        try:
            result = subprocess.run(
                ["cscript", "//Nologo", path, "/dstatus"],
                capture_output=True,
                text=True,
                timeout=10
            )
            output = result.stdout
            if "LICENSE NAME" in output:
                license_name_match = re.search(r"LICENSE NAME:\s*(.+)", output)
                license_status_match = re.search(r"LICENSE STATUS:\s*(.+)", output)

                license_name = license_name_match.group(1) if license_name_match else "Unknown"
                license_status = license_status_match.group(1) if license_status_match else "Unknown"

                if "Retail" in license_name:
                    edition_type = "Retail"
                    warning = "⚠️ RETAIL Edition Detected:\nCID activation will NOT work with this Office install.\n"
                elif "Volume" in license_name or "VL" in license_name:
                    edition_type = "Volume"
                    warning = "✅ VOLUME Edition Detected:\nThis Office install supports CID activation.\n"
                else:
                    edition_type = "Unknown"
                    warning = "⚠️ Unable to determine if Office is Retail or Volume.\n"

                result_window = tk.Toplevel(root)
                result_window.title("Office Activation Info")
                result_window.geometry("500x300")
                result_window.resizable(False, False)

                lbl_result = tk.Label(result_window, text="Activation Details", font=("Segoe UI", 12, "bold"))
                lbl_result.pack(pady=10)

                txt_output = scrolledtext.ScrolledText(result_window, wrap=tk.WORD, font=("Consolas", 10))
                txt_output.insert(tk.END, f"{warning}\n")
                txt_output.insert(tk.END, f"Detected Edition Type: {edition_type}\n\n")
                txt_output.insert(tk.END, f"LICENSE NAME   : {license_name}\n")
                txt_output.insert(tk.END, f"LICENSE STATUS : {license_status}\n\n")
                txt_output.insert(tk.END, output.strip())
                txt_output.config(state='disabled')
                txt_output.pack(expand=True, fill="both", padx=10, pady=5)

                return
        except Exception:
            continue

    messagebox.showerror("Error", "Office not found or unable to retrieve activation info.")

    

# GUI Setup
root = tk.Tk()
root.title("Activation Tool" )

frame = tk.Frame(root, padx=20, pady=20)
frame.pack()

tk.Label(frame, text="Activation Tool", font=("Arial", 16)).pack(pady=10)
tk.Button(frame, text="📡 Connect to WiFi (XYZ)", width=30, command=connect_to_wifi).pack(pady=5)
tk.Button(frame, text="🔁 Upgrade Windows to Pro", width=30, command=upgrade_windows_to_pro).pack(pady=5)
tk.Button(frame, text="🪟 Activate Windows", width=30, command=activate_windows).pack(pady=5)
tk.Label(frame, text="Office Activation", font=("Arial", 14, "bold")).pack(pady=(20, 5))
tk.Button(root, text="Inject Office Key", command=fetch_and_set_office_key).pack(pady=5)
tk.Button(frame, text="📥 1. Generate Installation ID (IID)", width=40,
          command=generate_office_iid).pack(pady=5)
tk.Button(frame, text="⚡ Get IID & Open GetCID", width=40, command=fast_office_iid_flow).pack(pady=5)


# tk.Button(frame, text="🔑 3. Enter Confirmation ID (CID)", width=40,
#           command=activate_office_with_cid).pack(pady=5)

tk.Button(frame, text="🔍 Check Office Activation", width=30, command=check_office_activation).pack(pady=5)
# tk.Button(frame, text="🌐 Open CID Website (getcid.info)", width=30, command=lambda: webbrowser.open("https://getcid.info")).pack(pady=5)
# # tk.Button(root, text="Install Drivers", command=install_drivers_automatically).pack(pady=5)
# # driver_button = tk.Button(root, text="Install Wifi Driver", command=install_exe_drivers_from_usb)
# # driver_button.pack(pady=10)
# tk.Button(root, text="Install USB WLAN Driver", command=install_driver_from_usb).pack(pady=10)

# # install_exe_drivers_from_usb.pack(pady=10)
tk.Button(root, text="📂 Pick USB Driver to Install", command=pick_and_install_usb_driver).pack(pady=10)





def on_enter_key(event):
    widget = root.focus_get()
    if isinstance(widget, tk.Button):
        widget.invoke()

root.bind("<Return>", on_enter_key)


# btn_check_office.pack(pady=10)

root.mainloop()
