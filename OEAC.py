import subprocess
import tkinter as tk
from tkinter import messagebox, scrolledtext
import re

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
                # Parse license info
                license_name_match = re.search(r"LICENSE NAME:\s*(.+)", output)
                license_status_match = re.search(r"LICENSE STATUS:\s*(.+)", output)

                license_name = license_name_match.group(1) if license_name_match else "Unknown"
                license_status = license_status_match.group(1) if license_status_match else "Unknown"

                # Detect edition type
                if "Retail" in license_name:
                    edition_type = "Retail"
                    warning = "⚠️ RETAIL Edition Detected:\nCID activation will NOT work with this Office install.\n"
                elif "Volume" in license_name or "VL" in license_name:
                    edition_type = "Volume"
                    warning = "✅ VOLUME Edition Detected:\nThis Office install supports CID activation.\n"
                else:
                    edition_type = "Unknown"
                    warning = "⚠️ Unable to determine if Office is Retail or Volume.\n"

                # Display in a separate window
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
root.title("Office Edition & Activation Checker")
root.geometry("400x200")
root.resizable(False, False)

lbl_title = tk.Label(root, text="Office Activation Check Tool", font=("Segoe UI", 14, "bold"))
lbl_title.pack(pady=15)

lbl_sub = tk.Label(root, text="Check if Office is Volume or Retail Edition\nand see current activation status", font=("Segoe UI", 9))
lbl_sub.pack(pady=5)

btn_check = tk.Button(root, text="Check Office Status", command=check_office_activation, font=("Segoe UI", 10), width=25)
btn_check.pack(pady=20)

root.mainloop()
