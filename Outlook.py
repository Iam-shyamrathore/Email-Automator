import tkinter as tk
from tkinter import ttk
from tkinter import messagebox, LabelFrame
from datetime import datetime
import win32com.client as win32
import configparser
import os
import sys
from PIL import ImageGrab, ImageTk

# --- Function to handle file paths for bundled data ---
def resource_path(relative_path):
    try: base_path = sys._MEIPASS
    except Exception: base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# --- Logic to save config.ini in a persistent location ---
if getattr(sys, 'frozen', False): base_path = os.path.dirname(sys.executable)
else: base_path = os.path.abspath(".")
CONFIG_FILE = os.path.join(base_path, 'config.ini')
SCREENSHOT_PATH = os.path.join(base_path, "custody_form_screenshot.png")

def save_config(name, mobile):
    config = configparser.ConfigParser(); config['UserDetails'] = {'Name': name, 'Mobile': mobile}
    with open(CONFIG_FILE, 'w') as configfile: config.write(configfile)

def load_config():
    if not os.path.exists(CONFIG_FILE): return '', ''
    config = configparser.ConfigParser(); config.read(CONFIG_FILE)
    name = config.get('UserDetails', 'Name', fallback=''); mobile = config.get('UserDetails', 'Mobile', fallback='')
    return name, mobile

# --- The logo file must be in the SAME FOLDER as the script ---
LOGO_IMAGE_PATH = resource_path("image002.png")

# Asset Definitions
ASSETS = ["Laptop + Charger", "Desktop", "Keyboard + Mouse", "Headphones", "Webcams", "DP Cable", "HDMI Cable", "Monitor", "Docker", "Mouse"]
ASSET_ID_MAPPING = {"Laptop + Charger": "LT", "Desktop": "DT", "Webcams": "WC", "Monitor": "MO", "Docker": "DS"}
PART1_OPTIONS = ["", "N", "H", "M"]

def send_email(template_path, data, subject):
    """Attaches images and displays the email."""
    try:
        with open(resource_path(template_path), "r", encoding="utf-8") as f: html_body = f.read()
        for key, value in data.items(): html_body = html_body.replace(f"{{{{{key}}}}}", str(value))
        outlook = win32.Dispatch("Outlook.Application"); mail = outlook.CreateItem(0)
        logo_attachment = mail.Attachments.Add(LOGO_IMAGE_PATH)
        logo_attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "companylogo")
        if os.path.exists(SCREENSHOT_PATH):
            ss_attachment = mail.Attachments.Add(SCREENSHOT_PATH)
            ss_attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "custodyformimage")
        mail.To = "OneIT@maqsoftware.com"; mail.Subject = subject; mail.HTMLBody = html_body
        mail.Display()
        return True
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")
        return False

# Dictionaries to hold our dynamic widgets and variables
issued_id_vars, returned_id_vars = {}, {}

def submit_form():
    """Gathers data, including screenshot info, and creates the email."""
    employee_name = entry_employee.get()
    sender_name = entry_sender_name.get()
    sender_mobile = entry_sender_mobile.get()
    if not all([employee_name, sender_name]):
        messagebox.showerror("Error", "Please fill in both Employee and Your Name fields."); return
    save_config(sender_name, sender_mobile)
    def format_asset_list(asset_vars, id_vars_dict):
        html_list = []
        for asset, var in asset_vars.items():
            if var.get():
                if asset in ASSET_ID_MAPPING:
                    part1, part3 = id_vars_dict[asset][0].get(), id_vars_dict[asset][1].get()
                    part2 = ASSET_ID_MAPPING[asset]
                    if part1 and part3: html_list.append(f"<li>{asset} ({part1}/{part2}/{part3})</li>")
                    else: html_list.append(f"<li>{asset} (ID Incomplete)</li>")
                else: html_list.append(f"<li>{asset}</li>")
        return "".join(html_list) or "<li>N/A</li>"
    issued_html = format_asset_list(issued_vars, issued_id_vars)
    returned_html = format_asset_list(returned_vars, returned_id_vars)
    possession_assets = [asset for asset, var in possession_vars.items() if var.get()]
    possession_html = "".join([f"<li>{a}</li>" for a in possession_assets]) or "<li>N/A</li>"
    mail_type = mail_type_var.get()
    template_path = f"{mail_type.lower()}_template.html"
    today_str = datetime.today().strftime("%B %d, %Y")
    subject_actions = {"Issue": "Issued hardware to the mentioned below employee", "Return": "Received hardware from mentioned below employee", "Swap": "Issued and Received hardware for the mentioned below employee"}
    action = subject_actions.get(mail_type, "Processed hardware for")
    subject = f"{action} as of {today_str}"
    data = {
        "EmployeeName": employee_name, "IssuedDate": today_str, "ReturnedDate": today_str,
        "SenderName": sender_name, "SenderMobile": sender_mobile,
        "InventoryLink": "https://itinventorydashboard.azurewebsites.net/", 
        "WFHInventoryLink": "https://testmaq-my.sharepoint.com/:x:/g/personal/amand_maqsoftware_com/EVyxR8YNUP5Di0HSTKmc7T8BKRlsDGE8MOvuaAZywoF51w?e=ilRcFQ&CID=4CB67C0F-4A7A-4522-83C4-03A5CF3F6CB9&wdLOR=cB5092A66-A6A4-4EC4-B938-368C9E7F19A1&xsdata=MDV8MDJ8c2h5YW1AbWFxc29mdHdhcmUuY29tfDM2OTJjZGFkZmMwODRhNGZlZmU1MDhkZGVlYzEwMmM1fGU0ZDk4ZGQyOTE5OTQyZTViYThiZGEzZTc2M2VkZTJlfDB8MHw2Mzg5MjkyMzQ5MzQ3NzQ3MDR8VW5rbm33bnxUV0ZwYkdac2IzZDhleUpGYlhCMGVVMWhjR2tpT25SeWRXVXNJbFlpT2lJd0xqQXVNREF3TUNJc0lsQWlPaUpYYVc0ek1pSXNJa0ZPSWpvaVRXRnBiQ0lzSWxkVUlqb3lmUT09fDB8fHw%3d&sdata=VEoyYnF5YkRXQjk5YWhKMkVlS2FpV3ZteUhxa3FmSjJqSmVNbXo5NEtJbz0%3d&CT=1758340307838&OR=Outlook-Body&CID=8F6E0104-4FA9-482F-80E9-E701760F02F2",
        "ClientDeviceInventoryLink": "https://itinventorydashboard.azurewebsites.net/#", 
        "InventoryUpdated": "Y", "WFHUpdated": "Y", "ClientDeviceUpdated": "Y", "HardwareChecked": "Y",
    }
    if os.path.exists(SCREENSHOT_PATH):
        data["CustodyFormImage"] = '<img src="cid:custodyformimage" alt="Custody Form Screenshot" width="500">'
    else: data["CustodyFormImage"] = 'N/A'
    data["HardwareIssued"] = issued_html; data["HardwareReturned"] = returned_html
    if mail_type in ["Return", "Swap"]: data.update({"RemainingHardware": possession_html})
    if mail_type in ["Issue", "Swap"]: data.update({"ExistingHardware": possession_html})
    if send_email(template_path, data, subject):
        messagebox.showinfo("Success", "Email draft created successfully!")

# --- GUI Section ---
root = tk.Tk(); root.title("Hardware Management Form"); root.geometry("600x800")

# NEW: Functions to handle screenshot pasting and focus
def paste_screenshot(event=None):
    try:
        image = ImageGrab.grabclipboard()
        if image:
            image.save(SCREENSHOT_PATH); image.thumbnail((200, 200)); photo = ImageTk.PhotoImage(image)
            screenshot_label.config(image=photo, text=""); screenshot_label.image = photo
        else: screenshot_label.config(text="No image found on clipboard.")
    except Exception as e: messagebox.showerror("Paste Error", f"Could not paste image.\n\nError: {e}")

def on_paste_box_focus(event):
    event.widget.focus_set()
    event.widget.config(highlightthickness=2, highlightcolor="#0078D7")

def on_paste_box_unfocus(event):
    event.widget.config(highlightthickness=1, highlightcolor="gray")

# Main scrollable layout
main_frame = tk.Frame(root); main_frame.pack(fill="both", expand=True)
main_canvas = tk.Canvas(main_frame); scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=main_canvas.yview)
scrollable_frame = ttk.Frame(main_canvas)
scrollable_frame.bind("<Configure>", lambda e: main_canvas.configure(scrollregion=main_canvas.bbox("all")))
scrollable_window = main_canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
def on_canvas_resize(event): main_canvas.itemconfig(scrollable_window, width=event.width)
main_canvas.bind("<Configure>", on_canvas_resize); main_canvas.configure(yscrollcommand=scrollbar.set)
def on_mouse_wheel(event): main_canvas.yview_scroll(int(-1*(event.delta/120)), "units")
main_canvas.bind_all("<MouseWheel>", on_mouse_wheel)
scrollbar.pack(side="right", fill="y"); main_canvas.pack(side="left", fill="both", expand=True)

# Helper function to create the dynamic asset frames
def create_asset_frame(parent, asset_vars, id_vars_dict):
    frame = LabelFrame(parent, padx=10, pady=5)
    def toggle_id_widgets(show, widgets):
        for widget in widgets:
            if show: widget.grid()
            else: widget.grid_remove()
    for i, asset_name in enumerate(ASSETS):
        var = tk.BooleanVar(); asset_vars[asset_name] = var; id_widgets = []
        if asset_name in ASSET_ID_MAPPING:
            part1_var, part3_var = tk.StringVar(), tk.StringVar()
            id_vars_dict[asset_name] = (part1_var, part3_var)
            combo1 = ttk.Combobox(frame, textvariable=part1_var, values=PART1_OPTIONS, width=3, state="readonly")
            label_s1, label_p2, label_s2 = tk.Label(frame, text="/"), tk.Label(frame, text=ASSET_ID_MAPPING[asset_name]), tk.Label(frame, text="/")
            entry_part3 = tk.Entry(frame, textvariable=part3_var, width=6)
            id_widgets.extend([combo1, label_s1, label_p2, label_s2, entry_part3])
        cb = tk.Checkbutton(frame, text=asset_name, variable=var, command=lambda s=var, w=id_widgets: toggle_id_widgets(s.get(), w))
        cb.grid(row=i, column=0, sticky="w", pady=1)
        if id_widgets:
            combo1.grid(row=i, column=1, padx=(10,0)); label_s1.grid(row=i, column=2); label_p2.grid(row=i, column=3)
            label_s2.grid(row=i, column=4); entry_part3.grid(row=i, column=5)
            toggle_id_widgets(False, id_widgets)
    return frame

# --- ALL WIDGETS ARE NOW CREATED WITH 'scrollable_frame' AS PARENT ---
type_frame = LabelFrame(scrollable_frame, text="Select Mail Type", padx=10, pady=10)
type_frame.pack(fill="x", padx=10, pady=10, expand=True)
mail_type_var = tk.StringVar(value="Issue")
tk.Radiobutton(type_frame, text="Issue", variable=mail_type_var, value="Issue", command=lambda: update_form_layout(mail_type_var.get())).pack(side="left", padx=10)
tk.Radiobutton(type_frame, text="Return", variable=mail_type_var, value="Return", command=lambda: update_form_layout(mail_type_var.get())).pack(side="left", padx=10)
tk.Radiobutton(type_frame, text="Swap", variable=mail_type_var, value="Swap", command=lambda: update_form_layout(mail_type_var.get())).pack(side="left", padx=10)
signature_frame = LabelFrame(scrollable_frame, text="Your Signature Details", padx=10, pady=10)
signature_frame.pack(fill="x", padx=10, pady=5, expand=True)
tk.Label(signature_frame, text="Your Name:").grid(row=0, column=0, sticky="w", pady=2)
entry_sender_name = tk.Entry(signature_frame); entry_sender_name.grid(row=0, column=1, sticky="ew")
tk.Label(signature_frame, text="Your Mobile:").grid(row=1, column=0, sticky="w", pady=2)
entry_sender_mobile = tk.Entry(signature_frame); entry_sender_mobile.grid(row=1, column=1, sticky="ew")
signature_frame.grid_columnconfigure(1, weight=1)
tk.Label(scrollable_frame, text="Employee Name:").pack(anchor="w", padx=10, pady=(10,0))
entry_employee = tk.Entry(scrollable_frame); entry_employee.pack(fill="x", padx=10, pady=5, expand=True)
issued_vars, returned_vars = {}, {}
issue_frame = create_asset_frame(scrollable_frame, issued_vars, issued_id_vars); issue_frame.config(text="Select Assets Issued")
return_frame = create_asset_frame(scrollable_frame, returned_vars, returned_id_vars); return_frame.config(text="Select Assets Returned")
possession_vars = {}
possession_frame = LabelFrame(scrollable_frame, text="Assets in Possession", padx=10, pady=5)
for asset in ASSETS: var = tk.BooleanVar(); possession_vars[asset] = var; tk.Checkbutton(possession_frame, text=asset, variable=var).pack(anchor="w")

# --- UPDATED: Frame for pasting the screenshot with new bindings ---
screenshot_frame = LabelFrame(scrollable_frame, text="Paste Asset Custody Screenshot", padx=10, pady=10)
screenshot_frame.pack(fill="x", padx=10, pady=10, expand=True)
screenshot_label = tk.Label(screenshot_frame, text="Click here and press Ctrl+V to paste", height=10, relief="sunken", bg="white",
                              highlightthickness=1, highlightbackground="gray")
screenshot_label.pack(fill="x", expand=True)
screenshot_label.bind("<Control-v>", paste_screenshot)
screenshot_label.bind("<Button-1>", on_paste_box_focus)
screenshot_label.bind("<FocusIn>", on_paste_box_focus)
screenshot_label.bind("<FocusOut>", on_paste_box_unfocus)

def update_form_layout(mail_type):
    issue_frame.pack_forget(); return_frame.pack_forget(); possession_frame.pack_forget()
    if mail_type == "Issue":
        issue_frame.pack(fill="x", padx=10, pady=5, expand=True); possession_frame.config(text="Assets Already in Possession"); possession_frame.pack(fill="x", padx=10, pady=5, expand=True)
    elif mail_type == "Return":
        return_frame.pack(fill="x", padx=10, pady=5, expand=True); possession_frame.config(text="Assets Still in Possession"); possession_frame.pack(fill="x", padx=10, pady=5, expand=True)
    elif mail_type == "Swap":
        issue_frame.pack(fill="x", padx=10, pady=5, expand=True); return_frame.pack(fill="x", padx=10, pady=5, expand=True)
        possession_frame.config(text="Assets Already in Possession (Before Swap)"); possession_frame.pack(fill="x", padx=10, pady=5, expand=True)

tk.Button(scrollable_frame, text="Create Email", command=submit_form, bg="#28a745", fg="white", font=("Helvetica", 10, "bold"), relief="flat", padx=10, pady=5).pack(pady=20)
saved_name, saved_mobile = load_config()
entry_sender_name.insert(0, saved_name); entry_sender_mobile.insert(0, saved_mobile)
update_form_layout(mail_type_var.get())
if os.path.exists(SCREENSHOT_PATH): os.remove(SCREENSHOT_PATH)
root.mainloop()