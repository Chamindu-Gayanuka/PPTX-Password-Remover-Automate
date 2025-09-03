"""
Apache License 2.0
Copyright (c) 2025 Chamindu Gayanuka

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.

Telegram Link : https://t.me/GwitcherG
Repo Link : https://github.com/Chamindu-Gayanuka/PPTX-Password-Remover-Automate
License Link : https://github.com/Chamindu-Gayanuka/PPTX-Password-Remover-Automate/blob/main/LICENSE
"""

from tkinter import messagebox, filedialog
import customtkinter as ctk
from PIL import Image
import os, re, shutil, zipfile
from tkinterdnd2 import DND_FILES, TkinterDnD

selected_file = None

def remove_pptx_modify_password(file_path, progress):
    try:
        progress.set(0)
        app.update_idletasks()

        temp_dir = "pptx_temp"
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)
        os.makedirs(temp_dir)

        progress.set(0.2)
        app.update_idletasks()

        with zipfile.ZipFile(file_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)

        progress.set(0.4)
        app.update_idletasks()

        presentation_path = os.path.join(temp_dir, "ppt", "presentation.xml")
        if not os.path.exists(presentation_path):
            messagebox.showerror("Error", "presentation.xml not found.")
            shutil.rmtree(temp_dir)
            return

        with open(presentation_path, "r", encoding="utf-8") as f:
            content = f.read()

        new_content = re.sub(r'<p:modifyVerifier[^>]*?/>', '', content)

        with open(presentation_path, "w", encoding="utf-8") as f:
            f.write(new_content)

        progress.set(0.7)
        app.update_idletasks()

        with zipfile.ZipFile(file_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for root_dir, dirs, files in os.walk(temp_dir):
                for file in files:
                    filepath = os.path.join(root_dir, file)
                    arcname = os.path.relpath(filepath, temp_dir)
                    zipf.write(filepath, arcname)

        shutil.rmtree(temp_dir)

        progress.set(1.0)
        app.update_idletasks()
        messagebox.showinfo("Done", f"Unlocked: {os.path.basename(file_path)}")
    except Exception as e:
        messagebox.showerror("Error", str(e))

def drop_file(event):
    file_path = event.data.strip('{}')
    set_selected_file(file_path)

def browse_file():
    file_path = filedialog.askopenfilename(
        title="Select PPTX file",
        filetypes=[("PowerPoint Files", "*.pptx")]
    )
    if file_path:
        set_selected_file(file_path)

def set_selected_file(file_path):
    global selected_file
    if file_path.lower().endswith(".pptx") and os.path.exists(file_path):
        selected_file = file_path
        file_label.configure(text=os.path.basename(file_path), text_color="green")
    else:
        messagebox.showerror("Invalid file", "Please select a valid PPTX file.")
        selected_file = None
        file_label.configure(text="No file selected", text_color="red")

def unlock_action():
    if not selected_file:
        messagebox.showwarning("No file", "Please select or drag a PPTX file first.")
        return
    remove_pptx_modify_password(selected_file, progress_bar)


"""
Apache License 2.0
Copyright (c) 2025 Chamindu Gayanuka
"""

# ================= UI =================
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("green")

app = TkinterDnD.Tk()
app.title(" PPTX Unlocker")
app.iconbitmap("asset/icon.ico")
width=500
height=400
x = (app.winfo_screenwidth() // 2) - (width // 2)
y = (app.winfo_screenheight() // 2) - (height // 2)
app.geometry(f'{width}x{height}+{x}+{y}')
app.resizable(False, False)

# Background
bg_image = ctk.CTkImage(
    light_image=Image.open("asset/background.png"),
    size=(500, 400)
)
bg_label = ctk.CTkLabel(app, image=bg_image, text="")
bg_label.place(relx=0.5, rely=0.5, anchor="center")

# Header
lock_icon = ctk.CTkImage(Image.open("asset/locker.png"), size=(30, 30))
header = ctk.CTkLabel(app, text=" PPTX Unlocker", image=lock_icon,
    compound="left", font=ctk.CTkFont(size=30, weight="bold", family="Times New Roman"), bg_color="#D9F4C6")
header.place(relx=0.5, rely=0.07, anchor="center")

sub_text = ctk.CTkLabel(app, text="Remove password protection from PowerPoint files",
    font=ctk.CTkFont(size=13, weight="bold"), text_color="gray50", bg_color="#D9F4C6")
sub_text.place(relx=0.5, rely=0.15, anchor="center")

# Drop area
file_icon = ctk.CTkImage(Image.open("asset/file.png"), size=(40, 40))
drop_frame = ctk.CTkFrame(app, width=400, height=120, fg_color="white", corner_radius=12, border_width=2, bg_color="#B7E6F3")
drop_frame.place(relx=0.5, rely=0.4, anchor="center")

drop_area = ctk.CTkButton(drop_frame,
    text="Drag & Drop PPTX Here\nor Click to Browse", image=file_icon,
    compound="top", fg_color="white", hover_color="#f0f0f0", text_color="gray40",
    width=400, height=130, command=browse_file, font=ctk.CTkFont(size=16, weight="bold"))
drop_area.pack(expand=True, fill="both", padx=10, pady=10)
drop_area.drop_target_register(DND_FILES)
drop_area.dnd_bind('<<Drop>>', drop_file)
drop_area.bind("<Button-1>", lambda e: browse_file())

file_label = ctk.CTkLabel(app, text="No file selected", text_color="red", font=ctk.CTkFont(size=14), bg_color="#B7E6F3")
file_label.place(relx=0.5, rely=0.65, anchor="center")

# Unlock button
unlock_icon = ctk.CTkImage(Image.open("asset/unlocker.png"), size=(25, 25))
unlock_btn = ctk.CTkButton(app, text=" Unlock", image=unlock_icon, compound="left",
    fg_color="#23c55f", hover_color="#00b050", font=ctk.CTkFont(size=25, weight="bold"),
    width=410, height=40, command=unlock_action, bg_color="#B7E6F3")
unlock_btn.place(relx=0.5, rely=0.73, anchor="center")

# Progress bar
progress_bar = ctk.CTkProgressBar(app, width=400, height=10, bg_color="#B7E6F3")
progress_bar.place(relx=0.5, rely=0.88, anchor="center")
progress_bar.set(0)

# Footer
"""
Don't remove the footer, it's my signature. 💗
Keep it or give credit. 😊
Telegram Link : https://t.me/GwitcherG
Repo Link : https://github.com/Chamindu-Gayanuka/PPTX-Password-Remover-Automate
License Link : https://github.com/Chamindu-Gayanuka/PPTX-Password-Remover-Automate/blob/main/LICENSE
Thank you! 🙏
"""
footer = ctk.CTkLabel(app, text="💗 Made by Chamindu Gayanuka 💗",
    font=ctk.CTkFont(size=18, family="Times New Roman"), text_color="red", bg_color="#b4bddd")
footer.place(relx=0.5, rely=0.95, anchor="center")

app.mainloop()