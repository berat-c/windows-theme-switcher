import os
import json
import winreg
import ctypes
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, colorchooser
from PIL import Image, ImageTk  # <-- New import for image preview
import subprocess

# --- Utility Functions ---

def hex_to_bgr(hex_color):
    rgb = int(hex_color.lstrip('#'), 16)
    bgr = ((rgb & 0xFF) << 16) | (rgb & 0xFF00) | ((rgb >> 16) & 0xFF)
    return bgr

def reverse_hex(hex_color):
    hex_color = hex_color.lstrip('#')
    r = hex_color[0:2]
    g = hex_color[2:4]
    b = hex_color[4:6]
    return f"#{b}{g}{r}"

# --- Core Theme Functions ---

def set_accent_color(hex_color):
    try:
        bgr_color_value = hex_to_bgr(hex_color)
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                            r"Software\Microsoft\Windows\CurrentVersion\Explorer\Accent",
                            0, winreg.KEY_SET_VALUE) as key:
            winreg.SetValueEx(key, "AccentColorMenu", 0, winreg.REG_DWORD, bgr_color_value)

        with winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                            r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize",
                            0, winreg.KEY_SET_VALUE) as key:
            winreg.SetValueEx(key, "ColorPrevalence", 0, winreg.REG_DWORD, 1)

        with winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                            r"Software\Microsoft\Windows\DWM",
                            0, winreg.KEY_SET_VALUE) as key:
            winreg.SetValueEx(key, "ColorizationColor", 0, winreg.REG_DWORD, bgr_color_value)
            winreg.SetValueEx(key, "ColorizationAfterglow", 0, winreg.REG_DWORD, bgr_color_value)
            winreg.SetValueEx(key, "AccentColor", 0, winreg.REG_DWORD, bgr_color_value)

        ctypes.windll.user32.SystemParametersInfoW(20, 0, 0, 0)
        ctypes.windll.user32.SendMessageTimeoutW(0xFFFF, 0x1A, 0, 0, 0, 1000, None)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to set accent color:\n{e}")

def set_accent_palette(hex_color):
    try:
        hex_color = hex_color.lstrip('#')
        r = int(hex_color[0:2], 16)
        g = int(hex_color[2:4], 16)
        b = int(hex_color[4:6], 16)
        bgra = bytes([b, g, r, 0xFF])
        palette = bgra * 8

        with winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                            r"Software\Microsoft\Windows\CurrentVersion\Explorer\Accent",
                            0, winreg.KEY_SET_VALUE) as key:
            winreg.SetValueEx(key, "AccentPalette", 0, winreg.REG_BINARY, palette)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to set accent palette:\n{e}")

def set_wallpaper(image_path):
    try:
        ctypes.windll.user32.SystemParametersInfoW(20, 0, image_path, 3)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to set wallpaper:\n{e}")

def restart_explorer():
    try:
        subprocess.run(["taskkill", "/f", "/im", "explorer.exe"], check=True)
        subprocess.run(["start", "explorer"], shell=True)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to restart Explorer:\n{e}")

# --- Preset Handling ---

def save_preset(accent_color, wallpaper):
    if not accent_color:
        messagebox.showerror("Error", "Please select an accent color before saving the preset.")
        return

    preset_name = simpledialog.askstring("Preset Name", "Enter a name for this preset:")
    if preset_name:
        preset = {
            "accent_color": accent_color,
            "wallpaper": wallpaper
        }
        preset_file = os.path.join(os.getcwd(), f"{preset_name}.json")
        try:
            with open(preset_file, "w") as file:
                json.dump(preset, file)
            messagebox.showinfo("Success", f"Preset '{preset_name}' saved successfully!")
            update_preset_viewer()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save preset:\n{e}")

def load_preset(preset_file):
    try:
        with open(preset_file, "r") as file:
            preset = json.load(file)
            if "accent_color" not in preset or "wallpaper" not in preset:
                messagebox.showerror("Error", "Preset file is corrupted.")
                return None
            return preset
    except Exception as e:
        messagebox.showerror("Error", f"Failed to load preset:\n{e}")
        return None

def apply_preset(preset_file):
    preset = load_preset(preset_file)
    if preset:
        accent_color = preset.get("accent_color", "").strip()
        wallpaper = preset.get("wallpaper", "").strip()

        if not accent_color or len(accent_color) != 7 or not accent_color.startswith("#"):
            messagebox.showerror("Error", "Invalid accent color in preset.")
            return

        reversed_color = reverse_hex(accent_color)
        set_accent_color(reversed_color)
        set_accent_palette(reversed_color)

        if wallpaper:
            set_wallpaper(wallpaper)
            wallpaper_path.set(wallpaper)
            wallpaper_label.config(text=os.path.basename(wallpaper))
            display_wallpaper_preview(wallpaper)
        else:
            wallpaper_path.set("")
            wallpaper_label.config(text="No file selected")
            wallpaper_preview_label.config(image="")

        restart_explorer()
        messagebox.showinfo("Preset Applied", f"Applied preset: {preset_file}")

def update_preset_viewer():
    for widget in preset_frame.winfo_children():
        widget.destroy()

    preset_files = [f for f in os.listdir(os.getcwd()) if f.endswith(".json")]
    for preset_file in preset_files:
        preset_name = os.path.splitext(preset_file)[0]
        btn = tk.Button(preset_frame, text=preset_name, width=25, height=2, bg="#e0e0e0",
                        font=("Segoe UI", 9), relief="flat", command=lambda pf=preset_file: apply_preset(pf))
        btn.pack(pady=4)

# --- GUI Callbacks ---

def choose_color():
    color_code = colorchooser.askcolor(title="Pick Accent Color")
    if color_code[1]:
        selected_color.set(color_code[1])
        preview_label.config(bg=color_code[1], text=f"Preview: {color_code[1]}")

def choose_wallpaper():
    path = filedialog.askopenfilename(title="Select Wallpaper", filetypes=[("Image Files", "*.jpg *.png *.bmp")])
    if path:
        wallpaper_path.set(path)
        wallpaper_label.config(text=os.path.basename(path))
        display_wallpaper_preview(path)

def display_wallpaper_preview(path):
    try:
        img = Image.open(path)
        img.thumbnail((300, 200))  # Resize for preview
        img_tk = ImageTk.PhotoImage(img)
        wallpaper_preview_label.config(image=img_tk)
        wallpaper_preview_label.image = img_tk
    except Exception as e:
        wallpaper_preview_label.config(text="Preview failed", image="")
        messagebox.showerror("Error", f"Could not preview image:\n{e}")

def confirm_and_apply():
    color = selected_color.get()
    if color:
        reversed_color = reverse_hex(color)
        set_accent_color(reversed_color)
        set_accent_palette(reversed_color)

        if wallpaper_path.get():
            set_wallpaper(wallpaper_path.get())

        restart_explorer()
        messagebox.showinfo("Accent Applied", f"Applied color: {color.upper()}")

# --- GUI Setup ---

root = tk.Tk()
root.title("Windows Theme Manager")
root.geometry("480x680")
root.resizable(False, False)
root.config(bg="#f3f3f3")

wallpaper_path = tk.StringVar()
selected_color = tk.StringVar()

theme_frame = tk.Frame(root, bg="#f3f3f3")
theme_frame.pack(pady=15)

# Color Picker
tk.Label(theme_frame, text="Accent Color Picker", font=("Segoe UI", 11), bg="#f3f3f3").pack()
tk.Button(theme_frame, text="Choose Color", command=choose_color, bg="#0078d7", fg="white", relief="flat").pack(pady=5)

preview_label = tk.Label(theme_frame, text="No color selected", font=("Segoe UI", 10), bg="#dcdcdc", width=30, height=2)
preview_label.pack(pady=5)

# Wallpaper Picker
tk.Label(theme_frame, text="Wallpaper:", font=("Segoe UI", 10), bg="#f3f3f3").pack()
wallpaper_label = tk.Label(theme_frame, text="No file selected", fg="gray", bg="#f3f3f3")
wallpaper_label.pack()

wallpaper_preview_label = tk.Label(theme_frame, bg="#f3f3f3")
wallpaper_preview_label.pack(pady=5)

tk.Button(theme_frame, text="Choose Wallpaper", command=choose_wallpaper, bg="#0078d7", fg="white", relief="flat").pack(pady=5)

# Apply Button
tk.Button(theme_frame, text="Confirm & Apply", command=confirm_and_apply, bg="#28a745", fg="white", relief="flat", height=2, width=25).pack(pady=10)

# Save Preset Button
tk.Button(theme_frame, text="Save Preset", command=lambda: save_preset(selected_color.get(), wallpaper_path.get()), relief="flat", bg="#0078d7", fg="white").pack(pady=5)

# Presets Section
tk.Label(root, text="Saved Presets", font=("Segoe UI", 11), bg="#f3f3f3").pack(pady=5)
preset_frame = tk.Frame(root, bg="#f3f3f3")
preset_frame.pack(pady=5)
update_preset_viewer()

root.mainloop()
