import os
import json
import winreg
import ctypes
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
import subprocess

# Function to convert hex to BGR format (used by Windows)
def hex_to_bgr(hex_color):
    rgb = int(hex_color.lstrip('#'), 16)
    bgr = ((rgb & 0xFF) << 16) | (rgb & 0xFF00) | ((rgb >> 16) & 0xFF)
    return bgr

# Function to set the accent palette (8 colors) based on a selected color
def set_accent_palette(hex_color):
    try:
        hex_color = hex_color.lstrip('#')
        r = int(hex_color[0:2], 16)
        g = int(hex_color[2:4], 16)
        b = int(hex_color[4:6], 16)
        bgra = bytes([b, g, r, 0xFF])
        palette = bgra * 8  # Set the palette with the same color repeated 8 times

        with winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                            r"Software\Microsoft\Windows\CurrentVersion\Explorer\Accent",
                            0, winreg.KEY_SET_VALUE) as key:
            winreg.SetValueEx(key, "AccentPalette", 0, winreg.REG_BINARY, palette)

    except Exception as e:
        messagebox.showerror("Error", f"Failed to set accent palette:\n{e}")

# Function to save current theme settings (Accent Color and Wallpaper) to a JSON file
def save_preset(accent_color, wallpaper):
    if not accent_color:  # Ensure accent color is selected
        messagebox.showerror("Error", "Please select an accent color before saving the preset.")
        return

    preset_name = simpledialog.askstring("Preset Name", "Enter a name for this preset:")
    if preset_name:
        preset = {
            "accent_color": accent_color,
            "wallpaper": wallpaper
        }

        # Save preset as a JSON file in the same directory
        preset_file = os.path.join(os.getcwd(), f"{preset_name}.json")
        try:
            with open(preset_file, "w") as file:
                json.dump(preset, file)
            messagebox.showinfo("Success", f"Preset '{preset_name}' saved successfully!")
            update_preset_viewer()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save preset:\n{e}")

# Function to load theme settings from a saved preset
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

# Function to apply a preset
def apply_preset(preset_file):
    preset = load_preset(preset_file)
    if preset:
        accent_color = preset.get("accent_color", "").strip()
        wallpaper = preset.get("wallpaper", "").strip()

        # Validate the accent color and apply it only if it's valid
        if not accent_color or len(accent_color) != 7 or not accent_color.startswith("#"):
            messagebox.showerror("Error", "Invalid accent color in preset.")
            return

        set_accent_color(accent_color)
        set_accent_palette(accent_color)

        # Apply wallpaper if available
        if wallpaper:
            set_wallpaper(wallpaper)
            wallpaper_path.set(wallpaper)
            wallpaper_label.config(text=os.path.basename(wallpaper))
        else:
            wallpaper_path.set("")
            wallpaper_label.config(text="No file selected")

        restart_explorer()
        messagebox.showinfo("Preset Applied", f"Applied preset: {preset_file}")

# Function to update the preset viewer
def update_preset_viewer():
    # Clear the current list
    for widget in preset_frame.winfo_children():
        widget.destroy()

    # List all saved preset files in the current directory
    preset_files = [f for f in os.listdir(os.getcwd()) if f.endswith(".json")]

    # Display the saved presets in the GUI
    for preset_file in preset_files:
        preset_name = os.path.splitext(preset_file)[0]
        preset_button = tk.Button(preset_frame, text=preset_name, width=20, height=2, command=lambda pf=preset_file: apply_preset(pf))
        preset_button.pack(pady=5)


# Function to set the accent color using the hex value
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

# Function to restart explorer for the theme to take effect
def restart_explorer():
    try:
        subprocess.run(["taskkill", "/f", "/im", "explorer.exe"], check=True)
        subprocess.run(["start", "explorer"], shell=True)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to restart Explorer:\n{e}")

# Function to set the wallpaper
def set_wallpaper(image_path):
    try:
        ctypes.windll.user32.SystemParametersInfoW(20, 0, image_path, 3)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to set wallpaper:\n{e}")

# Function to choose wallpaper
def choose_wallpaper():
    path = filedialog.askopenfilename(title="Select Wallpaper", filetypes=[("Image Files", "*.jpg *.png *.bmp")])
    if path:
        wallpaper_path.set(path)
        wallpaper_label.config(text=os.path.basename(path))

# Function to apply the theme (Accent Color and Wallpaper)
# Function to apply the theme (Accent Color and Wallpaper)
def apply_theme(hex_color):
    selected_color.set(hex_color)  # Save selected color for later saving

    set_accent_color(hex_color)
    set_accent_palette(hex_color)

    if wallpaper_path.get():
        set_wallpaper(wallpaper_path.get())

    restart_explorer()
    messagebox.showinfo("Accent Applied", f"Applied color: {hex_color.upper()}")

# Function to load and apply a saved preset
def load_and_apply_preset(preset_file):
    preset = load_preset(preset_file)
    if preset:
        set_accent_color(preset["accent_color"])
        set_accent_palette(preset["accent_color"])
        set_wallpaper(preset["wallpaper"])
        wallpaper_path.set(preset["wallpaper"])
        wallpaper_label.config(text=os.path.basename(preset["wallpaper"]))
        messagebox.showinfo("Preset Applied", f"Applied preset: {preset_file}")

# GUI Setup
root = tk.Tk()
root.title("Windows Theme Manager")
root.geometry("480x600")
root.resizable(False, False)
root.config(bg="#f3f3f3")  # Set a background color for smoother look

wallpaper_path = tk.StringVar()
selected_color = tk.StringVar()

# UI Frame for Theme Management
theme_frame = tk.Frame(root, bg="#f3f3f3")
theme_frame.pack(pady=10)

# Accent color section
tk.Label(theme_frame, text="Choose Accent Color:", font=("Segoe UI", 10), bg="#f3f3f3").grid(row=0, column=0, padx=5)
color_buttons_frame = tk.Frame(theme_frame, bg="#f3f3f3")
color_buttons_frame.grid(row=1, column=0, pady=5)

display_colors = [
    "#f25050", "#7f36c0", "#fcd116", "#16c60c", "#3a96dd",
    "#0078d7", "#6b69d6", "#b146c2", "#ff8c00", "#e74856",
    "#0099bc", "#7a7574", "#767676", "#4c4a48", "#000000",
    "#118811", "#118899", "#991188", "#aa5500", "#0055aa"
]

color_map = {
    "#f25050": "#5050f2",
    "#7f36c0": "#c0367f",
    "#fcd116": "#16d1fc",
    "#16c60c": "#0cc616",
    "#3a96dd": "#dd963a",
    "#0078d7": "#d77800",
    "#6b69d6": "#d6696b",
    "#b146c2": "#c246b1",
    "#ff8c00": "#008cff",
    "#e74856": "#5648e7",
    "#0099bc": "#bc9900",
    "#7a7574": "#74757a",
    "#767676": "#767676",
    "#4c4a48": "#484a4c",
    "#000000": "#000000",
    "#118811": "#118811",
    "#118899": "#998811",
    "#991188": "#881199",
    "#aa5500": "#0055aa",
    "#0055aa": "#aa5500"
}

# Creating color buttons
for i, display_color in enumerate(display_colors):
    apply_color = color_map.get(display_color, display_color)
    btn = tk.Button(color_buttons_frame, bg=display_color, width=4, height=2,
                    command=lambda c=apply_color: apply_theme(c), relief="flat")
    btn.grid(row=i // 8, column=i % 8, padx=4, pady=4)

# Wallpaper section
tk.Label(theme_frame, text="Wallpaper:", font=("Segoe UI", 10), bg="#f3f3f3").grid(row=2, column=0, padx=5)
wallpaper_label = tk.Label(theme_frame, text="No file selected", fg="gray", bg="#f3f3f3")
wallpaper_label.grid(row=3, column=0)
tk.Button(theme_frame, text="Choose Wallpaper", command=choose_wallpaper, relief="flat", bg="#0078d7", fg="white").grid(row=4, column=0, pady=5)

# Save Preset Button
tk.Button(theme_frame, text="Save Preset", command=lambda: save_preset(selected_color.get(), wallpaper_path.get()), relief="flat", bg="#0078d7", fg="white").grid(row=5, column=0, pady=5)

# Preset Viewer Section
preset_frame = tk.Frame(root, bg="#f3f3f3")
preset_frame.pack(pady=10)
update_preset_viewer()

root.mainloop()
