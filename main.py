import os
import json
import winreg
import ctypes
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, colorchooser
from PIL import Image, ImageTk
import subprocess
from functools import partial


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

def hex_to_bgra_bytes(hex_color):
    hex_color = hex_color.lstrip('#')
    r = int(hex_color[0:2], 16)
    g = int(hex_color[2:4], 16)
    b = int(hex_color[4:6], 16)
    return bytes([b, g, r, 0xAA])

def darken_color(hex_color, factor=0.85):
    hex_color = hex_color.lstrip("#")
    r = max(0, int(int(hex_color[0:2], 16) * factor))
    g = max(0, int(int(hex_color[2:4], 16) * factor))
    b = max(0, int(int(hex_color[4:6], 16) * factor))
    return f"#{r:02x}{g:02x}{b:02x}"

def add_hover_effect(widget, base_color):
    darker = darken_color(base_color)
    widget.bind("<Enter>", lambda e: widget.config(bg=darker))
    widget.bind("<Leave>", lambda e: widget.config(bg=base_color))

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

def set_accent_palette(accent_color, optional_colors):
    try:
        reversed_accent = reverse_hex(accent_color)
        main_bytes = hex_to_bgra_bytes(reversed_accent)

        optional_bytes = []
        for color in optional_colors:
            if color:
                reversed_color = reverse_hex(color)
                optional_bytes.append(hex_to_bgra_bytes(reversed_color))
            else:
                optional_bytes.append(bytes([0, 0, 0, 0xAA]))

        palette = (
            optional_bytes[0] +
            optional_bytes[1] +
            optional_bytes[2] +
            optional_bytes[3] +
            optional_bytes[4] +
            main_bytes +
            main_bytes +
            main_bytes
        )

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
def save_preset(accent_color, wallpaper, optional_colors):
    preset_files = [f for f in os.listdir(os.getcwd()) if f.endswith(".json")]
    if len(preset_files) >= 12:
        messagebox.showerror("Limit Reached", "You can only save up to 12 presets.")
        return

    if not accent_color:
        messagebox.showerror("Error", "Please select an accent color before saving the preset.")
        return

    def show_custom_preset_naming_popup(callback):
        popup = tk.Toplevel(root)
        popup.title("Save Preset")
        popup.geometry("300x150")
        popup.configure(bg="#f9f9f9")
        popup.resizable(False, False)
        popup.transient(root)
        popup.grab_set()

        # Center the popup
        x = root.winfo_rootx() + root.winfo_width() // 2 - 150
        y = root.winfo_rooty() + root.winfo_height() // 2 - 75
        popup.geometry(f"+{x}+{y}")

        tk.Label(popup, text="Enter a name for your preset:", bg="#f9f9f9", font=("Segoe UI", 10)).pack(pady=(15, 5))

        entry = tk.Entry(popup, font=("Segoe UI", 10), justify="center")
        entry.pack(pady=5, ipadx=10)

        def on_confirm():
            name = entry.get().strip()
            if name:
                filename = f"{name}.json"
                preset_data = {
                    "accent_color": accent_color,
                    "wallpaper": wallpaper,
                    "optional_colors": optional_colors
                }
                try:
                    with open(filename, "w") as f:
                        json.dump(preset_data, f, indent=4)
                    update_preset_viewer()
                    messagebox.showinfo("Preset Saved", f"Saved preset: {filename}")
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to save preset:\n{e}")
                popup.destroy()

        def on_cancel():
            popup.destroy()

        button_frame = tk.Frame(popup, bg="#f9f9f9")
        button_frame.pack(pady=(10, 0))

        confirm_btn = tk.Button(button_frame, text="Save", command=on_confirm, font=("Segoe UI", 10), bg="#0078D4",
                                fg="white", relief="flat", width=10)
        confirm_btn.pack(side="left", padx=5)

        cancel_btn = tk.Button(button_frame, text="Cancel", command=on_cancel, font=("Segoe UI", 10), bg="#e0e0e0",
                               relief="flat", width=10)
        cancel_btn.pack(side="right", padx=5)

        entry.focus()

    show_custom_preset_naming_popup(None)

def load_preset(preset_file):
    try:
        with open(preset_file, "r") as file:
            return json.load(file)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to load preset:\n{e}")
        return None

def apply_preset(preset_file):
    preset = load_preset(preset_file)
    if preset:
        accent_color = preset.get("accent_color", "").strip()
        wallpaper = preset.get("wallpaper", "").strip()
        optional_colors = preset.get("optional_colors", [""] * 5)

        if not accent_color or len(accent_color) != 7 or not accent_color.startswith("#"):
            messagebox.showerror("Error", "Invalid accent color in preset.")
            return

        selected_color.set(accent_color)
        preview_label.config(bg=accent_color, text=f"Preview: {accent_color}")

        for i, color in enumerate(optional_colors):
            optional_color_vars[i].set(color)
            optional_preview_labels[i].config(bg=color if color else "#dcdcdc", text=color if color else "Not set")

        set_accent_color(reverse_hex(accent_color))
        set_accent_palette(accent_color, optional_colors)

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
        messagebox.showinfo("Preset Applied", f"Applied preset: {os.path.splitext(preset_file)[0]}")

def update_preset_viewer():
    for widget in preset_frame.winfo_children():
        widget.destroy()

    preset_files = [f for f in os.listdir(os.getcwd()) if f.endswith(".json")]
    preset_count_label.config(text=f"Saved Presets ({len(preset_files)}/12)")

    for i, preset_file in enumerate(preset_files):
        preset_name = os.path.splitext(preset_file)[0]
        btn = tk.Button(preset_frame, text=preset_name, width=20, height=2, bg="#e0e0e0",
                        font=("Segoe UI", 9), relief="flat", command=lambda pf=preset_file: apply_preset(pf))
        btn.grid(row=i // 4, column=i % 4, padx=5, pady=5)
        add_hover_effect(btn, "#e0e0e0")

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
        img.thumbnail((300, 200))
        img_tk = ImageTk.PhotoImage(img)
        wallpaper_preview_label.config(image=img_tk, text="", bg="#dcdcdc")
        wallpaper_preview_label.image = img_tk
    except Exception as e:
        wallpaper_preview_label.config(text="Preview failed", image="", fg="red", bg="#dcdcdc")
        wallpaper_preview_label.image = None
        messagebox.showerror("Error", f"Could not preview image:\n{e}")

def confirm_and_apply():
    color = selected_color.get()
    if color:
        optional_colors = [var.get() for var in optional_color_vars]
        set_accent_color(reverse_hex(color))
        set_accent_palette(color, optional_colors)

        if wallpaper_path.get():
            set_wallpaper(wallpaper_path.get())

        restart_explorer()
        messagebox.showinfo("Accent Applied", f"Applied color: {color.upper()}")

# --- GUI Setup ---

root = tk.Tk()
root.title("Windows Theme Manager")
root.geometry("900x700")
root.configure(bg="#f3f3f3")

wallpaper_path = tk.StringVar()
selected_color = tk.StringVar()
optional_color_vars = [tk.StringVar() for _ in range(5)]
optional_preview_labels = []

main_frame = tk.Frame(root, bg="#f3f3f3")
main_frame.pack(pady=20)

left_frame = tk.Frame(main_frame, bg="#f3f3f3")
left_frame.pack(side="left", padx=40)

tk.Label(left_frame, text="Accent Color Picker", font=("Segoe UI", 11), bg="#f3f3f3").pack(pady=(0, 5))
choose_color_btn = tk.Button(left_frame, text="Choose Color", command=choose_color, width=20, bg="#0078d7", fg="white", relief="flat")
choose_color_btn.pack()
add_hover_effect(choose_color_btn, "#0078d7")

preview_label = tk.Label(left_frame, text="No color selected", font=("Segoe UI", 10), bg="#dcdcdc", width=30, height=2)
preview_label.pack(pady=5)

tk.Label(left_frame, text="Wallpaper:", font=("Segoe UI", 10), bg="#f3f3f3").pack(pady=(15, 0))
wallpaper_label = tk.Label(left_frame, text="No file selected", fg="gray", bg="#f3f3f3")
wallpaper_label.pack()

wallpaper_preview_frame = tk.Frame(left_frame, width=300, height=160, bg="#dcdcdc", bd=1, relief="sunken")
wallpaper_preview_frame.pack_propagate(False)  # Prevent resizing to contents
wallpaper_preview_frame.pack(pady=5)

wallpaper_preview_label = tk.Label(wallpaper_preview_frame, text="No Preview", fg="gray", bg="#dcdcdc", font=("Segoe UI", 10))
wallpaper_preview_label.pack(expand=True)


choose_wallpaper_btn = tk.Button(left_frame, text="Choose Wallpaper", command=choose_wallpaper, width=20, bg="#0078d7", fg="white", relief="flat")
choose_wallpaper_btn.pack()
add_hover_effect(choose_wallpaper_btn, "#0078d7")

right_frame = tk.Frame(main_frame, bg="#f3f3f3")
right_frame.pack(side="right", padx=40)

tk.Label(right_frame, text="Optional Colors", font=("Segoe UI", 11), bg="#f3f3f3").pack(pady=(0, 5))

optional_labels = [
    "Link and microphone color",
    "Taskbar focus / Alt+Tab highlight color",
    "Start button hover (W11 Pro only)",
    "Settings icons and buttons",
    "Pop-up window color"
]

def choose_optional_color(index):
    color_code = colorchooser.askcolor(title=f"Choose color for {optional_labels[index]}")
    if color_code[1]:
        optional_color_vars[index].set(color_code[1])
        optional_preview_labels[index].config(bg=color_code[1], text=color_code[1])

for i, label_text in enumerate(optional_labels):
    row = tk.Frame(right_frame, bg="#f3f3f3")
    row.pack(pady=3, anchor="w")
    tk.Label(row, text=label_text, bg="#f3f3f3", width=36, anchor="w", font=("Segoe UI", 10)).pack(side="left")
    btn = tk.Button(row, text="Choose", bg="#0078d7", fg="white", relief="flat",
                    command=lambda idx=i: choose_optional_color(idx))
    btn.pack(side="left", padx=5)
    add_hover_effect(btn, "#0078d7")
    preview = tk.Label(row, text="Not set", bg="#dcdcdc", width=10, font=("Segoe UI", 10))
    preview.pack(side="left")
    optional_preview_labels.append(preview)

button_section = tk.Frame(root, bg="#f3f3f3")
button_section.pack(pady=10)

apply_btn = tk.Button(button_section, text="Confirm & Apply", command=confirm_and_apply,
                      bg="#28a745", fg="white", relief="flat", height=2, width=25)
apply_btn.pack(side="left", padx=20)
add_hover_effect(apply_btn, "#28a745")

save_btn = tk.Button(button_section, text="Save Preset", command=lambda: save_preset(
    selected_color.get(), wallpaper_path.get(), [var.get() for var in optional_color_vars]),
    relief="flat", bg="#0078d7", fg="white", height=2, width=25)
save_btn.pack(side="right", padx=20)
add_hover_effect(save_btn, "#0078d7")

presets_container = tk.Frame(root, bg="#f3f3f3")
presets_container.pack(pady=(20, 10), fill="x")

preset_count_label = tk.Label(presets_container, text="Saved Presets (0/12)", font=("Segoe UI", 11), bg="#f3f3f3")
preset_count_label.pack(side="top", pady=(0, 10))

preset_frame = tk.Frame(presets_container, bg="#f3f3f3")
preset_frame.pack()

try:
    update_preset_viewer()
except:
    pass

root.mainloop()
