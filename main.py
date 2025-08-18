import os
import json
import winreg
import ctypes
import tkinter as tk
from tkinter import BooleanVar
import tkinter.ttk as ttk
from tkinter import filedialog, messagebox, colorchooser
from PIL import Image, ImageTk
import subprocess
from functools import partial

# Utility Functions

def hex_to_bgr(hex_color: str) -> int:
    rgb = int(hex_color.lstrip('#'), 16)
    bgr = ((rgb & 0xFF) << 16) | (rgb & 0xFF00) | ((rgb >> 16) & 0xFF)
    return bgr

def reverse_hex(hex_color: str) -> str:
    hex_color = hex_color.lstrip('#')
    r = hex_color[0:2]
    g = hex_color[2:4]
    b = hex_color[4:6]
    return f"#{b}{g}{r}"

def hex_to_bgra_bytes(hex_color: str) -> bytes:
    hex_color = hex_color.lstrip('#')
    r = int(hex_color[0:2], 16)
    g = int(hex_color[2:4], 16)
    b = int(hex_color[4:6], 16)
    return bytes([b, g, r, 0xAA])

def darken_color(hex_color: str, factor=0.85) -> str:
    hex_color = hex_color.lstrip("#")
    r = max(0, int(int(hex_color[0:2], 16) * factor))
    g = max(0, int(int(hex_color[2:4], 16) * factor))
    b = max(0, int(int(hex_color[4:6], 16) * factor))
    return f"#{r:02x}{g:02x}{b:02x}"

def add_hover_effect(widget, base_color: str):
    darker = darken_color(base_color)
    widget.bind("<Enter>", lambda e: widget.config(bg=darker))
    widget.bind("<Leave>", lambda e: widget.config(bg=base_color))

def is_valid_hex(s: str) -> bool:
    if not s or len(s) != 7 or not s.startswith("#"): 
        return False
    try:
        int(s[1:], 16)
        return True
    except ValueError:
        return False

def hex_to_rgb_string(h: str) -> str:
    """#RRGGBB -> 'R G B'"""
    h = h.lstrip("#")
    r = int(h[0:2], 16)
    g = int(h[2:4], 16)
    b = int(h[4:6], 16)
    return f"{r} {g} {b}"

def rgb_string_to_hex(s: str) -> str:
    """'R G B' -> #RRGGBB"""
    try:
        parts = [int(p) for p in s.split()]
        if len(parts) != 3:
            return "#000000"
        r, g, b = parts
        return f"#{r:02x}{g:02x}{b:02x}"
    except Exception:
        return "#000000"


# Modern Color Picker Popup

def open_modern_color_picker(initial_color="#0078D7", callback=None):
    """Fresh-looking popup: hex entry + preview + swatches + optional full custom dialog"""
    popup = tk.Toplevel(root)
    popup.title("Pick Color")
    popup.geometry("320x440")
    popup.configure(bg="#f9f9f9")
    popup.resizable(False, False)
    popup.transient(root)
    popup.grab_set()

    # Center popup relative to root
    root.update_idletasks()
    x = root.winfo_rootx() + root.winfo_width() // 2 - 160
    y = root.winfo_rooty() + root.winfo_height() // 2 - 220
    popup.geometry(f"+{x}+{y}")

    color_var = tk.StringVar(value=initial_color if is_valid_hex(initial_color) else "#0078D7")

    title = tk.Label(popup, text="Choose a color", bg="#f9f9f9", font=("Segoe UI", 11, "bold"))
    title.pack(pady=(12, 4))

    preview = tk.Label(popup, bg=color_var.get(), width=22, height=2, relief="flat", bd=0)
    preview.pack(pady=(6, 10))

    # Hex entry
    hex_frame = tk.Frame(popup, bg="#f9f9f9")
    hex_frame.pack(pady=6)
    tk.Label(hex_frame, text="Hex", bg="#f9f9f9", font=("Segoe UI", 10)).pack(side="left", padx=(0, 6))

    # Styled entry (ttk)
    style = ttk.Style(popup)
    try:
        style.theme_use("clam")
    except Exception:
        pass
    style.configure("Rounded.TEntry", padding=(8, 6, 8, 6))

    hex_entry = ttk.Entry(hex_frame, textvariable=color_var, width=12, justify="center", style="Rounded.TEntry")
    hex_entry.pack(side="left")
    hex_entry.focus_set()

    warn_label = tk.Label(popup, text="", fg="#cc0000", bg="#f9f9f9", font=("Segoe UI", 9))
    warn_label.pack(pady=(2, 0))

    def on_color_change(*_):
        val = color_var.get().strip()
        if is_valid_hex(val):
            warn_label.config(text="")
            try:
                preview.config(bg=val)
            except Exception:
                pass
        else:
            warn_label.config(text="Enter a valid hex like #0078D7")

    color_var.trace_add("write", on_color_change)

    # Swatches
    swatch_title = tk.Label(popup, text="Quick swatches", bg="#f9f9f9", font=("Segoe UI", 10))
    swatch_title.pack(pady=(12, 4))

    preset_colors = [
        "#0078D7", "#D83B01", "#107C10", "#FFB900", "#5C2D91",
        "#E81123", "#00B7C3", "#B4009E", "#FF8C00", "#498205"
    ]
    swatch_frame = tk.Frame(popup, bg="#f9f9f9")
    swatch_frame.pack(pady=6)
    for i, col in enumerate(preset_colors):
        btn = tk.Button(swatch_frame, bg=col, width=3, height=2, relief="flat", bd=0,
                        command=lambda c=col: color_var.set(c), activebackground=col)
        btn.grid(row=i//5, column=i%5, padx=6, pady=6)
        add_hover_effect(btn, col)

    def pick_custom():
        chosen = colorchooser.askcolor(color=color_var.get())[1]
        if chosen and is_valid_hex(chosen):
            color_var.set(chosen)

    # Buttons
    btn_frame = tk.Frame(popup, bg="#f9f9f9")
    btn_frame.pack(pady=16)

    def do_ok():
        if not is_valid_hex(color_var.get()):
            warn_label.config(text="Enter a valid hex like #0078D7")
            return
        if callback:
            callback(color_var.get())
        popup.destroy()

    custom_btn = tk.Button(btn_frame, text="Custom…", command=pick_custom,
                           bg="#e6e6e6", relief="flat", width=10)
    add_hover_effect(custom_btn, "#e6e6e6")
    custom_btn.pack(side="left", padx=6)

    ok_btn = tk.Button(btn_frame, text="OK", command=do_ok, bg="#0078D7", fg="white", relief="flat", width=10)
    add_hover_effect(ok_btn, "#0078D7")
    ok_btn.pack(side="left", padx=6)

    cancel_btn = tk.Button(btn_frame, text="Cancel", command=popup.destroy, bg="#e6e6e6", relief="flat", width=10)
    add_hover_effect(cancel_btn, "#e6e6e6")
    cancel_btn.pack(side="left", padx=6)


# Wallpaper Engine Integration

wallpaper_engine_path_file = "wallpaper_engine_path.txt"

def get_wallpaper_engine_exe():
    default_path = r"C:\Program Files (x86)\Steam\steamapps\common\wallpaper_engine\wallpaper64.exe"
    if os.path.isfile(default_path):
        return default_path
    if os.path.exists(wallpaper_engine_path_file):
        try:
            with open(wallpaper_engine_path_file, "r", encoding="utf-8") as f:
                path = f.read().strip()
                if os.path.isfile(path):
                    return path
        except Exception:
            pass
    return None

def prompt_for_wallpaper_engine():
    path = filedialog.askopenfilename(title="Locate Wallpaper Engine Executable",
                                      filetypes=[("Executable Files", "*.exe")])
    if path and os.path.isfile(path):
        with open(wallpaper_engine_path_file, "w", encoding="utf-8") as f:
            f.write(path)
        return path
    return None

def launch_wallpaper_engine():
    exe_path = get_wallpaper_engine_exe()
    if not exe_path:
        exe_path = prompt_for_wallpaper_engine()
    if exe_path:
        try:
            subprocess.Popen([exe_path],
                             creationflags=subprocess.DETACHED_PROCESS | subprocess.CREATE_NO_WINDOW,
                             stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to launch Wallpaper Engine:\n{e}")
    else:
        messagebox.showwarning("Wallpaper Engine Not Found", "Could not find or launch Wallpaper Engine.")

def get_wallpaper_engine_config_path():
    exe_path = get_wallpaper_engine_exe()
    if exe_path:
        exe_dir = os.path.dirname(exe_path)
        steam_config = os.path.join(exe_dir, "config.json")
        if os.path.exists(steam_config):
            return steam_config
    doc_config = os.path.expanduser(r"~\Documents\Wallpaper Engine\config.json")
    if os.path.exists(doc_config):
        return doc_config
    doc_startup = os.path.expanduser(r"~\Documents\Wallpaper Engine\startup.json")
    if os.path.exists(doc_startup):
        return doc_startup
    if exe_path:
        return os.path.join(os.path.dirname(exe_path), "config.json")
    return None

def read_we_config():
    path = get_wallpaper_engine_config_path()
    if not path or not os.path.exists(path):
        return None
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        print("Failed to read Wallpaper Engine config:", e)
        return None

def write_we_config(config_dict):
    path = get_wallpaper_engine_config_path()
    if not path:
        raise FileNotFoundError("Wallpaper Engine config path not found.")
    tmp = path + ".tmp"
    try:
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(config_dict, f, indent=4)
        os.replace(tmp, path)
    except Exception as e:
        try:
            if os.path.exists(tmp):
                os.remove(tmp)
        except Exception:
            pass
        raise

def reload_wallpaper_engine(exe_path=None):
    if exe_path is None:
        exe_path = get_wallpaper_engine_exe()
    if not exe_path:
        return False
    try:
        subprocess.run(["taskkill", "/IM", "wallpaper64.exe", "/F"], check=False,
                       creationflags=subprocess.DETACHED_PROCESS | subprocess.CREATE_NO_WINDOW,
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        subprocess.run(["taskkill", "/IM", "wallpaper32.exe", "/F"], check=False,
                       creationflags=subprocess.DETACHED_PROCESS | subprocess.CREATE_NO_WINDOW,
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        subprocess.Popen([exe_path],
                         creationflags=subprocess.DETACHED_PROCESS | subprocess.CREATE_NO_WINDOW,
                         stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        return True
    except Exception as e:
        print("Failed to reload Wallpaper Engine:", e)
        return False

# Core Theme Functions

def set_accent_color(hex_color):
    """hex_color MUST be reversed (BGR byte order) as caller already does reverse_hex(color)."""
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

        # Light refresh
        ctypes.windll.user32.SystemParametersInfoW(20, 0, 0, 0)  # SPI_SETDESKWALLPAPER broadcast
        ctypes.windll.user32.SendMessageTimeoutW(0xFFFF, 0x1A, 0, 0, 0, 1000, None)  # WM_SETTINGCHANGE
        ctypes.windll.user32.SendMessageTimeoutW(0xFFFF, 0x031A, 0, 0, 0, 1000, None)  # WM_THEMECHANGED
    except Exception as e:
        messagebox.showerror("Error", f"Failed to set accent color:\n{e}")

def set_accent_palette(accent_color, optional_colors):
    try:
        reversed_accent = reverse_hex(accent_color)
        main_bytes = hex_to_bgra_bytes(reversed_accent)

        optional_bytes = []
        for color in optional_colors:
            if color and is_valid_hex(color):
                reversed_color = reverse_hex(color)
                optional_bytes.append(hex_to_bgra_bytes(reversed_color))
            else:
                optional_bytes.append(bytes([0, 0, 0, 0xAA]))

        # 8 * 4 bytes = 32 bytes binary palette
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

def restart_explorer_smooth():
    """Broadcast theme changes first; then restart Explorer to force apply if needed."""
    try:
        # Broadcast change
        ctypes.windll.user32.SendMessageTimeoutW(0xFFFF, 0x1A, 0, 0, 0, 1000, None)  # WM_SETTINGCHANGE
        ctypes.windll.user32.SendMessageTimeoutW(0xFFFF, 0x031A, 0, 0, 0, 1000, None)  # WM_THEMECHANGED
    except Exception:
        pass
    try:
        subprocess.run(["taskkill", "/f", "/im", "explorer.exe"], check=True,
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        subprocess.Popen(["explorer.exe"], shell=False,
                         stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to restart Explorer:\n{e}")

# Transparency toggle
def toggle_transparency():
    key_path = r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_SET_VALUE) as key:
            value = 1 if transparency_var.get() else 0
            winreg.SetValueEx(key, "EnableTransparency", 0, winreg.REG_DWORD, value)
    except Exception as e:
        print("Failed to update transparency:", e)

def load_transparency_setting():
    key_path = r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path) as key:
            value, _ = winreg.QueryValueEx(key, "EnableTransparency")
            transparency_var.set(bool(value))
    except FileNotFoundError:
        transparency_var.set(True)  # Default to on

# Presets

PRESET_LIMIT = 12

def list_valid_presets():
    files = []
    for f in os.listdir(os.getcwd()):
        if f.lower().endswith(".json"):
            path = os.path.join(os.getcwd(), f)
            try:
                with open(path, "r", encoding="utf-8") as fh:
                    data = json.load(fh)
                if isinstance(data, dict) and "accent_color" in data:
                    files.append(f)
            except Exception:
                continue
    return sorted(files)

def save_preset(accent_color, wallpaper, optional_colors, control_panel_colors):
    preset_files = list_valid_presets()
    if len(preset_files) >= PRESET_LIMIT:
        messagebox.showerror("Limit Reached", f"You can only save up to {PRESET_LIMIT} presets.")
        return

    if not accent_color or not is_valid_hex(accent_color):
        messagebox.showerror("Error", "Please select a valid accent color before saving the preset.")
        return

    def show_custom_preset_naming_popup():
        popup = tk.Toplevel(root)
        popup.title("Save Preset")
        popup.geometry("320x160")
        popup.configure(bg="#f9f9f9")
        popup.resizable(False, False)
        popup.transient(root)
        popup.grab_set()

        x = root.winfo_rootx() + root.winfo_width() // 2 - 160
        y = root.winfo_rooty() + root.winfo_height() // 2 - 80
        popup.geometry(f"+{x}+{y}")

        tk.Label(popup, text="Enter a name for your preset:", bg="#f9f9f9", font=("Segoe UI", 10)).pack(pady=(15, 6))
        entry = ttk.Entry(popup, font=("Segoe UI", 10), justify="center", width=26)
        entry.pack(pady=6, ipadx=6)
        entry.focus_set()

        def on_confirm():
            name = entry.get().strip()
            if name:
                filename = f"{name}.json"
                if os.path.exists(filename):
                    if not messagebox.askyesno("Overwrite?", f"'{name}' already exists.\nDo you want to overwrite it?"):
                        return
                preset_data = {
                    "accent_color": accent_color,
                    "wallpaper": wallpaper,
                    "optional_colors": optional_colors,
                    "control_panel_colors": control_panel_colors,
                    "we_config": read_we_config()
                }
                try:
                    with open(filename, "w", encoding="utf-8") as f:
                        json.dump(preset_data, f, indent=4)
                    update_preset_viewer()
                    messagebox.showinfo("Preset Saved", f"Saved preset: {filename}")
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to save preset:\n{e}")
                finally:
                    popup.destroy()

        def on_cancel():
            popup.destroy()

        button_frame = tk.Frame(popup, bg="#f9f9f9")
        button_frame.pack(pady=(12, 10))

        confirm_btn = tk.Button(button_frame, text="Save", command=on_confirm, font=("Segoe UI", 10), bg="#0078D4",
                                fg="white", relief="flat", width=10)
        confirm_btn.pack(side="left", padx=6)
        add_hover_effect(confirm_btn, "#0078D4")

        cancel_btn = tk.Button(button_frame, text="Cancel", command=on_cancel, font=("Segoe UI", 10), bg="#e0e0e0",
                               relief="flat", width=10)
        cancel_btn.pack(side="left", padx=6)
        add_hover_effect(cancel_btn, "#e0e0e0")

    show_custom_preset_naming_popup()

def load_preset(preset_file):
    try:
        with open(preset_file, "r", encoding="utf-8") as file:
            return json.load(file)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to load preset:\n{e}")
        return None

def delete_preset(preset_file):
    if not os.path.exists(preset_file):
        return
    if messagebox.askyesno("Delete Preset", f"Are you sure you want to delete '{os.path.basename(preset_file)}'?"):
        try:
            os.remove(preset_file)
            update_preset_viewer()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to delete preset:\n{e}")
#yarrak
def apply_preset(preset_file):
    we_exe = get_wallpaper_engine_exe()
    if we_exe:
        subprocess.Popen([we_exe], creationflags=subprocess.DETACHED_PROCESS | subprocess.CREATE_NO_WINDOW,
                         stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)  # Ensure WE is running

    preset = load_preset(preset_file)
    if preset:
        # Restore WE config first if any
        if "we_config" in preset and preset["we_config"]:
            try:
                write_we_config(preset["we_config"])
                we_exe = get_wallpaper_engine_exe()
                if we_exe:
                    reload_wallpaper_engine(we_exe)
            except Exception as e:
                messagebox.showerror("Error", f"Failed to restore Wallpaper Engine config:\n{e}")

        accent_color = preset.get("accent_color", "").strip()
        wallpaper = preset.get("wallpaper", "").strip()
        optional_colors = preset.get("optional_colors", [""] * 5)
        cp_colors = preset.get("control_panel_colors", {})

        if not is_valid_hex(accent_color):
            messagebox.showerror("Error", "Invalid accent color in preset.")
            return

        # UI sync
        selected_color.set(accent_color)
        preview_label.config(bg=accent_color, text=f"Preview: {accent_color}")

        for i, color in enumerate(optional_colors):
            optional_color_vars[i].set(color)
            optional_preview_labels[i].config(bg=color if is_valid_hex(color) else "#dcdcdc",
                                              text=color if is_valid_hex(color) else "Not set")

        # Control Panel colors
        for key in control_panel_color_vars.keys():
            rgb = cp_colors.get(key, "")
            control_panel_color_vars[key].set(rgb)
            hx = rgb_string_to_hex(rgb) if rgb else "#dcdcdc"
            control_panel_preview_labels[key].config(bg=hx if is_valid_hex(hx) else "#dcdcdc",
                                                     text=rgb if rgb else "Not set")

        # Apply
        set_accent_color(reverse_hex(accent_color))
        set_accent_palette(accent_color, optional_colors)

        # Write Control Panel colors
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Control Panel\Colors", 0, winreg.KEY_SET_VALUE) as key:
                for k, v in cp_colors.items():
                    if v:
                        winreg.SetValueEx(key, k, 0, winreg.REG_SZ, v)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to set Control Panel colors:\n{e}")

        if wallpaper:
            set_wallpaper(wallpaper)
            wallpaper_path.set(wallpaper)
            wallpaper_label.config(text=os.path.basename(wallpaper))
            display_wallpaper_preview(wallpaper)
        else:
            wallpaper_path.set("")
            wallpaper_label.config(text="No file selected")
            wallpaper_preview_label.config(image="")
            wallpaper_preview_label.image = None

        restart_explorer_smooth()
        messagebox.showinfo("Preset Applied", f"Applied preset: {os.path.splitext(preset_file)[0]}")

def update_preset_viewer():
    for widget in preset_frame.winfo_children():
        widget.destroy()

    preset_files = list_valid_presets()
    preset_count_label.config(text=f"Saved Presets ({len(preset_files)}/{PRESET_LIMIT})")

    for i, preset_file in enumerate(preset_files):
        preset_name = os.path.splitext(preset_file)[0]

        row = tk.Frame(preset_frame, bg="#f3f3f3")
        row.grid(row=i // 2, column=i % 2, padx=6, pady=6, sticky="w")

        btn = tk.Button(row, text=preset_name, width=22, height=2, bg="#e0e0e0",
                        font=("Segoe UI", 9), relief="flat",
                        command=lambda pf=preset_file: apply_preset(pf))
        btn.pack(side="left", padx=(0, 6))
        add_hover_effect(btn, "#e0e0e0")

        del_btn = tk.Button(row, text="X", bg="#E81123", fg="white", relief="flat", width=2,
                            command=lambda pf=preset_file: delete_preset(pf))
        del_btn.pack(side="left")
        add_hover_effect(del_btn, "#E81123")

# GUI Callbacks

def choose_accent_color():
    open_modern_color_picker(
        initial_color=selected_color.get() or "#0078D7",
        callback=lambda c: (
            selected_color.set(c),
            preview_label.config(bg=c, text=f"Preview: {c}")
        )
    )

def choose_optional_color(index):
    open_modern_color_picker(
        initial_color=optional_color_vars[index].get() or "#0078D7",
        callback=lambda c: (
            optional_color_vars[index].set(c),
            optional_preview_labels[index].config(bg=c, text=c)
        )
    )

def choose_control_panel_color(key_name):
    # initial from current var -> hex
    current_rgb = control_panel_color_vars[key_name].get()
    initial_hex = rgb_string_to_hex(current_rgb) if current_rgb else "#0078D7"

    def on_pick(hex_color):
        rgb_string = hex_to_rgb_string(hex_color)
        control_panel_color_vars[key_name].set(rgb_string)
        control_panel_preview_labels[key_name].config(bg=hex_color, text=rgb_string)
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Control Panel\Colors", 0, winreg.KEY_SET_VALUE) as key:
                winreg.SetValueEx(key, key_name, 0, winreg.REG_SZ, rgb_string)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to set {key_name}:\n{e}")

    open_modern_color_picker(initial_color=initial_hex, callback=on_pick)

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
    if color and is_valid_hex(color):
        optional_colors = [var.get() for var in optional_color_vars]
        set_accent_color(reverse_hex(color))
        set_accent_palette(color, optional_colors)

        # Write Control Panel colors
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Control Panel\Colors", 0, winreg.KEY_SET_VALUE) as key:
                for k, v in control_panel_color_vars.items():
                    val = v.get()
                    if val:
                        winreg.SetValueEx(key, k, 0, winreg.REG_SZ, val)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to set Control Panel colors:\n{e}")

        if wallpaper_path.get():
            set_wallpaper(wallpaper_path.get())
        restart_explorer_smooth()
        messagebox.showinfo("Accent Applied", f"Applied color: {color.upper()}")
    else:
        messagebox.showerror("Error", "Please choose a valid accent color.")

def launch_we():
    launch_wallpaper_engine()

# GUI Setup

root = tk.Tk()
root.title("Windows Theme Manager")
root.geometry("980x740")
root.configure(bg="#f3f3f3")

# Vars
wallpaper_path = tk.StringVar()
selected_color = tk.StringVar(value="#0078D7")
optional_color_vars = [tk.StringVar() for _ in range(5)]
optional_preview_labels = []
transparency_var = BooleanVar()

# Layout frames
main_frame = tk.Frame(root, bg="#f3f3f3")
main_frame.pack(pady=20)

left_frame = tk.Frame(main_frame, bg="#f3f3f3")
left_frame.pack(side="left", padx=40, anchor="n")

right_frame = tk.Frame(main_frame, bg="#f3f3f3")
right_frame.pack(side="right", padx=40, anchor="n")

# Left side (Accent + Wallpaper + WE)
tk.Label(left_frame, text="Accent Color", font=("Segoe UI", 12, "bold"), bg="#f3f3f3").pack(pady=(0, 6))

choose_color_btn = tk.Button(left_frame, text="Choose Accent", command=choose_accent_color, width=20, bg="#0078d7", fg="white", relief="flat")
choose_color_btn.pack()
add_hover_effect(choose_color_btn, "#0078d7")

preview_label = tk.Label(left_frame, text="Preview: #0078D7", font=("Segoe UI", 10), bg="#0078D7", fg="white", width=30, height=2)
preview_label.pack(pady=8)

tk.Label(left_frame, text="Wallpaper", font=("Segoe UI", 12, "bold"), bg="#f3f3f3").pack(pady=(16, 4))

wallpaper_label = tk.Label(left_frame, text="No file selected", fg="gray", bg="#f3f3f3")
wallpaper_label.pack()

wallpaper_preview_frame = tk.Frame(left_frame, width=300, height=160, bg="#dcdcdc", bd=1, relief="sunken")
wallpaper_preview_frame.pack_propagate(False)
wallpaper_preview_frame.pack(pady=6)

wallpaper_preview_label = tk.Label(wallpaper_preview_frame, text="No Preview", fg="gray", bg="#dcdcdc", font=("Segoe UI", 10))
wallpaper_preview_label.pack(expand=True)

choose_wallpaper_btn = tk.Button(left_frame, text="Choose Wallpaper", command=choose_wallpaper, width=20, bg="#0078d7", fg="white", relief="flat")
choose_wallpaper_btn.pack(pady=(0, 6))
add_hover_effect(choose_wallpaper_btn, "#0078d7")

open_we_btn = tk.Button(left_frame, text="Open Wallpaper Engine", command=launch_we, width=20, bg="#6f42c1", fg="white", relief="flat")
open_we_btn.pack()
add_hover_effect(open_we_btn, "#6f42c1")

# Right side (Side Colors + Transparency)
tk.Label(right_frame, text="Side Colors", font=("Segoe UI", 12, "bold"), bg="#f3f3f3").pack(pady=(0, 8))

# Optional side colors
optional_labels = [
    "Link and microphone color",
    "Taskbar focus / Alt+Tab highlight color",
    "Start button hover (W11 Pro only)",
    "Settings icons and buttons",
    "Pop-up window color"
]

for i, label_text in enumerate(optional_labels):
    row = tk.Frame(right_frame, bg="#f3f3f3")
    row.pack(pady=3, anchor="w")
    tk.Label(row, text=label_text, bg="#f3f3f3", width=36, anchor="w", font=("Segoe UI", 10)).pack(side="left")
    btn = tk.Button(row, text="Choose", bg="#0078d7", fg="white", relief="flat",
                    command=lambda idx=i: choose_optional_color(idx))
    btn.pack(side="left", padx=6)
    add_hover_effect(btn, "#0078d7")
    preview = tk.Label(row, text="Not set", bg="#dcdcdc", width=12, font=("Segoe UI", 10))
    preview.pack(side="left")
    optional_preview_labels.append(preview)

# Control Panel related side colors
control_panel_color_labels = [
    ("Hilight", "Highlight color (selected items)"),
    ("HotTrackingColor", "Hyperlink hover color"),
    ("MenuHilight", "Menu selection highlight")
]

control_panel_color_vars = {}
control_panel_preview_labels = {}

for key, label in control_panel_color_labels:
    control_panel_color_vars[key] = tk.StringVar()
    row = tk.Frame(right_frame, bg="#f3f3f3")
    row.pack(pady=3, anchor="w")
    tk.Label(row, text=label, bg="#f3f3f3", width=36, anchor="w", font=("Segoe UI", 10)).pack(side="left")
    btn = tk.Button(row, text="Choose", bg="#0078d7", fg="white", relief="flat",
                    command=partial(choose_control_panel_color, key))
    btn.pack(side="left", padx=6)
    add_hover_effect(btn, "#0078d7")
    preview = tk.Label(row, text="Not set", bg="#dcdcdc", width=12, font=("Segoe UI", 10))
    preview.pack(side="left")
    control_panel_preview_labels[key] = preview

# Transparency (keep exactly like old one)
load_transparency_setting()
transparency_check = ttk.Checkbutton(
    right_frame,
    text="Transparency",
    variable=transparency_var,
    command=toggle_transparency,
    style="Switch.TCheckbutton"  # old style preserved
)
transparency_check.pack(pady=8)

# Bottom buttons
button_section = tk.Frame(root, bg="#f3f3f3")
button_section.pack(pady=12)

apply_btn = tk.Button(button_section, text="Confirm & Apply", command=confirm_and_apply,
                      bg="#28a745", fg="white", relief="flat", height=2, width=25)
apply_btn.pack(side="left", padx=16)
add_hover_effect(apply_btn, "#28a745")

def do_save_preset():
    cp_colors = {k: v.get() for k, v in control_panel_color_vars.items()}
    save_preset(selected_color.get(), wallpaper_path.get(), [v.get() for v in optional_color_vars], cp_colors)

save_btn = tk.Button(button_section, text="Save Preset", command=do_save_preset,
                     relief="flat", bg="#0078d7", fg="white", height=2, width=25)
save_btn.pack(side="left", padx=16)
add_hover_effect(save_btn, "#0078d7")

# Presets list
presets_container = tk.Frame(root, bg="#f3f3f3")
presets_container.pack(pady=(16, 10), fill="x")

preset_count_label = tk.Label(presets_container, text=f"Saved Presets (0/{PRESET_LIMIT})", font=("Segoe UI", 11, "bold"), bg="#f3f3f3")
preset_count_label.pack(side="top", pady=(0, 8))

preset_frame = tk.Frame(presets_container, bg="#f3f3f3")
preset_frame.pack()

update_preset_viewer()

root.mainloop()
