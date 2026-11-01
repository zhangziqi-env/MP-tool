import copy
import json
import logging
import math
import os
import queue
import random
import sys
import threading
import traceback
import tkinter as tk
from io import BytesIO
from tkinter import messagebox, filedialog, simpledialog, ttk

import openpyxl
from docx import Document
from docx.shared import Inches, Pt
from PIL import Image, ImageDraw, ImageFont, ImageTk

# ------------------ Constants ------------------
APP_TITLE = "MP Label-V1"
ZOOM_MIN = 0.1
ZOOM_MAX = 10.0
UNDO_LIMIT = 50
PAN_MARGIN_PX = 220          # extra pixels rendered around the viewport for smooth panning
AUTOSAVE_DELAY_MS = 1500
APP_DIR = os.path.join(os.environ.get("LOCALAPPDATA") or os.path.expanduser("~"), "ENVer")
LOG_PATH = os.path.join(APP_DIR, "error.log")
AUTOSAVE_PATH = os.path.join(APP_DIR, "autosave.json")

# ------------------ Global Variables ------------------
image_paths = []
current_idx = 0
current_img = None
canvas = None
status_label = None
depth_label = None
root = None

marks_dict = {}
depth_dict = {}

# Zoom and drag related
canvas_scale = 1.0
canvas_offset_x = 0
canvas_offset_y = 0
drag_start_x = 0
drag_start_y = 0
is_dragging = False

undo_stack = []
redo_stack = []
# Lasso region storage
# each element: {"depth", "img_path", "bbox":(x1,y1,x2,y2), "polygon":[(x,y)...],
#                "judgment", "type", "subtype", "test_subtype"}
lasso_regions = []

# Preset type options
TYPE_OPTIONS = ['fibre', 'film', 'fragment', 'foram', 'bead', 'pellet', 'custom']
TEST_SUBTYPE_OPTIONS = ['label', 'unlabel']

# UI language: "en" for the English build, "zh" for the Chinese build.
LANG = "en"

# Colour annotation. Internal values stay stable in project JSON; labels are localized.
COLOR_OPTIONS = [
    ("", "Not set", "未设置"),
    ("transparent", "Transparent", "透明"),
    ("white", "White", "白色"),
    ("black", "Black", "黑色"),
    ("blue", "Blue", "蓝色"),
    ("red", "Red", "红色"),
    ("yellow", "Yellow", "黄色"),
    ("green", "Green", "绿色"),
    ("brown", "Brown", "棕色"),
    ("grey", "Grey", "灰色"),
    ("other", "Other", "其他"),
]

# Lasso related variables
lasso_points = []  # polygon vertices (store canvas coordinates)
lasso_rect_start = None  # rectangle start point (canvas coordinates)
lasso_rect_item = None  # canvas item id of the temporary rectangle preview

# Subtype/remark storage
subtype_dict = {}  # {(image_path, annotation_id): subtype_remark}

# Currently selected tool mode
tool_mode = "mark"  # "mark" or "rect_lasso" or "poly_lasso"

# Project / autosave state
project_path = None
dirty = False
autosave_job = None

# Rendering cache (only the latest rendered viewport is kept)
_render_cache = {"key": None, "photo": None}
_render_crop = None  # image-coordinate crop currently on the canvas
_configure_job = None

# ------------------ Helper Functions ------------------
def canvas_to_image_coords(x, y):
    """Convert canvas coordinates to original image coordinates"""
    img_x = int((x - canvas_offset_x) / canvas_scale)
    img_y = int((y - canvas_offset_y) / canvas_scale)
    return img_x, img_y

def image_to_canvas_coords(x, y):
    """Convert original image coordinates to canvas coordinates"""
    canvas_x = x * canvas_scale + canvas_offset_x
    canvas_y = y * canvas_scale + canvas_offset_y
    return canvas_x, canvas_y


def _loc(en_text, zh_text):
    """Return the text for the current UI language."""
    return zh_text if LANG == "zh" else en_text


def _color_display(value):
    for key, en_label, zh_label in COLOR_OPTIONS:
        if key == (value or ""):
            return _loc(en_label, zh_label)
    return value or _loc("Not set", "未设置")


def _color_choices():
    return [_color_display(key) for key, _, _ in COLOR_OPTIONS]


def _color_key(label):
    for key, en_label, zh_label in COLOR_OPTIONS:
        if label == _loc(en_label, zh_label):
            return key
    return ""


def _color_selector(parent, initial=""):
    """Build a `Colour: [combobox]` row and return (frame, StringVar)."""
    row = tk.Frame(parent)
    tk.Label(row, text=_loc("Colour:", "颜色："), font=('Arial', 11)).pack(side=tk.LEFT)
    var = tk.StringVar(value=_color_display(initial))
    combo = ttk.Combobox(row, textvariable=var, values=_color_choices(), state='readonly', width=16)
    combo.pack(side=tk.LEFT, padx=5)
    return row, var


def _center_dialog(dialog, width=None, height=None):
    """Place a Toplevel in the center of the screen (or its current monitor)."""
    try:
        dialog.update_idletasks()
        w = width or dialog.winfo_reqwidth()
        h = height or dialog.winfo_reqheight()
        screen_w = dialog.winfo_screenwidth()
        screen_h = dialog.winfo_screenheight()
        x = max(0, (screen_w - w) // 2)
        y = max(0, (screen_h - h) // 2)
        dialog.geometry(f"{w}x{h}+{x}+{y}")
    except tk.TclError:
        pass


def _draw_canvas_tips(canvas_w, canvas_h):
    if tool_mode == "mark":
        left_text = "Mode: 【Annotation Mode】Left click add | Ctrl+Left click delete"
    elif tool_mode == "rect_lasso":
        left_text = "Mode: 【Rectangle Lasso】Hold left button and drag to select area"
    else:
        left_text = "Mode: 【Polygon Lasso】Left click add vertex | Right click complete"
    canvas.create_text(10, canvas_h - 10, text=left_text, anchor='sw', fill='white', font=('Arial', 9), tags=("tips",))
    canvas.create_text(canvas_w - 10, canvas_h - 10, text="Scroll zoom | Middle button drag",
                       anchor='se', fill='white', font=('Arial', 9), tags=("tips",))


def _visible_render_box():
    """Image-coordinate crop that covers the viewport plus a pan margin."""
    if not current_img or not canvas:
        return None
    canvas_w = canvas.winfo_width()
    canvas_h = canvas.winfo_height()
    if canvas_w <= 1 or canvas_h <= 1:
        canvas_w, canvas_h = 800, 600
    scale = max(canvas_scale, 1e-6)
    img_w, img_h = current_img.size
    margin = PAN_MARGIN_PX / scale
    x0 = int(math.floor((0 - canvas_offset_x) / scale - margin))
    y0 = int(math.floor((0 - canvas_offset_y) / scale - margin))
    x1 = int(math.ceil((canvas_w - canvas_offset_x) / scale + margin))
    y1 = int(math.ceil((canvas_h - canvas_offset_y) / scale + margin))
    x0 = max(0, min(img_w, x0))
    y0 = max(0, min(img_h, y0))
    x1 = max(0, min(img_w, x1))
    y1 = max(0, min(img_h, y1))
    if x1 <= x0 or y1 <= y0:
        return None
    return x0, y0, x1, y1


def _clear_render_cache():
    global _render_crop
    _render_cache["key"] = None
    _render_cache["photo"] = None
    _render_crop = None


def redraw_canvas(event=None):
    global lasso_rect_item, _render_crop
    if not current_img or not canvas:
        if canvas is not None:
            canvas.delete("all")
        _clear_render_cache()
        return

    canvas_w = canvas.winfo_width()
    canvas_h = canvas.winfo_height()
    if canvas_w <= 1 or canvas_h <= 1:
        canvas_w, canvas_h = 800, 600

    box = _visible_render_box()
    if box is None:
        canvas.delete("all")
        _clear_render_cache()
        _draw_canvas_tips(canvas_w, canvas_h)
        return

    x0, y0, x1, y1 = box
    key = (image_paths[current_idx] if image_paths else "", round(canvas_scale, 6), x0, y0, x1, y1)
    if _render_cache["key"] == key and _render_cache["photo"] is not None:
        photo = _render_cache["photo"]
    else:
        disp_w = max(1, int(round((x1 - x0) * canvas_scale)))
        disp_h = max(1, int(round((y1 - y0) * canvas_scale)))
        crop = current_img.crop((x0, y0, x1, y1))
        if crop.size != (disp_w, disp_h):
            crop = crop.resize((disp_w, disp_h), Image.Resampling.LANCZOS)
        photo = ImageTk.PhotoImage(crop)
        _render_cache["key"] = key
        _render_cache["photo"] = photo

    canvas.delete("all")
    canvas.create_image(canvas_offset_x + x0 * canvas_scale,
                        canvas_offset_y + y0 * canvas_scale,
                        image=photo, anchor='nw', tags=("content",))
    canvas.image = photo
    _render_crop = box

    current_path = image_paths[current_idx]
    current_marks = marks_dict.get(current_path, [])
    show_text = canvas_scale >= 0.35
    radius = max(8, min(24, int(round(14 * (canvas_scale ** 0.25)))))

    for m in current_marks:
        cx = m['x'] * canvas_scale + canvas_offset_x
        cy = m['y'] * canvas_scale + canvas_offset_y
        if cx < -radius * 2 or cx > canvas_w + radius * 2 or cy < -radius * 2 or cy > canvas_h + radius * 2:
            continue
        canvas.create_oval(cx - radius, cy - radius, cx + radius, cy + radius,
                           fill='red', outline='white', width=2,
                           tags=("mark", f"mark_{m['id']}", "content"))
        canvas.create_text(cx, cy, text=str(m['id']), fill='white',
                           font=('Arial', 11, 'bold'), tags=("mark_text", "content"))
        if show_text:
            display_text = f"{m['type']}:{m['judgment']}"
            canvas.create_text(cx + radius + 4, cy - 8, text=display_text, fill='yellow',
                               font=('Arial', 10), anchor='w', tags=("mark_text", "content"))
            sub_key = (current_path, m['id'])
            has_sub = sub_key in subtype_dict and subtype_dict[sub_key]
            if has_sub:
                canvas.create_text(cx + radius + 4, cy + 5, text=subtype_dict[sub_key][:12],
                                   fill='cyan', font=('Arial', 8), anchor='w',
                                   tags=("mark_text", "content"))
            if m.get("color"):
                canvas.create_text(cx + radius + 4, cy + (18 if has_sub else 5),
                                   text=_color_display(m["color"]), fill='orange',
                                   font=('Arial', 8), anchor='w',
                                   tags=("mark_text", "content"))

    current_depth = depth_dict.get(current_path, _loc("Not set", "未设置"))
    if depth_label is not None:
        depth_label.config(text=f"Image {current_idx+1}/{len(image_paths)} | Depth: {current_depth}")

    _draw_canvas_tips(canvas_w, canvas_h)

    lasso_rect_item = None
    if tool_mode == "rect_lasso" and lasso_rect_start:
        try:
            mx = canvas.canvasx(canvas.winfo_pointerx() - canvas.winfo_rootx())
            my = canvas.canvasy(canvas.winfo_pointery() - canvas.winfo_rooty())
        except tk.TclError:
            mx, my = lasso_rect_start
        x1c, y1c = lasso_rect_start
        lasso_rect_item = canvas.create_rectangle(x1c, y1c, mx, my, outline='lime', width=2,
                                                  dash=(4, 4), tags=("temp_lasso",))

    if tool_mode == "poly_lasso" and len(lasso_points) >= 2:
        for i in range(len(lasso_points) - 1):
            canvas.create_line(lasso_points[i][0], lasso_points[i][1],
                               lasso_points[i+1][0], lasso_points[i+1][1],
                               fill='lime', width=2, tags=("temp_lasso",))
        for p in lasso_points:
            canvas.create_oval(p[0]-3, p[1]-3, p[0]+3, p[1]+3, fill='lime',
                               outline='white', tags=("temp_lasso",))
        try:
            mx = canvas.canvasx(canvas.winfo_pointerx() - canvas.winfo_rootx())
            my = canvas.canvasy(canvas.winfo_pointery() - canvas.winfo_rooty())
        except tk.TclError:
            mx, my = lasso_points[-1]
        canvas.create_line(lasso_points[-1][0], lasso_points[-1][1], mx, my,
                           fill='lime', width=2, dash=(4, 4), tags=("temp_lasso",))

def zoom(event):
    """Mouse wheel zoom (keeps the point under the cursor fixed)."""
    global canvas_scale, canvas_offset_x, canvas_offset_y

    if not current_img or not canvas:
        return

    delta = int(getattr(event, "delta", 0) or 0)
    if delta:
        scale_factor = 1.1 ** (delta / 120.0)
    else:
        num = int(getattr(event, "num", 0) or 0)
        if num == 4:
            scale_factor = 1.1
        elif num == 5:
            scale_factor = 0.9
        else:
            return

    new_scale = max(ZOOM_MIN, min(ZOOM_MAX, canvas_scale * scale_factor))
    if abs(new_scale - canvas_scale) < 1e-9:
        return

    mouse_x = canvas.canvasx(event.x)
    mouse_y = canvas.canvasy(event.y)
    img_x = (mouse_x - canvas_offset_x) / canvas_scale
    img_y = (mouse_y - canvas_offset_y) / canvas_scale

    canvas_scale = new_scale
    canvas_offset_x = mouse_x - img_x * canvas_scale
    canvas_offset_y = mouse_y - img_y * canvas_scale
    redraw_canvas()

def start_drag(event):
    """Start panning with the middle mouse button."""
    global drag_start_x, drag_start_y, is_dragging
    if int(getattr(event, "num", 0) or 0) != 2:
        return
    is_dragging = True
    drag_start_x = event.x
    drag_start_y = event.y
    if canvas is not None:
        canvas.config(cursor="fleur")
        try:
            canvas.grab_set()
        except tk.TclError:
            pass

def on_drag(event):
    """Drag canvas (move the cached render when possible)."""
    global canvas_offset_x, canvas_offset_y, drag_start_x, drag_start_y
    if not is_dragging:
        return
    dx = event.x - drag_start_x
    dy = event.y - drag_start_y
    if dx == 0 and dy == 0:
        return
    drag_start_x = event.x
    drag_start_y = event.y
    canvas_offset_x += dx
    canvas_offset_y += dy
    if canvas is not None and _pan_keeps_cached_render():
        canvas.move("content", dx, dy)
    else:
        redraw_canvas()


def _pan_keeps_cached_render():
    """True if the current viewport is still inside the rendered crop."""
    if not current_img or _render_crop is None or not canvas:
        return False
    scale = max(canvas_scale, 1e-6)
    canvas_w = canvas.winfo_width() if canvas.winfo_width() > 1 else 800
    canvas_h = canvas.winfo_height() if canvas.winfo_height() > 1 else 600
    img_w, img_h = current_img.size
    x0, y0, x1, y1 = _render_crop
    if (x0, y0, x1, y1) == (0, 0, img_w, img_h):
        return True
    vx0 = (0 - canvas_offset_x) / scale
    vy0 = (0 - canvas_offset_y) / scale
    vx1 = (canvas_w - canvas_offset_x) / scale
    vy1 = (canvas_h - canvas_offset_y) / scale
    return x0 <= vx0 and y0 <= vy0 and x1 >= vx1 and y1 >= vy1

def stop_drag(event):
    """End canvas drag."""
    global is_dragging
    is_dragging = False
    if canvas is not None:
        canvas.config(cursor="")
        try:
            canvas.grab_release()
        except tk.TclError:
            pass

def _reset_lasso_state():
    global lasso_points, lasso_rect_start, lasso_rect_item
    lasso_points = []
    lasso_rect_start = None
    lasso_rect_item = None
    if canvas is not None:
        canvas.delete("temp_lasso")
        canvas.unbind("<B1-Motion>")
        canvas.unbind("<ButtonRelease-1>")
        try:
            canvas.grab_release()
        except tk.TclError:
            pass


def set_tool_mark():
    """Switch to annotation mode."""
    global tool_mode
    tool_mode = "mark"
    _reset_lasso_state()
    if status_label is not None:
        status_label.config(text="Switched to 【Annotation Mode】")
    redraw_canvas()


def set_tool_rect_lasso():
    """Switch to rectangle lasso mode."""
    global tool_mode
    tool_mode = "rect_lasso"
    _reset_lasso_state()
    if status_label is not None:
        status_label.config(text="Switched to 【Rectangle Lasso Mode】- Hold left button and drag to select area")
    redraw_canvas()


def set_tool_poly_lasso():
    """Switch to polygon lasso mode."""
    global tool_mode
    tool_mode = "poly_lasso"
    _reset_lasso_state()
    if status_label is not None:
        status_label.config(text="Switched to 【Polygon Lasso Mode】- Left click add vertex, right click to complete")
    redraw_canvas()



def point_in_polygon(x, y, poly):
    """Ray casting algorithm to check if point is inside polygon"""
    inside = False
    n = len(poly)
    for i in range(n):
        x1, y1 = poly[i]
        x2, y2 = poly[(i + 1) % n]
        if ((y1 > y) != (y2 > y)) and (x < (x2 - x1) * (y - y1) / (y2 - y1) + x1):
            inside = not inside
    return inside


def _drop_regions_for(img_path):
    """Remove every saved lasso region that belongs to img_path."""
    lasso_regions[:] = [r for r in lasso_regions if r.get("img_path") != img_path]


def _retag_regions_depth(img_path, new_depth):
    """Keep saved lasso regions in sync when a depth value changes."""
    for r in lasso_regions:
        if r.get("img_path") == img_path:
            r["depth"] = new_depth


def batch_add_marks_in_region_rect(bbox):
    """Batch add annotation points within rectangle area"""
    current_path = image_paths[current_idx]
    current_depth = depth_dict.get(current_path, "")

    if not current_depth:
        messagebox.showinfo("Info", "Please set depth first")
        return

    left, top, right, bottom = bbox
    img_w, img_h = current_img.size
    left = max(0, min(left, img_w - 1))
    top = max(0, min(top, img_h - 1))
    right = max(left, min(right, img_w - 1))
    bottom = max(top, min(bottom, img_h - 1))

    dialog = tk.Toplevel(root)
    dialog.title("Batch Add Annotations in Rectangle Area")
    dialog.geometry("400x540")
    dialog.transient(root)
    dialog.grab_set()
    _center_dialog(dialog, 400, 540)

    tk.Label(dialog, text=f"Rectangle area: ({left},{top}) to ({right},{bottom})").pack(pady=10)

    tk.Label(dialog, text="How many annotation points to add?").pack(pady=(10,5))
    count_var = tk.StringVar(value="1")
    tk.Spinbox(dialog, from_=1, to=100, textvariable=count_var, width=10).pack(pady=5)

    tk.Label(dialog, text="Type:").pack(pady=(10,5))
    type_var = tk.StringVar()
    type_combo = ttk.Combobox(dialog, textvariable=type_var, values=TYPE_OPTIONS, state='readonly', width=20)
    type_combo.pack(pady=5)
    type_var.set('fibre')

    tk.Label(dialog, text="Judgment:").pack(pady=(10,5))
    judgment_var = tk.StringVar()
    judgment_combo = ttk.Combobox(dialog, textvariable=judgment_var, values=['MP', 'NoMP', 'Test'], state='readonly', width=20)
    judgment_combo.pack(pady=5)
    judgment_var.set('MP')

    test_frame = tk.Frame(dialog)
    tk.Label(test_frame, text="Test subtype:", font=('Arial', 10)).pack(side=tk.LEFT, padx=5)
    test_subtype_var = tk.StringVar()
    test_combo = ttk.Combobox(test_frame, textvariable=test_subtype_var, values=['label', 'unlabel'], state='disabled', width=10)
    test_combo.pack(side=tk.LEFT, padx=5)
    test_frame.pack(pady=5)

    def on_judgment_change(*args):
        if judgment_var.get() == 'Test':
            test_combo.config(state='readonly')
            test_subtype_var.set('label')
        else:
            test_combo.config(state='disabled')
            test_subtype_var.set('')

    judgment_var.trace_add('write', on_judgment_change)
    on_judgment_change()

    tk.Label(dialog, text="Remark (optional):").pack(pady=(10,5))
    subtype_entry = tk.Entry(dialog, width=30)
    subtype_entry.pack(pady=5)

    color_row, color_var = _color_selector(dialog)
    color_row.pack(pady=5)

    def apply():
        try:
            count = max(1, int(count_var.get()))
        except Exception:
            count = 1

        type_val = type_var.get()
        judgment_val = judgment_var.get()

        test_subtype_val = ""
        if judgment_val == "Test":
            test_subtype_val = test_subtype_var.get()
            if not test_subtype_val:
                messagebox.showwarning("Warning", "Please select Test subtype (label/unlabel)")
                return

        remark = subtype_entry.get().strip()
        color_val = _color_key(color_var.get())
        save_state()

        depth_marks = []
        for p, marks in marks_dict.items():
            if depth_dict.get(p, "") == current_depth:
                depth_marks.extend(marks)
        next_id = max([m['id'] for m in depth_marks], default=0) + 1
        current_marks = marks_dict.get(current_path, [])

        for i in range(count):
            rand_x = random.randint(left, right)
            rand_y = random.randint(top, bottom)
            current_marks.append({
                "id": next_id + i,
                "x": rand_x,
                "y": rand_y,
                "type": type_val,
                "judgment": judgment_val,
                "test_subtype": test_subtype_val,
                "color": color_val
            })
            if remark:
                subtype_dict[(current_path, next_id + i)] = remark
            elif test_subtype_val:
                subtype_dict[(current_path, next_id + i)] = f"Test-{test_subtype_val}"

        marks_dict[current_path] = current_marks
        lasso_regions.append({
            "depth": current_depth,
            "img_path": current_path,
            "bbox": (left, top, right, bottom),
            "judgment": judgment_val,
            "type": type_val,
            "subtype": remark,
            "test_subtype": test_subtype_val,
            "color": color_val
        })

        mark_dirty()
        redraw_canvas()
        status_label.config(text=f"Added {count} annotation points in rectangle area")
        dialog.destroy()

    tk.Button(dialog, text="Add Annotations", command=apply, bg="lightgreen", width=15).pack(pady=20)

def add_mark_dialog(img_x, img_y):
    """Popup dialog for type, judgment and colour selection"""
    dialog = tk.Toplevel(root)
    dialog.title("Add Annotation")
    dialog.geometry("450x540")
    dialog.transient(root)
    dialog.grab_set()
    _center_dialog(dialog, 450, 540)

    tk.Label(dialog, text=f"Position: ({img_x}, {img_y})", font=('Arial', 10)).pack(pady=(10, 5))
    
    tk.Label(dialog, text="Type:", font=('Arial', 11)).pack(pady=(5, 5))
    type_var = tk.StringVar()
    type_combo = ttk.Combobox(dialog, textvariable=type_var, values=TYPE_OPTIONS, state='readonly', width=20)
    type_combo.pack(pady=5)
    type_combo.current(0)
    
    custom_frame = tk.Frame(dialog)
    tk.Label(custom_frame, text="Custom type:").pack(side=tk.LEFT)
    custom_entry = tk.Entry(custom_frame, width=15)
    custom_entry.pack(side=tk.LEFT, padx=5)
    
    def on_type_change(*args):
        if type_var.get() == "custom":
            custom_frame.pack(pady=5)
        else:
            custom_frame.pack_forget()
    
    type_var.trace_add('write', on_type_change)
    
    tk.Label(dialog, text="Judgment:", font=('Arial', 11)).pack(pady=(5, 5))
    judgment_var = tk.StringVar()
    judgment_combo = ttk.Combobox(dialog, textvariable=judgment_var, values=['MP', 'NoMP', 'Test'], state='readonly', width=20)
    judgment_combo.pack(pady=5)
    judgment_combo.current(0)
    
    test_subtype_frame = tk.Frame(dialog)
    tk.Label(test_subtype_frame, text="Test subtype:", font=('Arial', 10)).pack(side=tk.LEFT, padx=5)
    test_subtype_var = tk.StringVar()
    test_subtype_combo = ttk.Combobox(test_subtype_frame, textvariable=test_subtype_var, 
                                       values=TEST_SUBTYPE_OPTIONS, state='readonly', width=10)
    test_subtype_combo.pack(side=tk.LEFT, padx=5)
    
    def on_judgment_change(*args):
        if judgment_var.get() == "Test":
            test_subtype_frame.pack(pady=5)
        else:
            test_subtype_frame.pack_forget()
    
    judgment_var.trace_add('write', on_judgment_change)

    color_row, color_var = _color_selector(dialog)
    color_row.pack(pady=(5, 5))
    
    tk.Label(dialog, text="Remark (optional):", font=('Arial', 11)).pack(pady=(5, 5))
    subtype_entry = tk.Entry(dialog, width=30)
    subtype_entry.pack(pady=5)
    
    result = [None, None, None, None, None]
    
    def on_ok():
        type_val = type_var.get()
        if type_val == "custom":
            type_val = custom_entry.get().strip()
            if not type_val:
                messagebox.showwarning("Warning", "Please enter custom type")
                return
        judgment_val = judgment_var.get()
        
        test_subtype_val = ""
        if judgment_val == "Test":
            test_subtype_val = test_subtype_var.get()
            if not test_subtype_val:
                messagebox.showwarning("Warning", "Please select Test subtype (label/unlabel)")
                return
        
        remark_val = subtype_entry.get().strip()
        result[0] = type_val
        result[1] = judgment_val
        result[2] = test_subtype_val
        result[3] = remark_val
        result[4] = _color_key(color_var.get())
        dialog.destroy()
    
    def on_cancel():
        dialog.destroy()
    
    btn_frame = tk.Frame(dialog)
    btn_frame.pack(pady=20)
    tk.Button(btn_frame, text="OK", command=on_ok, width=10).pack(side=tk.LEFT, padx=10)
    tk.Button(btn_frame, text="Cancel", command=on_cancel, width=10).pack(side=tk.LEFT, padx=10)
    
    dialog.wait_window()
    return result[0], result[1], result[2], result[3], result[4]

def on_rect_release(event):
    """Rectangle lasso release, popup to add annotation points"""
    global lasso_rect_start, lasso_rect_item, tool_mode

    if lasso_rect_start is None:
        return

    if not current_img:
        lasso_rect_start = None
        lasso_rect_item = None
        canvas.delete("temp_lasso")
        try:
            canvas.grab_release()
        except tk.TclError:
            pass
        return

    x1, y1 = lasso_rect_start
    x2, y2 = event.x, event.y

    img_x1, img_y1 = canvas_to_image_coords(x1, y1)
    img_x2, img_y2 = canvas_to_image_coords(x2, y2)

    left = max(0, min(img_x1, img_x2))
    top = max(0, min(img_y1, img_y2))
    right = min(current_img.width, max(img_x1, img_x2))
    bottom = min(current_img.height, max(img_y1, img_y2))

    lasso_rect_start = None
    lasso_rect_item = None
    canvas.delete("temp_lasso")
    canvas.unbind("<B1-Motion>")
    canvas.unbind("<ButtonRelease-1>")
    try:
        canvas.grab_release()
    except tk.TclError:
        pass

    if right > left and bottom > top:
        batch_add_marks_in_region_rect((left, top, right, bottom))
    else:
        status_label.config(text="Invalid rectangle area, please reselect")


def on_canvas_click(event):
    """Canvas click event"""
    global lasso_rect_start, lasso_rect_item, lasso_points, tool_mode

    if not current_img:
        return

    if tool_mode == "rect_lasso":
        if event.num == 1:
            lasso_rect_start = (event.x, event.y)
            lasso_rect_item = None
            canvas.delete("temp_lasso")
            lasso_rect_item = canvas.create_rectangle(event.x, event.y, event.x, event.y,
                                                      outline='lime', width=2, dash=(4, 4),
                                                      tags=("temp_lasso",))
            canvas.bind("<B1-Motion>", on_rect_drag)
            canvas.bind("<ButtonRelease-1>", on_rect_release)
            try:
                canvas.grab_set()
            except tk.TclError:
                pass
            status_label.config(text="Rectangle lasso: drag to select area...")

    elif tool_mode == "poly_lasso":
        if event.num == 1:
            lasso_points.append((event.x, event.y))
            redraw_canvas()
            status_label.config(text=f"Polygon lasso: added {len(lasso_points)} vertices, right click to complete")
        elif event.num == 3:
            finish_polygon_lasso()

    else:  # Annotation mode
        if event.state & 0x0004:
            delete_mark_at_position(event)
        elif event.num == 1:
            add_new_mark(event)


def on_rect_drag(event):
    """Rectangle lasso dragging (only move the preview rectangle)."""
    if lasso_rect_start is None:
        return
    if lasso_rect_item is None:
        redraw_canvas()
        return
    x1, y1 = lasso_rect_start
    try:
        canvas.coords(lasso_rect_item, x1, y1, event.x, event.y)
    except tk.TclError:
        redraw_canvas()





def finish_polygon_lasso(self=None):
    """Complete polygon lasso, batch add annotation points within polygon area and save region"""
    global lasso_points, lasso_regions, tool_mode

    if len(lasso_points) < 3:
        status_label.config(text="Polygon requires at least 3 vertices, continue adding")
        return

    img_polygon = []
    for px, py in lasso_points:
        img_x, img_y = canvas_to_image_coords(px, py)
        img_polygon.append((img_x, img_y))

    xs = [p[0] for p in img_polygon]
    ys = [p[1] for p in img_polygon]
    min_x, max_x = max(0, min(xs)), min(current_img.width, max(xs))
    min_y, max_y = max(0, min(ys)), min(current_img.height, max(ys))

    canvas.delete("temp_lasso")

    current_path = image_paths[current_idx]
    current_depth = depth_dict.get(current_path, "")

    if not current_depth:
        messagebox.showinfo("Info", "Please set depth first")
        lasso_points = []
        redraw_canvas()
        return

    dialog = tk.Toplevel(root)
    dialog.title("Batch Add Annotations in Polygon Area")
    dialog.geometry("450x610")
    dialog.transient(root)
    dialog.grab_set()
    _center_dialog(dialog, 450, 610)

    tk.Label(dialog, text=f"Polygon area ({len(img_polygon)} vertices)", font=('Arial', 10)).pack(pady=10)

    tk.Label(dialog, text="How many annotation points to add?", font=('Arial', 11)).pack(pady=(10,5))
    count_var = tk.StringVar(value="1")
    count_spin = tk.Spinbox(dialog, from_=1, to=50, textvariable=count_var, width=10)
    count_spin.pack(pady=5)

    tk.Label(dialog, text="Type:", font=('Arial', 11)).pack(pady=(10,5))
    type_var = tk.StringVar()
    type_combo = ttk.Combobox(dialog, textvariable=type_var, values=TYPE_OPTIONS, state='readonly', width=20)
    type_combo.pack(pady=5)
    type_combo.current(0)

    custom_frame = tk.Frame(dialog)
    tk.Label(custom_frame, text="Custom type:").pack(side=tk.LEFT)
    custom_entry = tk.Entry(custom_frame, width=15)
    custom_entry.pack(side=tk.LEFT, padx=5)

    def on_type_change(*args):
        if type_var.get() == "custom":
            custom_frame.pack(pady=5)
        else:
            custom_frame.pack_forget()
    type_var.trace_add('write', on_type_change)

    tk.Label(dialog, text="Judgment:", font=('Arial', 11)).pack(pady=(10,5))
    judgment_var = tk.StringVar()
    judgment_combo = ttk.Combobox(dialog, textvariable=judgment_var, values=['MP', 'NoMP', 'Test'], state='readonly', width=20)
    judgment_combo.pack(pady=5)
    judgment_combo.current(0)

    test_subtype_frame = tk.Frame(dialog)
    tk.Label(test_subtype_frame, text="Test subtype:", font=('Arial', 10)).pack(side=tk.LEFT, padx=5)
    test_subtype_var = tk.StringVar()
    test_subtype_combo = ttk.Combobox(test_subtype_frame, textvariable=test_subtype_var,
                                      values=['label', 'unlabel'], state='readonly', width=10)
    test_subtype_combo.pack(side=tk.LEFT, padx=5)

    def on_judgment_change(*args):
        if judgment_var.get() == "Test":
            test_subtype_frame.pack(pady=5)
        else:
            test_subtype_frame.pack_forget()

    judgment_var.trace_add('write', on_judgment_change)

    tk.Label(dialog, text="Subtype/Remark (optional):", font=('Arial', 11)).pack(pady=(10,5))
    subtype_entry = tk.Entry(dialog, width=30)
    subtype_entry.pack(pady=5)

    color_row, color_var = _color_selector(dialog)
    color_row.pack(pady=5)

    def apply_batch():
        try:
            count = max(1, int(count_var.get()))
        except Exception:
            count = 1

        type_val = type_var.get()
        if type_val == "custom":
            type_val = custom_entry.get().strip()
            if not type_val:
                messagebox.showwarning("Warning", "Please enter custom type")
                return
        judgment_val = judgment_var.get()

        test_subtype_val = ""
        if judgment_val == "Test":
            test_subtype_val = test_subtype_var.get()
            if not test_subtype_val:
                messagebox.showwarning("Warning", "Please select Test subtype (label/unlabel)")
                return

        subtype_val = subtype_entry.get().strip()
        color_val = _color_key(color_var.get())
        save_state()

        depth_marks = []
        for p, marks in marks_dict.items():
            if depth_dict.get(p, "") == current_depth:
                depth_marks.extend(marks)
        next_id = max([m['id'] for m in depth_marks], default=0) + 1
        current_marks = marks_dict.get(current_path, [])

        max_sample_x = max(min_x, min(max_x, current_img.width - 1))
        max_sample_y = max(min_y, min(max_y, current_img.height - 1))

        added = 0
        attempts = 0
        while added < count and attempts < 500:
            rand_x = random.randint(min_x, max_sample_x)
            rand_y = random.randint(min_y, max_sample_y)
            if point_in_polygon(rand_x, rand_y, img_polygon):
                new_mark = {
                    "id": next_id + added,
                    "x": rand_x,
                    "y": rand_y,
                    "type": type_val,
                    "judgment": judgment_val,
                    "test_subtype": test_subtype_val,
                    "color": color_val
                }
                current_marks.append(new_mark)
                if subtype_val:
                    subtype_dict[(current_path, next_id + added)] = subtype_val
                elif test_subtype_val:
                    subtype_dict[(current_path, next_id + added)] = f"Test-{test_subtype_val}"
                added += 1
            attempts += 1

        marks_dict[current_path] = current_marks
        lasso_regions.append({
            "depth": current_depth,
            "img_path": current_path,
            "bbox": (min_x, min_y, max_x, max_y),
            "polygon": img_polygon,
            "judgment": judgment_val,
            "type": type_val,
            "subtype": subtype_val,
            "test_subtype": test_subtype_val,
            "color": color_val
        })

        mark_dirty()
        redraw_canvas()
        if added:
            status_label.config(text=f"Polygon area saved! Total saved {len(lasso_regions)} areas | Added {added} annotation points, IDs {next_id} - {next_id+added-1}")
        else:
            status_label.config(text=f"Polygon area saved, but no valid positions were found (requested {count}) | Total saved {len(lasso_regions)} areas")
        if added < count and added:
            status_label.config(text=f"Only {added} valid positions found in polygon (requested {count}) | Total saved {len(lasso_regions)} areas")
        dialog.destroy()

    tk.Button(dialog, text="Add Annotations and Save Area", command=apply_batch, bg="lightgreen", width=20).pack(pady=20)

    lasso_points = []
    redraw_canvas()

def delete_mark_at_position(event):
    """Delete annotation point at click position"""
    current_path = image_paths[current_idx]
    current_marks = marks_dict.get(current_path, [])

    min_dist = max(4.0, min(250.0, 20.0 * canvas_scale))
    target_idx = -1
    for i, mark in enumerate(current_marks):
        mx, my = image_to_canvas_coords(mark['x'], mark['y'])
        dist = ((event.x - mx) ** 2 + (event.y - my) ** 2) ** 0.5
        if dist < min_dist:
            min_dist = dist
            target_idx = i

    if target_idx >= 0:
        save_state()
        deleted_mark = current_marks.pop(target_idx)
        sub_key = (current_path, deleted_mark['id'])
        if sub_key in subtype_dict:
            del subtype_dict[sub_key]
        if current_marks:
            marks_dict[current_path] = current_marks
        else:
            marks_dict.pop(current_path, None)
        mark_dirty()
        redraw_canvas()
        status_label.config(text=f"Deleted ID {deleted_mark['id']}")
    else:
        status_label.config(text="No annotation point found nearby")

def add_new_mark(event):
    """Add new annotation point"""
    if not current_img:
        return
    orig_x, orig_y = canvas_to_image_coords(event.x, event.y)
    if orig_x < 0 or orig_y < 0 or orig_x >= current_img.width or orig_y >= current_img.height:
        return

    current_path = image_paths[current_idx]
    current_depth = depth_dict.get(current_path, "")

    depth_marks = []
    for p, marks in marks_dict.items():
        if depth_dict.get(p, "") == current_depth:
            depth_marks.extend(marks)

    new_id = (max((m['id'] for m in depth_marks), default=0) + 1)

    type_val, judgment_val, test_subtype_val, remark_val, color_val = add_mark_dialog(orig_x, orig_y)
    if type_val is None:
        return

    save_state()
    current_marks = marks_dict.get(current_path, [])
    current_marks.append({
        "id": new_id,
        "x": orig_x,
        "y": orig_y,
        "type": type_val,
        "judgment": judgment_val,
        "test_subtype": test_subtype_val,
        "color": color_val
    })
    marks_dict[current_path] = current_marks

    if remark_val:
        subtype_dict[(current_path, new_id)] = remark_val
    elif test_subtype_val:
        subtype_dict[(current_path, new_id)] = f"Test-{test_subtype_val}"

    mark_dirty()
    redraw_canvas()
    status_label.config(text=f"Added ID {new_id} | Type: {type_val} | Judgment: {judgment_val}" +
                       (f" ({test_subtype_val})" if test_subtype_val else "") +
                       (f" | {_loc('Colour', '颜色')}: {_color_display(color_val)}" if color_val else ""))

# ------------------ Original Functionality Functions ------------------
def save_marks_to_image():
    global current_idx
    if not current_img:
        messagebox.showinfo("Info", "No image")
        return

    current_path = image_paths[current_idx]
    current_marks = marks_dict.get(current_path, [])

    if not current_marks:
        if not messagebox.askyesno("Info", "Current image has no annotations. Save anyway?"):
            return

    img_copy = current_img.copy()
    draw = ImageDraw.Draw(img_copy)
    try:
        font = ImageFont.truetype("arial.ttf", 20)
        font_small = ImageFont.truetype("arial.ttf", 18)
    except Exception:
        font = ImageFont.load_default()
        font_small = font

    for mark in current_marks:
        x, y = mark['x'], mark['y']
        r = 15
        draw.ellipse([x-r, y-r, x+r, y+r], fill='red', outline='white', width=2)
        draw.text((x-8, y-12), str(mark['id']), fill='white', font=font)
        display_text = f"{mark['type']}:{mark['judgment']}"
        draw.text((x+15, y-10), display_text, fill='yellow', font=font_small)

        sub_key = (current_path, mark['id'])
        if sub_key in subtype_dict and subtype_dict[sub_key]:
            draw.text((x+15, y+5), subtype_dict[sub_key][:10], fill='cyan', font=font_small)
        if mark.get("color"):
            draw.text((x+15, y+20), _color_display(mark["color"]), fill='orange', font=font_small)

    current_depth = depth_dict.get(current_path, "")
    if current_depth:
        draw.text((10, 10), f"Depth: {current_depth}", fill='white', font=font)

    dir_name = os.path.dirname(current_path)
    base_name = os.path.splitext(os.path.basename(current_path))[0]
    save_path = os.path.join(dir_name, f"{base_name}_annotated.png")
    try:
        img_copy.save(save_path)
    except Exception as exc:
        logging.exception("Failed to save annotated image")
        messagebox.showerror("Save failed", f"Could not save annotated image:\n{save_path}\n\n{exc}")
        return
    messagebox.showinfo("Success", f"Saved to: {save_path}")

    answer = messagebox.askyesno(
        "Add to list",
        "Add this annotated image to the list for further editing?\n\n"
        "It will be treated as a separate image and counted again in exports.\n"
        "Choose No if you only need the annotated PNG file.",
        default=messagebox.NO)
    if not answer:
        return

    save_state()
    # Deep-copy the marks, assign fresh IDs and copy remarks, so the copy never
    # shares mutable state with the original image.
    depth_marks = []
    for p, marks in marks_dict.items():
        if depth_dict.get(p, "") == current_depth:
            depth_marks.extend(marks)
    next_id = max([m['id'] for m in depth_marks], default=0) + 1

    new_marks = copy.deepcopy(current_marks)
    old_to_new = {}
    for m in new_marks:
        old_to_new[m['id']] = next_id
        m['id'] = next_id
        next_id += 1

    for (p, i) in list(subtype_dict.keys()):
        if p == save_path:
            del subtype_dict[(p, i)]
    for old_id, new_id in old_to_new.items():
        if (current_path, old_id) in subtype_dict:
            subtype_dict[(save_path, new_id)] = subtype_dict[(current_path, old_id)]

    if save_path in image_paths:
        current_idx = image_paths.index(save_path)
    else:
        image_paths.append(save_path)
        current_idx = len(image_paths) - 1
    marks_dict[save_path] = new_marks
    depth_dict[save_path] = current_depth
    mark_dirty()
    load_image_by_index()

def renumber_current_depth():
    if not image_paths:
        return

    current_path = image_paths[current_idx]
    current_depth = depth_dict.get(current_path, "")
    if not current_depth:
        messagebox.showinfo("Info", "Current image has no depth set. Please set depth first")
        return

    depth_images = []
    for img_path, marks in marks_dict.items():
        if depth_dict.get(img_path, "") == current_depth:
            depth_images.append((img_path, marks))

    if not depth_images:
        messagebox.showinfo("Info", f"No annotation data under depth「{current_depth}」")
        return

    if not messagebox.askyesno("Confirm", f"Renumber all annotations under depth「{current_depth}」(total {sum(len(m) for _, m in depth_images)}). Continue?"):
        return

    save_state()
    old_to_new = {}
    new_id = 1
    depth_paths = {p for p, _ in depth_images}

    for img_path, marks in sorted(depth_images, key=lambda x: x[0]):
        for mark in marks:
            old_to_new[(img_path, mark['id'])] = new_id
            mark['id'] = new_id
            new_id += 1
        marks_dict[img_path] = marks

    # subtype_dict is keyed by (image, annotation id), so it must be remapped too.
    remapped = {}
    for (img_path, old_id), value in subtype_dict.items():
        if img_path in depth_paths:
            new_id_for_mark = old_to_new.get((img_path, old_id))
            if new_id_for_mark is None:
                continue
            remapped[(img_path, new_id_for_mark)] = value
        else:
            remapped[(img_path, old_id)] = value
    subtype_dict.clear()
    subtype_dict.update(remapped)

    mark_dirty()
    redraw_canvas()
    messagebox.showinfo("Success", f"Depth「{current_depth}」renumbered. Total {new_id-1} annotations")

def continue_numbering_current_depth():
    if not image_paths:
        return
    
    current_path = image_paths[current_idx]
    current_depth = depth_dict.get(current_path, "")
    if not current_depth:
        messagebox.showinfo("Info", "Current image has no depth set. Please set depth first")
        return
    
    depth_marks = []
    for img_path, marks in marks_dict.items():
        if depth_dict.get(img_path, "") == current_depth:
            depth_marks.extend(marks)
    
    if not depth_marks:
        messagebox.showinfo("Info", f"No annotation data under depth「{current_depth}」. Please add annotations first")
        return
    
    max_id = max(m['id'] for m in depth_marks)
    status_label.config(text=f"Current depth「{current_depth}」max ID is {max_id}, new annotations will start from {max_id+1}")

def delete_mark_by_id():
    if not image_paths:
        return

    current_path = image_paths[current_idx]
    current_marks = marks_dict.get(current_path, [])

    if not current_marks:
        status_label.config(text="Current image has no annotations")
        return

    ids = sorted([m['id'] for m in current_marks])

    dialog = tk.Toplevel(root)
    dialog.title("Delete Annotation")
    dialog.geometry("300x150")
    dialog.transient(root)
    dialog.grab_set()
    _center_dialog(dialog, 300, 150)

    tk.Label(dialog, text=f"Current image IDs: {ids}", font=('Arial', 10)).pack(pady=10)
    tk.Label(dialog, text="Enter ID to delete:").pack(pady=5)

    id_var = tk.StringVar()
    entry = tk.Entry(dialog, textvariable=id_var, width=10)
    entry.pack(pady=5)
    entry.focus()

    def do_delete():
        try:
            del_id = int(id_var.get())
        except ValueError:
            messagebox.showwarning("Warning", "Please enter a numeric ID")
            return

        for i, mark in enumerate(current_marks):
            if mark['id'] == del_id:
                save_state()
                current_marks.pop(i)
                if current_marks:
                    marks_dict[current_path] = current_marks
                else:
                    marks_dict.pop(current_path, None)
                sub_key = (current_path, del_id)
                if sub_key in subtype_dict:
                    del subtype_dict[sub_key]
                mark_dirty()
                redraw_canvas()
                status_label.config(text=f"Deleted ID {del_id}")
                dialog.destroy()
                return

        messagebox.showwarning("Warning", f"ID {del_id} not found")

    tk.Button(dialog, text="Delete", command=do_delete, bg="#FF6666", width=10).pack(pady=10)

def edit_mark_subtype():
    if not image_paths:
        return

    current_path = image_paths[current_idx]
    current_marks = marks_dict.get(current_path, [])

    if not current_marks:
        status_label.config(text="Current image has no annotations")
        return

    ids = sorted([m['id'] for m in current_marks])

    dialog = tk.Toplevel(root)
    dialog.title("Edit Subtype/Remark")
    dialog.geometry("350x200")
    dialog.transient(root)
    dialog.grab_set()
    _center_dialog(dialog, 350, 200)

    tk.Label(dialog, text=f"Current image IDs: {ids}", font=('Arial', 10)).pack(pady=10)
    tk.Label(dialog, text="Enter ID:").pack(pady=5)

    id_var = tk.StringVar()
    entry = tk.Entry(dialog, textvariable=id_var, width=10)
    entry.pack(pady=5)

    tk.Label(dialog, text="Subtype/Remark:").pack(pady=5)
    subtype_entry = tk.Entry(dialog, width=30)
    subtype_entry.pack(pady=5)

    def do_edit():
        try:
            edit_id = int(id_var.get())
        except ValueError:
            messagebox.showwarning("Warning", "Please enter a numeric ID")
            return

        for mark in current_marks:
            if mark['id'] == edit_id:
                save_state()
                sub_key = (current_path, edit_id)
                new_subtype = subtype_entry.get().strip()
                if new_subtype:
                    subtype_dict[sub_key] = new_subtype
                elif sub_key in subtype_dict:
                    del subtype_dict[sub_key]
                mark_dirty()
                redraw_canvas()
                status_label.config(text=f"Updated subtype for ID {edit_id}")
                dialog.destroy()
                return

        messagebox.showwarning("Warning", f"ID {edit_id} not found")

    tk.Button(dialog, text="Save", command=do_edit, bg="lightgreen", width=10).pack(pady=10)

def edit_mark_color():
    """Edit the colour annotation of an existing mark."""
    if not image_paths:
        return

    current_path = image_paths[current_idx]
    current_marks = marks_dict.get(current_path, [])

    if not current_marks:
        status_label.config(text="Current image has no annotations")
        return

    ids = sorted([m['id'] for m in current_marks])

    dialog = tk.Toplevel(root)
    dialog.title("Edit Colour")
    dialog.geometry("350x230")
    dialog.transient(root)
    dialog.grab_set()
    _center_dialog(dialog, 350, 230)

    tk.Label(dialog, text=f"Current image IDs: {ids}", font=('Arial', 10)).pack(pady=10)
    tk.Label(dialog, text="Enter ID:").pack(pady=5)

    id_var = tk.StringVar()
    entry = tk.Entry(dialog, textvariable=id_var, width=10)
    entry.pack(pady=5)

    color_row, color_var = _color_selector(dialog)
    color_row.pack(pady=10)

    def do_edit():
        try:
            edit_id = int(id_var.get())
        except ValueError:
            messagebox.showwarning("Warning", "Please enter a numeric ID")
            return

        for mark in current_marks:
            if mark['id'] == edit_id:
                save_state()
                mark["color"] = _color_key(color_var.get())
                mark_dirty()
                redraw_canvas()
                status_label.config(text=f"Updated colour for ID {edit_id}")
                dialog.destroy()
                return

        messagebox.showwarning("Warning", f"ID {edit_id} not found")

    tk.Button(dialog, text="Save", command=do_edit, bg="lightgreen", width=10).pack(pady=5)


def set_depth():
    global depth_dict
    if not image_paths:
        messagebox.showinfo("Info", "Please load images first")
        return

    current_path = image_paths[current_idx]
    current_depth = depth_dict.get(current_path, "")

    depth = simpledialog.askstring("Set Depth", f"Current depth: {current_depth}\nEnter depth information:", initialvalue=current_depth)
    if depth is None:
        return
    depth = depth.strip()
    if depth == current_depth:
        return
    save_state()
    depth_dict[current_path] = depth
    _retag_regions_depth(current_path, depth)
    mark_dirty()
    redraw_canvas()
    status_label.config(text=f"Depth set to: {depth}")

def clear_current_marks():
    if not image_paths:
        return
    if messagebox.askyesno("Confirm", "Clear all annotations on current image?"):
        save_state()
        current_path = image_paths[current_idx]
        if current_path in marks_dict:
            del marks_dict[current_path]
        for key in [k for k in subtype_dict if k[0] == current_path]:
            del subtype_dict[key]
        _drop_regions_for(current_path)
        mark_dirty()
        redraw_canvas()
        status_label.config(text="Cleared all annotations on current image")

def _ensure_app_dir():
    try:
        os.makedirs(APP_DIR, exist_ok=True)
    except OSError:
        pass


def _setup_logging():
    _ensure_app_dir()
    try:
        logging.basicConfig(filename=LOG_PATH, level=logging.ERROR,
                            format="%(asctime)s %(levelname)s %(message)s",
                            encoding="utf-8")
    except Exception:
        pass


def _report_exception(exc, val, tb):
    """Log unhandled exceptions and show them even in --windowed builds."""
    try:
        logging.error("Unhandled exception", exc_info=(exc, val, tb))
    except Exception:
        pass
    try:
        messagebox.showerror("Program error",
                             f"{getattr(exc, '__name__', exc)}: {val}\n\nDetails were written to:\n{LOG_PATH}")
    except Exception:
        pass


def _update_title():
    if root is None:
        return
    if project_path:
        suffix = os.path.basename(project_path)
    elif dirty:
        suffix = "Unsaved project"
    else:
        suffix = ""
    try:
        root.title(APP_TITLE + (f" - {suffix}" if suffix else ""))
    except tk.TclError:
        pass


def _snapshot():
    return {
        "marks_dict": copy.deepcopy(marks_dict),
        "depth_dict": dict(depth_dict),
        "subtype_dict": dict(subtype_dict),
        "lasso_regions": copy.deepcopy(lasso_regions),
        "image_paths": list(image_paths),
        "current_idx": current_idx,
    }


def _restore_state(state):
    global marks_dict, depth_dict, subtype_dict, lasso_regions, image_paths, current_idx, current_img
    marks_dict = state["marks_dict"]
    depth_dict = state["depth_dict"]
    subtype_dict = state["subtype_dict"]
    lasso_regions = state["lasso_regions"]
    image_paths = state["image_paths"]
    current_idx = int(state.get("current_idx", 0) or 0)
    if image_paths:
        current_idx = max(0, min(current_idx, len(image_paths) - 1))
        load_image_by_index()
    else:
        current_idx = 0
        current_img = None
        _clear_render_cache()
        if canvas is not None:
            canvas.delete("all")
        if status_label is not None:
            status_label.config(text="No images loaded")
        if depth_label is not None:
            depth_label.config(text="Image 0/0 | Depth: Not set")


def mark_dirty():
    global dirty, autosave_job
    dirty = True
    _update_title()
    if root is None:
        return
    try:
        if autosave_job is not None:
            root.after_cancel(autosave_job)
        autosave_job = root.after(AUTOSAVE_DELAY_MS, autosave_project)
    except tk.TclError:
        autosave_job = None


def _project_to_dict():
    return {
        "app": "ENVer",
        "version": 1,
        "image_paths": list(image_paths),
        "current_idx": current_idx,
        "marks_dict": {p: [dict(m) for m in ms] for p, ms in marks_dict.items()},
        "depth_dict": dict(depth_dict),
        "subtype_dict": [[p, int(i), v] for (p, i), v in subtype_dict.items()],
        "lasso_regions": [
            {
                "depth": r.get("depth", ""),
                "img_path": r.get("img_path", ""),
                "bbox": list(r.get("bbox", (0, 0, 0, 0))),
                "polygon": [list(pt) for pt in r.get("polygon", [])] if r.get("polygon") else None,
                "judgment": r.get("judgment", ""),
                "type": r.get("type", ""),
                "subtype": r.get("subtype", ""),
                "test_subtype": r.get("test_subtype", ""),
                "color": r.get("color", ""),
            }
            for r in lasso_regions
        ],
    }


def _normalise_loaded_project(data):
    """Validate a loaded project dict and drop entries whose images are gone."""
    if not isinstance(data, dict) or "image_paths" not in data:
        raise ValueError("Not an ENVer project file")

    raw_paths = [p for p in data.get("image_paths", []) if isinstance(p, str)]
    existing = [p for p in raw_paths if os.path.exists(p)]
    missing = [p for p in raw_paths if p not in existing]
    if missing:
        messagebox.showwarning(
            "Missing images",
            f"{len(missing)} image(s) from this project no longer exist and will be skipped.\n\n"
            + "\n".join(os.path.basename(p) for p in missing[:8])
            + ("\n..." if len(missing) > 8 else ""))

    image_paths.clear()
    image_paths.extend(existing)
    marks_dict.clear()
    for p, ms in (data.get("marks_dict") or {}).items():
        if p not in existing or not isinstance(ms, list):
            continue
        clean = []
        for m in ms:
            try:
                item = {
                    "id": int(m["id"]),
                    "x": int(m["x"]),
                    "y": int(m["y"]),
                    "type": str(m["type"]),
                    "judgment": str(m["judgment"]),
                }
                if m.get("test_subtype"):
                    item["test_subtype"] = str(m["test_subtype"])
                if m.get("color"):
                    item["color"] = str(m["color"])
                clean.append(item)
            except (KeyError, TypeError, ValueError):
                continue
        if clean:
            marks_dict[p] = clean

    depth_dict.clear()
    for p, d in (data.get("depth_dict") or {}).items():
        if p in existing:
            depth_dict[p] = str(d)

    subtype_dict.clear()
    for item in data.get("subtype_dict") or []:
        try:
            p, i, v = item
            if p in existing:
                subtype_dict[(p, int(i))] = str(v)
        except (TypeError, ValueError):
            continue

    lasso_regions.clear()
    for r in data.get("lasso_regions") or []:
        try:
            p = r.get("img_path")
            if p not in existing:
                continue
            bbox = tuple(int(v) for v in r.get("bbox", (0, 0, 0, 0)))
            polygon = tuple((int(pt[0]), int(pt[1])) for pt in r.get("polygon") or []) or None
            lasso_regions.append({
                "depth": str(r.get("depth", "")),
                "img_path": p,
                "bbox": bbox,
                "polygon": polygon,
                "judgment": str(r.get("judgment", "")),
                "type": str(r.get("type", "")),
                "subtype": str(r.get("subtype", "")),
                "test_subtype": str(r.get("test_subtype", "")),
                "color": str(r.get("color", "")),
            })
        except (TypeError, ValueError):
            continue

    idx = int(data.get("current_idx", 0) or 0)
    return max(0, min(idx, len(image_paths) - 1)) if image_paths else 0


def _load_project_from_path(path, is_autosave=False):
    global current_idx, project_path, dirty
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        new_idx = _normalise_loaded_project(data)
    except Exception as exc:
        logging.exception("Failed to load project")
        messagebox.showerror("Load failed", f"Could not load project:\n{path}\n\n{exc}")
        return False

    current_idx = new_idx
    undo_stack.clear()
    redo_stack.clear()
    project_path = None if is_autosave else path
    dirty = bool(is_autosave)
    _update_title()

    if image_paths:
        load_image_by_index()
        if status_label is not None:
            status_label.config(text=f"Loaded {len(image_paths)} image(s) from {os.path.basename(path)}"
                                     + (" (auto-save)" if is_autosave else ""))
    else:
        global current_img
        current_img = None
        _clear_render_cache()
        if canvas is not None:
            canvas.delete("all")
        if status_label is not None:
            status_label.config(text="Project loaded, but no existing images")
        if depth_label is not None:
            depth_label.config(text="Image 0/0 | Depth: Not set")
    if not is_autosave:
        _delete_autosave()
    return True


def save_project():
    """Save the current project (asks for a path the first time)."""
    global project_path, dirty
    if not image_paths:
        messagebox.showinfo("Info", "No images loaded, nothing to save")
        return False
    path = project_path
    if not path:
        path = filedialog.asksaveasfilename(
            title="Save ENVer Project",
            defaultextension=".json",
            filetypes=[("ENVer Project", "*.json"), ("All files", "*.*")])
        if not path:
            return False
    try:
        _ensure_app_dir()
        data = _project_to_dict()
        tmp_path = path + ".tmp"
        with open(tmp_path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=1)
        os.replace(tmp_path, path)
    except Exception as exc:
        logging.exception("Failed to save project")
        messagebox.showerror("Save failed", f"Could not save project:\n{path}\n\n{exc}")
        return False
    project_path = path
    dirty = False
    _delete_autosave()
    _update_title()
    if status_label is not None:
        status_label.config(text=f"Project saved: {path}")
    return True


def load_project(select_path=None, is_autosave=False):
    """Open a project file, optionally asking the user to confirm discarding changes."""
    if not is_autosave and dirty:
        if not messagebox.askyesno("Confirm",
                                   "The current project has unsaved changes.\nOpen another project anyway?"):
            return False
    path = select_path
    if not path:
        path = filedialog.askopenfilename(
            title="Open ENVer Project",
            filetypes=[("ENVer Project", "*.json"), ("All files", "*.*")])
        if not path:
            return False
    return _load_project_from_path(path, is_autosave=is_autosave)


def _delete_autosave():
    try:
        if os.path.exists(AUTOSAVE_PATH):
            os.remove(AUTOSAVE_PATH)
    except OSError:
        pass


def autosave_project():
    """Write a lightweight recovery file without bothering the user."""
    global autosave_job
    autosave_job = None
    if not dirty or not image_paths:
        return
    try:
        _ensure_app_dir()
        data = _project_to_dict()
        tmp_path = AUTOSAVE_PATH + ".tmp"
        with open(tmp_path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=1)
        os.replace(tmp_path, AUTOSAVE_PATH)
    except Exception:
        logging.exception("Autosave failed")


def _maybe_restore_autosave():
    if not os.path.exists(AUTOSAVE_PATH):
        return
    try:
        if messagebox.askyesno(
                "Recover session",
                "An auto-saved session from a previous run was found.\nRestore it?"):
            load_project(AUTOSAVE_PATH, is_autosave=True)
        else:
            _delete_autosave()
    except Exception:
        logging.exception("Autosave restore failed")


def on_close():
    global dirty
    try:
        saved = False
        if dirty and image_paths:
            answer = messagebox.askyesnocancel(
                "Quit",
                "The current project has unsaved changes.\n\nSave before quitting?")
            if answer is None:
                return
            if answer:
                if not save_project():
                    return
                saved = True
        if not dirty or saved or not image_paths:
            _delete_autosave()
    except Exception:
        logging.exception("Error while closing")
    try:
        root.destroy()
    except tk.TclError:
        pass


def save_state():
    global undo_stack, redo_stack
    undo_stack.append(_snapshot())
    if len(undo_stack) > UNDO_LIMIT:
        undo_stack.pop(0)
    redo_stack.clear()
    mark_dirty()

def undo():
    global undo_stack, redo_stack, marks_dict, depth_dict, subtype_dict, lasso_regions, image_paths, current_idx
    if not undo_stack:
        messagebox.showinfo("Info", "No undoable actions")
        return
    redo_stack.append(_snapshot())
    last_state = undo_stack.pop()
    _restore_state(last_state)
    mark_dirty()
    if status_label is not None:
        status_label.config(text=f"Total {len(image_paths)} images | Undo")

def redo():
    global undo_stack, redo_stack, marks_dict, depth_dict, subtype_dict, lasso_regions, image_paths, current_idx
    if not redo_stack:
        messagebox.showinfo("Info", "No redoable actions")
        return
    undo_stack.append(_snapshot())
    next_state = redo_stack.pop()
    _restore_state(next_state)
    mark_dirty()
    if status_label is not None:
        status_label.config(text=f"Total {len(image_paths)} images | Redo")

def add_images():
    files = filedialog.askopenfilenames(title="Select Images", filetypes=[("Images", "*.jpg *.jpeg *.png *.bmp *.tif *.tiff")])
    if not files:
        return
    global image_paths, current_idx
    new_paths = [f for f in files if f not in image_paths]
    if not new_paths:
        messagebox.showinfo("Info", "All images already loaded")
        return
    save_state()
    old_len = len(image_paths)
    image_paths.extend(new_paths)
    current_idx = old_len
    _clear_render_cache()
    load_image_by_index()
    mark_dirty()
    status_label.config(text=f"Total {len(image_paths)} images")


def delete_current_image():
    global image_paths, current_idx, marks_dict, depth_dict, subtype_dict, current_img
    if not image_paths:
        messagebox.showinfo("Info", "No images to delete")
        return

    current_path = image_paths[current_idx]
    current_name = os.path.basename(current_path)

    if messagebox.askyesno("Confirm Delete", f"Delete image「{current_name}」and its annotation data?"):
        save_state()
        image_paths.pop(current_idx)
        marks_dict.pop(current_path, None)
        depth_dict.pop(current_path, None)
        for key in [k for k in subtype_dict if k[0] == current_path]:
            del subtype_dict[key]
        _drop_regions_for(current_path)
        _clear_render_cache()

        if not image_paths:
            current_idx = 0
            current_img = None
            canvas.delete("all")
            status_label.config(text="No images loaded")
            depth_label.config(text="Image 0/0 | Depth: Not set")
        else:
            if current_idx >= len(image_paths):
                current_idx = len(image_paths) - 1
            load_image_by_index()

        mark_dirty()
        status_label.config(text=f"Deleted image「{current_name}」, total {len(image_paths)} images")

def load_image_by_index():
    global current_img, canvas_scale, canvas_offset_x, canvas_offset_y
    if not image_paths or current_idx >= len(image_paths) or current_idx < 0:
        return
    path = image_paths[current_idx]
    try:
        with Image.open(path) as im:
            im.load()
            current_img = im.copy()
    except Exception as exc:
        logging.exception("Failed to open image: %s", path)
        current_img = None
        _clear_render_cache()
        if canvas is not None:
            canvas.delete("all")
        messagebox.showerror("Load failed", f"Could not open image:\n{path}\n\n{exc}")
        status_label.config(text=f"Failed to load: {os.path.basename(path)}")
        return

    canvas_w = canvas.winfo_width() if canvas.winfo_width() > 1 else 800
    canvas_h = canvas.winfo_height() if canvas.winfo_height() > 1 else 600
    img_w, img_h = current_img.size
    if img_w <= 0 or img_h <= 0:
        messagebox.showerror("Load failed", f"Invalid image size:\n{path}")
        return
    canvas_scale = min(canvas_w / img_w, canvas_h / img_h)
    canvas_offset_x = int((canvas_w - img_w * canvas_scale) // 2)
    canvas_offset_y = int((canvas_h - img_h * canvas_scale) // 2)
    _clear_render_cache()
    redraw_canvas()
    status_label.config(text=f"Current: {os.path.basename(path)}")
    _update_title()

def next_img():
    global current_idx
    if not image_paths:
        return
    if current_idx < len(image_paths)-1:
        current_idx += 1
        load_image_by_index()

def prev_img():
    global current_idx
    if not image_paths:
        return
    if current_idx > 0:
        current_idx -= 1
        load_image_by_index()

# ------------------ Export helpers ------------------
def _natural_key(value):
    """Natural sort key so that 2 comes before 10."""
    text = str(value)
    parts = []
    number = ""
    for ch in text:
        if ch.isdigit():
            number += ch
        else:
            if number:
                parts.append((0, int(number)))
                number = ""
            parts.append((1, ch.lower()))
    if number:
        parts.append((0, int(number)))
    return parts


def _depth_sort_key(value):
    return (_natural_key(value), str(value))


def _crop_region_image(region):
    """Return the full-resolution crop for a saved lasso region."""
    img_path = region.get("img_path", "")
    if not os.path.exists(img_path):
        raise FileNotFoundError(img_path)
    with Image.open(img_path) as im:
        im.load()
        img = im.copy()
    x1, y1, x2, y2 = region.get("bbox", (0, 0, img.width, img.height))
    x1 = max(0, min(x1, img.width))
    x2 = max(0, min(x2, img.width))
    y1 = max(0, min(y1, img.height))
    y2 = max(0, min(y2, img.height))
    left, right = min(x1, x2), max(x1, x2)
    top, bottom = min(y1, y2), max(y1, y2)
    if right > left and bottom > top:
        return img.crop((left, top, right, bottom))
    return img


def _add_picture(paragraph, pil_img, width_inches):
    """Embed an image losslessly without touching the source directory.

    Deliberately no resize: the Word report must keep the native pixel
    resolution of the original image (user requirement).
    """
    buf = BytesIO()
    pil_img.save(buf, format="PNG")
    buf.seek(0)
    paragraph.add_run().add_picture(buf, width=Inches(width_inches))


def _build_excel_detail(save_path, marks, depths, subtypes):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Annotation Data"
    ws.append(["Depth", "Item ID", "Type", "Judgment", "Subtype/Remark", "Test Subtype", "Colour"])

    rows = []
    for img_path, img_marks in marks.items():
        depth = depths.get(img_path, "")
        for mark in img_marks:
            sub_key = (img_path, mark["id"])
            rows.append((
                depth,
                int(mark["id"]),
                mark.get("type", ""),
                mark.get("judgment", ""),
                subtypes.get(sub_key, ""),
                mark.get("test_subtype", ""),
                _color_display(mark.get("color", "")) if mark.get("color") else "",
                img_path,
            ))
    rows.sort(key=lambda r: (_depth_sort_key(r[0]), r[7].lower(), r[1]))
    for row in rows:
        ws.append(row[:7])

    for col, width in zip("ABCDEFG", (14, 10, 14, 12, 32, 14, 14)):
        ws.column_dimensions[col].width = width
    ws.freeze_panes = "A2"
    wb.save(save_path)
    return len(rows)


def _build_excel_stats(save_path, marks, depths):
    stats = {}
    for img_path, img_marks in marks.items():
        if not img_marks:
            continue
        depth = depths.get(img_path, "Depth not set")
        bucket = stats.setdefault(depth, {})
        for mark in img_marks:
            type_name = mark.get("type", "")
            bucket[type_name] = bucket.get(type_name, 0) + 1

    all_types = sorted({t for bucket in stats.values() for t in bucket})

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Type Statistics"
    ws.cell(row=1, column=1, value="Depth")
    for col, type_name in enumerate(all_types, start=2):
        ws.cell(row=1, column=col, value=type_name)
    ws.cell(row=1, column=len(all_types) + 2, value="Total")

    row = 2
    for depth in sorted(stats, key=_depth_sort_key):
        type_counts = stats[depth]
        ws.cell(row=row, column=1, value=depth)
        total = 0
        for col, type_name in enumerate(all_types, start=2):
            count = type_counts.get(type_name, 0)
            ws.cell(row=row, column=col, value=count)
            total += count
        ws.cell(row=row, column=len(all_types) + 2, value=total)
        row += 1

    ws.freeze_panes = "A2"
    wb.save(save_path)
    return len(stats)


def _build_word_all(save_path, marks, depths, subtypes):
    doc = Document()
    doc.add_heading('Annotated Image Report (Full Images)', 0)

    depth_groups = {}
    for img_path, img_marks in marks.items():
        if not img_marks:
            continue
        depth = depths.get(img_path, "Depth not set")
        depth_groups.setdefault(depth, []).append((img_path, img_marks))

    exported_count = 0
    depth_names = sorted(depth_groups, key=_depth_sort_key)

    for depth_index, depth in enumerate(depth_names):
        doc.add_heading(f'Depth：{depth}', level=1)

        for img_path, img_marks in sorted(depth_groups[depth], key=lambda item: item[0].lower()):
            try:
                if not os.path.exists(img_path):
                    doc.add_paragraph(f"[Image does not exist: {img_path}]")
                    continue

                with Image.open(img_path) as im:
                    im.load()
                    img = im.copy()
                draw = ImageDraw.Draw(img)

                try:
                    font_large = ImageFont.truetype("arial.ttf", 48)
                    font_small = ImageFont.truetype("arial.ttf", 32)
                except Exception:
                    font_large = ImageFont.load_default()
                    font_small = font_large

                for mark in img_marks:
                    x, y = mark['x'], mark['y']
                    r = 35
                    draw.ellipse([x-r, y-r, x+r, y+r], outline='white', width=4, fill='red')
                    draw.text((x-25, y-25), str(mark['id']), fill='white', font=font_large)

                    display_text = f"{mark['type']}:{mark['judgment']}"
                    if mark.get("test_subtype"):
                        display_text += f" ({mark['test_subtype']})"
                    if mark.get("color"):
                        display_text += f" [{_color_display(mark['color'])}]"
                    draw.text((x+40, y-35), display_text, fill='yellow', font=font_small)

                    sub_key = (img_path, mark['id'])
                    if sub_key in subtypes and subtypes[sub_key]:
                        draw.text((x+40, y+10), subtypes[sub_key][:20], fill='cyan', font=font_small)

                try:
                    depth_font = ImageFont.truetype("arial.ttf", 60)
                except Exception:
                    depth_font = ImageFont.load_default()
                draw.text((20, 20), f"Depth: {depth}", fill='white', font=depth_font)

                info_table = doc.add_table(rows=1, cols=3)
                info_table.style = 'Table Grid'
                hdr_cells = info_table.rows[0].cells
                hdr_cells[0].text = 'Image Name'
                hdr_cells[1].text = 'Annotation Count'
                hdr_cells[2].text = 'ID Range'

                ids = sorted([m['id'] for m in img_marks])
                id_range = f"{ids[0]}-{ids[-1]}" if len(ids) > 1 else str(ids[0])

                row_cells = info_table.add_row().cells
                row_cells[0].text = os.path.basename(img_path)
                row_cells[1].text = str(len(img_marks))
                row_cells[2].text = id_range

                paragraph = doc.add_paragraph()
                paragraph.alignment = 1
                _add_picture(paragraph, img, 6)

                doc.add_paragraph()
                exported_count += 1
            except Exception as exc:
                logging.exception("Word export failed for %s", img_path)
                doc.add_paragraph(f"[Image processing failed: {os.path.basename(img_path)} - {str(exc)}]")

        if depth_index < len(depth_names) - 1:
            doc.add_page_break()

    if exported_count == 0:
        raise RuntimeError("No images were exported")
    doc.save(save_path)
    return exported_count


def add_region_to_document(doc, region, index):
    """Add a full-resolution lasso region to a Word document."""
    try:
        img_path = region.get('img_path', '')
        if not os.path.exists(img_path):
            doc.add_paragraph(f"[Error] Image does not exist: {img_path}")
            return False

        cropped = _crop_region_image(region)
        table = doc.add_table(rows=2, cols=2)
        table.style = 'Table Grid'
        table.cell(0, 0).text = 'ID'
        table.cell(0, 1).text = str(index)
        table.cell(1, 0).text = 'Type/Judgment'
        judgment_text = region.get('judgment', '')
        if judgment_text == 'Test':
            judgment_text += f"-{region.get('test_subtype', '')}"
        table.cell(1, 1).text = f"{region.get('type', '')} / {judgment_text}"

        if region.get('subtype'):
            row = table.add_row().cells
            row[0].text = 'Remark'
            row[1].text = region['subtype']

        if region.get('color'):
            row = table.add_row().cells
            row[0].text = _loc('Colour', '颜色')
            row[1].text = _color_display(region['color'])

        paragraph = doc.add_paragraph()
        _add_picture(paragraph, cropped, 3)
        doc.add_paragraph()
        return True
    except Exception as exc:
        doc.add_paragraph(f"[Export failed: {os.path.basename(region.get('img_path', ''))} - {str(exc)}]")
        logging.exception("Region export failed")
        return False


def add_region_to_cell(cell, region, index):
    """Add a full-resolution lasso region to a table cell. Returns success flag."""
    try:
        if not os.path.exists(region.get('img_path', '')):
            cell.text = f"[Image does not exist: {region.get('img_path', '')}]"
            return False

        cropped = _crop_region_image(region)
        cell.text = ""
        paragraph = cell.paragraphs[0]
        _add_picture(paragraph, cropped, 2.5)

        paragraph = cell.add_paragraph()
        judgment_text = region.get('judgment', '')
        if judgment_text == 'Test':
            judgment_text += f"-{region.get('test_subtype', '')}"
        paragraph.add_run(f"#{index}\n{region.get('type', '')}\n{judgment_text}").font.size = Pt(9)

        if region.get('subtype'):
            p = cell.add_paragraph()
            p.add_run(f"Remark: {region['subtype'][:20]}").font.size = Pt(8)

        if region.get('color'):
            p = cell.add_paragraph()
            p.add_run(f"{_loc('Colour', '颜色')}: {_color_display(region['color'])}").font.size = Pt(8)
        return True
    except Exception as exc:
        cell.text = f"[Error: {str(exc)[:40]}]"
        logging.exception("Region cell export failed")
        return False


def _build_word_judgment(save_path, regions, selected_judgment):
    doc = Document()
    if selected_judgment == 'Test':
        doc.add_heading('Test Type Lasso Region Export', 0)
    else:
        doc.add_heading(f'Lasso Region Export - {selected_judgment}', 0)

    depth_groups = {}
    for region in regions:
        if selected_judgment != 'All' and region.get('judgment') != selected_judgment:
            continue
        depth = region.get('depth', '')
        categories = depth_groups.setdefault(
            depth, {'MP': [], 'NoMP': [], 'Test_label': [], 'Test_unlabel': []})
        if region.get('judgment') == 'Test':
            if region.get('test_subtype') == 'label':
                categories['Test_label'].append(region)
            else:
                categories['Test_unlabel'].append(region)
        else:
            judgment = region.get('judgment')
            if judgment in categories:
                categories[judgment].append(region)

    exported_count = 0
    for depth in sorted(depth_groups, key=_depth_sort_key):
        categories = depth_groups[depth]
        doc.add_heading(f'Depth：{depth}', level=1)

        for judgment in ['MP', 'NoMP']:
            if categories.get(judgment) and selected_judgment in ['All', judgment]:
                doc.add_heading(f'{judgment}', level=2)
                for i, region in enumerate(categories[judgment]):
                    if add_region_to_document(doc, region, i + 1):
                        exported_count += 1

        if (selected_judgment in ('Test', 'All')) and (categories['Test_label'] or categories['Test_unlabel']):
            doc.add_heading('Test', level=2)
            table = doc.add_table(rows=1, cols=2)
            table.style = 'Table Grid'
            hdr_cells = table.rows[0].cells
            hdr_cells[0].text = 'Test-Label'
            hdr_cells[1].text = 'Test-Unlabel'

            labels = categories['Test_label']
            unlabels = categories['Test_unlabel']
            max_rows = max(len(labels), len(unlabels))
            for row_idx in range(max_rows):
                row_cells = table.add_row().cells
                if row_idx < len(labels):
                    if add_region_to_cell(row_cells[0], labels[row_idx], row_idx + 1):
                        exported_count += 1
                else:
                    row_cells[0].text = ""
                if row_idx < len(unlabels):
                    if add_region_to_cell(row_cells[1], unlabels[row_idx], row_idx + 1):
                        exported_count += 1
                else:
                    row_cells[1].text = ""
            doc.add_paragraph()

    if exported_count == 0:
        raise RuntimeError("No matching lasso regions were exported")
    doc.save(save_path)
    return exported_count


def _run_export(title, work, success_message):
    """Run a document build in a worker thread so the Tk window stays responsive."""
    if root is None:
        return
    progress = tk.Toplevel(root)
    progress.title(title)
    progress.geometry("380x120")
    progress.transient(root)
    _center_dialog(progress, 380, 120)
    progress.resizable(False, False)
    progress.protocol("WM_DELETE_WINDOW", lambda: None)
    tk.Label(progress, text=title, font=('Arial', 11)).pack(pady=14)
    bar = ttk.Progressbar(progress, mode='indeterminate', length=320)
    bar.pack(pady=4)
    bar.start(12)
    try:
        progress.grab_set()
    except tk.TclError:
        pass

    messages = queue.Queue()

    def runner():
        try:
            messages.put(("ok", work()))
        except Exception as exc:
            messages.put(("err", (exc, traceback.format_exc())))

    threading.Thread(target=runner, daemon=True).start()

    def poll():
        try:
            kind, payload = messages.get_nowait()
        except queue.Empty:
            try:
                if progress.winfo_exists():
                    root.after(100, poll)
            except tk.TclError:
                pass
            return
        try:
            bar.stop()
            progress.grab_release()
            progress.destroy()
        except tk.TclError:
            pass
        if kind == "ok":
            messagebox.showinfo("Success", success_message(payload))
        else:
            exc, tb = payload
            try:
                logging.error("Export failed: %s\n%s", exc, tb)
            except Exception:
                pass
            messagebox.showerror("Export failed", f"{exc}\n\nDetails were written to:\n{LOG_PATH}")

    root.after(100, poll)


def export_excel():
    if not any(marks_dict.values()):
        messagebox.showinfo("Info", "No annotation data to export")
        return
    save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel File", "*.xlsx")])
    if not save_path:
        return
    marks = copy.deepcopy(marks_dict)
    depths = dict(depth_dict)
    subtypes = dict(subtype_dict)
    _run_export(
        "Exporting Excel detail...",
        lambda: _build_excel_detail(save_path, marks, depths, subtypes),
        lambda count: f"Exported {count} rows to:\n{save_path}")


def export_excel_statistics():
    if not any(marks_dict.values()):
        messagebox.showinfo("Info", "No annotation data to export")
        return
    save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel File", "*.xlsx")])
    if not save_path:
        return
    marks = copy.deepcopy(marks_dict)
    depths = dict(depth_dict)
    _run_export(
        "Exporting Excel statistics...",
        lambda: _build_excel_stats(save_path, marks, depths),
        lambda count: f"Statistics exported ({count} depth rows) to:\n{save_path}")


def export_word():
    """Export Word document: complete full-resolution images per depth with annotations."""
    if not any(marks_dict.values()):
        messagebox.showinfo("Info", "No annotation data to export")
        return
    save_path = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word Document", "*.docx")])
    if not save_path:
        return
    marks = copy.deepcopy(marks_dict)
    depths = dict(depth_dict)
    subtypes = dict(subtype_dict)
    _run_export(
        "Exporting Word report (full-resolution images)...",
        lambda: _build_word_all(save_path, marks, depths, subtypes),
        lambda count: f"Exported {count} annotated images to:\n{save_path}")


def export_word_filter_by_judgment():
    """Export lasso regions to Word by judgment type - Test split into label/unlabel columns."""
    if not lasso_regions:
        messagebox.showinfo("Info", "No saved lasso regions. Please use rectangle/polygon lasso to select and save areas first")
        return

    dialog = tk.Toplevel(root)
    dialog.title("Select Judgment Type to Export")
    dialog.geometry("300x200")
    dialog.transient(root)
    dialog.grab_set()
    _center_dialog(dialog, 300, 200)

    tk.Label(dialog, text="Select judgment type to export:", font=('Arial', 11)).pack(pady=15)

    judgment_var = tk.StringVar()
    judgment_combo = ttk.Combobox(dialog, textvariable=judgment_var,
                                  values=['MP', 'NoMP', 'Test', 'All'],
                                  state='readonly', width=15)
    judgment_combo.pack(pady=10)
    judgment_combo.current(2)
    judgment_combo.focus()

    result = [None]

    def confirm():
        result[0] = judgment_var.get()
        dialog.destroy()

    dialog.bind('<Return>', lambda event: confirm())
    tk.Button(dialog, text="Export Word", command=confirm, bg="lightgreen", width=10).pack(pady=20)
    dialog.wait_window()

    if result[0] is None:
        return

    selected_judgment = result[0]
    matching = [r for r in lasso_regions
                if selected_judgment == 'All' or r.get('judgment') == selected_judgment]
    if not matching:
        messagebox.showinfo("Info", f"No lasso regions with judgment「{selected_judgment}」")
        return
    save_path = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word Document", "*.docx")])
    if not save_path:
        return
    regions = copy.deepcopy(lasso_regions)
    _run_export(
        "Exporting Word lasso regions (full resolution)...",
        lambda: _build_word_judgment(save_path, regions, selected_judgment),
        lambda count: f"Exported {count} lasso regions to:\n{save_path}")









def reset_view():
    """Reset view to fit window"""
    global canvas_scale, canvas_offset_x, canvas_offset_y
    if not current_img:
        return
    canvas_w = canvas.winfo_width()
    canvas_h = canvas.winfo_height()
    if canvas_w <= 1:
        canvas_w, canvas_h = 800, 600
    img_w, img_h = current_img.size
    canvas_scale = min(canvas_w/img_w, canvas_h/img_h)
    canvas_offset_x = (canvas_w - img_w * canvas_scale) // 2
    canvas_offset_y = (canvas_h - img_h * canvas_scale) // 2
    redraw_canvas()
    status_label.config(text="View reset")

# ------------------ GUI ------------------
def _schedule_redraw(event=None):
    """Debounce <Configure> so window resizing does not redraw on every pixel."""
    global _configure_job
    if root is None:
        return
    try:
        if _configure_job is not None:
            root.after_cancel(_configure_job)
        _configure_job = root.after(80, _run_scheduled_redraw)
    except tk.TclError:
        _configure_job = None


def _run_scheduled_redraw():
    global _configure_job
    _configure_job = None
    redraw_canvas()


def cancel_lasso(event=None):
    """Cancel an in-progress rectangle/polygon lasso."""
    _reset_lasso_state()
    redraw_canvas()


def main():
    global root, canvas, status_label, depth_label

    _setup_logging()
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass

    root = tk.Tk()
    root.report_callback_exception = _report_exception
    sys.excepthook = _report_exception

    screen_w = root.winfo_screenwidth()
    screen_h = root.winfo_screenheight()
    win_w = min(1000, max(760, int(screen_w * 0.9)))
    win_h = min(800, max(560, int(screen_h * 0.85)))
    root.geometry(f"{win_w}x{win_h}")
    root.minsize(760, 560)
    root.title(APP_TITLE)

    root.bind("<Control-z>", lambda e: undo())
    root.bind("<Control-Z>", lambda e: undo())
    root.bind("<Control-y>", lambda e: redo())
    root.bind("<Control-Y>", lambda e: redo())
    root.bind("<Control-r>", lambda e: reset_view())
    root.bind("<Control-s>", lambda e: save_project())
    root.bind("<Control-o>", lambda e: load_project())
    root.bind("<Control-Shift-S>", lambda e: save_marks_to_image())
    root.bind("<Escape>", cancel_lasso)

    toolbar1 = tk.Frame(root)
    toolbar1.pack(fill=tk.X, padx=5, pady=2)

    tk.Button(toolbar1, text="Add", command=add_images, bg="lightblue", width=7).pack(side=tk.LEFT, padx=2)
    tk.Button(toolbar1, text="Delete", command=delete_current_image, bg="#FF6666", width=7).pack(side=tk.LEFT, padx=2)
    tk.Button(toolbar1, text="Prev", command=prev_img, width=5).pack(side=tk.LEFT, padx=2)
    tk.Button(toolbar1, text="Next", command=next_img, width=5).pack(side=tk.LEFT, padx=2)
    tk.Button(toolbar1, text="Save", command=save_marks_to_image, bg="lightgreen", width=7).pack(side=tk.LEFT, padx=2)
    tk.Button(toolbar1, text="Clear", command=clear_current_marks, bg="pink", width=7).pack(side=tk.LEFT, padx=2)
    tk.Button(toolbar1, text="Undo", command=undo, bg="#CCCCCC", width=5).pack(side=tk.LEFT, padx=2)
    tk.Button(toolbar1, text="Redo", command=redo, bg="#CCCCCC", width=5).pack(side=tk.LEFT, padx=2)
    tk.Button(toolbar1, text="Reset View", command=reset_view, bg="#DDDDDD", width=9).pack(side=tk.LEFT, padx=2)
    tk.Button(toolbar1, text="Open Proj", command=load_project, bg="#E6E6FA", width=9).pack(side=tk.LEFT, padx=2)
    tk.Button(toolbar1, text="Save Proj", command=save_project, bg="#E6E6FA", width=9).pack(side=tk.LEFT, padx=2)

    toolbar2 = tk.Frame(root)
    toolbar2.pack(fill=tk.X, padx=5, pady=2)

    tk.Button(toolbar2, text="Annotation Mode", command=set_tool_mark, bg="#90EE90", width=15).pack(side=tk.LEFT, padx=2)
    tk.Button(toolbar2, text="Rectangle Lasso", command=set_tool_rect_lasso, bg="#FFD700", width=15).pack(side=tk.LEFT, padx=2)
    tk.Button(toolbar2, text="Polygon Lasso", command=set_tool_poly_lasso, bg="#FFA500", width=15).pack(side=tk.LEFT, padx=2)

    toolbar3 = tk.Frame(root)
    toolbar3.pack(fill=tk.X, padx=5, pady=2)

    tk.Button(toolbar3, text="Depth", command=set_depth, bg="lightyellow", width=8).pack(side=tk.LEFT, padx=2)
    tk.Button(toolbar3, text="Renumber", command=renumber_current_depth, bg="lightcyan", width=10).pack(side=tk.LEFT, padx=2)
    tk.Button(toolbar3, text="Continue Num", command=continue_numbering_current_depth, bg="#FFD700", width=12).pack(side=tk.LEFT, padx=2)
    tk.Button(toolbar3, text="Delete ID", command=delete_mark_by_id, bg="#FF8888", width=10).pack(side=tk.LEFT, padx=2)
    tk.Button(toolbar3, text="Edit Remark", command=edit_mark_subtype, bg="#DDA0DD", width=10).pack(side=tk.LEFT, padx=2)
    tk.Button(toolbar3, text="Edit Colour", command=edit_mark_color, bg="#FFDAB9", width=10).pack(side=tk.LEFT, padx=2)

    toolbar4 = tk.Frame(root)
    toolbar4.pack(fill=tk.X, padx=5, pady=2)

    tk.Button(toolbar4, text="Excel Detail", command=export_excel, bg="#90EE90", width=12).pack(side=tk.LEFT, padx=2)
    tk.Button(toolbar4, text="Excel Stats", command=export_excel_statistics, bg="#FFB6C1", width=12).pack(side=tk.LEFT, padx=2)
    tk.Button(toolbar4, text="Word Report(All)", command=export_word, bg="#87CEEB", width=15).pack(side=tk.LEFT, padx=2)
    tk.Button(toolbar4, text="Word Export Test", command=export_word_filter_by_judgment, bg="#98FB98", width=15).pack(side=tk.LEFT, padx=2)

    status_frame = tk.Frame(root)
    status_frame.pack(fill=tk.X, padx=5)

    status_label = tk.Label(status_frame, text="No images loaded", fg="gray")
    status_label.pack(side=tk.LEFT, padx=5)

    depth_label = tk.Label(status_frame, text="Image 0/0 | Depth: ")
    depth_label.pack(side=tk.RIGHT, padx=5)

    canvas = tk.Canvas(root, bg='gray')
    canvas.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

    canvas.bind("<Button-1>", on_canvas_click)
    canvas.bind("<Button-3>", on_canvas_click)
    canvas.bind("<MouseWheel>", zoom)
    canvas.bind("<Button-4>", zoom)
    canvas.bind("<Button-5>", zoom)
    canvas.bind("<Button-2>", start_drag)
    canvas.bind("<B2-Motion>", on_drag)
    canvas.bind("<ButtonRelease-2>", stop_drag)
    canvas.bind("<Configure>", _schedule_redraw)

    info = tk.Label(
        root,
        text="【Annotation Mode】Left click add | Ctrl+Left click delete | "
             "【Rectangle Lasso】Left drag select | 【Polygon Lasso】Left click add point Right click complete | "
             "Scroll zoom | Middle button drag | Ctrl+S save project | Esc cancel lasso",
        fg="darkgreen", font=('Arial', 9))
    info.pack(pady=5)

    root.protocol("WM_DELETE_WINDOW", on_close)

    load_image_by_index()
    root.after(300, _maybe_restore_autosave)
    root.mainloop()


if __name__ == "__main__":
    main()
