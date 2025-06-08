import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, Toplevel
from tkinter.ttk import Scrollbar, Treeview
from PIL import Image, ImageTk, ImageFilter, ImageDraw, UnidentifiedImageError
import os
import math
import json
import datetime
import cv2
import numpy as np
import traceback

# Assuming filters, modes, utils, models are in the same package/directory
from .filters import FilterHandler
from .modes import ModeHandler # This file was pre-existing
from .utils import AppUtils    # This file was pre-existing
from .models import MeasurementModel

class ImageAnalyzer:
    def __init__(self, root):
        self.root = root
        self.root.title("Image Analyzer")
        self.root.geometry("1400x900")

        # Initialize handlers and models
        self.measurement_model = MeasurementModel()
        self.filter_handler = FilterHandler(self) # FilterHandler needs app instance
        self.mode_handler = ModeHandler(self)     # ModeHandler needs app instance
        self.utils = AppUtils(self)               # AppUtils needs app instance

        self.ZOOM_BOX_SIZE = 180
        self.ZOOM_BOX_FACTOR = 4

        self.file_path = None
        self.image_files = []
        self.current_index = 0
        self.zoom_factor = 1.0
        self.img_original = None
        self.img_filtered = None # Will hold filtered image if any filter is applied
        self.photo = None
        self.zoom_box_photo = None
        
        self.keep_calibration_dots_fixed = False # Flag to keep calibration dots when changing images

        # --- Mode Flags (managed by ImageAnalyzer or ModeHandler) ---
        # These flags control UI interaction states.
        self.edge_detection_active = False # Legacy FIND_EDGES filter flag (consider removing if Canny is primary)
        self.global_canny_active = False   # This is now primarily managed by FilterHandler.is_global_canny_active
                                           # but ImageAnalyzer might need a mirrored flag for UI button state if not directly tied.
                                           # For now, let FilterHandler be the source of truth.
        self.edge_selection_mode = False  # ROI Selection mode for legacy FIND_EDGES
        self.canny_selection_mode = False # ROI Selection mode for Canny
        self.zoom_box_mode = False
        self.calibration_mode = False
        self.artery_mode = False
        self.angle_mode = False
        self.line_mode = False

        # --- Selection Rectangles (for display on canvas, original image coords for filters) ---
        self.selection_rect_display = None # For FIND_EDGES ROI on canvas
        self.selection_start_display = None
        self.selection_end_display = None

        self.canny_rect_display = None      # For Canny ROI on canvas
        self.canny_start_display = None
        self.canny_end_display = None
        
        self.zoom_box = None

        self.undo_stack = []
        self.redo_stack = []

        # --- StringVars for Labels ---
        self.path_text = tk.StringVar(value="Path: No image loaded")
        self.pixel_info = tk.StringVar(value="Pixel: ")
        self.measurement = tk.StringVar(value="Distance: ")
        self.name_var = tk.StringVar(value="")
        self.diameter_var = tk.StringVar(value="")
        
        # --- IntVars for Canny Thresholds (link to FilterHandler) ---
        self.canny_low = tk.IntVar(value=self.filter_handler.canny_low_thresh)
        self.canny_high = tk.IntVar(value=self.filter_handler.canny_high_thresh)
        
        # Link slider changes to update FilterHandler and re-apply filters
        self.canny_low.trace_add("write", lambda *args: self.filter_handler.update_canny_thresholds(self.canny_low.get(), self.canny_high.get()))
        self.canny_high.trace_add("write", lambda *args: self.filter_handler.update_canny_thresholds(self.canny_low.get(), self.canny_high.get()))

        self.measurement_table = None
        self.image_frame = None

        self.create_gui()
        self.bind_events()

    def create_gui(self):
        # This method remains largely the same as in the original ImageAnalyzer.py
        # It sets up the Tkinter widgets.
        self.main_frame = tk.Frame(self.root)
        self.main_frame.pack(fill=tk.BOTH, expand=True)

        self.top_frame = tk.Frame(self.main_frame)
        self.top_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        self.button_area = tk.Frame(self.top_frame, borderwidth=2, relief=tk.SOLID)
        self.button_area.pack(side=tk.LEFT, fill=tk.Y, padx=5, pady=5)
        self.button_area.pack_propagate(False) 
        self.button_area.config(width=210) 

        self.button_canvas = tk.Canvas(self.button_area, borderwidth=0, highlightthickness=0)
        self.button_scrollbar = Scrollbar(self.button_area, orient=tk.VERTICAL, command=self.button_canvas.yview)
        self.button_canvas.configure(yscrollcommand=self.button_scrollbar.set)
        self.button_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.button_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.button_frame = tk.Frame(self.button_canvas)
        self.button_canvas.create_window((0, 0), window=self.button_frame, anchor='nw', tags="button_frame")
        self.button_frame.bind('<Configure>', self._on_button_frame_configure)

        self.image_frame = tk.Frame(self.top_frame)
        self.image_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.image_canvas = tk.Canvas(self.image_frame, bg="gray")
        self.image_canvas.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        if self.image_frame:
            self.zoom_box = tk.Canvas(self.image_frame, width=self.ZOOM_BOX_SIZE, height=self.ZOOM_BOX_SIZE,
                                    bg="black", highlightthickness=1, highlightbackground="white")
        else:
             print("Error: image_frame not created before zoom_box initialization.")
        
        self.middle_section_frame = tk.Frame(self.main_frame)
        self.middle_section_frame.pack(side=tk.TOP, fill=tk.X, padx=5, pady=5)

        self.dot_coords_frame = tk.Frame(self.middle_section_frame)
        self.dot_coords_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(5, 2), pady=5)
        tk.Label(self.dot_coords_frame, text="Dot Coordinates:").pack(side=tk.TOP, anchor=tk.W)
        self.dot_coords_text = tk.Text(self.dot_coords_frame, height=10, width=40)
        self.dot_coords_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.dot_coords_text.config(state=tk.DISABLED)
        dot_scrollbar = Scrollbar(self.dot_coords_frame, orient=tk.VERTICAL, command=self.dot_coords_text.yview)
        dot_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.dot_coords_text.configure(yscrollcommand=dot_scrollbar.set)

        self.table_frame = tk.Frame(self.middle_section_frame)
        self.table_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(2, 5), pady=5)
        tk.Label(self.table_frame, text="Measurements Summary:").pack(side=tk.TOP, anchor=tk.W)
        self.measurement_table = Treeview(self.table_frame, columns=("Type", "Pixel Distance", "Real Distance (mm)", "Angle (deg)"), show="headings", height=10)
        self.measurement_table.heading("Type", text="Type")
        self.measurement_table.heading("Pixel Distance", text="Pixel Distance")
        self.measurement_table.heading("Real Distance (mm)", text="Real Distance (mm)")
        self.measurement_table.heading("Angle (deg)", text="Angle (deg)")
        self.measurement_table.column("Type", width=80, anchor=tk.W)
        self.measurement_table.column("Pixel Distance", width=150, anchor=tk.W)
        self.measurement_table.column("Real Distance (mm)", width=150, anchor=tk.W)
        self.measurement_table.column("Angle (deg)", width=100, anchor=tk.W)
        self.measurement_table.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        table_scrollbar = Scrollbar(self.table_frame, orient=tk.VERTICAL, command=self.measurement_table.yview)
        table_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.measurement_table.configure(yscrollcommand=table_scrollbar.set)

        self.status_frame = tk.Frame(self.main_frame, bg="black", bd=1, relief=tk.SUNKEN)
        self.status_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=5, pady=3)
        self.path_label = tk.Label(self.status_frame, textvariable=self.path_text, bg="black", fg="white", anchor=tk.W)
        self.path_label.pack(side=tk.LEFT, padx=5)
        self.pixel_label = tk.Label(self.status_frame, textvariable=self.pixel_info, bg="black", fg="white", anchor=tk.W)
        self.pixel_label.pack(side=tk.LEFT, padx=10)
        self.measurement_label = tk.Label(self.status_frame, textvariable=self.measurement, bg="black", fg="white", anchor=tk.W)
        self.measurement_label.pack(side=tk.LEFT, padx=10)

        self.create_buttons()

        self.button_canvas.bind("<Enter>", self._bind_mousewheel_button_area)
        self.button_canvas.bind("<Leave>", self._unbind_mousewheel_button_area)
        self.button_frame.bind("<Enter>", self._bind_mousewheel_button_area)
        self.button_frame.bind("<Leave>", self._unbind_mousewheel_button_area)

    def _on_button_frame_configure(self, event=None):
        if hasattr(self, 'button_canvas') and self.button_canvas and self.button_canvas.winfo_exists():
            self.button_canvas.configure(scrollregion=self.button_canvas.bbox('all'))

    def _bind_mousewheel_button_area(self, event):
        self.root.bind_all("<MouseWheel>", self._on_mousewheel)
        self.root.bind_all("<Button-4>", self._on_mousewheel) 
        self.root.bind_all("<Button-5>", self._on_mousewheel) 

    def _unbind_mousewheel_button_area(self, event):
        self.root.unbind_all("<MouseWheel>")
        self.root.unbind_all("<Button-4>")
        self.root.unbind_all("<Button-5>")

    def _on_mousewheel(self, event):
        bx, by, bw, bh = self.button_area.winfo_rootx(), self.button_area.winfo_rooty(), self.button_area.winfo_width(), self.button_area.winfo_height()
        if not (bx <= event.x_root < bx + bw and by <= event.y_root < by + bh):
             return 

        delta = 0
        if event.num == 4: delta = -1 
        elif event.num == 5: delta = 1  
        elif hasattr(event, 'delta'): 
             if event.delta > 0: delta = -1 
             elif event.delta < 0: delta = 1  

        if delta != 0 and hasattr(self, 'button_canvas') and self.button_canvas:
            self.button_canvas.yview_scroll(delta, "units")
            return "break"

    def create_buttons(self):
        # This method also remains largely the same, defining buttons and linking them
        # to methods in ImageAnalyzer, ModeHandler, FilterHandler, or AppUtils.
        self.buttons = {}
        pad_options = {'fill': tk.X, 'padx': 3, 'pady': 2}

        file_frame = tk.LabelFrame(self.button_frame, text="File", bd=2, relief=tk.GROOVE)
        file_frame.pack(fill=tk.X, padx=3, pady=3)
        self.buttons["Load Image"] = tk.Button(file_frame, text="Load Image", command=self.load_image)
        self.buttons["Load Image"].pack(**pad_options)
        self.buttons["Export Image"] = tk.Button(file_frame, text="Export Image", command=self.utils.export_annotated_image) # Uses AppUtils
        self.buttons["Export Image"].pack(**pad_options)

        artery_frame = tk.LabelFrame(self.button_frame, text="Dots Mode (Distance/Angle)", bd=2, relief=tk.GROOVE)
        artery_frame.pack(fill=tk.X, padx=3, pady=3)
        self.buttons["Artery Mode"] = tk.Button(artery_frame, text="Dots Mode", command=self.mode_handler.toggle_artery_mode)
        self.buttons["Artery Mode"].pack(**pad_options)
        self.buttons["Reset Artery"] = tk.Button(artery_frame, text="Reset Dots", command=self.reset_artery_mode) # Local method
        self.buttons["Reset Artery"].pack(**pad_options)
        self.buttons["Delete Last Pair"] = tk.Button(artery_frame, text="Delete Last Pair", command=self.delete_last_artery_pair) # Local method
        self.buttons["Delete Last Pair"].pack(**pad_options)

        calib_frame = tk.LabelFrame(self.button_frame, text="Calibration", bd=2, relief=tk.GROOVE)
        calib_frame.pack(fill=tk.X, padx=3, pady=3)
        self.buttons["Calibrate"] = tk.Button(calib_frame, text="Calibrate", command=self.mode_handler.toggle_calibration_mode)
        self.buttons["Calibrate"].pack(**pad_options)
        self.buttons["Reset Calibration"] = tk.Button(calib_frame, text="Reset Calibration", command=self.reset_calibration) # Local method
        self.buttons["Reset Calibration"].pack(**pad_options)
        self.buttons["Keep Dots Fixed"] = tk.Button(calib_frame, text="Keep Dots Fixed", command=self.toggle_keep_dots_fixed) # Local method
        self.buttons["Keep Dots Fixed"].pack(**pad_options)

        angle_frame = tk.LabelFrame(self.button_frame, text="Angle Measurement", bd=2, relief=tk.GROOVE)
        angle_frame.pack(fill=tk.X, padx=3, pady=3)
        self.buttons["Angle Mode"] = tk.Button(angle_frame, text="Angle Mode", command=self.mode_handler.toggle_angle_mode)
        self.buttons["Angle Mode"].pack(**pad_options)
        self.buttons["Reset Angle"] = tk.Button(angle_frame, text="Reset Angle", command=self.reset_angle_mode) # Local method
        self.buttons["Reset Angle"].pack(**pad_options)

        line_frame = tk.LabelFrame(self.button_frame, text="Line Mode (Parallel)", bd=2, relief=tk.GROOVE)
        line_frame.pack(fill=tk.X, padx=3, pady=3)
        self.buttons["Line Mode"] = tk.Button(line_frame, text="Line Mode", command=self.mode_handler.toggle_line_mode)
        self.buttons["Line Mode"].pack(**pad_options)
        self.buttons["Reset Lines"] = tk.Button(line_frame, text="Reset Lines", command=self.reset_lines) # Local method
        self.buttons["Reset Lines"].pack(**pad_options)
        self.buttons["Show Line Measurements"] = tk.Button(line_frame, text="Show Line Measurements", command=self.show_line_measurements) # Local method
        self.buttons["Show Line Measurements"].pack(**pad_options)

        filter_frame = tk.LabelFrame(self.button_frame, text="Filters", bd=2, relief=tk.GROOVE)
        filter_frame.pack(fill=tk.X, padx=3, pady=3)
        
        # Global Canny Button - Toggles FilterHandler's global Canny state
        self.buttons["Global Canny"] = tk.Button(filter_frame, text="Global Canny Filter", command=self.toggle_global_canny_filter) # Local method to interact with FilterHandler
        self.buttons["Global Canny"].pack(**pad_options)

        # Canny ROI Button - Toggles Canny ROI selection mode
        self.buttons["Canny Selection"] = tk.Button(filter_frame, text="Canny ROI Selection", command=self.toggle_canny_roi_selection_mode) # Local method
        self.buttons["Canny Selection"].pack(**pad_options)

        canny_params_frame = tk.Frame(filter_frame)
        canny_params_frame.pack(fill=tk.X, padx=3, pady=3)
        low_frame = tk.Frame(canny_params_frame)
        low_frame.pack(fill=tk.X, pady=1)
        tk.Label(low_frame, text="Low:", width=4).pack(side=tk.LEFT, padx=(0,2))
        self.canny_low_slider = tk.Scale(low_frame, from_=0, to=255, orient=tk.HORIZONTAL, variable=self.canny_low, length=140, showvalue=True)
        self.canny_low_slider.pack(side=tk.LEFT, fill=tk.X, expand=True)
        high_frame = tk.Frame(canny_params_frame)
        high_frame.pack(fill=tk.X, pady=1)
        tk.Label(high_frame, text="High:", width=4).pack(side=tk.LEFT, padx=(0,2))
        self.canny_high_slider = tk.Scale(high_frame, from_=0, to=255, orient=tk.HORIZONTAL, variable=self.canny_high, length=140, showvalue=True)
        self.canny_high_slider.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.buttons["Reset Filters"] = tk.Button(filter_frame, text="Reset Filters", command=self.reset_all_filters) # Local method to call FilterHandler
        self.buttons["Reset Filters"].pack(**pad_options)

        zoom_frame = tk.LabelFrame(self.button_frame, text="Zoom", bd=2, relief=tk.GROOVE)
        zoom_frame.pack(fill=tk.X, padx=3, pady=3)
        self.buttons["Zoom In Box"] = tk.Button(zoom_frame, text="Zoom In Box", command=self.toggle_zoom_box)
        self.buttons["Zoom In Box"].pack(**pad_options)

        history_frame = tk.LabelFrame(self.button_frame, text="History", bd=2, relief=tk.GROOVE)
        history_frame.pack(fill=tk.X, padx=3, pady=3)
        self.buttons["Undo"] = tk.Button(history_frame, text="Undo (Ctrl+Z)", command=self.undo)
        self.buttons["Undo"].pack(**pad_options)
        self.buttons["Redo"] = tk.Button(history_frame, text="Redo (Ctrl+Y)", command=self.redo)
        self.buttons["Redo"].pack(**pad_options)

        meas_frame = tk.LabelFrame(self.button_frame, text="Measurements", bd=2, relief=tk.GROOVE)
        meas_frame.pack(fill=tk.X, padx=3, pady=3)
        name_frame = tk.Frame(meas_frame)
        name_frame.pack(fill=tk.X, padx=3, pady=1)
        tk.Label(name_frame, text="Name:").pack(side=tk.LEFT, padx=(0,3))
        self.name_entry = tk.Entry(name_frame, textvariable=self.name_var)
        self.name_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        diameter_frame = tk.Frame(meas_frame)
        diameter_frame.pack(fill=tk.X, padx=3, pady=1)
        tk.Label(diameter_frame, text="Real Ø (mm):").pack(side=tk.LEFT, padx=(0,3))
        self.diameter_entry = tk.Entry(diameter_frame, textvariable=self.diameter_var)
        self.diameter_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.save_button = tk.Button(meas_frame, text="Save Measurements", command=self.utils.save_measurements_to_json) # Uses AppUtils
        self.save_button.pack(**pad_options)

        self.root.update_idletasks()
        self._on_button_frame_configure()

    def bind_events(self):
        # This method remains largely the same.
        self.image_canvas.bind("<Button-1>", self.on_press)
        self.image_canvas.bind("<B1-Motion>", self.on_motion)
        self.image_canvas.bind("<ButtonRelease-1>", self.on_release)

        if self.root.tk.call('tk', 'windowingsystem') == 'aqua':
            self.image_canvas.bind("<MouseWheel>", self.zoom_mouse)
            self.image_canvas.bind("<Button-4>", self.zoom_mouse) 
            self.image_canvas.bind("<Button-5>", self.zoom_mouse) 
        elif self.root.tk.call('tk', 'windowingsystem') == 'x11':
             self.image_canvas.bind("<Button-4>", self.zoom_mouse)
             self.image_canvas.bind("<Button-5>", self.zoom_mouse)
        else: 
            self.image_canvas.bind("<MouseWheel>", self.zoom_mouse)
        
        self.root.bind("<plus>", self.zoom_in_center)
        self.root.bind("<minus>", self.zoom_out_center)
        self.root.bind("<KeyPress-Right>", self.next_image)
        self.root.bind("<KeyPress-Left>", self.prev_image)
        self.root.bind("<Control-z>", self.undo)
        self.root.bind("<Control-y>", self.redo)
        self.image_canvas.bind("<Motion>", self.update_zoom_box_and_pixel)
        self.image_canvas.bind("<Configure>", self.on_canvas_resize)

    def on_canvas_resize(self, event=None):
         self.display_image()

    def reset_image_state(self, reset_zoom=True):
        if reset_zoom:
            self.zoom_factor = 1.0
        
        self.mode_handler._reset_all_modes() # Resets UI mode flags in ImageAnalyzer via ModeHandler
        self.img_filtered = None # Clear any active filter effect

        # Reset measurement data via MeasurementModel
        self.measurement_model.reset_all_data(keep_calibration=self.keep_calibration_dots_fixed)

        # Reset filter states via FilterHandler (includes Canny ROI, global Canny state)
        # This will also reset self.filter_handler.canny_roi_start_orig etc.
        self.filter_handler.reset_filter_states() 
        # Ensure UI elements like Canny sliders are reset by FilterHandler or here
        self.canny_low.set(self.filter_handler.canny_low_thresh)
        self.canny_high.set(self.filter_handler.canny_high_thresh)


        # Reset local display ROI rectangle coordinates
        self.selection_start_display = self.selection_end_display = None
        self.canny_start_display = self.canny_end_display = None
        if hasattr(self, 'image_canvas') and self.image_canvas.winfo_exists():
            self.image_canvas.delete("selection_rect") # Legacy
            self.image_canvas.delete("canny_rect")
        self.selection_rect_display = None
        self.canny_rect_display = None
        
        self.photo = None
        self.zoom_box_photo = None

        if "Zoom In Box" in self.buttons: self.buttons["Zoom In Box"].config(relief=tk.RAISED)
        # Global Canny button state is handled by toggle_global_canny_filter / reset_all_filters

        self.utils.update_dot_coords_display() # AppUtils will need to use measurement_model
        self.utils.update_tables()             # AppUtils will need to use measurement_model
        self.measurement.set("Status: Ready" if self.img_original else "Status: Load Image")
        self.pixel_info.set("Mode: None | Pixel: | Zoom: OFF")

        if self.zoom_box_mode and self.zoom_box:
            self.update_zoom_box_content(None)

    def load_image(self):
        # This method remains largely the same, but calls self.reset_image_state()
        # which now correctly uses MeasurementModel and FilterHandler for resets.
        file_path = filedialog.askopenfilename(
            title="Select Image File",
            filetypes=[("Image Files", "*.png;*.jpg;*.jpeg;*.bmp;*.gif;*.tif;*.tiff"), ("All Files", "*.*")]
        )
        if not file_path: return

        try:
            img_test = Image.open(file_path)
            img_test.verify() 
            img_test.close() 

            self.file_path = file_path
            folder = os.path.dirname(file_path)
            try:
                self.image_files = sorted([
                    f for f in os.listdir(folder)
                    if f.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp', '.gif', '.tif', '.tiff'))
                       and os.path.isfile(os.path.join(folder, f))
                ])
                self.current_index = self.image_files.index(os.path.basename(file_path))
            except ValueError: 
                self.image_files = [os.path.basename(file_path)]
                self.current_index = 0
            except OSError:
                 self.image_files = [os.path.basename(file_path)]
                 self.current_index = 0
                 messagebox.showwarning("Folder Access", "Could not read image folder content. Navigation disabled.")

            self.path_text.set(f"Path: {os.path.basename(file_path)}")
            self.img_original = Image.open(file_path).convert("RGBA")
            self.reset_image_state(reset_zoom=True) 
            self.display_image()

            if self.zoom_box_mode and self.zoom_box:
                 try:
                     self.zoom_box.place(in_=self.image_frame, anchor='se', relx=1.0, rely=1.0, x=-10, y=-10)
                     self.update_zoom_box_content(None) 
                 except tk.TclError as e:
                      print(f"Error re-placing zoom box: {e}")
                      self.zoom_box_mode = False 
                      self.buttons["Zoom In Box"].config(relief=tk.RAISED)
        except FileNotFoundError:
            messagebox.showerror("Error", f"File not found:\n{file_path}")
            self.file_path = None
        except UnidentifiedImageError:
             messagebox.showerror("Error", f"Cannot identify image file:\n{file_path}\nMay be corrupted or unsupported format.")
             self.file_path = None
        except Exception as e:
            messagebox.showerror("Error Loading Image", f"An unexpected error occurred:\n{e}")
            print(traceback.format_exc())
            self.file_path = None
            self.img_original = None
            self.reset_image_state(reset_zoom=True)
            self.display_image()

    def change_image(self, direction):
        # This method also remains largely the same.
        if not self.file_path or not self.image_files or len(self.image_files) < 2:
             self.measurement.set("Status: No other images in folder.")
             return

        original_index = self.current_index
        num_files = len(self.image_files)
        attempt = 0

        while attempt < num_files:
            attempt += 1
            if direction == "next": next_idx = (original_index + attempt) % num_files
            elif direction == "previous": next_idx = (original_index - attempt + num_files) % num_files
            else: return 

            if next_idx == original_index and attempt > 0: break 

            new_file_name = self.image_files[next_idx]
            new_file_path = os.path.join(os.path.dirname(self.file_path), new_file_name)

            try:
                img_test = Image.open(new_file_path)
                img_test.verify()
                img_test.close()

                self.current_index = next_idx 
                self.file_path = new_file_path
                self.path_text.set(f"Path: {new_file_name}")
                self.img_original = Image.open(self.file_path).convert("RGBA")
                self.reset_image_state(reset_zoom=False) # Keeps current zoom
                self.display_image()
                if self.zoom_box_mode and self.zoom_box:
                     self.update_zoom_box_content(None)
                return 
            except (FileNotFoundError, UnidentifiedImageError, OSError) as e:
                print(f"Skipping file '{new_file_name}': {e}")
            except Exception as e:
                messagebox.showerror("Error Changing Image", f"An unexpected error occurred loading '{new_file_name}':\n{e}")
                print(traceback.format_exc())
                return
        messagebox.showinfo("Image Navigation", "No other valid images found in the folder.")

    def next_image(self, event=None): self.change_image("next")
    def prev_image(self, event=None): self.change_image("previous")

    def display_image(self):
        # Uses self.measurement_model for drawing dots/lines.
        if not self.root or not self.root.winfo_exists() or not self.image_canvas or not self.image_canvas.winfo_exists():
            return
        if not self.img_original:
            self.image_canvas.delete("all")
            self.image_canvas.config(scrollregion=(0, 0, 1, 1))
            try:
                canvas_width = self.image_canvas.winfo_width()
                canvas_height = self.image_canvas.winfo_height()
                if canvas_width > 1 and canvas_height > 1:
                    self.image_canvas.create_text(canvas_width / 2, canvas_height / 2, text="No Image Loaded", fill="white", font=("Arial", 16))
            except tk.TclError: pass
            return

        img_display_base = self.img_filtered if self.img_filtered is not None else self.img_original
        width, height = img_display_base.size
        new_width = int(width * self.zoom_factor)
        new_height = int(height * self.zoom_factor)
        if new_width <= 0 or new_height <= 0: return

        try:
            resized_img = img_display_base.resize((new_width, new_height), Image.Resampling.LANCZOS)
            if self.root and self.root.winfo_exists(): self.photo = ImageTk.PhotoImage(resized_img)
            else: return
        except Exception: # Fallback to NEAREST
            try:
                resized_img = img_display_base.resize((new_width, new_height), Image.Resampling.NEAREST)
                if self.root and self.root.winfo_exists(): self.photo = ImageTk.PhotoImage(resized_img)
                else: return
            except Exception: return

        self.image_canvas.delete("all")
        if self.photo:
            self.image_canvas.create_image(0, 0, anchor=tk.NW, image=self.photo)
            self.image_canvas.config(scrollregion=(0, 0, new_width, new_height))
        else: return

        def scale_pt(pt): return (pt[0] * self.zoom_factor, pt[1] * self.zoom_factor)
        dot_radius, line_width = 3, 2

        for dot in self.measurement_model.calibration_dots:
            sx, sy = scale_pt(dot)
            self.image_canvas.create_oval(sx - dot_radius, sy - dot_radius, sx + dot_radius, sy + dot_radius, fill="cyan", outline="black")
        
        for i in range(len(self.measurement_model.artery_dots)):
            sx, sy = scale_pt(self.measurement_model.artery_dots[i])
            self.image_canvas.create_oval(sx - dot_radius, sy - dot_radius, sx + dot_radius, sy + dot_radius, fill="yellow", outline="black")
            if i % 2 == 1:
                sx_prev, sy_prev = scale_pt(self.measurement_model.artery_dots[i-1])
                self.image_canvas.create_line(sx_prev, sy_prev, sx, sy, fill="yellow", width=line_width)

        for i in range(len(self.measurement_model.line_points)):
            sx, sy = scale_pt(self.measurement_model.line_points[i])
            self.image_canvas.create_oval(sx - dot_radius, sy - dot_radius, sx + dot_radius, sy + dot_radius, fill="magenta", outline="black")
            if i % 2 == 1:
                sx_prev, sy_prev = scale_pt(self.measurement_model.line_points[i-1])
                self.image_canvas.create_line(sx_prev, sy_prev, sx, sy, fill="magenta", width=line_width)

        if self.measurement_model.angle_points:
            for i in range(len(self.measurement_model.angle_points)):
                sp = scale_pt(self.measurement_model.angle_points[i])
                self.image_canvas.create_oval(sp[0]-dot_radius, sp[1]-dot_radius, sp[0]+dot_radius, sp[1]+dot_radius, fill="lime green", outline="black")
            if len(self.measurement_model.angle_points) > 1:
                scaled_points = [scale_pt(p) for p in self.measurement_model.angle_points]
                self.image_canvas.create_line(scaled_points, fill="lime green", width=line_width, dash=(4, 2))
        
        tick_radius = 2
        for pt1, pt2 in self.measurement_model.line_measurement_visualization_points:
            sx1, sy1 = scale_pt(pt1); sx2, sy2 = scale_pt(pt2)
            self.image_canvas.create_oval(sx1 - tick_radius, sy1 - tick_radius, sx1 + tick_radius, sy1 + tick_radius, fill="red", outline="red")
            self.image_canvas.create_line(sx1, sy1, sx2, sy2, fill="red", dash=(2, 2))

        # Draw Canny ROI rectangle (display coordinates)
        if self.canny_start_display and self.canny_end_display and not self.canny_selection_mode:
            self.image_canvas.delete("canny_rect_display_tag") 
            self.image_canvas.create_rectangle(
                self.canny_start_display[0], self.canny_start_display[1],
                self.canny_end_display[0], self.canny_end_display[1],
                outline="blue", dash=(4, 4), width=1, tags="canny_rect_display_tag"
            )
        elif not self.canny_selection_mode: # Clear if not selecting and no finalized ROI
             self.image_canvas.delete("canny_rect_display_tag")
        
        # Legacy FIND_EDGES ROI rectangle
        if self.selection_start_display and self.selection_end_display and not self.edge_selection_mode:
             self.image_canvas.delete("selection_rect_display_tag")
             self.image_canvas.create_rectangle(
                 self.selection_start_display[0], self.selection_start_display[1],
                 self.selection_end_display[0], self.selection_end_display[1],
                 outline="red", dash=(4, 4), width=1, tags="selection_rect_display_tag"
             )
        elif not self.edge_selection_mode:
             self.image_canvas.delete("selection_rect_display_tag")


    def update_zoom_box_and_pixel(self, event=None):
        # This method remains largely the same.
        if not self.img_original or not event or not self.image_canvas or not self.image_canvas.winfo_exists(): return
        try:
            canvas_x = self.image_canvas.canvasx(event.x); canvas_y = self.image_canvas.canvasy(event.y)
            orig_x = int(canvas_x / self.zoom_factor); orig_y = int(canvas_y / self.zoom_factor)
            pixel_str_part = "Pixel:"
            if 0 <= orig_x < self.img_original.width and 0 <= orig_y < self.img_original.height:
                try:
                    pixel_value = self.img_original.getpixel((orig_x, orig_y))
                    if isinstance(pixel_value, tuple): 
                        pixel_str = f"RGB:({pixel_value[0]},{pixel_value[1]},{pixel_value[2]})"
                        if len(pixel_value) == 4: pixel_str += f" A:{pixel_value[3]}"
                    else: pixel_str = f"Gray:{pixel_value}"
                    pixel_str_part = f"Pixel @ ({orig_x}, {orig_y}): {pixel_str}"
                except Exception: pixel_str_part = f"Pixel @ ({orig_x}, {orig_y}): Error"
            else: pixel_str_part = "Pixel: Outside Image"
            try:
                current_info = self.pixel_info.get(); parts = current_info.split('|')
                mode_info = parts[0].strip() if parts else "Mode: Unknown"
                zoom_info = parts[-1].strip() if parts else "Zoom: ?"
                self.pixel_info.set(f"{mode_info} | {pixel_str_part} | {zoom_info}")
            except Exception: self.pixel_info.set(f"Mode: Unknown | {pixel_str_part} | Zoom: {'ON' if self.zoom_box_mode else 'OFF'}")
            if self.zoom_box_mode and self.zoom_box and self.zoom_box.winfo_exists(): self.update_zoom_box_content(event)
        except tk.TclError: pass

    def update_zoom_box_content(self, event=None):
        # Uses self.measurement_model for drawing dots/lines in zoom box.
        if not self.root or not self.root.winfo_exists() or not self.zoom_box or not self.zoom_box.winfo_exists(): return
        if not self.zoom_box_mode or not self.img_original: return

        try:
            if event:
                canvas_x = self.image_canvas.canvasx(event.x); canvas_y = self.image_canvas.canvasy(event.y)
                orig_x = int(canvas_x / self.zoom_factor); orig_y = int(canvas_y / self.zoom_factor)
            else:
                canvas_width = self.image_canvas.winfo_width(); canvas_height = self.image_canvas.winfo_height()
                scroll_x = self.image_canvas.canvasx(0); scroll_y = self.image_canvas.canvasy(0)
                center_canvas_x = scroll_x + canvas_width / 2; center_canvas_y = scroll_y + canvas_height / 2
                orig_x = int(center_canvas_x / self.zoom_factor); orig_y = int(center_canvas_y / self.zoom_factor)

            crop_width_orig = self.ZOOM_BOX_SIZE / self.ZOOM_BOX_FACTOR; crop_height_orig = self.ZOOM_BOX_SIZE / self.ZOOM_BOX_FACTOR
            left = int(orig_x - crop_width_orig / 2); top = int(orig_y - crop_height_orig / 2)
            right = int(left + crop_width_orig); bottom = int(top + crop_height_orig)
            left = max(0, left); top = max(0, top)
            right = min(self.img_original.width, right); bottom = min(self.img_original.height, bottom)

            if right <= left or bottom <= top:
                self.zoom_box.delete("all"); self.zoom_box.create_text(self.ZOOM_BOX_SIZE / 2, self.ZOOM_BOX_SIZE / 2, text="Invalid Area", fill="red")
                return

            img_source = self.img_filtered if self.img_filtered is not None else self.img_original
            cropped_image = img_source.crop((left, top, right, bottom))
            zoomed = cropped_image.resize((self.ZOOM_BOX_SIZE, self.ZOOM_BOX_SIZE), Image.Resampling.NEAREST)
            if self.root and self.root.winfo_exists(): self.zoom_box_photo = ImageTk.PhotoImage(zoomed)
            else: return

            self.zoom_box.delete("all")
            self.zoom_box.create_image(0, 0, anchor=tk.NW, image=self.zoom_box_photo)
            
            dot_radius_zoom, line_width_zoom = 2, 1
            def scale_to_zoom(pt_orig):
                if left <= pt_orig[0] < right and top <= pt_orig[1] < bottom:
                    return (pt_orig[0] - left) * self.ZOOM_BOX_FACTOR, (pt_orig[1] - top) * self.ZOOM_BOX_FACTOR
                return None

            for dot_orig in self.measurement_model.calibration_dots:
                sp = scale_to_zoom(dot_orig)
                if sp: self.zoom_box.create_oval(sp[0]-dot_radius_zoom, sp[1]-dot_radius_zoom, sp[0]+dot_radius_zoom, sp[1]+dot_radius_zoom, fill="cyan", outline="black")
            
            for i in range(len(self.measurement_model.artery_dots)):
                sp = scale_to_zoom(self.measurement_model.artery_dots[i])
                if sp:
                    self.zoom_box.create_oval(sp[0]-dot_radius_zoom, sp[1]-dot_radius_zoom, sp[0]+dot_radius_zoom, sp[1]+dot_radius_zoom, fill="yellow", outline="black")
                    if i % 2 == 1:
                        sp_prev = scale_to_zoom(self.measurement_model.artery_dots[i-1])
                        if sp_prev: self.zoom_box.create_line(sp_prev[0], sp_prev[1], sp[0], sp[1], fill="yellow", width=line_width_zoom)
            
            # ... (similar updates for line_points, angle_points using self.measurement_model) ...

            center = self.ZOOM_BOX_SIZE / 2; offset = 5; lw = 1
            self.zoom_box.create_oval(center-offset, center-offset, center+offset, center+offset, outline="red", width=lw)
            self.zoom_box.create_line(center, center-offset*0.6, center, center+offset*0.6, fill="red", width=lw)
            self.zoom_box.create_line(center-offset*0.6, center, center+offset*0.6, center, fill="red", width=lw)
        except Exception as e:
            print(f"Error updating zoom box: {e}, {traceback.format_exc()}")
            try: self.zoom_box.delete("all"); self.zoom_box.create_text(self.ZOOM_BOX_SIZE / 2, self.ZOOM_BOX_SIZE / 2, text="Error", fill="red")
            except tk.TclError: pass


    def zoom(self, factor, event=None):
        # This method remains largely the same.
        if not self.img_original: return
        old_zoom = self.zoom_factor; new_zoom = max(0.05, min(old_zoom * factor, 50.0))
        if abs(new_zoom - old_zoom) < 0.001: return
        
        mouse_x, mouse_y, target_canvas_x, target_canvas_y, valid_coords = 0,0,0,0,False
        if self.image_canvas and self.image_canvas.winfo_exists():
            try:
                canvas_width = self.image_canvas.winfo_width(); canvas_height = self.image_canvas.winfo_height()
                if event:
                     mouse_x = self.image_canvas.canvasx(event.x); mouse_y = self.image_canvas.canvasy(event.y)
                     target_canvas_x = event.x; target_canvas_y = event.y; valid_coords = True
                elif canvas_width > 0 and canvas_height > 0: 
                     mouse_x = self.image_canvas.canvasx(canvas_width / 2); mouse_y = self.image_canvas.canvasy(canvas_height / 2)
                     target_canvas_x = canvas_width / 2; target_canvas_y = canvas_height / 2; valid_coords = True
            except tk.TclError: return
        else: return
        if not valid_coords: return

        img_coord_x = mouse_x / old_zoom; img_coord_y = mouse_y / old_zoom
        self.zoom_factor = new_zoom
        self.display_image()

        if not self.image_canvas or not self.image_canvas.winfo_exists(): return
        new_mouse_x = img_coord_x * new_zoom; new_mouse_y = img_coord_y * new_zoom
        scroll_x = new_mouse_x - target_canvas_x; scroll_y = new_mouse_y - target_canvas_y
        
        current_scroll_region = self.image_canvas.cget('scrollregion'); img_width_new, img_height_new = 1,1
        if current_scroll_region:
            try: sr_parts = current_scroll_region.split(); img_width_new = float(sr_parts[2]); img_height_new = float(sr_parts[3])
            except: pass
        
        scroll_x_frac = scroll_x / img_width_new if img_width_new > 0 else 0
        scroll_y_frac = scroll_y / img_height_new if img_height_new > 0 else 0
        try:
            self.image_canvas.xview_moveto(max(0.0, min(scroll_x_frac, 1.0)))
            self.image_canvas.yview_moveto(max(0.0, min(scroll_y_frac, 1.0)))
        except tk.TclError: pass
        if self.zoom_box_mode and self.zoom_box: self.update_zoom_box_content(event)

    def zoom_in(self, event=None): self.zoom(1.2, event)
    def zoom_out(self, event=None): self.zoom(1 / 1.2, event)
    def zoom_in_center(self, event=None): self.zoom(1.2, None)
    def zoom_out_center(self, event=None): self.zoom(1 / 1.2, None)
    def zoom_mouse(self, event):
        # This method remains largely the same.
        if not self.image_canvas or not self.image_canvas.winfo_exists(): return
        try:
            widget = self.root.winfo_containing(event.x_root, event.y_root); is_over_image_canvas = False
            while widget is not None:
                if widget == self.image_canvas: is_over_image_canvas = True; break
                if widget == self.root: break
                widget = widget.master
        except tk.TclError: return
        if not is_over_image_canvas: return
        factor = 1.0
        if event.num == 4: factor = 1.2       
        elif event.num == 5: factor = 1 / 1.2 
        elif hasattr(event, 'delta'): 
            if event.delta > 0: factor = 1.2
            elif event.delta < 0: factor = 1 / 1.2
        if factor != 1.0: self.zoom(factor, event)

    def on_press(self, event):
        # Uses self.measurement_model to add points and log measurements.
        if not self.img_original or not self.image_canvas or not self.image_canvas.winfo_exists(): return
        self.save_state()
        canvas_x = self.image_canvas.canvasx(event.x); canvas_y = self.image_canvas.canvasy(event.y)
        orig_x = canvas_x / self.zoom_factor; orig_y = canvas_y / self.zoom_factor

        if not (0 <= orig_x < self.img_original.width and 0 <= orig_y < self.img_original.height):
             self.measurement.set("Status: Click outside image bounds."); return

        if self.edge_selection_mode: # Legacy FIND_EDGES ROI
            self.selection_start_display = (canvas_x, canvas_y)
            self.selection_end_display = (canvas_x, canvas_y)
            self.image_canvas.delete("selection_rect_display_tag")
            self.selection_rect_display = self.image_canvas.create_rectangle(canvas_x, canvas_y, canvas_x, canvas_y, outline="red", dash=(4, 4), tags="selection_rect_display_tag")
        elif self.canny_selection_mode: # Canny ROI
            self.canny_start_display = (canvas_x, canvas_y) # Store display coords for drawing rect
            self.canny_end_display = (canvas_x, canvas_y)
            self.image_canvas.delete("canny_rect_display_tag")
            self.canny_rect_display = self.image_canvas.create_rectangle(canvas_x, canvas_y, canvas_x, canvas_y, outline="blue", dash=(4, 4), tags="canny_rect_display_tag")
        elif self.calibration_mode:
            if self.measurement_model.add_calibration_dot((orig_x, orig_y)) == 2:
                p1 = self.measurement_model.calibration_dots[0]
                p2 = self.measurement_model.calibration_dots[1]
                dist_px = math.sqrt((p2[0] - p1[0])**2 + (p2[1] - p1[1])**2)
                self.prompt_for_calibration(dist_px)
            else: self.measurement.set("Calibration: Click second point.")
            self.utils.update_dot_coords_display()
        elif self.artery_mode:
            self.measurement_model.add_artery_dot((orig_x, orig_y))
            self.utils.update_dot_coords_display()
            if len(self.measurement_model.artery_dots) % 2 == 0:
                last_meas = self.measurement_model.measurements_log[-1] if self.measurement_model.measurements_log and self.measurement_model.measurements_log[-1]['type'] == 'artery_distance' else {}
                dist_px = last_meas.get('distance_px', 0)
                angle_deg = last_meas.get('angle_deg', 0)
                status_text = f"Pair {len(self.measurement_model.artery_dots)//2}: {dist_px:.2f}px, {angle_deg:.1f}°"
                if self.measurement_model.calibration_done:
                    dist_mm = last_meas.get('distance_mm', 0)
                    status_text += f" = {dist_mm:.3f}mm"
                else: status_text += " (Uncalibrated)"
                self.measurement.set(status_text)
                self.utils.update_tables()
            else: self.measurement.set(f"Pair {len(self.measurement_model.artery_dots)//2 + 1}: Click second point.")
        elif self.angle_mode:
            num_pts = self.measurement_model.add_angle_point((orig_x, orig_y))
            self.utils.update_dot_coords_display()
            if num_pts == 1: self.measurement.set("Angle: Click vertex point (2nd).")
            elif num_pts == 2: self.measurement.set("Angle: Click final point (3rd).")
            elif num_pts == 3:
                last_meas = self.measurement_model.measurements_log[-1] if self.measurement_model.measurements_log and self.measurement_model.measurements_log[-1]['type'] == 'angle' else {}
                angle_deg = last_meas.get('angle_deg', 0)
                self.measurement.set(f"Angle Measured: {angle_deg:.2f}°. Click to start new angle.")
                self.utils.update_tables()
        elif self.line_mode:
            num_pts = self.measurement_model.add_line_mode_point((orig_x, orig_y))
            self.utils.update_dot_coords_display()
            if num_pts < 4: self.measurement.set(f"Line Mode: Click {4-num_pts} more point(s).")
            else: # 4 points placed, measurement logged by model
                last_meas = self.measurement_model.measurements_log[-1] if self.measurement_model.measurements_log and self.measurement_model.measurements_log[-1]['type'] == 'line_measurement' else {}
                l1px = last_meas.get('length1_px',0); l2px = last_meas.get('length2_px',0)
                angle = last_meas.get('angle_deviation_deg',0); avg_px = last_meas.get('avg_distance_px',0)
                status = f"Lines: L1={l1px:.1f}px, L2={l2px:.1f}px, Angle={angle:.1f}°, AvgDist={avg_px:.2f}px"
                if self.measurement_model.calibration_done:
                    l1mm = last_meas.get('length1_mm',0); l2mm = last_meas.get('length2_mm',0); avg_mm = last_meas.get('avg_distance_mm',0)
                    status = f"Lines: L1={l1mm:.2f}mm, L2={l2mm:.2f}mm, Angle={angle:.1f}°, AvgDist={avg_mm:.3f}mm"
                self.measurement.set(f"{status} (Click 'Reset Lines' for new measurement)")
                self.utils.update_tables()
        self.display_image()
        if self.zoom_box_mode: self.update_zoom_box_content(event)

    def on_motion(self, event):
        # This method remains largely the same.
        if not self.image_canvas or not self.image_canvas.winfo_exists(): return
        try:
            canvas_x = self.image_canvas.canvasx(event.x); canvas_y = self.image_canvas.canvasy(event.y)
            if self.edge_selection_mode and self.selection_start_display:
                self.selection_end_display = (canvas_x, canvas_y)
                if self.selection_rect_display: self.image_canvas.coords(self.selection_rect_display, self.selection_start_display[0], self.selection_start_display[1], canvas_x, canvas_y)
                self.update_zoom_box_and_pixel(event); return
            elif self.canny_selection_mode and self.canny_start_display:
                self.canny_end_display = (canvas_x, canvas_y)
                if self.canny_rect_display: self.image_canvas.coords(self.canny_rect_display, self.canny_start_display[0], self.canny_start_display[1], canvas_x, canvas_y)
                # For live Canny ROI preview, convert display coords to original image coords
                orig_x1 = self.canny_start_display[0] / self.zoom_factor
                orig_y1 = self.canny_start_display[1] / self.zoom_factor
                orig_x2 = canvas_x / self.zoom_factor # current mouse pos in orig coords
                orig_y2 = canvas_y / self.zoom_factor # current mouse pos in orig coords
                self.filter_handler.set_canny_roi_original_coords((orig_x1, orig_y1), (orig_x2, orig_y2))
                self.filter_handler.apply_filters_and_update_display() # Apply Canny to this temp ROI
                self.update_zoom_box_and_pixel(event); return # update_zoom_box_and_pixel is separate
            self.update_zoom_box_and_pixel(event)
        except tk.TclError: pass

    def on_release(self, event):
        # Uses FilterHandler to set final ROI.
        if not self.image_canvas or not self.image_canvas.winfo_exists(): return
        try:
            canvas_x = self.image_canvas.canvasx(event.x); canvas_y = self.image_canvas.canvasy(event.y)
            if self.edge_selection_mode and self.selection_start_display:
                self.selection_end_display = (canvas_x, canvas_y)
                # Finalize legacy selection_start/end if needed by other parts, or remove if fully deprecated
                # For now, just update display rect
                if abs(self.selection_start_display[0] - self.selection_end_display[0]) < 1 or abs(self.selection_start_display[1] - self.selection_end_display[1]) < 1:
                    self.selection_start_display = self.selection_end_display = None
                    self.image_canvas.delete("selection_rect_display_tag")
                    self.measurement.set("Status: ROI selection cancelled (zero size).")
                else:
                    x1,y1,x2,y2 = min(self.selection_start_display[0], self.selection_end_display[0]), min(self.selection_start_display[1], self.selection_end_display[1]), max(self.selection_start_display[0], self.selection_end_display[0]), max(self.selection_start_display[1], self.selection_end_display[1])
                    self.selection_start_display=(x1,y1); self.selection_end_display=(x2,y2)
                    self.image_canvas.coords("selection_rect_display_tag", x1,y1,x2,y2) # Finalize display rect
                    self.measurement.set("Status: ROI selected for Edge Detection.")
                self.edge_selection_mode = False
                if "ROI Selection" in self.buttons: self.buttons["ROI Selection"].config(relief=tk.RAISED)
            elif self.canny_selection_mode and self.canny_start_display:
                self.canny_end_display = (canvas_x, canvas_y)
                if abs(self.canny_start_display[0] - self.canny_end_display[0]) < 1 or abs(self.canny_start_display[1] - self.canny_end_display[1]) < 1:
                    self.canny_start_display = self.canny_end_display = None
                    self.filter_handler.set_canny_roi_original_coords(None, None) # Clear in filter handler
                    self.image_canvas.delete("canny_rect_display_tag")
                    self.img_filtered = None # Clear filter effect
                    self.measurement.set("Status: Canny ROI cancelled (zero size).")
                else:
                    # Convert final display ROI to original image coordinates for FilterHandler
                    orig_x1 = min(self.canny_start_display[0], self.canny_end_display[0]) / self.zoom_factor
                    orig_y1 = min(self.canny_start_display[1], self.canny_end_display[1]) / self.zoom_factor
                    orig_x2 = max(self.canny_start_display[0], self.canny_end_display[0]) / self.zoom_factor
                    orig_y2 = max(self.canny_start_display[1], self.canny_end_display[1]) / self.zoom_factor
                    self.filter_handler.set_canny_roi_original_coords((orig_x1, orig_y1), (orig_x2, orig_y2))
                    # Finalize display rect coords
                    self.canny_start_display = (min(self.canny_start_display[0], self.canny_end_display[0]), min(self.canny_start_display[1], self.canny_end_display[1]))
                    self.canny_end_display = (max(self.canny_start_display[0], self.canny_end_display[0]), max(self.selection_start_display[1], self.canny_end_display[1]))
                    self.image_canvas.coords("canny_rect_display_tag", self.canny_start_display[0], self.canny_start_display[1], self.canny_end_display[0], self.canny_end_display[1])
                    self.measurement.set(f"Status: Canny filter applied to ROI (Thresh: {self.canny_low.get()}/{self.canny_high.get()}).")
                
                self.filter_handler.apply_filters_and_update_display() # Apply final Canny ROI
                self.canny_selection_mode = False # Turn off selection mode
                if "Canny Selection" in self.buttons: self.buttons["Canny Selection"].config(relief=tk.RAISED)
            
            self.display_image() # Ensure final state is drawn
            if self.zoom_box_mode: self.update_zoom_box_content(event)
        except tk.TclError: pass

    # _reset_all_modes is now primarily handled by ModeHandler.
    # ImageAnalyzer methods will call ModeHandler methods.

    def delete_last_artery_pair(self):
        self.save_state()
        if self.measurement_model.delete_last_artery_pair():
            self.measurement.set("Status: Last dot pair and measurement deleted.")
        else:
            self.measurement.set("Status: No dots to delete.")
        self.display_image()
        self.utils.update_dot_coords_display()
        self.utils.update_tables()

    def reset_artery_mode(self):
        self.save_state()
        if self.artery_mode: self.mode_handler.toggle_artery_mode() # Turn off if active
        self.measurement_model.reset_artery_dots_and_measurements()
        self.measurement.set("Status: Dots Mode reset.")
        self.utils.update_dot_coords_display()
        self.utils.update_tables()
        self.display_image()

    def prompt_for_calibration(self, distance_px):
        # Uses self.measurement_model to set calibration.
        self.root.attributes('-topmost', 1)
        real_value_str = simpledialog.askstring("Calibration Input", f"Measured {distance_px:.2f} pixels.\nEnter REAL distance (mm):", parent=self.root)
        self.root.attributes('-topmost', 0)
        if real_value_str:
            try:
                real_value = float(real_value_str)
                if self.measurement_model.set_calibration(distance_px, real_value):
                    self.utils.update_tables()
                    self.utils.update_dot_coords_display()
                    self.mode_handler._reset_all_modes() # Exit calibration mode
                    self.measurement.set(f"Calibrated: {self.measurement_model.calibration_factor:.4f} px/mm")
                    messagebox.showinfo("Calibration Success", f"Factor: {self.measurement_model.calibration_factor:.4f} px/mm", parent=self.root)
                else:
                    messagebox.showerror("Calibration Error", "Real distance must be positive.", parent=self.root)
                    if self.measurement_model.calibration_dots: self.measurement_model.calibration_dots.pop() # Remove last dot if invalid
                    self.measurement.set("Calibration Error: Enter positive distance.")
            except ValueError:
                messagebox.showerror("Calibration Error", "Invalid number.", parent=self.root)
                if self.measurement_model.calibration_dots: self.measurement_model.calibration_dots.pop()
                self.measurement.set("Calibration Error: Invalid input.")
        else: # Cancelled
            if self.measurement_model.calibration_dots: self.measurement_model.calibration_dots.pop()
            self.measurement.set("Calibration: Cancelled. Click second point again.")
        self.display_image() # Redraw with potentially removed dot

    def reset_calibration(self, ask_confirm=True):
        # Uses self.measurement_model.
        if ask_confirm:
             if not self.measurement_model.calibration_done and not self.measurement_model.calibration_dots:
                  messagebox.showinfo("Calibration Reset", "No calibration data to reset.", parent=self.root); return
             if not messagebox.askyesno("Confirm Reset", "Reset current calibration data?", parent=self.root): return
        self.save_state()
        self.measurement_model.reset_calibration()
        self.keep_calibration_dots_fixed = False # Also reset this flag
        if "Keep Dots Fixed" in self.buttons:
            try: self.buttons["Keep Dots Fixed"].config(relief=tk.RAISED)
            except tk.TclError: pass
        if self.calibration_mode: self.mode_handler._reset_all_modes()
        else: self.measurement.set("Status: Calibration reset.")
        self.utils.update_dot_coords_display(); self.utils.update_tables(); self.display_image()

    def toggle_keep_dots_fixed(self):
        self.save_state()
        self.keep_calibration_dots_fixed = not self.keep_calibration_dots_fixed
        status = "ON" if self.keep_calibration_dots_fixed else "OFF"
        relief_style = tk.SUNKEN if self.keep_calibration_dots_fixed else tk.RAISED
        if "Keep Dots Fixed" in self.buttons:
            try: self.buttons["Keep Dots Fixed"].config(relief=relief_style)
            except tk.TclError: pass
        self.measurement.set(f"Status: Keep Calibration Dots Fixed is {status}.")

    # --- Filter Control Methods ---
    def toggle_global_canny_filter(self):
        self.save_state()
        self.filter_handler.is_global_canny_active = not self.filter_handler.is_global_canny_active
        if self.filter_handler.is_global_canny_active:
            # If activating global, ensure ROI selection mode is off and ROI data is cleared in FilterHandler
            if self.canny_selection_mode:
                self.mode_handler._set_mode('canny_selection_mode', False) # Update app's mode flag
                # self.mode_handler._reset_all_modes() # Or more general reset
            self.filter_handler.set_canny_roi_original_coords(None, None) # Clear ROI in filter handler
            self.canny_start_display = self.canny_end_display = None # Clear display ROI
            if hasattr(self, 'image_canvas'): self.image_canvas.delete("canny_rect_display_tag")
        
        # Update button appearance
        if "Global Canny" in self.buttons:
            self.buttons["Global Canny"].config(relief=tk.SUNKEN if self.filter_handler.is_global_canny_active else tk.RAISED)
        
        self.filter_handler.apply_filters_and_update_display()

    def toggle_canny_roi_selection_mode(self):
        self.save_state()
        new_state = not self.canny_selection_mode # Toggle app's selection mode flag
        
        if new_state: # Entering ROI selection mode
            if self.filter_handler.is_global_canny_active:
                self.filter_handler.is_global_canny_active = False # Turn off global Canny
                if "Global Canny" in self.buttons: self.buttons["Global Canny"].config(relief=tk.RAISED)
            # Clear any existing ROI in FilterHandler to start fresh selection
            self.filter_handler.set_canny_roi_original_coords(None, None)
            self.canny_start_display = self.canny_end_display = None # Clear display ROI
            if hasattr(self, 'image_canvas'): self.image_canvas.delete("canny_rect_display_tag")
            self.filter_handler.apply_filters_and_update_display() # Remove global Canny effect if it was on

        self.mode_handler._set_mode('canny_selection_mode', new_state) # Use ModeHandler to manage mode state and button relief

    def reset_all_filters(self):
        self.save_state()
        self.filter_handler.reset_filter_states() # This handles logic and calls apply_filters_and_update_display
        # Ensure app's selection mode flags are also reset if FilterHandler doesn't do it
        if self.canny_selection_mode:
            self.mode_handler._set_mode('canny_selection_mode', False)
        # Reset legacy edge detection if any
        self.edge_detection_active = False
        self.selection_start_display = self.selection_end_display = None
        if hasattr(self, 'image_canvas'): self.image_canvas.delete("selection_rect_display_tag")
        self.measurement.set("Status: All filters reset.") # FilterHandler might set its own message

    def toggle_zoom_box(self):
        # This method remains largely the same.
        if not self.root or not self.root.winfo_exists() or not self.image_frame or not self.image_frame.winfo_exists(): return
        self.save_state(); self.zoom_box_mode = not self.zoom_box_mode
        if self.zoom_box_mode:
            self.buttons["Zoom In Box"].config(relief=tk.SUNKEN)
            self.measurement.set("Status: Zoom Box ON. Move cursor over image.")
            if not self.zoom_box or not self.zoom_box.winfo_exists():
                 try: self.zoom_box = tk.Canvas(self.image_frame, width=self.ZOOM_BOX_SIZE, height=self.ZOOM_BOX_SIZE, bg="black", highlightthickness=1, highlightbackground="white")
                 except tk.TclError: self.zoom_box_mode = False; self.buttons["Zoom In Box"].config(relief=tk.RAISED); self.measurement.set("Status: Error creating Zoom Box."); return
            try: self.zoom_box.place(in_=self.image_frame, anchor='se', relx=1.0, rely=1.0, x=-10, y=-10); self.update_zoom_box_content(None)
            except tk.TclError: self.zoom_box_mode = False; self.buttons["Zoom In Box"].config(relief=tk.RAISED); self.measurement.set("Status: Zoom Box Error.")
        else: 
            self.buttons["Zoom In Box"].config(relief=tk.RAISED); self.measurement.set("Status: Zoom Box OFF.")
            if self.zoom_box and self.zoom_box.winfo_exists():
                try: self.zoom_box.place_forget()
                except tk.TclError: pass
        try: # Update pixel info string
            current_info = self.pixel_info.get(); parts = current_info.split('|')
            mode_str = parts[0].strip() if parts else "Mode: Unknown"
            pixel_str = parts[1].strip() if len(parts) > 1 else "Pixel:"
            self.pixel_info.set(f"{mode_str} | {pixel_str} | Zoom: {'ON' if self.zoom_box_mode else 'OFF'}")
        except Exception: self.pixel_info.set(f"Mode: Unknown | Zoom: {'ON' if self.zoom_box_mode else 'OFF'}")

    def reset_angle_mode(self):
        self.save_state()
        self.measurement_model.reset_angle_points_and_measurements()
        if self.angle_mode: self.measurement.set("Angle: Click first point.") # Stay in mode if active
        else: self.measurement.set("Status: Angle measurements cleared.")
        self.utils.update_dot_coords_display(); self.utils.update_tables(); self.display_image()

    def reset_lines(self):
        self.save_state()
        if self.line_mode: self.mode_handler.toggle_line_mode() # Deactivate mode
        self.measurement_model.reset_line_points_and_measurements()
        self.measurement.set("Status: Line Mode reset.")
        self.utils.update_dot_coords_display(); self.utils.update_tables(); self.display_image()
    
    def show_line_measurements(self):
        # Uses self.measurement_model.
        line_measurement = None
        for meas in reversed(self.measurement_model.get_all_measurements()): # Use getter
            if meas.get("type") == "line_measurement": line_measurement = meas; break
        if not line_measurement: messagebox.showinfo("No Line Measurement", "No Line Mode measurement found.", parent=self.root); return
        
        window = Toplevel(self.root); window.title("Line Mode - Measurement Details"); window.geometry("450x400"); window.transient(self.root); window.grab_set()
        text_frame = tk.Frame(window); text_frame.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)
        text_widget = tk.Text(text_frame, height=15, width=50, wrap=tk.WORD, font=("Courier New", 10))
        scrollbar = Scrollbar(text_frame, orient=tk.VERTICAL, command=text_widget.yview); text_widget.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y); text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        text_widget.insert(tk.END, "--- Line Mode Summary ---\n\n")
        len1_px = line_measurement.get('length1_px'); len2_px = line_measurement.get('length2_px'); angle = line_measurement.get('angle_deviation_deg')
        text_widget.insert(tk.END, f"Line 1 Length: {len1_px:>8.2f} px" if len1_px is not None else "N/A")
        if self.measurement_model.calibration_done and 'length1_mm' in line_measurement: text_widget.insert(tk.END, f"  ({line_measurement['length1_mm']:.3f} mm)\n")
        else: text_widget.insert(tk.END, "\n")
        # ... (similar for line 2 and angle) ...
        distances_px = line_measurement.get('distances_px', []); distances_mm = line_measurement.get('distances_mm', [])
        text_widget.insert(tk.END, "\n--- Distance Measurements ---\n # |   Pixels  |    mm\n---|-----------|-----------\n")
        for i, dist_px in enumerate(distances_px):
             line = f"{i+1:>2} | {dist_px:>9.2f} |"
             if self.measurement_model.calibration_done and i < len(distances_mm): line += f" {distances_mm[i]:>9.3f}\n"
             else: line += "   N/A\n"
             text_widget.insert(tk.END, line)
        # ... (stats: avg, min, max) ...
        text_widget.config(state=tk.DISABLED)
        tk.Button(window, text="Close", command=window.destroy).pack(pady=5)
        window.wait_window()

    def save_state(self):
        if not self.img_original: return
        state = {
            "measurement_model_state": self.measurement_model.to_dict(),
            "filter_handler_state": { # Basic filter state, FilterHandler could have its own to_dict
                "canny_low_thresh": self.filter_handler.canny_low_thresh,
                "canny_high_thresh": self.filter_handler.canny_high_thresh,
                "canny_roi_start_orig": self.filter_handler.canny_roi_start_orig,
                "canny_roi_end_orig": self.filter_handler.canny_roi_end_orig,
                "is_global_canny_active": self.filter_handler.is_global_canny_active,
            },
            "img_filtered_pil": self.img_filtered.copy() if self.img_filtered else None, # Save PIL image directly
            "keep_calibration_dots_fixed": self.keep_calibration_dots_fixed,
            # UI specific states not in models/handlers
            "canny_selection_mode": self.canny_selection_mode, # App's mode flag
            "canny_start_display": self.canny_start_display, # For UI rect
            "canny_end_display": self.canny_end_display,     # For UI rect
        }
        self.undo_stack.append(state)
        self.redo_stack.clear()
        if len(self.undo_stack) > 50: self.undo_stack.pop(0)

    def _restore_state(self, state):
        self.measurement_model.from_dict(state.get("measurement_model_state", {}))
        
        fh_state = state.get("filter_handler_state", {})
        self.filter_handler.canny_low_thresh = fh_state.get("canny_low_thresh", 100)
        self.filter_handler.canny_high_thresh = fh_state.get("canny_high_thresh", 200)
        self.filter_handler.canny_roi_start_orig = fh_state.get("canny_roi_start_orig")
        self.filter_handler.canny_roi_end_orig = fh_state.get("canny_roi_end_orig")
        self.filter_handler.is_global_canny_active = fh_state.get("is_global_canny_active", False)
        self.canny_low.set(self.filter_handler.canny_low_thresh) # Update UI slider
        self.canny_high.set(self.filter_handler.canny_high_thresh) # Update UI slider

        img_filt_pil = state.get("img_filtered_pil")
        self.img_filtered = img_filt_pil.copy() if img_filt_pil else None
        
        self.keep_calibration_dots_fixed = state.get("keep_calibration_dots_fixed", False)
        self.canny_selection_mode = state.get("canny_selection_mode", False)
        self.canny_start_display = state.get("canny_start_display")
        self.canny_end_display = state.get("canny_end_display")

        # Update button states based on restored flags
        if "Global Canny" in self.buttons: self.buttons["Global Canny"].config(relief=tk.SUNKEN if self.filter_handler.is_global_canny_active else tk.RAISED)
        if "Keep Dots Fixed" in self.buttons: self.buttons["Keep Dots Fixed"].config(relief=tk.SUNKEN if self.keep_calibration_dots_fixed else tk.RAISED)
        if "Canny Selection" in self.buttons: self.buttons["Canny Selection"].config(relief=tk.SUNKEN if self.canny_selection_mode else tk.RAISED)

        self.display_image() # This will draw ROI rects based on display coords if they exist
        self.utils.update_dot_coords_display()
        self.utils.update_tables()
        self.mode_handler._reset_all_modes() # Resets UI modes, not data
        # Set appropriate status message
        if self.measurement_model.calibration_done: self.measurement.set(f"Calibrated: {self.measurement_model.calibration_factor:.4f} px/mm")
        elif self.img_filtered: self.measurement.set("Status: Filter applied.") # Generic
        else: self.measurement.set("Status: Ready.")


    def undo(self, event=None):
        if not self.undo_stack: self.measurement.set("Status: Nothing to undo."); return
        # Create current state for redo before popping from undo_stack
        current_state_for_redo = {
            "measurement_model_state": self.measurement_model.to_dict(),
            "filter_handler_state": {
                "canny_low_thresh": self.filter_handler.canny_low_thresh,
                "canny_high_thresh": self.filter_handler.canny_high_thresh,
                "canny_roi_start_orig": self.filter_handler.canny_roi_start_orig,
                "canny_roi_end_orig": self.filter_handler.canny_roi_end_orig,
                "is_global_canny_active": self.filter_handler.is_global_canny_active,
            },
            "img_filtered_pil": self.img_filtered.copy() if self.img_filtered else None,
            "keep_calibration_dots_fixed": self.keep_calibration_dots_fixed,
            "canny_selection_mode": self.canny_selection_mode,
            "canny_start_display": self.canny_start_display,
            "canny_end_display": self.canny_end_display,
        }
        self.redo_stack.append(current_state_for_redo)
        state_to_restore = self.undo_stack.pop()
        self._restore_state(state_to_restore)
        self.measurement.set("Status: Undo successful.")

    def redo(self, event=None):
        if not self.redo_stack: self.measurement.set("Status: Nothing to redo."); return
        # Create current state for undo before popping from redo_stack
        current_state_for_undo = {
             "measurement_model_state": self.measurement_model.to_dict(),
             "filter_handler_state": {
                "canny_low_thresh": self.filter_handler.canny_low_thresh,
                "canny_high_thresh": self.filter_handler.canny_high_thresh,
                "canny_roi_start_orig": self.filter_handler.canny_roi_start_orig,
                "canny_roi_end_orig": self.filter_handler.canny_roi_end_orig,
                "is_global_canny_active": self.filter_handler.is_global_canny_active,
            },
             "img_filtered_pil": self.img_filtered.copy() if self.img_filtered else None,
             "keep_calibration_dots_fixed": self.keep_calibration_dots_fixed,
             "canny_selection_mode": self.canny_selection_mode,
             "canny_start_display": self.canny_start_display,
             "canny_end_display": self.canny_end_display,
         }
        self.undo_stack.append(current_state_for_undo)
        state_to_restore = self.redo_stack.pop()
        self._restore_state(state_to_restore)
        self.measurement.set("Status: Redo successful.")

    # AppUtils methods (update_dot_coords_display, update_tables, save_measurements_to_json, export_annotated_image)
    # will be updated in utils.py to use self.app.measurement_model where 'app' is the ImageAnalyzer instance.
    # For now, calls like self.utils.update_tables() are assumed to work once utils.py is updated.

# Main execution
if __name__ == "__main__":
    root = None 
    try:
        root = tk.Tk()
        try:
             from ttkthemes import ThemedTk
             if root: root.destroy()
             root = ThemedTk(theme="arc") 
        except ImportError:
             print("ttkthemes not found, using default Tk theme.")
             if not root or not root.winfo_exists(): root = tk.Tk()
        except tk.TclError as theme_error:
             print(f"Error setting ttk theme: {theme_error}. Using default Tk theme.")
             if not root or not root.winfo_exists(): root = tk.Tk()
        
        app = ImageAnalyzer(root)
        root.mainloop()
    except Exception as e:
        print(f"--- FATAL ERROR IN MAIN --- \n{traceback.format_exc()}")
        try: messagebox.showerror("Fatal Error", f"An unrecoverable error occurred:\n\n{e}\n\nSee console for details.")
        except Exception: pass
