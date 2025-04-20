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

class ImageAnalyzer:
    def __init__(self, root):
        self.root = root
        self.root.title("Image Analyzer")
        self.root.geometry("1400x900") # Increased default height slightly

        self.ZOOM_BOX_SIZE = 180
        self.ZOOM_BOX_FACTOR = 4
        # Position is now relative to image_frame, set in toggle_zoom_box

        self.file_path = None
        self.image_files = []
        self.current_index = 0
        self.zoom_factor = 1.0 # Start at 1.0 zoom
        self.img_original = None
        self.img_filtered = None # Will hold filtered image if any filter is applied
        self.photo = None # Reference to PhotoImage for main canvas
        self.zoom_box_photo = None # Reference to PhotoImage for zoom box

        self.calibration_dots = []
        self.artery_dots = []
        self.line_points = []  # For Line Mode
        self.measurements = []
        self.line_measurements = []  # For Line Mode measurements (pixel distances list)
        self.line_measurement_points = []  # For visualization of tick markers in Line Mode
        self.calibration_factor = 1.0
        self.calibration_done = False
        self.angle_points = []

        # --- Mode Flags ---
        self.edge_detection_active = False # Legacy FIND_EDGES filter flag
        self.global_canny_active = False # Global Canny Toggle
        self.edge_selection_mode = False  # ROI Selection mode for FIND_EDGES
        self.canny_selection_mode = False # ROI Selection mode for Canny
        self.zoom_box_mode = False
        self.calibration_mode = False
        self.artery_mode = False
        self.angle_mode = False
        self.line_mode = False  # New Line Mode

        # --- Selection Rectangles ---
        self.selection_rect = None # For FIND_EDGES ROI
        self.selection_start = None
        self.selection_end = None
        self.canny_rect = None      # For Canny ROI
        self.canny_start = None
        self.canny_end = None
        self.zoom_box = None # Will be created in create_gui

        self.undo_stack = []
        self.redo_stack = []

        # --- StringVars for Labels ---
        self.path_text = tk.StringVar(value="Path: No image loaded")
        self.pixel_info = tk.StringVar(value="Pixel: ")
        self.measurement = tk.StringVar(value="Distance: ")
        self.name_var = tk.StringVar(value="")
        self.diameter_var = tk.StringVar(value="")
        # --- IntVars for Canny Thresholds ---
        self.canny_low = tk.IntVar(value=100)
        self.canny_high = tk.IntVar(value=200)
        # Link slider changes to update the display
        self.canny_low.trace_add("write", self.apply_filters_and_display)
        self.canny_high.trace_add("write", self.apply_filters_and_display)


        self.measurement_table = None
        self.image_frame = None # Initialize image_frame attribute

        self.create_gui()
        self.bind_events()


    def create_gui(self):
        self.main_frame = tk.Frame(self.root)
        self.main_frame.pack(fill=tk.BOTH, expand=True)

        # Top frame holds button panel and image canvas side by side
        self.top_frame = tk.Frame(self.main_frame)
        self.top_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        # --- Scrollable Button Frame Setup ---
        # Container for the canvas and scrollbar
        self.button_area = tk.Frame(self.top_frame, borderwidth=2, relief=tk.SOLID)
        self.button_area.pack(side=tk.LEFT, fill=tk.Y, padx=5, pady=5)
        self.button_area.pack_propagate(False) # Prevent container resizing
        self.button_area.config(width=210) # Give the container a fixed width

        # Canvas to hold the buttons
        self.button_canvas = tk.Canvas(self.button_area, borderwidth=0, highlightthickness=0)

        # Scrollbar linked to the canvas
        self.button_scrollbar = Scrollbar(self.button_area, orient=tk.VERTICAL, command=self.button_canvas.yview)
        self.button_canvas.configure(yscrollcommand=self.button_scrollbar.set)

        # Pack canvas and scrollbar
        self.button_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.button_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # The ACTUAL frame where buttons will be placed (INSIDE the canvas)
        self.button_frame = tk.Frame(self.button_canvas)
        self.button_canvas.create_window((0, 0), window=self.button_frame, anchor='nw', tags="button_frame")

        # Update scrollregion when button_frame size changes
        self.button_frame.bind('<Configure>', self._on_button_frame_configure)
        # --- End Scrollable Button Frame Setup ---


        # Image frame on the right of buttons
        # Make sure self.image_frame is assigned here
        self.image_frame = tk.Frame(self.top_frame)
        self.image_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.image_canvas = tk.Canvas(self.image_frame, bg="gray") # Set background color
        self.image_canvas.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        # --- Create Zoom Box (now child of image_frame) ---
        # Ensure image_frame exists before creating zoom_box as its child
        if self.image_frame:
            self.zoom_box = tk.Canvas(self.image_frame, width=self.ZOOM_BOX_SIZE, height=self.ZOOM_BOX_SIZE,
                                    bg="black", highlightthickness=1, highlightbackground="white")
        else:
             print("Error: image_frame not created before zoom_box initialization.")
        # Note: We use .place() later in toggle_zoom_box, no pack/grid here.

        # --- Middle Section (Dot Coords & Table) ---
        self.middle_section_frame = tk.Frame(self.main_frame)
        self.middle_section_frame.pack(side=tk.TOP, fill=tk.X, padx=5, pady=5)

        # Dot Coordinates Text Box
        self.dot_coords_frame = tk.Frame(self.middle_section_frame)
        self.dot_coords_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(5, 2), pady=5)
        tk.Label(self.dot_coords_frame, text="Dot Coordinates:").pack(side=tk.TOP, anchor=tk.W)
        self.dot_coords_text = tk.Text(self.dot_coords_frame, height=10, width=40)
        self.dot_coords_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.dot_coords_text.config(state=tk.DISABLED)
        dot_scrollbar = Scrollbar(self.dot_coords_frame, orient=tk.VERTICAL, command=self.dot_coords_text.yview)
        dot_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.dot_coords_text.configure(yscrollcommand=dot_scrollbar.set)

        # Measurement Table
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
        # --- End Middle Section ---


        # --- Status Bar ---
        self.status_frame = tk.Frame(self.main_frame, bg="black", bd=1, relief=tk.SUNKEN)
        self.status_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=5, pady=3)
        self.path_label = tk.Label(self.status_frame, textvariable=self.path_text, bg="black", fg="white", anchor=tk.W)
        self.path_label.pack(side=tk.LEFT, padx=5)
        self.pixel_label = tk.Label(self.status_frame, textvariable=self.pixel_info, bg="black", fg="white", anchor=tk.W)
        self.pixel_label.pack(side=tk.LEFT, padx=10)
        self.measurement_label = tk.Label(self.status_frame, textvariable=self.measurement, bg="black", fg="white", anchor=tk.W)
        self.measurement_label.pack(side=tk.LEFT, padx=10)
        # --- End Status Bar ---

        self.create_buttons() # Buttons are now created *after* button_frame exists

        # Bind mousewheel scrolling to the button canvas/frame
        # Bind directly to the canvas and frame where scrolling should happen
        # These bindings attach/detach the global scroll binding when mouse enters/leaves
        self.button_canvas.bind("<Enter>", self._bind_mousewheel_button_area)
        self.button_canvas.bind("<Leave>", self._unbind_mousewheel_button_area)
        self.button_frame.bind("<Enter>", self._bind_mousewheel_button_area)
        self.button_frame.bind("<Leave>", self._unbind_mousewheel_button_area)


    # --- Helper Method for Scrollbar ---
    def _on_button_frame_configure(self, event=None):
        """Update scroll region of button_canvas when button_frame resizes."""
        if hasattr(self, 'button_canvas') and self.button_canvas and self.button_canvas.winfo_exists():
            self.button_canvas.configure(scrollregion=self.button_canvas.bbox('all'))

    # --- Helper Methods for Mouse Wheel ---
    def _bind_mousewheel_button_area(self, event):
        """Bind mousewheel globally when pointer enters button area."""
        self.root.bind_all("<MouseWheel>", self._on_mousewheel)
        self.root.bind_all("<Button-4>", self._on_mousewheel) # Linux scroll up
        self.root.bind_all("<Button-5>", self._on_mousewheel) # Linux scroll down

    def _unbind_mousewheel_button_area(self, event):
        """Unbind global mousewheel when pointer leaves button area."""
        self.root.unbind_all("<MouseWheel>")
        self.root.unbind_all("<Button-4>")
        self.root.unbind_all("<Button-5>")

    def _on_mousewheel(self, event):
        """Scroll the button canvas using the mouse wheel IF event is over button area."""
        # Check if the event happened directly over the canvas or frame
        # Alternative: Check if pointer is currently within the button_area bounds
        bx, by, bw, bh = self.button_area.winfo_rootx(), self.button_area.winfo_rooty(), self.button_area.winfo_width(), self.button_area.winfo_height()
        if not (bx <= event.x_root < bx + bw and by <= event.y_root < by + bh):
             return # Event not within button area, let other bindings handle it (e.g., zoom)

        # Determine scroll direction based on platform
        delta = 0
        if event.num == 4: delta = -1 # Linux scroll up
        elif event.num == 5: delta = 1  # Linux scroll down
        elif hasattr(event, 'delta'): # Windows/macOS
             if event.delta > 0: delta = -1 # Windows/macOS scroll up
             elif event.delta < 0: delta = 1  # Windows/macOS scroll down

        if delta != 0 and hasattr(self, 'button_canvas') and self.button_canvas:
            self.button_canvas.yview_scroll(delta, "units")
            return "break" # Consume the event to prevent other bindings


    def create_buttons(self):
        # IMPORTANT: All buttons now have self.button_frame as their parent
        self.buttons = {}
        pad_options = {'fill': tk.X, 'padx': 3, 'pady': 2} # Consistent padding

        # --- File ---
        file_frame = tk.LabelFrame(self.button_frame, text="File", bd=2, relief=tk.GROOVE)
        file_frame.pack(fill=tk.X, padx=3, pady=3)
        self.buttons["Load Image"] = tk.Button(file_frame, text="Load Image", command=self.load_image)
        self.buttons["Load Image"].pack(**pad_options)
        self.buttons["Export Image"] = tk.Button(file_frame, text="Export Image", command=self.export_annotated_image)
        self.buttons["Export Image"].pack(**pad_options)

        # --- Dots Mode ---
        artery_frame = tk.LabelFrame(self.button_frame, text="Dots Mode (Distance/Angle)", bd=2, relief=tk.GROOVE)
        artery_frame.pack(fill=tk.X, padx=3, pady=3)
        self.buttons["Artery Mode"] = tk.Button(artery_frame, text="Dots Mode", command=self.toggle_artery_mode)
        self.buttons["Artery Mode"].pack(**pad_options)
        self.buttons["Reset Artery"] = tk.Button(artery_frame, text="Reset Dots", command=self.reset_artery_mode)
        self.buttons["Reset Artery"].pack(**pad_options)
        self.buttons["Delete Last Pair"] = tk.Button(artery_frame, text="Delete Last Pair", command=self.delete_last_pair)
        self.buttons["Delete Last Pair"].pack(**pad_options)

        # --- Calibration ---
        calib_frame = tk.LabelFrame(self.button_frame, text="Calibration", bd=2, relief=tk.GROOVE)
        calib_frame.pack(fill=tk.X, padx=3, pady=3)
        self.buttons["Calibrate"] = tk.Button(calib_frame, text="Calibrate", command=self.toggle_calibration_mode)
        self.buttons["Calibrate"].pack(**pad_options)
        self.buttons["Reset Calibration"] = tk.Button(calib_frame, text="Reset Calibration", command=self.reset_calibration)
        self.buttons["Reset Calibration"].pack(**pad_options)

        # --- Angle Measurement ---
        angle_frame = tk.LabelFrame(self.button_frame, text="Angle Measurement", bd=2, relief=tk.GROOVE)
        angle_frame.pack(fill=tk.X, padx=3, pady=3)
        self.buttons["Angle Mode"] = tk.Button(angle_frame, text="Angle Mode", command=self.toggle_angle_mode)
        self.buttons["Angle Mode"].pack(**pad_options)

        # --- Line Mode ---
        line_frame = tk.LabelFrame(self.button_frame, text="Line Mode (Parallel)", bd=2, relief=tk.GROOVE)
        line_frame.pack(fill=tk.X, padx=3, pady=3)
        self.buttons["Line Mode"] = tk.Button(line_frame, text="Line Mode", command=self.toggle_line_mode)
        self.buttons["Line Mode"].pack(**pad_options)
        self.buttons["Reset Lines"] = tk.Button(line_frame, text="Reset Lines", command=self.reset_lines)
        self.buttons["Reset Lines"].pack(**pad_options)
        self.buttons["Show Line Measurements"] = tk.Button(line_frame, text="Show Line Measurements", command=self.show_line_measurements)
        self.buttons["Show Line Measurements"].pack(**pad_options)

        # --- Filters ---
        filter_frame = tk.LabelFrame(self.button_frame, text="Filters", bd=2, relief=tk.GROOVE)
        filter_frame.pack(fill=tk.X, padx=3, pady=3)

        # --- Global Canny ---
        self.buttons["Global Canny"] = tk.Button(filter_frame, text="Global Canny Filter", command=self.toggle_global_canny)
        self.buttons["Global Canny"].pack(**pad_options)
        # --- End Global Canny ---

        # --- Canny ROI ---
        self.buttons["Canny Selection"] = tk.Button(filter_frame, text="Canny ROI Selection", command=self.toggle_canny_selection)
        self.buttons["Canny Selection"].pack(**pad_options)

        canny_params_frame = tk.Frame(filter_frame)
        canny_params_frame.pack(fill=tk.X, padx=3, pady=3)

        low_frame = tk.Frame(canny_params_frame)
        low_frame.pack(fill=tk.X, pady=1)
        tk.Label(low_frame, text="Low:", width=4).pack(side=tk.LEFT, padx=(0,2))
        self.canny_low_slider = tk.Scale(low_frame, from_=0, to=255, orient=tk.HORIZONTAL,
                                         variable=self.canny_low, length=140, showvalue=True)
        self.canny_low_slider.pack(side=tk.LEFT, fill=tk.X, expand=True)

        high_frame = tk.Frame(canny_params_frame)
        high_frame.pack(fill=tk.X, pady=1)
        tk.Label(high_frame, text="High:", width=4).pack(side=tk.LEFT, padx=(0,2))
        self.canny_high_slider = tk.Scale(high_frame, from_=0, to=255, orient=tk.HORIZONTAL,
                                          variable=self.canny_high, length=140, showvalue=True)
        self.canny_high_slider.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.buttons["Reset Filters"] = tk.Button(filter_frame, text="Reset Filters", command=self.reset_filters)
        self.buttons["Reset Filters"].pack(**pad_options)


        # --- Zoom ---
        zoom_frame = tk.LabelFrame(self.button_frame, text="Zoom", bd=2, relief=tk.GROOVE)
        zoom_frame.pack(fill=tk.X, padx=3, pady=3)
        self.buttons["Zoom In Box"] = tk.Button(zoom_frame, text="Zoom In Box", command=self.toggle_zoom_box)
        self.buttons["Zoom In Box"].pack(**pad_options)

        # --- History ---
        history_frame = tk.LabelFrame(self.button_frame, text="History", bd=2, relief=tk.GROOVE)
        history_frame.pack(fill=tk.X, padx=3, pady=3)
        self.buttons["Undo"] = tk.Button(history_frame, text="Undo (Ctrl+Z)", command=self.undo)
        self.buttons["Undo"].pack(**pad_options)
        self.buttons["Redo"] = tk.Button(history_frame, text="Redo (Ctrl+Y)", command=self.redo)
        self.buttons["Redo"].pack(**pad_options)

        # --- Measurements Save ---
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

        self.save_button = tk.Button(meas_frame, text="Save Measurements", command=self.save_measurements_to_json)
        self.save_button.pack(**pad_options)

        # Trigger the scrollregion calculation after buttons are created
        self.root.update_idletasks()
        self._on_button_frame_configure()


    def bind_events(self):
        self.image_canvas.bind("<Button-1>", self.on_press)
        self.image_canvas.bind("<B1-Motion>", self.on_motion)
        self.image_canvas.bind("<ButtonRelease-1>", self.on_release)

        # Use platform-specific mouse wheel binding for IMAGE CANVAS ZOOMING
        if self.root.tk.call('tk', 'windowingsystem') == 'aqua': # macOS
            self.image_canvas.bind("<MouseWheel>", self.zoom_mouse) # For trackpad pinch
            self.image_canvas.bind("<Button-4>", self.zoom_mouse) # For mouse wheel scroll up
            self.image_canvas.bind("<Button-5>", self.zoom_mouse) # For mouse wheel scroll down
        elif self.root.tk.call('tk', 'windowingsystem') == 'x11': # Linux
             self.image_canvas.bind("<Button-4>", self.zoom_mouse)
             self.image_canvas.bind("<Button-5>", self.zoom_mouse)
        else: # Windows
            self.image_canvas.bind("<MouseWheel>", self.zoom_mouse)
        # NOTE: Scrolling for the button panel is handled via Enter/Leave bindings

        self.root.bind("<plus>", self.zoom_in_center) # Zoom towards center
        self.root.bind("<minus>", self.zoom_out_center) # Zoom out from center
        self.root.bind("<KeyPress-Right>", self.next_image)
        self.root.bind("<KeyPress-Left>", self.prev_image)
        self.root.bind("<Control-z>", self.undo)
        self.root.bind("<Control-y>", self.redo)
        self.image_canvas.bind("<Motion>", self.update_zoom_box_and_pixel) # Combined update

        # Bind canvas resizing to update scroll region (Keep this)
        self.image_canvas.bind("<Configure>", self.on_canvas_resize)

    def on_canvas_resize(self, event=None):
         """Callback when the image canvas is resized."""
         self.display_image() # Re-display to fit potentially new canvas size

    def update_dot_coords_display(self):
        """Updates the text box showing coordinates and measurements."""
        if not hasattr(self, 'dot_coords_text') or not self.dot_coords_text.winfo_exists():
            return # Avoid errors if widget doesn't exist yet

        text = ""
        if self.calibration_dots:
            text += "--- Calibration ---\n"
            for i, (x, y) in enumerate(self.calibration_dots):
                text += f"  Dot {i+1}: ({x:.1f}, {y:.1f})\n"
            if len(self.calibration_dots) >= 2 and self.calibration_done:
                 dist_px = math.sqrt((self.calibration_dots[1][0] - self.calibration_dots[0][0])**2 +
                                     (self.calibration_dots[1][1] - self.calibration_dots[0][1])**2)
                 # Find the calibration entry to get the real distance entered
                 calib_entry = next((m for m in self.measurements if m.get("type") == "calibration" and len(m.get("points", []))==2 and m["points"][0]==self.calibration_dots[0]), None)
                 real_dist = calib_entry.get("real_value_mm", 0) if calib_entry else (dist_px / self.calibration_factor)
                 text += f"  -> Dist: {dist_px:.2f}px = {real_dist:.2f}mm (Factor: {self.calibration_factor:.4f} px/mm)\n"
            elif len(self.calibration_dots) >= 2:
                 dist_px = math.sqrt((self.calibration_dots[1][0] - self.calibration_dots[0][0])**2 +
                                     (self.calibration_dots[1][1] - self.calibration_dots[0][1])**2)
                 text += f"  -> Dist: {dist_px:.2f}px (Pending Calibration)\n"


        if self.artery_dots:
            text += "\n--- Dots Mode Measurements ---\n"
            pair_count = 1
            i = 0
            while i < len(self.artery_dots):
                if i + 1 < len(self.artery_dots):
                    x1, y1 = self.artery_dots[i]
                    x2, y2 = self.artery_dots[i+1]
                    dx = x2 - x1
                    dy = y2 - y1
                    dist_px = math.sqrt(dx**2 + dy**2)
                    # Calculate angle relative to positive X-axis
                    angle = math.degrees(math.atan2(-dy, dx)) # Use -dy because Y increases downwards
                    if angle < 0: angle += 360 # Normalize to 0-360

                    text += f"  Pair {pair_count}: ({x1:.1f},{y1:.1f}) -> ({x2:.1f},{y2:.1f})\n"
                    if self.calibration_done:
                        dist_mm = dist_px / self.calibration_factor
                        text += f"    Dist: {dist_px:.2f}px = {dist_mm:.3f}mm | Angle: {angle:.1f}°\n"
                    else:
                        text += f"    Dist: {dist_px:.2f}px | Angle: {angle:.1f}° (Uncalibrated)\n"
                    i += 2
                else:
                    # Single unpaired dot
                    x, y = self.artery_dots[i]
                    text += f"  Pair {pair_count} (Pending): ({x:.1f}, {y:.1f})\n"
                    i += 1
                pair_count += 1

        if self.line_points:
            text += "\n--- Line Mode Points ---\n"
            for idx, (x, y) in enumerate(self.line_points):
                 text += f"  Point {idx+1}: ({x:.1f}, {y:.1f})\n"
            if len(self.line_points) == 4:
                 text += "  (Ready for 'Reset Lines' or new mode)\n"

        if self.angle_points:
            text += "\n--- Angle Mode Points ---\n"
            for idx, (x, y) in enumerate(self.angle_points):
                 text += f"  Point {idx+1}: ({x:.1f}, {y:.1f})\n"


        self.dot_coords_text.config(state=tk.NORMAL)
        self.dot_coords_text.delete("1.0", tk.END)
        self.dot_coords_text.insert(tk.END, text if text else "No points placed yet.")
        self.dot_coords_text.config(state=tk.DISABLED)
        self.dot_coords_text.yview_moveto(1.0) # Scroll to end

    def reset_image_state(self, reset_zoom=True):
        """Resets most state variables associated with the current image."""
        if reset_zoom:
            self.zoom_factor = 1.0 # Reset zoom to 100%
        self._reset_all_modes()
        self.img_filtered = None
        self.calibration_dots = []
        self.artery_dots = []
        self.line_points = []
        self.measurements = []
        self.line_measurements = []
        self.line_measurement_points = []
        self.calibration_done = False
        self.calibration_factor = 1.0
        self.edge_detection_active = False # Reset legacy filter
        self.global_canny_active = False # Reset global canny
        self.selection_rect = None
        self.selection_start = None
        self.selection_end = None
        self.canny_rect = None      # Reset Canny selection
        self.canny_start = None
        self.canny_end = None
        self.angle_points = []
        self.photo = None # Clear image references
        self.zoom_box_photo = None

        # Reset button states that might be sunken
        if "Zoom In Box" in self.buttons: self.buttons["Zoom In Box"].config(relief=tk.RAISED)
        if "Global Canny" in self.buttons: self.buttons["Global Canny"].config(relief=tk.RAISED)


        # Reset UI elements
        self.update_dot_coords_display()
        self.update_tables()
        self.measurement.set("Status: Ready" if self.img_original else "Status: Load Image")
        self.pixel_info.set("Mode: None | Pixel: | Zoom: OFF") # Reset pixel info string format

        # Update zoom box content if active (will likely just clear it)
        if self.zoom_box_mode and self.zoom_box:
            self.update_zoom_box_content(None) # Pass None for event


    def load_image(self):
        """Loads an image file and prepares the application state."""
        file_path = filedialog.askopenfilename(
            title="Select Image File",
            filetypes=[("Image Files", "*.png;*.jpg;*.jpeg;*.bmp;*.gif;*.tif;*.tiff"), ("All Files", "*.*")]
        )
        if not file_path:
            return # User cancelled

        try:
            # Check if it's a valid image file before proceeding
            img_test = Image.open(file_path)
            img_test.verify() # Verify checks integrity without loading full data
            img_test.close() # Close the test image

            self.file_path = file_path
            folder = os.path.dirname(file_path)
            try:
                self.image_files = sorted([
                    f for f in os.listdir(folder)
                    if f.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp', '.gif', '.tif', '.tiff'))
                       and os.path.isfile(os.path.join(folder, f)) # Ensure it's a file
                ])
                self.current_index = self.image_files.index(os.path.basename(file_path))
            except ValueError: # If file not found in listing (e.g. restricted folder)
                self.image_files = [os.path.basename(file_path)]
                self.current_index = 0
            except OSError: # Handle cases like restricted folder access
                 self.image_files = [os.path.basename(file_path)]
                 self.current_index = 0
                 messagebox.showwarning("Folder Access", "Could not read image folder content. Navigation disabled.")


            self.path_text.set(f"Path: {os.path.basename(file_path)}") # Show only filename
            self.img_original = Image.open(file_path).convert("RGBA") # Convert to RGBA for consistency
            self.reset_image_state(reset_zoom=True) # Full reset for new image
            self.display_image()

            # Keep zoom box state as it was (on or off)
            if self.zoom_box_mode and self.zoom_box:
                 try:
                     self.zoom_box.place(in_=self.image_frame, anchor='se', relx=1.0, rely=1.0, x=-10, y=-10)
                     self.update_zoom_box_content(None) # Update content
                 except tk.TclError as e:
                      print(f"Error re-placing zoom box: {e}")
                      self.zoom_box_mode = False # Turn off if error
                      self.buttons["Zoom In Box"].config(relief=tk.RAISED)


        except FileNotFoundError:
            messagebox.showerror("Error", f"File not found:\n{file_path}")
            self.file_path = None
        except UnidentifiedImageError:
             messagebox.showerror("Error", f"Cannot identify image file:\n{file_path}\nMay be corrupted or unsupported format.")
             self.file_path = None
        except Exception as e:
            messagebox.showerror("Error Loading Image", f"An unexpected error occurred:\n{e}")
            print(traceback.format_exc()) # Print detailed traceback to console
            self.file_path = None
            self.img_original = None
            self.reset_image_state(reset_zoom=True)
            self.display_image() # Display empty canvas


    def change_image(self, direction):
        """Changes to the next or previous image in the folder."""
        if not self.file_path or not self.image_files or len(self.image_files) < 2:
             self.measurement.set("Status: No other images in folder.")
             return

        original_index = self.current_index
        num_files = len(self.image_files)
        attempt = 0

        while attempt < num_files:
            attempt += 1
            if direction == "next":
                next_idx = (original_index + attempt) % num_files
            elif direction == "previous":
                next_idx = (original_index - attempt + num_files) % num_files
            else:
                return # Should not happen

            if next_idx == original_index and attempt > 0: # Wrapped around completely
                 break # Avoid infinite loop if only one valid file

            new_file_name = self.image_files[next_idx]
            new_file_path = os.path.join(os.path.dirname(self.file_path), new_file_name)

            try:
                # Verify the next image before loading fully
                img_test = Image.open(new_file_path)
                img_test.verify()
                img_test.close()

                # Successfully verified, load it
                self.current_index = next_idx # Update index only on success
                self.file_path = new_file_path
                self.path_text.set(f"Path: {new_file_name}")
                self.img_original = Image.open(self.file_path).convert("RGBA") # Load and convert
                # Reset state but keep zoom level
                self.reset_image_state(reset_zoom=False)
                self.display_image()
                # Update zoom box content if active
                if self.zoom_box_mode and self.zoom_box:
                     self.update_zoom_box_content(None)
                return # Success, exit loop

            except (FileNotFoundError, UnidentifiedImageError, OSError) as e:
                print(f"Skipping file '{new_file_name}': {e}")
                # Continue loop to try the next file

            except Exception as e:
                messagebox.showerror("Error Changing Image", f"An unexpected error occurred loading '{new_file_name}':\n{e}")
                print(traceback.format_exc())
                # Stop on unexpected error, keep current image
                return

        # If loop finishes without success
        messagebox.showinfo("Image Navigation", "No other valid images found in the folder.")


    def next_image(self, event=None):
        """Event handler for next image."""
        self.change_image("next")

    def prev_image(self, event=None):
        """Event handler for previous image."""
        self.change_image("previous")

    def apply_filters_and_display(self, *args):
        """Applies selected filters (Global Canny OR ROI Canny) and then calls display_image."""
        if not self.img_original:
            return

        # Start with the original image
        img_to_process = self.img_original.copy()
        filter_applied = False
        processed_image = None # Will hold the result of filtering

        # --- Apply Global Canny FIRST ---
        if self.global_canny_active:
            try:
                # Convert to grayscale numpy array
                img_np_rgb = np.array(img_to_process.convert("RGB"))
                img_np_gray = cv2.cvtColor(img_np_rgb, cv2.COLOR_RGB2GRAY)

                # Apply Canny
                edges_np = cv2.Canny(img_np_gray, self.canny_low.get(), self.canny_high.get())

                # Convert grayscale edges back to RGBA PIL Image
                processed_image = Image.fromarray(edges_np).convert("RGBA")
                filter_applied = True
                # Update status only if slider change isn't causing it
                if not args: # args is empty if called directly, not by slider trace
                    self.measurement.set(f"Status: Global Canny Filter ON (Thresh: {self.canny_low.get()}/{self.canny_high.get()}).")

            except Exception as e:
                print(f"Error applying Global Canny: {e}")
                print(traceback.format_exc())
                processed_image = img_to_process # Fallback to original on error
                filter_applied = False
                self.measurement.set("Status: Error applying Global Canny.")

        # --- Apply ROI Canny ONLY if Global is OFF and ROI is defined ---
        elif self.canny_start and self.canny_end:
            # Make a copy to paste onto if applying ROI filter
            processed_image = img_to_process.copy() # Start with original for ROI paste

            # Convert canvas coords to original image coords
            x1_orig = int(min(self.canny_start[0], self.canny_end[0]) / self.zoom_factor)
            y1_orig = int(min(self.canny_start[1], self.canny_end[1]) / self.zoom_factor)
            x2_orig = int(max(self.canny_start[0], self.canny_end[0]) / self.zoom_factor)
            y2_orig = int(max(self.canny_start[1], self.canny_end[1]) / self.zoom_factor)

            # Clamp coordinates to image bounds
            x1_orig = max(0, x1_orig)
            y1_orig = max(0, y1_orig)
            x2_orig = min(img_to_process.width, x2_orig)
            y2_orig = min(img_to_process.height, y2_orig)

            if x2_orig > x1_orig and y2_orig > y1_orig: # Check for valid region
                try:
                    # Crop the region from the original image for processing
                    cropped_pil = img_to_process.crop((x1_orig, y1_orig, x2_orig, y2_orig))

                    # Convert cropped PIL image to NumPy array -> Grayscale
                    cropped_np_rgb = np.array(cropped_pil.convert("RGB"))
                    cropped_np_gray = cv2.cvtColor(cropped_np_rgb, cv2.COLOR_RGB2GRAY)

                    # Apply Canny edge detection
                    edges_np = cv2.Canny(cropped_np_gray, self.canny_low.get(), self.canny_high.get())

                    # Create a mask from edges (white edges, black background)
                    mask = Image.fromarray(edges_np).convert("L")

                    # Create colored overlay (green edges)
                    colored_edges = Image.new("RGBA", mask.size, (0, 255, 0, 255)) # Green edges

                    # Paste the colored edges onto the processed_image copy using the mask
                    if processed_image.mode != 'RGBA':
                         processed_image = processed_image.convert('RGBA')
                    processed_image.paste(colored_edges, (x1_orig, y1_orig), mask=mask)

                    filter_applied = True
                    # Update status only if ROI selection isn't actively happening
                    if not self.canny_selection_mode and not args:
                        self.measurement.set(f"Status: Canny filter applied to ROI (Thresh: {self.canny_low.get()}/{self.canny_high.get()}).")

                except Exception as e:
                    print(f"Error applying ROI Canny: {e}")
                    print(traceback.format_exc())
                    # Keep processed_image as the original copy
                    filter_applied = False
                    self.measurement.set("Status: Error applying ROI Canny.")
            else:
                # ROI defined but has zero area, treat as no filter applied
                 processed_image = img_to_process
                 filter_applied = False

        # --- Update the filtered image attribute ---
        # If a filter was applied, store the result, otherwise clear img_filtered
        self.img_filtered = processed_image if filter_applied else None

        # --- Display Result ---
        self.display_image()
        # Update zoom box content as well
        if self.zoom_box_mode and self.zoom_box:
            self.update_zoom_box_content(None)


    def display_image(self):
        """Displays the current image (original or filtered) on the canvas with overlays."""
        # --- Safeguard ---
        if not self.root or not self.root.winfo_exists() or not self.image_canvas or not self.image_canvas.winfo_exists():
            # print("Debug: display_image called too early or widgets destroyed.")
            return
        # --- End Safeguard ---

        if not self.img_original:
            self.image_canvas.delete("all")
            self.image_canvas.config(scrollregion=(0, 0, 1, 1)) # Reset scroll
            try:
                canvas_width = self.image_canvas.winfo_width()
                canvas_height = self.image_canvas.winfo_height()
                if canvas_width > 1 and canvas_height > 1:
                    self.image_canvas.create_text(
                        canvas_width / 2, canvas_height / 2,
                        text="No Image Loaded", fill="white", font=("Arial", 16)
                    )
            except tk.TclError: pass
            return

        # Use the filtered image if available, otherwise the original
        img_display_base = self.img_filtered if self.img_filtered is not None else self.img_original

        # Resize the base image for display
        width, height = img_display_base.size
        new_width = int(width * self.zoom_factor)
        new_height = int(height * self.zoom_factor)

        # Prevent zero size errors if zoom factor is too small
        if new_width <= 0 or new_height <= 0:
            return

        try:
            resized_img = img_display_base.resize((new_width, new_height), Image.Resampling.LANCZOS)
            if self.root and self.root.winfo_exists():
                self.photo = ImageTk.PhotoImage(resized_img) # Store reference
            else:
                # print("Debug: Root window not ready for PhotoImage creation.")
                return # Cannot proceed
        except tk.TclError as e:
             # print(f"Tkinter Error resizing/creating PhotoImage: {e}")
             return # Stop if image cannot be created
        except Exception as e:
             print(f"Error resizing image: {e}")
             try:
                 resized_img = img_display_base.resize((new_width, new_height), Image.Resampling.NEAREST)
                 if self.root and self.root.winfo_exists():
                     self.photo = ImageTk.PhotoImage(resized_img)
                 else:
                      # print("Debug: Root window not ready for PhotoImage creation (Nearest).")
                      return
             except tk.TclError as e_near:
                 # print(f"Tkinter Error resizing/creating PhotoImage (Nearest): {e_near}")
                 return
             except Exception as e_near_gen:
                 print(f"Failed to resize even with NEAREST: {e_near_gen}")
                 return # Cannot display

        # Clear previous drawings
        self.image_canvas.delete("all")

        # Draw the image
        if self.photo:
            self.image_canvas.create_image(0, 0, anchor=tk.NW, image=self.photo)
            self.image_canvas.config(scrollregion=(0, 0, new_width, new_height))
        else:
             # print("Error: self.photo is None, cannot draw image.")
             return # Skip drawing overlays if image failed

        # --- Draw Overlays ---
        def scale_pt(pt):
            return (pt[0] * self.zoom_factor, pt[1] * self.zoom_factor)

        dot_radius = 3
        line_width = 2

        # Draw calibration dots
        for dot in self.calibration_dots:
            sx, sy = scale_pt(dot)
            self.image_canvas.create_oval(sx - dot_radius, sy - dot_radius, sx + dot_radius, sy + dot_radius, fill="cyan", outline="black")

        # Draw artery lines and dots
        for i in range(0, len(self.artery_dots)):
            sx, sy = scale_pt(self.artery_dots[i])
            self.image_canvas.create_oval(sx - dot_radius, sy - dot_radius, sx + dot_radius, sy + dot_radius, fill="yellow", outline="black")
            if i % 2 == 1:
                sx_prev, sy_prev = scale_pt(self.artery_dots[i-1])
                self.image_canvas.create_line(sx_prev, sy_prev, sx, sy, fill="yellow", width=line_width)

        # Draw line mode lines and points (PERSISTENT)
        for i in range(0, len(self.line_points)):
            sx, sy = scale_pt(self.line_points[i])
            self.image_canvas.create_oval(sx - dot_radius, sy - dot_radius, sx + dot_radius, sy + dot_radius, fill="magenta", outline="black")
            if i % 2 == 1: # Draw line for pairs
                sx_prev, sy_prev = scale_pt(self.line_points[i-1])
                self.image_canvas.create_line(sx_prev, sy_prev, sx, sy, fill="magenta", width=line_width)

        # Draw angle points and lines
        if self.angle_points:
             sp1_a, sp2_a, sp3_a = None, None, None # Use temp vars
             if len(self.angle_points) >= 1:
                 sp1_a = scale_pt(self.angle_points[0])
                 self.image_canvas.create_oval(sp1_a[0]-dot_radius, sp1_a[1]-dot_radius, sp1_a[0]+dot_radius, sp1_a[1]+dot_radius, fill="lime green", outline="black")
             if len(self.angle_points) >= 2:
                 sp2_a = scale_pt(self.angle_points[1])
                 self.image_canvas.create_oval(sp2_a[0]-dot_radius, sp2_a[1]-dot_radius, sp2_a[0]+dot_radius, sp2_a[1]+dot_radius, fill="lime green", outline="black")
                 if sp1_a: self.image_canvas.create_line(sp1_a[0], sp1_a[1], sp2_a[0], sp2_a[1], fill="lime green", width=line_width, dash=(4, 2))
             if len(self.angle_points) == 3:
                 sp3_a = scale_pt(self.angle_points[2])
                 self.image_canvas.create_oval(sp3_a[0]-dot_radius, sp3_a[1]-dot_radius, sp3_a[0]+dot_radius, sp3_a[1]+dot_radius, fill="lime green", outline="black")
                 if sp2_a: self.image_canvas.create_line(sp2_a[0], sp2_a[1], sp3_a[0], sp3_a[1], fill="lime green", width=line_width, dash=(4, 2))


        # Draw the tick markers for Line Mode measurements if available
        tick_radius = 2
        for pt1, pt2 in self.line_measurement_points:
            sx1, sy1 = scale_pt(pt1)
            sx2, sy2 = scale_pt(pt2)
            self.image_canvas.create_oval(sx1 - tick_radius, sy1 - tick_radius, sx1 + tick_radius, sy1 + tick_radius, fill="red", outline="red")
            self.image_canvas.create_line(sx1, sy1, sx2, sy2, fill="red", dash=(2, 2))

        # --- Draw Selection Rectangles ---
        # Draw completed Canny rectangle if selection is done
        if self.canny_start and self.canny_end and not self.canny_selection_mode:
            self.image_canvas.delete("canny_rect") # Delete potential old one during drag
            self.image_canvas.create_rectangle(
                self.canny_start[0], self.canny_start[1],
                self.canny_end[0], self.canny_end[1],
                outline="blue", dash=(4, 4), width=1, tags="canny_rect"
            )
        elif not self.canny_selection_mode:
             self.image_canvas.delete("canny_rect")

        # Draw completed FIND_EDGES rectangle (legacy)
        if self.selection_start and self.selection_end and not self.edge_selection_mode:
             self.image_canvas.delete("selection_rect")
             self.image_canvas.create_rectangle(
                 self.selection_start[0], self.selection_start[1],
                 self.selection_end[0], self.selection_end[1],
                 outline="red", dash=(4, 4), width=1, tags="selection_rect"
             )
        elif not self.edge_selection_mode:
             self.image_canvas.delete("selection_rect")


    def update_zoom_box_and_pixel(self, event=None):
         """Updates pixel info and zoom box based on mouse position."""
         if not self.img_original or not event or not self.image_canvas or not self.image_canvas.winfo_exists():
             return

         try:
             canvas_x = self.image_canvas.canvasx(event.x)
             canvas_y = self.image_canvas.canvasy(event.y)
             orig_x = int(canvas_x / self.zoom_factor)
             orig_y = int(canvas_y / self.zoom_factor)

             # Update Pixel Info Label
             pixel_str_part = "Pixel:" # Default
             if 0 <= orig_x < self.img_original.width and 0 <= orig_y < self.img_original.height:
                 try:
                     pixel_value = self.img_original.getpixel((orig_x, orig_y))
                     if isinstance(pixel_value, tuple): # RGBA or RGB
                         pixel_str = f"RGB:({pixel_value[0]},{pixel_value[1]},{pixel_value[2]})"
                         if len(pixel_value) == 4: pixel_str += f" A:{pixel_value[3]}"
                     else: # Grayscale
                         pixel_str = f"Gray:{pixel_value}"
                     pixel_str_part = f"Pixel @ ({orig_x}, {orig_y}): {pixel_str}"
                 except Exception:
                     pixel_str_part = f"Pixel @ ({orig_x}, {orig_y}): Error"
             else:
                 pixel_str_part = "Pixel: Outside Image"

             # Combine with current mode info safely
             try:
                 current_info = self.pixel_info.get()
                 parts = current_info.split('|')
                 mode_info = parts[0].strip() if parts else "Mode: Unknown"
                 zoom_info = parts[-1].strip() if parts else "Zoom: ?"
                 self.pixel_info.set(f"{mode_info} | {pixel_str_part} | {zoom_info}")
             except Exception: # Fallback if pixel_info string is unexpected
                 self.pixel_info.set(f"Mode: Unknown | {pixel_str_part} | Zoom: {'ON' if self.zoom_box_mode else 'OFF'}")


             # Update Zoom Box if active
             if self.zoom_box_mode and self.zoom_box and self.zoom_box.winfo_exists():
                 self.update_zoom_box_content(event)

         except tk.TclError:
             pass # Handle potential errors if canvas is destroyed during motion event


    def update_zoom_box_content(self, event=None):
        """Updates the content of the zoom box canvas."""
        # --- Safeguard ---
        if not self.root or not self.root.winfo_exists() or not self.zoom_box or not self.zoom_box.winfo_exists():
            return # Cannot update if widgets aren't ready
        # --- End Safeguard ---

        if not self.zoom_box_mode or not self.img_original:
            return

        try:
            # Determine center point for zoom box
            if event:
                canvas_x = self.image_canvas.canvasx(event.x)
                canvas_y = self.image_canvas.canvasy(event.y)
                orig_x = int(canvas_x / self.zoom_factor)
                orig_y = int(canvas_y / self.zoom_factor)
            else:
                # If no event, center on the canvas view's center (approx)
                canvas_width = self.image_canvas.winfo_width()
                canvas_height = self.image_canvas.winfo_height()
                scroll_x = self.image_canvas.canvasx(0) # Get current view top-left
                scroll_y = self.image_canvas.canvasy(0)
                center_canvas_x = scroll_x + canvas_width / 2
                center_canvas_y = scroll_y + canvas_height / 2
                orig_x = int(center_canvas_x / self.zoom_factor)
                orig_y = int(center_canvas_y / self.zoom_factor)

            # Calculate the region in the original image to crop
            crop_width_orig = self.ZOOM_BOX_SIZE / self.ZOOM_BOX_FACTOR
            crop_height_orig = self.ZOOM_BOX_SIZE / self.ZOOM_BOX_FACTOR

            left = int(orig_x - crop_width_orig / 2)
            top = int(orig_y - crop_height_orig / 2)
            right = int(left + crop_width_orig)
            bottom = int(top + crop_height_orig)

            # Clamp coordinates to image boundaries
            left = max(0, left)
            top = max(0, top)
            right = min(self.img_original.width, right)
            bottom = min(self.img_original.height, bottom)

            # Ensure valid crop dimensions
            if right <= left or bottom <= top:
                self.zoom_box.delete("all")
                self.zoom_box.create_text(self.ZOOM_BOX_SIZE / 2, self.ZOOM_BOX_SIZE / 2, text="Invalid Area", fill="red")
                return

            # Use filtered image if available, else original
            img_source = self.img_filtered if self.img_filtered is not None else self.img_original

            # Crop the calculated region
            cropped_image = img_source.crop((left, top, right, bottom))

            # Resize the cropped region to fit the zoom box
            zoomed = cropped_image.resize((self.ZOOM_BOX_SIZE, self.ZOOM_BOX_SIZE), Image.Resampling.NEAREST) # Use NEAREST for sharp pixels
            # Check Tkinter is ready before creating PhotoImage
            if self.root and self.root.winfo_exists():
                 self.zoom_box_photo = ImageTk.PhotoImage(zoomed) # Store reference
            else:
                # print("Debug: Root window not ready for zoom_box_photo creation.")
                return # Cannot proceed

            # Display the zoomed image in the zoom box
            self.zoom_box.delete("all")
            self.zoom_box.create_image(0, 0, anchor=tk.NW, image=self.zoom_box_photo)

            # --- Draw Overlays in Zoom Box ---
            dot_radius_zoom = 2
            line_width_zoom = 1

            def scale_to_zoom(pt_orig):
                 if left <= pt_orig[0] < right and top <= pt_orig[1] < bottom:
                     zoom_x = (pt_orig[0] - left) * self.ZOOM_BOX_FACTOR
                     zoom_y = (pt_orig[1] - top) * self.ZOOM_BOX_FACTOR
                     return zoom_x, zoom_y
                 return None

            # Calibration dots
            for dot in self.calibration_dots:
                 sp = scale_to_zoom(dot)
                 if sp: self.zoom_box.create_oval(sp[0]-dot_radius_zoom, sp[1]-dot_radius_zoom, sp[0]+dot_radius_zoom, sp[1]+dot_radius_zoom, fill="cyan", outline="black")

            # Artery dots and lines
            for i in range(0, len(self.artery_dots)):
                 sp = scale_to_zoom(self.artery_dots[i])
                 if sp:
                     self.zoom_box.create_oval(sp[0]-dot_radius_zoom, sp[1]-dot_radius_zoom, sp[0]+dot_radius_zoom, sp[1]+dot_radius_zoom, fill="yellow", outline="black")
                     if i % 2 == 1:
                         sp_prev = scale_to_zoom(self.artery_dots[i-1])
                         if sp_prev: self.zoom_box.create_line(sp_prev[0], sp_prev[1], sp[0], sp[1], fill="yellow", width=line_width_zoom)

            # Line mode points and lines
            for i in range(0, len(self.line_points)):
                 sp = scale_to_zoom(self.line_points[i])
                 if sp:
                     self.zoom_box.create_oval(sp[0]-dot_radius_zoom, sp[1]-dot_radius_zoom, sp[0]+dot_radius_zoom, sp[1]+dot_radius_zoom, fill="magenta", outline="black")
                     if i % 2 == 1:
                         sp_prev = scale_to_zoom(self.line_points[i-1])
                         if sp_prev: self.zoom_box.create_line(sp_prev[0], sp_prev[1], sp[0], sp[1], fill="magenta", width=line_width_zoom)

            # Angle points and lines
            if self.angle_points:
                sp1_z, sp2_z, sp3_z = None, None, None
                if len(self.angle_points) >= 1: sp1_z = scale_to_zoom(self.angle_points[0])
                if len(self.angle_points) >= 2: sp2_z = scale_to_zoom(self.angle_points[1])
                if len(self.angle_points) >= 3: sp3_z = scale_to_zoom(self.angle_points[2])

                if sp1_z: self.zoom_box.create_oval(sp1_z[0]-dot_radius_zoom, sp1_z[1]-dot_radius_zoom, sp1_z[0]+dot_radius_zoom, sp1_z[1]+dot_radius_zoom, fill="lime green", outline="black")
                if sp2_z:
                    self.zoom_box.create_oval(sp2_z[0]-dot_radius_zoom, sp2_z[1]-dot_radius_zoom, sp2_z[0]+dot_radius_zoom, sp2_z[1]+dot_radius_zoom, fill="lime green", outline="black")
                    if sp1_z: self.zoom_box.create_line(sp1_z[0], sp1_z[1], sp2_z[0], sp2_z[1], fill="lime green", width=line_width_zoom, dash=(3,1))
                if sp3_z:
                    self.zoom_box.create_oval(sp3_z[0]-dot_radius_zoom, sp3_z[1]-dot_radius_zoom, sp3_z[0]+dot_radius_zoom, sp3_z[1]+dot_radius_zoom, fill="lime green", outline="black")
                    if sp2_z: self.zoom_box.create_line(sp2_z[0], sp2_z[1], sp3_z[0], sp3_z[1], fill="lime green", width=line_width_zoom, dash=(3,1))

            # Draw crosshair at the center
            center = self.ZOOM_BOX_SIZE / 2
            offset = 5
            lw = 1
            self.zoom_box.create_oval(center-offset, center-offset, center+offset, center+offset, outline="red", width=lw)
            self.zoom_box.create_line(center, center-offset*0.6, center, center+offset*0.6, fill="red", width=lw)
            self.zoom_box.create_line(center-offset*0.6, center, center+offset*0.6, center, fill="red", width=lw)

        except tk.TclError as e:
             # print(f"Tkinter Error during zoom box update: {e}")
             try:
                 self.zoom_box.delete("all")
                 self.zoom_box.create_text(self.ZOOM_BOX_SIZE / 2, self.ZOOM_BOX_SIZE / 2, text="Tk Err", fill="red")
             except tk.TclError: pass # Ignore if zoom_box itself is destroyed
        except ValueError as e:
             # print(f"ValueError during zoom box update: {e}")
             try:
                 self.zoom_box.delete("all")
                 self.zoom_box.create_text(self.ZOOM_BOX_SIZE / 2, self.ZOOM_BOX_SIZE / 2, text="Size Err", fill="red")
             except tk.TclError: pass
        except Exception as e:
             print(f"Error updating zoom box content: {e}")
             print(traceback.format_exc())
             try:
                 self.zoom_box.delete("all")
                 self.zoom_box.create_text(self.ZOOM_BOX_SIZE / 2, self.ZOOM_BOX_SIZE / 2, text="Error", fill="red")
             except tk.TclError: pass


    def zoom(self, factor, event=None):
        """Zooms the image view by a given factor, optionally centering on the event coordinates."""
        if not self.img_original:
            return

        old_zoom = self.zoom_factor
        new_zoom = max(0.05, min(old_zoom * factor, 50.0))
        if abs(new_zoom - old_zoom) < 0.001:
             return

        # Get mouse position relative to canvas content (before zoom)
        mouse_x, mouse_y = 0, 0 # Defaults
        target_canvas_x, target_canvas_y = 0, 0 # Defaults
        valid_coords = False
        if self.image_canvas and self.image_canvas.winfo_exists():
            try:
                canvas_width = self.image_canvas.winfo_width()
                canvas_height = self.image_canvas.winfo_height()
                if event:
                     mouse_x = self.image_canvas.canvasx(event.x)
                     mouse_y = self.image_canvas.canvasy(event.y)
                     target_canvas_x = event.x
                     target_canvas_y = event.y
                     valid_coords = True
                elif canvas_width > 0 and canvas_height > 0: # Center zoom if no event
                     mouse_x = self.image_canvas.canvasx(canvas_width / 2)
                     mouse_y = self.image_canvas.canvasy(canvas_height / 2)
                     target_canvas_x = canvas_width / 2
                     target_canvas_y = canvas_height / 2
                     valid_coords = True
            except tk.TclError: # Handle case where canvas might be destroyed
                  return
        else: # No canvas exists
            return

        if not valid_coords: return # Cannot zoom without reference point

        # Calculate the image coordinate under the mouse *before* zooming
        img_coord_x = mouse_x / old_zoom
        img_coord_y = mouse_y / old_zoom

        # Apply the new zoom factor
        self.zoom_factor = new_zoom

        # Update the image display (this also updates scrollregion)
        self.display_image()

        # --- Recenter View ---
        if not self.image_canvas or not self.image_canvas.winfo_exists():
             return # Check if canvas still exists

        # Calculate where the same image coordinate should be *after* zooming
        new_mouse_x = img_coord_x * new_zoom
        new_mouse_y = img_coord_y * new_zoom

        # Calculate the required top-left corner of the view (scroll position)
        scroll_x = new_mouse_x - target_canvas_x
        scroll_y = new_mouse_y - target_canvas_y

        # Convert scroll position to fraction for xview_moveto/yview_moveto
        current_scroll_region = self.image_canvas.cget('scrollregion')
        img_width_new, img_height_new = 1, 1 # Fallbacks
        if current_scroll_region:
            try:
                sr_parts = current_scroll_region.split()
                img_width_new = float(sr_parts[2])
                img_height_new = float(sr_parts[3])
            except (IndexError, ValueError, tk.TclError): pass

        scroll_x_frac = scroll_x / img_width_new if img_width_new > 0 else 0
        scroll_y_frac = scroll_y / img_height_new if img_height_new > 0 else 0

        # Apply the scroll, clamping to valid range [0, 1]
        try:
            self.image_canvas.xview_moveto(max(0.0, min(scroll_x_frac, 1.0)))
            self.image_canvas.yview_moveto(max(0.0, min(scroll_y_frac, 1.0)))
        except tk.TclError:
             pass # Ignore errors if canvas is destroyed during zoom

        # Update zoom box content after zooming
        if self.zoom_box_mode and self.zoom_box:
            self.update_zoom_box_content(event) # Use event to center zoom box correctly


    def zoom_in(self, event=None):
        self.zoom(1.2, event)

    def zoom_out(self, event=None):
        self.zoom(1 / 1.2, event)

    def zoom_in_center(self, event=None):
         self.zoom(1.2, None) # Pass None for event

    def zoom_out_center(self, event=None):
        self.zoom(1 / 1.2, None) # Pass None for event


    def zoom_mouse(self, event):
        """Handles mouse wheel zooming ONLY on the image canvas."""
        if not self.image_canvas or not self.image_canvas.winfo_exists(): return

        # Check if the event occurred within the image canvas bounds
        try:
            widget = self.root.winfo_containing(event.x_root, event.y_root)
            is_over_image_canvas = False
            while widget is not None:
                if widget == self.image_canvas:
                    is_over_image_canvas = True
                    break
                # Prevent infinite loop if root is the widget
                if widget == self.root: break
                widget = widget.master
        except tk.TclError: # Handle if winfo_containing fails
             return

        if not is_over_image_canvas:
             return # Don't zoom if mouse isn't over the image canvas

        factor = 1.0
        if event.num == 4: factor = 1.2       # Linux scroll up
        elif event.num == 5: factor = 1 / 1.2 # Linux scroll down
        elif hasattr(event, 'delta'): # Windows/macOS
            if event.delta > 0: factor = 1.2
            elif event.delta < 0: factor = 1 / 1.2

        if factor != 1.0:
            self.zoom(factor, event)


    def on_press(self, event):
        """Handles mouse button press events on the canvas."""
        if not self.img_original or not self.image_canvas or not self.image_canvas.winfo_exists():
            return

        self.save_state() # Save state before modification
        canvas_x = self.image_canvas.canvasx(event.x)
        canvas_y = self.image_canvas.canvasy(event.y)
        orig_x = canvas_x / self.zoom_factor
        orig_y = canvas_y / self.zoom_factor

        if not (0 <= orig_x < self.img_original.width and 0 <= orig_y < self.img_original.height):
             self.measurement.set("Status: Click outside image bounds.")
             return

        current_mode_action = False

        if self.edge_selection_mode: # Legacy FIND_EDGES ROI start
            self.selection_start = (canvas_x, canvas_y)
            self.selection_end = (canvas_x, canvas_y)
            self.image_canvas.delete("selection_rect")
            self.selection_rect = self.image_canvas.create_rectangle(
                canvas_x, canvas_y, canvas_x, canvas_y,
                outline="red", dash=(4, 4), tags="selection_rect"
            )
            current_mode_action = True

        elif self.canny_selection_mode: # Canny ROI start
            self.canny_start = (canvas_x, canvas_y)
            self.canny_end = (canvas_x, canvas_y)
            self.image_canvas.delete("canny_rect")
            self.canny_rect = self.image_canvas.create_rectangle(
                canvas_x, canvas_y, canvas_x, canvas_y,
                outline="blue", dash=(4, 4), tags="canny_rect"
            )
            current_mode_action = True

        elif self.calibration_mode:
            if len(self.calibration_dots) < 2:
                self.calibration_dots.append((orig_x, orig_y))
                if len(self.calibration_dots) == 2:
                    p1 = self.calibration_dots[0]
                    p2 = self.calibration_dots[1]
                    distance_px = math.sqrt((p2[0] - p1[0])**2 + (p2[1] - p1[1])**2)
                    self.prompt_for_calibration(distance_px)
                else:
                    self.measurement.set("Calibration: Click second point.")
            else:
                self.measurement.set("Calibration: Reset to add new points.")
            self.update_dot_coords_display()
            current_mode_action = True

        elif self.artery_mode:
            self.artery_dots.append((orig_x, orig_y))
            self.update_dot_coords_display()
            if len(self.artery_dots) % 2 == 0:
                dot1_x, dot1_y = self.artery_dots[-2]
                dot2_x, dot2_y = self.artery_dots[-1]
                dx = dot2_x - dot1_x
                dy = dot2_y - dot1_y
                distance_px = math.sqrt(dx**2 + dy**2)
                angle = math.degrees(math.atan2(-dy, dx))
                if angle < 0: angle += 360

                meas_info = {
                    "type": "artery",
                    "points": [(dot1_x, dot1_y), (dot2_x, dot2_y)],
                    "distance_px": distance_px,
                    "angle_deg": angle
                }
                status_text = f"Pair {len(self.artery_dots)//2}: {distance_px:.2f}px, {angle:.1f}°"
                if self.calibration_done:
                    real_distance = distance_px / self.calibration_factor
                    meas_info["distance_mm"] = real_distance
                    status_text += f" = {real_distance:.3f}mm"
                else:
                     status_text += " (Uncalibrated)"

                self.measurements.append(meas_info)
                self.measurement.set(status_text)
                self.update_tables()
            else:
                self.measurement.set(f"Pair {len(self.artery_dots)//2 + 1}: Click second point.")
            current_mode_action = True

        elif self.angle_mode:
            if len(self.angle_points) < 3:
                self.angle_points.append((orig_x, orig_y))
                self.update_dot_coords_display()
                if len(self.angle_points) == 1:
                    self.measurement.set("Angle: Click vertex point (2nd).")
                elif len(self.angle_points) == 2:
                     self.measurement.set("Angle: Click final point (3rd).")
                elif len(self.angle_points) == 3:
                    p1, p2, p3 = self.angle_points
                    v1 = (p1[0] - p2[0], p1[1] - p2[1])
                    v2 = (p3[0] - p2[0], p3[1] - p2[1])
                    mag1_sq = v1[0]**2 + v1[1]**2
                    mag2_sq = v2[0]**2 + v2[1]**2

                    if mag1_sq < 1e-9 or mag2_sq < 1e-9:
                        angle_deg = 0.0
                        self.measurement.set("Angle Error: Points coincide.")
                    else:
                        angle1 = math.atan2(v1[1], v1[0])
                        angle2 = math.atan2(v2[1], v2[0])
                        angle_rad = angle2 - angle1
                        while angle_rad > math.pi: angle_rad -= 2 * math.pi
                        while angle_rad <= -math.pi: angle_rad += 2 * math.pi
                        angle_deg = abs(math.degrees(angle_rad))
                        if angle_deg > 180: angle_deg = 360 - angle_deg

                        self.measurement.set(f"Angle Measured: {angle_deg:.2f}°")
                        self.measurements.append({
                            "type": "angle",
                            "points": self.angle_points.copy(),
                            "angle_deg": angle_deg
                        })
                        self.update_tables()
                        self.angle_points = []
                        self.measurement.set("Angle: Click first point for new angle.")
            else:
                self.angle_points = [(orig_x, orig_y)]
                self.measurement.set("Angle: Click vertex point (2nd).")
                self.update_dot_coords_display()
            current_mode_action = True

        elif self.line_mode:
            if len(self.line_points) < 4:
                self.line_points.append((orig_x, orig_y))
                self.update_dot_coords_display()
                pts_needed = 4 - len(self.line_points)
                if pts_needed > 0:
                     self.measurement.set(f"Line Mode: Click {pts_needed} more point(s).")
                else: # 4 points placed
                    meas = self.calculate_line_measurements()
                    if meas:
                        self.measurements.append(meas)
                        self.update_tables()
                        avg_dist_px = sum(meas["distances_px"]) / len(meas["distances_px"]) if meas["distances_px"] else 0
                        status = f"Lines: L1={meas['length1_px']:.1f}px, L2={meas['length2_px']:.1f}px, Angle={meas['angle_deg']:.1f}°, AvgDist={avg_dist_px:.2f}px"
                        if self.calibration_done:
                             # Ensure mm distances are present before calculating average
                             if meas.get("distances_mm"):
                                 avg_dist_mm = sum(meas["distances_mm"]) / len(meas["distances_mm"])
                                 status = f"Lines: L1={meas['length1_mm']:.2f}mm, L2={meas['length2_mm']:.2f}mm, Angle={meas['angle_deg']:.1f}°, AvgDist={avg_dist_mm:.3f}mm"
                             else: # Should not happen if calculation worked, but safety check
                                 status += " (mm N/A)"


                        self.measurement.set(f"{status} (Click 'Reset Lines' for new measurement)")
                        # self.line_points = [] # POINTS ARE NOW PERSISTENT
                    else:
                        self.measurement.set("Line Mode Error: Calculation failed. Resetting.")
                        self.line_points = []
                        self.line_measurement_points = []
                        self.update_dot_coords_display()
            else: # User clicks again after 4 points are already placed
                 # Reset and start new line measurement
                 self.line_points = [(orig_x, orig_y)] # Start with the new click
                 self.line_measurements = []
                 self.line_measurement_points = []
                 # Remove the previous line measurement result if it exists
                 self.measurements = [m for m in self.measurements if m.get("type") != "line" or m.get("points") != self.line_points[:-1]] # Rough check
                 self.measurement.set("Line Mode: Reset. Click 3 more points for new line.")
                 self.update_dot_coords_display()
                 self.update_tables() # Update table after removing old line measurement
            current_mode_action = True

        # Final redraw after action
        self.display_image()
        if self.zoom_box_mode:
            self.update_zoom_box_content(event)


    def on_motion(self, event):
        """Handles mouse motion events on the canvas (dragging)."""
        if not self.image_canvas or not self.image_canvas.winfo_exists():
            return

        try:
            canvas_x = self.image_canvas.canvasx(event.x)
            canvas_y = self.image_canvas.canvasy(event.y)

            # --- Handle ROI Dragging ---
            if self.edge_selection_mode and self.selection_start:
                self.selection_end = (canvas_x, canvas_y)
                if self.selection_rect:
                    self.image_canvas.coords(self.selection_rect,
                                             self.selection_start[0], self.selection_start[1],
                                             canvas_x, canvas_y)
                self.update_zoom_box_and_pixel(event)
                return

            elif self.canny_selection_mode and self.canny_start:
                self.canny_end = (canvas_x, canvas_y)
                if self.canny_rect:
                    self.image_canvas.coords(self.canny_rect,
                                             self.canny_start[0], self.canny_start[1],
                                             canvas_x, canvas_y)
                # Apply filter live during drag and update display
                self.apply_filters_and_display()
                # No need to call update_zoom_box_and_pixel here, it's handled by apply_filters_and_display
                return

            # --- Default: Update Pixel Info and Zoom Box ---
            self.update_zoom_box_and_pixel(event)

        except tk.TclError:
            pass # Handle errors if canvas is destroyed during motion


    def on_release(self, event):
        """Handles mouse button release events on the canvas."""
        if not self.image_canvas or not self.image_canvas.winfo_exists():
            return

        try:
            canvas_x = self.image_canvas.canvasx(event.x)
            canvas_y = self.image_canvas.canvasy(event.y)

            # --- Finalize ROI Selection ---
            if self.edge_selection_mode and self.selection_start:
                self.selection_end = (canvas_x, canvas_y)
                if abs(self.selection_start[0] - self.selection_end[0]) < 1 or \
                   abs(self.selection_start[1] - self.selection_end[1]) < 1:
                    self.selection_start = None
                    self.selection_end = None
                    self.image_canvas.delete("selection_rect")
                    self.measurement.set("Status: ROI selection cancelled (zero size).")
                else:
                    x1 = min(self.selection_start[0], self.selection_end[0])
                    y1 = min(self.selection_start[1], self.selection_end[1])
                    x2 = max(self.selection_start[0], self.selection_end[0])
                    y2 = max(self.selection_start[1], self.selection_end[1])
                    self.selection_start = (x1, y1)
                    self.selection_end = (x2, y2)
                    self.measurement.set("Status: ROI selected for Edge Detection.")
                    self.image_canvas.delete("selection_rect") # Delete temp rect
                    self.image_canvas.create_rectangle(x1, y1, x2, y2, outline="red", dash=(4, 4), tags="selection_rect")

                self.edge_selection_mode = False
                if "ROI Selection" in self.buttons: self.buttons["ROI Selection"].config(relief=tk.RAISED)

            elif self.canny_selection_mode and self.canny_start:
                self.canny_end = (canvas_x, canvas_y)
                if abs(self.canny_start[0] - self.canny_end[0]) < 1 or \
                   abs(self.canny_start[1] - self.canny_end[1]) < 1:
                    self.canny_start = None
                    self.canny_end = None
                    self.image_canvas.delete("canny_rect")
                    self.img_filtered = None
                    self.measurement.set("Status: Canny ROI cancelled (zero size).")
                    self.display_image()
                else:
                    x1 = min(self.canny_start[0], self.canny_end[0])
                    y1 = min(self.canny_start[1], self.canny_end[1])
                    x2 = max(self.canny_start[0], self.canny_end[0])
                    y2 = max(self.canny_start[1], self.canny_end[1])
                    self.canny_start = (x1, y1)
                    self.canny_end = (x2, y2)
                    # Filter applied live, just update status and finalize rect display
                    self.measurement.set(f"Status: Canny filter applied to ROI (Thresh: {self.canny_low.get()}/{self.canny_high.get()}).")
                    self.image_canvas.delete("canny_rect") # Delete temp rect
                    self.image_canvas.create_rectangle(x1, y1, x2, y2, outline="blue", dash=(4, 4), tags="canny_rect")

                self.canny_selection_mode = False
                if "Canny Selection" in self.buttons: self.buttons["Canny Selection"].config(relief=tk.RAISED)

            # Update zoom box regardless
            if self.zoom_box_mode:
                self.update_zoom_box_content(event)

        except tk.TclError:
             pass # Ignore errors if canvas is destroyed during release


    def _reset_all_modes(self, active_mode_attr=None):
        """Deactivates all measurement/selection modes and updates button states."""
        mode_info = {
            "artery_mode": ("Artery Mode", "Dots Mode: Click to measure distance/angle."),
            "calibration_mode": ("Calibrate", "Calibration: Click two known points."),
            "edge_selection_mode": ("ROI Selection", "ROI Selection (Edges): Drag area for FIND_EDGES."),
            "canny_selection_mode": ("Canny Selection", "Canny ROI Selection: Drag area for Canny filter."),
            "angle_mode": ("Angle Mode", "Angle Mode: Click 3 points (point, vertex, point)."),
            "line_mode": ("Line Mode", "Line Mode: Click 4 points for parallel lines.")
        }

        status_message = "Status: Ready"
        active_mode_display = "None"

        for mode_attr, (button_key, msg) in mode_info.items():
            is_active = (mode_attr == active_mode_attr)
            setattr(self, mode_attr, is_active)

            if button_key in self.buttons:
                 try:
                     self.buttons[button_key].config(relief=tk.SUNKEN if is_active else tk.RAISED)
                 except tk.TclError: pass

            if is_active:
                status_message = msg
                active_mode_display = mode_attr.split('_')[0].capitalize()
                if mode_attr == "angle_mode": self.angle_points = []
                if mode_attr == "line_mode":
                    # Don't reset line_points here, allow persistence
                    # self.line_points = []
                    self.line_measurement_points = [] # Clear ticks when entering mode
                if mode_attr == "calibration_mode":
                    # Dots are reset when toggling calibration mode ON
                    pass

        self.measurement.set(status_message)
        try:
             current_info = self.pixel_info.get()
             parts = current_info.split('|')
             pixel_part = parts[1].strip() if len(parts) > 1 else "Pixel:"
             zoom_part = parts[-1].strip() if len(parts) > 0 else "Zoom: ?"
             self.pixel_info.set(f"Mode: {active_mode_display} | {pixel_part} | {zoom_part}")
        except Exception:
             self.pixel_info.set(f"Mode: {active_mode_display} | Zoom: {'ON' if self.zoom_box_mode else 'OFF'}")


    def toggle_artery_mode(self):
        self.save_state()
        new_state = not self.artery_mode
        self._reset_all_modes("artery_mode" if new_state else None)

    def reset_artery_mode(self):
        self.save_state()
        if self.artery_mode:
            self._reset_all_modes()
        self.artery_dots = []
        self.measurements = [m for m in self.measurements if m.get("type") != "artery"]
        self.measurement.set("Status: Dots Mode reset.")
        self.update_dot_coords_display()
        self.update_tables()
        self.display_image()

    def toggle_calibration_mode(self):
        self.save_state()
        new_state = not self.calibration_mode
        if new_state:
             self.calibration_dots = [] # Always reset dots when entering mode
             if self.calibration_done:
                  if not messagebox.askyesno("Recalibrate?", "Calibration already exists. Reset and recalibrate?", parent=self.root):
                      return
                  else:
                     self.reset_calibration(ask_confirm=False)
             self._reset_all_modes("calibration_mode")
             # Status message set by _reset_all_modes
             self.display_image()
             self.update_dot_coords_display()
        else:
             self._reset_all_modes()


    def prompt_for_calibration(self, distance_px):
        """Asks the user for the real distance corresponding to the pixel distance."""
        self.root.attributes('-topmost', 1)
        real_value_str = simpledialog.askstring("Calibration Input",
                                              f"Measured {distance_px:.2f} pixels between points.\nEnter the REAL distance (e.g., in mm):",
                                              parent=self.root)
        self.root.attributes('-topmost', 0)

        if real_value_str:
            try:
                real_value = float(real_value_str)
                if real_value > 0:
                    self.calibration_factor = distance_px / real_value
                    self.calibration_done = True
                    # Remove any previous calibration measurements before adding new one
                    self.measurements = [m for m in self.measurements if m.get("type") != "calibration"]
                    self.measurements.append({
                        "type": "calibration",
                        "points": self.calibration_dots.copy(),
                        "distance_px": distance_px,
                        "real_value_mm": real_value,
                        "calibration_factor": self.calibration_factor
                    })
                    self.update_tables()
                    self.update_dot_coords_display()
                    self._reset_all_modes()
                    self.measurement.set(f"Calibrated: {self.calibration_factor:.4f} px/mm")
                    messagebox.showinfo("Calibration Success", f"Calibration successful!\nFactor: {self.calibration_factor:.4f} pixels/mm", parent=self.root)
                    self.display_image()

                else:
                    messagebox.showerror("Calibration Error", "Real distance must be positive.", parent=self.root)
                    if self.calibration_dots: self.calibration_dots.pop()
                    self.measurement.set("Calibration Error: Enter positive distance.")
                    self.display_image()
                    self.update_dot_coords_display()
            except ValueError:
                messagebox.showerror("Calibration Error", "Invalid number entered.", parent=self.root)
                if self.calibration_dots: self.calibration_dots.pop()
                self.measurement.set("Calibration Error: Invalid input.")
                self.display_image()
                self.update_dot_coords_display()
        else:
            if self.calibration_dots: self.calibration_dots.pop()
            self.measurement.set("Calibration: Cancelled. Click second point again.")
            # Stay in calibration mode
            self.display_image()
            self.update_dot_coords_display()


    def delete_last_pair(self):
        """Removes the last PAIR of artery dots and their measurement."""
        self.save_state()
        if len(self.artery_dots) >= 2:
            removed_pt1 = self.artery_dots.pop()
            removed_pt2 = self.artery_dots.pop()
            last_artery_meas_index = -1
            for i in range(len(self.measurements) - 1, -1, -1):
                 meas = self.measurements[i]
                 if meas.get("type") == "artery":
                     m_pts = meas.get("points", [])
                     if len(m_pts) == 2 and m_pts[0] == removed_pt2 and m_pts[1] == removed_pt1:
                           last_artery_meas_index = i
                           break
            if last_artery_meas_index != -1:
                del self.measurements[last_artery_meas_index]
                self.measurement.set("Status: Last dot pair and measurement deleted.")
            else:
                 self.measurement.set("Status: Last dot pair deleted (no matching measurement).")

            self.display_image()
            self.update_dot_coords_display()
            self.update_tables()
        elif len(self.artery_dots) == 1:
             self.artery_dots.pop()
             self.measurement.set("Status: Last pending dot deleted.")
             self.display_image()
             self.update_dot_coords_display()
             self.update_tables()
        else:
            self.measurement.set("Status: No dots to delete.")

    def reset_calibration(self, ask_confirm=True):
        """Resets calibration factor and removes calibration dots."""
        if ask_confirm:
             if not self.calibration_done and not self.calibration_dots:
                  messagebox.showinfo("Calibration Reset", "No calibration data to reset.", parent=self.root)
                  return
             if not messagebox.askyesno("Confirm Reset", "Reset current calibration data?", parent=self.root):
                return

        self.save_state()
        self.calibration_factor = 1.0
        self.calibration_done = False
        self.calibration_dots = []
        self.measurements = [m for m in self.measurements if m.get("type") != "calibration"]

        if self.calibration_mode:
            self._reset_all_modes()
        else:
             self.measurement.set("Status: Calibration reset.")

        self.update_dot_coords_display()
        self.update_tables()
        self.display_image()

    # --- ROI Selection (Legacy FIND_EDGES) ---
    def toggle_roi_selection(self):
        self.save_state()
        new_state = not self.edge_selection_mode
        self._reset_all_modes("edge_selection_mode" if new_state else None)

    def toggle_edge_detection(self):
        self.save_state()
        self.edge_detection_active = not self.edge_detection_active
        self.measurement.set(f"Edge Detection (Simple): {'ON' if self.edge_detection_active else 'OFF'}")
        print("Warning: Simple Edge Detection (FIND_EDGES) toggled. May conflict/overwrite Canny.")


    # --- Canny Selection ---
    def toggle_canny_selection(self):
        """Toggles the mode for selecting the Canny filter ROI."""
        self.save_state()
        new_state = not self.canny_selection_mode
        # If activating ROI mode, turn off global Canny first
        if new_state and self.global_canny_active:
             self.toggle_global_canny() # Turn off global filter
        self._reset_all_modes("canny_selection_mode" if new_state else None)


    # --- Global Canny ---
    def toggle_global_canny(self):
        """Toggles the global Canny filter on/off."""
        self.save_state()
        new_state = not self.global_canny_active
        # If activating global mode, turn off ROI mode/clear ROI first
        if new_state:
             if self.canny_selection_mode:
                 self._reset_all_modes() # Turn off ROI selection mode
             # Clear any existing ROI selection
             self.canny_start = None
             self.canny_end = None
             if hasattr(self, 'image_canvas') and self.image_canvas and self.image_canvas.winfo_exists():
                 self.image_canvas.delete("canny_rect")
             self.canny_rect = None

        self.global_canny_active = new_state

        # Update button appearance
        if "Global Canny" in self.buttons:
             try:
                 self.buttons["Global Canny"].config(relief=tk.SUNKEN if self.global_canny_active else tk.RAISED)
             except tk.TclError: pass # Ignore if button destroyed

        # Update status and apply/remove filter
        self.apply_filters_and_display() # This also sets the status message

    def toggle_zoom_box(self):
        """Toggles the visibility and functionality of the zoom box."""
        if not self.root or not self.root.winfo_exists() or not self.image_frame or not self.image_frame.winfo_exists():
             return

        self.save_state()
        self.zoom_box_mode = not self.zoom_box_mode

        if self.zoom_box_mode:
            self.buttons["Zoom In Box"].config(relief=tk.SUNKEN)
            self.measurement.set("Status: Zoom Box ON. Move cursor over image.")

            if not self.zoom_box or not self.zoom_box.winfo_exists():
                 print("Error: Zoom box widget does not exist! Recreating...")
                 try:
                     self.zoom_box = tk.Canvas(self.image_frame, width=self.ZOOM_BOX_SIZE, height=self.ZOOM_BOX_SIZE,
                                            bg="black", highlightthickness=1, highlightbackground="white")
                 except tk.TclError as e:
                     print(f"Fatal Error: Could not recreate zoom box: {e}")
                     self.zoom_box_mode = False
                     self.buttons["Zoom In Box"].config(relief=tk.RAISED)
                     self.measurement.set("Status: Error creating Zoom Box.")
                     return
                 if not self.zoom_box:
                     self.zoom_box_mode = False
                     self.buttons["Zoom In Box"].config(relief=tk.RAISED)
                     self.measurement.set("Status: Error creating Zoom Box.")
                     return

            try:
                self.zoom_box.place(in_=self.image_frame, anchor='se', relx=1.0, rely=1.0, x=-10, y=-10)
                self.update_zoom_box_content(None)
            except tk.TclError as e:
                 print(f"Error placing or updating zoom box: {e}")
                 self.zoom_box_mode = False
                 self.buttons["Zoom In Box"].config(relief=tk.RAISED)
                 self.measurement.set("Status: Zoom Box Error.")

        else: # Turning zoom box off
            self.buttons["Zoom In Box"].config(relief=tk.RAISED)
            self.measurement.set("Status: Zoom Box OFF.")
            if self.zoom_box and self.zoom_box.winfo_exists():
                try: self.zoom_box.place_forget()
                except tk.TclError: pass

        # Update pixel info mode display
        try:
            current_info = self.pixel_info.get()
            parts = current_info.split('|')
            mode_str = parts[0].strip() if parts else "Mode: Unknown"
            pixel_str = parts[1].strip() if len(parts) > 1 else "Pixel:"
            zoom_status = f"Zoom: {'ON' if self.zoom_box_mode else 'OFF'}"
            self.pixel_info.set(f"{mode_str} | {pixel_str} | {zoom_status}")
        except Exception:
             self.pixel_info.set(f"Mode: Unknown | Zoom: {'ON' if self.zoom_box_mode else 'OFF'}")


    def toggle_angle_mode(self):
        self.save_state()
        new_state = not self.angle_mode
        self._reset_all_modes("angle_mode" if new_state else None)

    def toggle_line_mode(self):
        self.save_state()
        new_state = not self.line_mode
        self._reset_all_modes("line_mode" if new_state else None)


    def reset_lines(self):
        """Resets the points and measurements for Line Mode."""
        self.save_state()
        if self.line_mode: # Deactivate mode if active
            self._reset_all_modes()
        self.line_points = []
        self.line_measurements = [] # Pixel distances list
        self.line_measurement_points = [] # Points for drawing ticks
        # Remove line measurements from the main list
        self.measurements = [m for m in self.measurements if m.get("type") != "line"]
        self.measurement.set("Status: Line Mode reset.")
        self.update_dot_coords_display()
        self.update_tables()
        self.display_image() # Redraw without lines/points


    def calculate_line_measurements(self):
        """Calculates distances between two parallel lines defined by 4 points."""
        if len(self.line_points) != 4:
             return None

        p1, p2 = self.line_points[0], self.line_points[1] # First line segment
        p3, p4 = self.line_points[2], self.line_points[3] # Second line segment

        # --- Calculate Line Lengths ---
        len1_px = math.sqrt((p2[0] - p1[0])**2 + (p2[1] - p1[1])**2)
        len2_px = math.sqrt((p4[0] - p3[0])**2 + (p4[1] - p3[1])**2)

        # --- Calculate Angle Between Lines ---
        v1 = (p2[0] - p1[0], p2[1] - p1[1])
        v2 = (p4[0] - p3[0], p4[1] - p3[1])
        mag1_sq = v1[0]**2 + v1[1]**2
        mag2_sq = v2[0]**2 + v2[1]**2
        angle_deg = 0.0

        if mag1_sq > 1e-9 and mag2_sq > 1e-9: # Avoid division by zero / zero length vectors
            mag1 = math.sqrt(mag1_sq)
            mag2 = math.sqrt(mag2_sq)
            dot_product = v1[0] * v2[0] + v1[1] * v2[1]
            # Clamp cos_theta to handle potential float errors slightly outside [-1, 1]
            cos_theta = max(-1.0, min(1.0, dot_product / (mag1 * mag2)))
            angle_deg = math.degrees(math.acos(cos_theta))
            if angle_deg > 90: angle_deg = 180.0 - angle_deg
        else:
             print("Warning: Line Mode - One or both line segments have zero length.")
             return None


        # --- Calculate Perpendicular Distances ---
        distances_px = []
        self.line_measurement_points = [] # Reset points for drawing ticks
        num_measures = 15

        line2_vec = v2

        if mag2_sq < 1e-9:
             print("Error: Line 2 magnitude squared is zero in distance calculation loop.")
             return None

        for i in range(num_measures + 1):
            t = i / num_measures
            pt_on_line1 = (p1[0] + t * v1[0], p1[1] + t * v1[1])
            p3_to_pt = (pt_on_line1[0] - p3[0], pt_on_line1[1] - p3[1])
            t_proj = (p3_to_pt[0] * line2_vec[0] + p3_to_pt[1] * line2_vec[1]) / mag2_sq
            closest_pt_on_inf_line2 = (p3[0] + t_proj * line2_vec[0], p3[1] + t_proj * line2_vec[1])
            dist = math.sqrt((pt_on_line1[0] - closest_pt_on_inf_line2[0])**2 +
                             (pt_on_line1[1] - closest_pt_on_inf_line2[1])**2)
            distances_px.append(dist)
            self.line_measurement_points.append((pt_on_line1, closest_pt_on_inf_line2))

        # --- Calculate Averages --- <<< ADDED
        avg_dist_px = sum(distances_px) / len(distances_px) if distances_px else 0
        avg_dist_mm = None # Initialize

        # --- Store results ---
        result = {
            "type": "line",
            "points": self.line_points.copy(),
            "length1_px": len1_px,
            "length2_px": len2_px,
            "angle_deg": angle_deg,
            "distances_px": distances_px,
            "avg_dist_px": avg_dist_px, # <<< ADDED AVG PX
        }

        # Add real distances AND averages if calibrated
        if self.calibration_done:
             result["length1_mm"] = len1_px / self.calibration_factor
             result["length2_mm"] = len2_px / self.calibration_factor
             distances_mm = [d / self.calibration_factor for d in distances_px]
             result["distances_mm"] = distances_mm
             # Calculate mm average only if mm distances were calculated
             if distances_mm:
                 avg_dist_mm = sum(distances_mm) / len(distances_mm)
                 result["avg_dist_mm"] = avg_dist_mm # <<< ADDED AVG MM

        self.line_measurements = distances_px # Legacy storage

        return result

    def show_line_measurements(self):
        """Displays detailed results of the last Line Mode measurement in a new window."""
        line_measurement = None
        for meas in reversed(self.measurements):
            if meas.get("type") == "line":
                line_measurement = meas
                break

        if not line_measurement:
            messagebox.showinfo("No Line Measurement", "No Line Mode measurement found.", parent=self.root)
            return

        window = Toplevel(self.root)
        window.title("Line Mode - Measurement Details")
        window.geometry("450x400")
        window.transient(self.root)
        window.grab_set()

        text_frame = tk.Frame(window)
        text_frame.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

        text_widget = tk.Text(text_frame, height=15, width=50, wrap=tk.WORD, font=("Courier New", 10))
        scrollbar = Scrollbar(text_frame, orient=tk.VERTICAL, command=text_widget.yview)
        text_widget.configure(yscrollcommand=scrollbar.set)

        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # --- Populate Text Widget ---
        text_widget.insert(tk.END, "--- Line Mode Summary ---\n\n")

        len1_px = line_measurement.get('length1_px')
        len2_px = line_measurement.get('length2_px')
        angle = line_measurement.get('angle_deg')
        text_widget.insert(tk.END, f"Line 1 Length: {len1_px:>8.2f} px" if len1_px is not None else "Line 1 Length: N/A")
        if self.calibration_done and 'length1_mm' in line_measurement:
             text_widget.insert(tk.END, f"  ({line_measurement['length1_mm']:.3f} mm)\n")
        else:
             text_widget.insert(tk.END, "\n")

        text_widget.insert(tk.END, f"Line 2 Length: {len2_px:>8.2f} px" if len2_px is not None else "Line 2 Length: N/A")
        if self.calibration_done and 'length2_mm' in line_measurement:
             text_widget.insert(tk.END, f"  ({line_measurement['length2_mm']:.3f} mm)\n")
        else:
             text_widget.insert(tk.END, "\n")

        text_widget.insert(tk.END, f"Angle Deviation: {angle:>6.2f}° (0° = parallel)\n\n" if angle is not None else "Angle Deviation: N/A\n\n")

        distances_px = line_measurement.get('distances_px', [])
        distances_mm = line_measurement.get('distances_mm', [])
        text_widget.insert(tk.END, "--- Distance Measurements ---\n")
        text_widget.insert(tk.END, " # |   Pixels  |    mm\n")
        text_widget.insert(tk.END, "---|-----------|-----------\n")

        for i, dist_px in enumerate(distances_px):
             line = f"{i+1:>2} | {dist_px:>9.2f} |"
             if self.calibration_done and i < len(distances_mm):
                 line += f" {distances_mm[i]:>9.3f}\n"
             else:
                 line += "   N/A\n"
             text_widget.insert(tk.END, line)

        avg_dist_px = sum(distances_px) / len(distances_px) if distances_px else 0
        min_dist_px = min(distances_px) if distances_px else 0
        max_dist_px = max(distances_px) if distances_px else 0

        text_widget.insert(tk.END, "\n--- Statistics ---\n")
        text_widget.insert(tk.END, f"Average Dist: {avg_dist_px:>8.2f} px")
        if self.calibration_done and distances_mm:
             avg_dist_mm = sum(distances_mm) / len(distances_mm)
             text_widget.insert(tk.END, f"  ({avg_dist_mm:.3f} mm)\n")
        else:
             text_widget.insert(tk.END, "\n")

        text_widget.insert(tk.END, f"Minimum Dist: {min_dist_px:>8.2f} px")
        if self.calibration_done and distances_mm:
             min_dist_mm = min(distances_mm)
             text_widget.insert(tk.END, f"  ({min_dist_mm:.3f} mm)\n")
        else:
             text_widget.insert(tk.END, "\n")

        text_widget.insert(tk.END, f"Maximum Dist: {max_dist_px:>8.2f} px")
        if self.calibration_done and distances_mm:
             max_dist_mm = max(distances_mm)
             text_widget.insert(tk.END, f"  ({max_dist_mm:.3f} mm)\n")
        else:
             text_widget.insert(tk.END, "\n")


        text_widget.config(state=tk.DISABLED)

        close_button = tk.Button(window, text="Close", command=window.destroy)
        close_button.pack(pady=5)

        window.wait_window()


    def reset_filters(self):
        """Resets all filter effects and selections."""
        self.save_state()
        self.img_filtered = None
        self.canny_start = None
        self.canny_end = None
        if hasattr(self, 'image_canvas') and self.image_canvas and self.image_canvas.winfo_exists():
             self.image_canvas.delete("canny_rect")
        self.canny_rect = None
        if self.canny_selection_mode:
             self._reset_all_modes()

        self.global_canny_active = False
        if "Global Canny" in self.buttons:
            try: self.buttons["Global Canny"].config(relief=tk.RAISED)
            except tk.TclError: pass

        self.edge_detection_active = False
        self.selection_start = None
        self.selection_end = None
        if hasattr(self, 'image_canvas') and self.image_canvas and self.image_canvas.winfo_exists():
             self.image_canvas.delete("selection_rect")
        self.selection_rect = None
        if self.edge_selection_mode: self._reset_all_modes()


        self.measurement.set("Status: Filters reset.")
        self.display_image()
        if self.zoom_box_mode:
             self.update_zoom_box_content(None)


    def save_state(self):
        """Saves the current relevant state for undo."""
        if not self.img_original: return

        state = {
            "calibration_dots": self.calibration_dots.copy(),
            "artery_dots": self.artery_dots.copy(),
            "line_points": self.line_points.copy(),
            "angle_points": self.angle_points.copy(),
            "measurements": [m.copy() for m in self.measurements],
            "line_measurements": self.line_measurements.copy(),
            "line_measurement_points": self.line_measurement_points.copy(),
            "img_filtered": self.img_filtered.copy() if self.img_filtered else None,
            "calibration_factor": self.calibration_factor,
            "calibration_done": self.calibration_done,
            # Filter states
            "canny_start": self.canny_start,
            "canny_end": self.canny_end,
            "selection_start": self.selection_start, # Legacy
            "selection_end": self.selection_end,     # Legacy
            "edge_detection_active": self.edge_detection_active, # Legacy
            "global_canny_active": self.global_canny_active, # Added
        }
        self.undo_stack.append(state)
        self.redo_stack.clear()

        max_undo = 50
        if len(self.undo_stack) > max_undo:
            self.undo_stack.pop(0)


    def _restore_state(self, state):
         """Restores the application state from a saved dictionary."""
         self.calibration_dots = state.get("calibration_dots", [])
         self.artery_dots = state.get("artery_dots", [])
         self.line_points = state.get("line_points", [])
         self.angle_points = state.get("angle_points", [])
         self.measurements = state.get("measurements", [])
         self.line_measurements = state.get("line_measurements", [])
         self.line_measurement_points = state.get("line_measurement_points", [])
         img_filt_data = state.get("img_filtered")
         self.img_filtered = img_filt_data.copy() if img_filt_data else None
         self.calibration_factor = state.get("calibration_factor", 1.0)
         self.calibration_done = state.get("calibration_done", False)
         # Filter states
         self.canny_start = state.get("canny_start")
         self.canny_end = state.get("canny_end")
         self.selection_start = state.get("selection_start")
         self.selection_end = state.get("selection_end")
         self.edge_detection_active = state.get("edge_detection_active", False)
         self.global_canny_active = state.get("global_canny_active", False) # Added

         # Update Global Canny button state
         if "Global Canny" in self.buttons:
             try:
                 self.buttons["Global Canny"].config(relief=tk.SUNKEN if self.global_canny_active else tk.RAISED)
             except tk.TclError: pass

         if hasattr(self, 'image_canvas') and self.image_canvas and self.image_canvas.winfo_exists():
              self.image_canvas.delete("canny_rect")
              self.image_canvas.delete("selection_rect")
              self.canny_rect = None
              self.selection_rect = None

         self.display_image()
         self.update_dot_coords_display()
         self.update_tables()
         self._reset_all_modes() # Simple reset, doesn't restore active mode, but cleans buttons
         if self.calibration_done:
              self.measurement.set(f"Calibrated: {self.calibration_factor:.4f} px/mm")
         elif self.img_filtered:
             if self.global_canny_active:
                 self.measurement.set("Status: Global Canny Filter ON.")
             elif self.canny_start:
                 self.measurement.set("Status: Canny ROI Filter Applied.")
             else:
                 self.measurement.set("Status: Filter applied.")
         # Keep status message generic after restore, specific message set in undo/redo


    def undo(self, event=None):
        """Reverts to the previous saved state."""
        if not self.undo_stack:
             self.measurement.set("Status: Nothing to undo.")
             return

        current_state = {
            "calibration_dots": self.calibration_dots.copy(),
            "artery_dots": self.artery_dots.copy(),
            "line_points": self.line_points.copy(),
            "angle_points": self.angle_points.copy(),
            "measurements": [m.copy() for m in self.measurements],
            "line_measurements": self.line_measurements.copy(),
            "line_measurement_points": self.line_measurement_points.copy(),
            "img_filtered": self.img_filtered.copy() if self.img_filtered else None,
            "calibration_factor": self.calibration_factor,
            "calibration_done": self.calibration_done,
            "canny_start": self.canny_start,
            "canny_end": self.canny_end,
            "selection_start": self.selection_start,
            "selection_end": self.selection_end,
            "edge_detection_active": self.edge_detection_active,
            "global_canny_active": self.global_canny_active, # Added
        }
        self.redo_stack.append(current_state)

        state_to_restore = self.undo_stack.pop()
        self._restore_state(state_to_restore)
        self.measurement.set("Status: Undo successful.")


    def redo(self, event=None):
        """Re-applies the last undone state."""
        if not self.redo_stack:
            self.measurement.set("Status: Nothing to redo.")
            return

        current_state = {
             "calibration_dots": self.calibration_dots.copy(),
             "artery_dots": self.artery_dots.copy(),
             "line_points": self.line_points.copy(),
             "angle_points": self.angle_points.copy(),
             "measurements": [m.copy() for m in self.measurements],
             "line_measurements": self.line_measurements.copy(),
             "line_measurement_points": self.line_measurement_points.copy(),
             "img_filtered": self.img_filtered.copy() if self.img_filtered else None,
             "calibration_factor": self.calibration_factor,
             "calibration_done": self.calibration_done,
             "canny_start": self.canny_start,
             "canny_end": self.canny_end,
             "selection_start": self.selection_start,
             "selection_end": self.selection_end,
             "edge_detection_active": self.edge_detection_active,
             "global_canny_active": self.global_canny_active, # Added
         }
        self.undo_stack.append(current_state)

        state_to_restore = self.redo_stack.pop()
        self._restore_state(state_to_restore)
        self.measurement.set("Status: Redo successful.")


    def export_annotated_image(self):
        """Exports the currently displayed image with overlays."""
        if not self.img_original:
            messagebox.showerror("Export Error", "No image loaded to export.", parent=self.root)
            return

        # Start with the currently displayed image (could be filtered or original)
        img_to_export = (self.img_filtered if self.img_filtered is not None else self.img_original).copy()
        # Ensure it's suitable for color drawing
        if img_to_export.mode not in ('RGB', 'RGBA'):
             img_to_export = img_to_export.convert('RGBA')
        elif img_to_export.mode == 'RGB':
             img_to_export = img_to_export.convert('RGBA') # Ensure alpha for drawing

        draw = ImageDraw.Draw(img_to_export)

        dot_radius = 3
        line_width = 2

        # Calibration dots
        for x, y in self.calibration_dots:
            draw.ellipse((x - dot_radius, y - dot_radius, x + dot_radius, y + dot_radius), fill="cyan", outline="black")

        # Artery dots and lines
        for i in range(0, len(self.artery_dots)):
            x, y = self.artery_dots[i]
            draw.ellipse((x - dot_radius, y - dot_radius, x + dot_radius, y + dot_radius), fill="yellow", outline="black")
            if i % 2 == 1:
                x_prev, y_prev = self.artery_dots[i-1]
                draw.line([(x_prev, y_prev), (x, y)], fill="yellow", width=line_width)

        # Line mode points and lines
        for i in range(0, len(self.line_points)):
            x, y = self.line_points[i]
            draw.ellipse((x - dot_radius, y - dot_radius, x + dot_radius, y + dot_radius), fill="magenta", outline="black")
            if i % 2 == 1:
                x_prev, y_prev = self.line_points[i-1]
                draw.line([(x_prev, y_prev), (x, y)], fill="magenta", width=line_width)

        # Angle points and lines
        if self.angle_points:
             p1_a, p2_a, p3_a = None, None, None
             if len(self.angle_points) >= 1: p1_a = self.angle_points[0]
             if len(self.angle_points) >= 2: p2_a = self.angle_points[1]
             if len(self.angle_points) >= 3: p3_a = self.angle_points[2]

             if p1_a: draw.ellipse((p1_a[0]-dot_radius, p1_a[1]-dot_radius, p1_a[0]+dot_radius, p1_a[1]+dot_radius), fill="lime green", outline="black")
             if p2_a:
                 draw.ellipse((p2_a[0]-dot_radius, p2_a[1]-dot_radius, p2_a[0]+dot_radius, p2_a[1]+dot_radius), fill="lime green", outline="black")
                 if p1_a: draw.line([p1_a, p2_a], fill="lime green", width=line_width, joint="curve")
             if p3_a:
                 draw.ellipse((p3_a[0]-dot_radius, p3_a[1]-dot_radius, p3_a[0]+dot_radius, p3_a[1]+dot_radius), fill="lime green", outline="black")
                 if p2_a: draw.line([p2_a, p3_a], fill="lime green", width=line_width, joint="curve")

        # Line mode measurement ticks
        tick_radius = 2
        for p1, p2 in self.line_measurement_points:
             draw.ellipse((p1[0]-tick_radius, p1[1]-tick_radius, p1[0]+tick_radius, p1[1]+tick_radius), fill="red", outline="red")
             draw.line([p1, p2], fill="red", width=1)

        # Ask for save file path
        base_name = os.path.splitext(os.path.basename(self.file_path))[0] if self.file_path else "image"
        save_path = filedialog.asksaveasfilename(
            title="Save Annotated Image As",
            initialfile=f"{base_name}_annotated.png",
            defaultextension=".png",
            filetypes=[("PNG Image", "*.png"), ("JPEG Image", "*.jpg"), ("Bitmap Image", "*.bmp"), ("TIFF Image", "*.tif"), ("All Files", "*.*")],
            parent=self.root
        )

        if save_path:
            try:
                save_format = os.path.splitext(save_path)[1].lower()
                final_image_to_save = img_to_export
                if save_format == ".jpg" or save_format == ".jpeg":
                     if final_image_to_save.mode == 'RGBA':
                          bg = Image.new("RGB", final_image_to_save.size, (255, 255, 255))
                          bg.paste(final_image_to_save, mask=final_image_to_save.split()[3])
                          final_image_to_save = bg
                     elif final_image_to_save.mode != 'RGB':
                          final_image_to_save = final_image_to_save.convert('RGB')

                final_image_to_save.save(save_path)
                messagebox.showinfo("Export Successful", f"Annotated image saved to:\n{save_path}", parent=self.root)
            except Exception as e:
                messagebox.showerror("Export Error", f"Failed to save image:\n{e}", parent=self.root)
                print(traceback.format_exc())


    def save_measurements_to_json(self):
        """Saves all collected measurement data, calibration, and metadata to a JSON file."""
        if not self.file_path:
            messagebox.showerror("Save Error", "No image loaded.", parent=self.root)
            return

        if not self.measurements:
             if not messagebox.askokcancel("Save?", "No measurements recorded. Save empty file?", parent=self.root):
                 return

        name = self.name_var.get().strip()
        diameter_str = self.diameter_var.get().strip()

        if not name:
            messagebox.showerror("Input Error", "Please enter a 'Name'.", parent=self.root)
            self.name_entry.focus_set()
            return
        if not diameter_str:
            messagebox.showerror("Input Error", "Please enter 'Real Ø (mm)'.", parent=self.root)
            self.diameter_entry.focus_set()
            return

        try:
            real_diameter = float(diameter_str)
            if real_diameter <= 0: raise ValueError("Diameter must be positive")
        except ValueError:
            messagebox.showerror("Input Error", "Enter a valid positive number for 'Real Ø (mm)'.", parent=self.root)
            self.diameter_entry.focus_set()
            return

        current_time = datetime.datetime.now()
        time_str = current_time.strftime("%Y-%m-%d_%H-%M-%S")
        time_iso = current_time.isoformat()

        data = {
            "metadata": {
                "source_image_path": self.file_path,
                "source_image_name": os.path.basename(self.file_path),
                "analysis_name": name,
                "analysis_timestamp_iso": time_iso,
                "expected_real_diameter_mm": real_diameter,
                "software_version": "ImageAnalyzer_1.4_JsonMmAvg" # Example version
            },
            "calibration": {
                "calibrated": self.calibration_done,
                "pixels_per_mm": round(self.calibration_factor, 6) if self.calibration_done else None,
                "calibration_points": [(round(p[0], 1), round(p[1], 1)) for p in self.calibration_dots]
            },
            "measurements": []
        }

        # Add measurements, ensuring consistent formatting
        for meas in self.measurements:
            clean_meas = meas.copy() # Work on a copy

            # Round floats for cleaner output - Revised to handle specific keys better
            for key, value in clean_meas.items():
                 if isinstance(value, float):
                     round_digits = 4 # Default rounding for floats
                     if key.endswith("_mm"): round_digits = 4
                     elif key.endswith("_px"): round_digits = 2
                     elif key.endswith("_deg"): round_digits = 2
                     elif key == "calibration_factor": round_digits = 6
                     # Apply rounding
                     clean_meas[key] = round(value, round_digits)

                 elif isinstance(value, list): # Handle lists (like points or distance lists)
                     if key == "points":
                          # Round points (tuples) within the list
                          clean_meas[key] = [(round(pt[0], 1), round(pt[1], 1))
                                             if isinstance(pt, (list, tuple)) and len(pt)==2
                                             else pt
                                             for pt in value]
                     elif key == "distances_px":
                         # Round pixel distances in the list
                         clean_meas[key] = [round(d, 2) if isinstance(d, (int, float)) else d for d in value]
                     elif key == "distances_mm":
                          # Round mm distances in the list
                          clean_meas[key] = [round(d, 4) if isinstance(d, (int, float)) else d for d in value]

            # Add the processed measurement to the final data
            data["measurements"].append(clean_meas)


        # Ask for save file location
        default_filename = f"analysis_{name}_{time_str}.json"
        filename = filedialog.asksaveasfilename(
            title="Save Analysis Data As",
            initialfile=default_filename,
            defaultextension=".json",
            filetypes=[("JSON Files", "*.json"), ("All Files", "*.*")],
            parent=self.root
        )

        if filename:
            try:
                with open(filename, 'w') as f:
                    json.dump(data, f, indent=4) # Use indent for readability
                messagebox.showinfo("Save Successful", f"Analysis data saved to:\n{filename}", parent=self.root)
            except Exception as e:
                messagebox.showerror("Save Error", f"Failed to save JSON file:\n{e}", parent=self.root)
                print(traceback.format_exc())

                
    def update_tables(self):
        """Updates the measurement summary table more robustly."""
        if not hasattr(self, 'measurement_table') or not self.measurement_table or not self.measurement_table.winfo_exists():
            return

        try:
            for item in self.measurement_table.get_children():
                self.measurement_table.delete(item)
        except tk.TclError as e:
            return

        for i, meas in enumerate(self.measurements):
            if not isinstance(meas, dict) or 'type' not in meas:
                print(f"Warning: Invalid measurement data found at index {i}: {meas}")
                try:
                    self.measurement_table.insert("", tk.END, iid=f"error_{i}", values=("Error", "Invalid Data", "", ""))
                except tk.TclError: pass
                continue

            m_type_orig = meas.get("type", "unknown")
            m_type_display = m_type_orig.capitalize()
            px_dist_str = "N/A"
            mm_dist_str = "N/A"
            angle_str = "N/A"

            try:
                if m_type_orig == "artery":
                    px_dist = meas.get('distance_px')
                    angle = meas.get('angle_deg')
                    px_dist_str = f"{px_dist:.2f}" if px_dist is not None else "N/A"
                    angle_str = f"{angle:.1f}" if angle is not None else "N/A"
                    if self.calibration_done and px_dist is not None:
                         mm_dist_recalc = px_dist / self.calibration_factor
                         mm_dist_str = f"{mm_dist_recalc:.3f}"
                    elif self.calibration_done: mm_dist_str = "Error"
                    else: mm_dist_str = "Uncalib."

                elif m_type_orig == "angle":
                    angle = meas.get('angle_deg')
                    angle_str = f"{angle:.2f}" if angle is not None else "N/A"

                elif m_type_orig == "line":
                    len1_px = meas.get('length1_px')
                    len2_px = meas.get('length2_px')
                    angle = meas.get('angle_deg')
                    dists_px = meas.get('distances_px', [])
                    avg_dist_px = sum(dists_px) / len(dists_px) if dists_px else 0
                    px_dist_str = f"Avg:{avg_dist_px:.2f} (L1:{len1_px:.1f}, L2:{len2_px:.1f})" if len1_px is not None and len2_px is not None else "N/A"
                    angle_str = f"{angle:.1f}" if angle is not None else "N/A"

                    if self.calibration_done:
                        if len1_px is not None and len2_px is not None and dists_px:
                             len1_mm = len1_px / self.calibration_factor
                             len2_mm = len2_px / self.calibration_factor
                             dists_mm = [d / self.calibration_factor for d in dists_px]
                             avg_dist_mm = sum(dists_mm) / len(dists_mm) if dists_mm else 0
                             mm_dist_str = f"Avg:{avg_dist_mm:.3f} (L1:{len1_mm:.2f}, L2:{len2_mm:.2f})"
                        else: mm_dist_str = "N/A"
                    else: mm_dist_str = "Uncalib."

                elif m_type_orig == "calibration":
                    px_dist = meas.get('distance_px')
                    mm_val = meas.get('real_value_mm')
                    factor = meas.get('calibration_factor')
                    px_dist_str = f"{px_dist:.2f}" if px_dist is not None else "N/A"
                    mm_dist_str = f"{mm_val:.3f}" if mm_val is not None else "N/A"
                    angle_str = f"{factor:.4f} px/mm" if factor is not None else "N/A"
                    m_type_display = f"Calib Set"

                else:
                    print(f"Warning: Unhandled measurement type '{m_type_orig}'.")
                    m_type_display = f"{m_type_orig.capitalize()} (?"

                try:
                    self.measurement_table.insert("", tk.END, iid=str(i), values=(m_type_display, px_dist_str, mm_dist_str, angle_str))
                except tk.TclError: pass

            except Exception as e:
                print(f"Error processing measurement for table display at index {i} (type: {m_type_orig}): {e}")
                print(traceback.format_exc())
                try:
                    self.measurement_table.insert("", tk.END, iid=f"proc_error_{i}", values=(m_type_display, "Proc Error", str(e), ""))
                except tk.TclError: pass


# Main execution
if __name__ == "__main__":
    root = None # Initialize root to None
    try:
        # Create the main window first
        root = tk.Tk()
        # Optional: Set a theme for a more modern look (requires ttkthemes)
        try:
             from ttkthemes import ThemedTk
             # Ensure root is destroyed before creating ThemedTk if necessary
             if root: root.destroy()
             root = ThemedTk(theme="arc") # Example theme: arc, plastique, clearlooks etc.
        except ImportError:
             print("ttkthemes not found, using default Tk theme.")
             if not root or not root.winfo_exists(): root = tk.Tk()
             pass
        except tk.TclError as theme_error:
             print(f"Error setting ttk theme: {theme_error}. Using default Tk theme.")
             if not root or not root.winfo_exists(): root = tk.Tk()
             pass

        # Now initialize the application with the root window
        app = ImageAnalyzer(root)
        root.mainloop()

    except Exception as e:
        print("--- FATAL ERROR ---")
        print(traceback.format_exc())
        try:
            messagebox.showerror("Fatal Error", f"An unrecoverable error occurred:\n\n{e}\n\nSee console for details.")
        except Exception as msg_e:
            print(f"Could not display error messagebox: {msg_e}")

# --- END OF FULL MODIFIED FILE ---