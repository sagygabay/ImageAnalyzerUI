import tkinter as tk
from tkinter import messagebox # simpledialog and math are not directly used by ModeHandler now

class ModeHandler:
    def __init__(self, app):
        """
        Manages the different interaction modes of the Image Analyzer application.

        Args:
            app: The main application instance (ImageAnalyzer).
        """
        self.app = app
        self.mode_config = {
            "artery_mode": {"button_key": "Artery Mode", "status_msg": "Dots Mode: Click to measure distance/angle."},
            "calibration_mode": {"button_key": "Calibrate", "status_msg": "Calibration: Click two known points."},
            "edge_selection_mode": {"button_key": "ROI Selection", "status_msg": "ROI Selection (Edges): Drag area for FIND_EDGES."}, # Legacy
            "canny_selection_mode": {"button_key": "Canny Selection", "status_msg": "Canny ROI Selection: Drag area for Canny filter."},
            "angle_mode": {"button_key": "Angle Mode", "status_msg": "Angle Mode: Click 3 points (point, vertex, point)."},
            "line_mode": {"button_key": "Line Mode", "status_msg": "Line Mode: Click 4 points for parallel lines."}
        }

    def _set_mode(self, mode_to_set, active_state):
        """
        Internal helper to set a specific mode's state and update UI.
        This will also deactivate other modes if active_state is True.

        Args:
            mode_to_set (str): The attribute name of the mode to set (e.g., "calibration_mode").
            active_state (bool): True to activate the mode, False to deactivate.
        """
        if active_state:
            # Deactivate all other modes first
            for mode_attr in self.mode_config.keys():
                if mode_attr != mode_to_set:
                    setattr(self.app, mode_attr, False)
                    if self.mode_config[mode_attr]["button_key"] in self.app.buttons:
                        try:
                            self.app.buttons[self.mode_config[mode_attr]["button_key"]].config(relief=tk.RAISED)
                        except tk.TclError: pass
        
        # Set the target mode's state
        setattr(self.app, mode_to_set, active_state)
        
        button_key = self.mode_config[mode_to_set]["button_key"]
        if button_key in self.app.buttons:
            try:
                self.app.buttons[button_key].config(relief=tk.SUNKEN if active_state else tk.RAISED)
            except tk.TclError: pass

        if active_state:
            self.app.measurement.set(self.mode_config[mode_to_set]["status_msg"])
            active_mode_display = mode_to_set.split('_')[0].capitalize()
            # Specific actions when entering a mode
            if mode_to_set == "angle_mode":
                # If there are already 3 points, starting angle mode again implies a new angle.
                # The model's add_angle_point handles clearing if len >= 3.
                pass
            elif mode_to_set == "line_mode":
                # Similar logic for line_mode, model handles clearing if len >=4.
                self.app.measurement_model.line_measurement_visualization_points = [] # Clear ticks
            elif mode_to_set == "calibration_mode":
                # Resetting calibration dots is handled in toggle_calibration_mode before calling _set_mode
                pass

        else: # Deactivating the specified mode (or if all are deactivated)
            # If no other mode was set to active, reset to default status
            is_any_mode_active = any(getattr(self.app, ma) for ma in self.mode_config.keys())
            if not is_any_mode_active:
                self.app.measurement.set("Status: Ready")
                active_mode_display = "None"
            else: # Another mode is active, its status message should prevail.
                # Find the currently active mode to set the display string.
                active_mode_display = "Unknown" # Fallback
                for ma_check, details_check in self.mode_config.items():
                    if getattr(self.app, ma_check):
                        active_mode_display = ma_check.split('_')[0].capitalize()
                        break


        # Update pixel info string
        try:
            current_pixel_info = self.app.pixel_info.get()
            parts = current_pixel_info.split('|')
            pixel_part = parts[1].strip() if len(parts) > 1 else "Pixel:"
            zoom_part = parts[-1].strip() if len(parts) > 0 else f"Zoom: {'ON' if self.app.zoom_box_mode else 'OFF'}"
            self.app.pixel_info.set(f"Mode: {active_mode_display} | {pixel_part} | {zoom_part}")
        except Exception: # Fallback
            self.app.pixel_info.set(f"Mode: {active_mode_display} | Zoom: {'ON' if self.app.zoom_box_mode else 'OFF'}")


    def _reset_all_modes(self, active_mode_attr=None):
        """
        Deactivates all measurement/selection modes and updates button states.
        If active_mode_attr is provided, that mode will be set as active.
        """
        # First, ensure all known mode flags on the app are False and buttons raised
        for mode_attr, config in self.mode_config.items():
            setattr(self.app, mode_attr, False)
            if config["button_key"] in self.app.buttons:
                try:
                    self.app.buttons[config["button_key"]].config(relief=tk.RAISED)
                except tk.TclError: pass
        
        # Now, if a specific mode needs to be activated, use _set_mode
        if active_mode_attr and active_mode_attr in self.mode_config:
            self._set_mode(active_mode_attr, True)
        else:
            # If no specific mode is activated, ensure status and pixel info are default
            self.app.measurement.set("Status: Ready")
            try:
                current_pixel_info = self.app.pixel_info.get()
                parts = current_pixel_info.split('|')
                pixel_part = parts[1].strip() if len(parts) > 1 else "Pixel:"
                zoom_part = parts[-1].strip() if len(parts) > 0 else f"Zoom: {'ON' if self.app.zoom_box_mode else 'OFF'}"
                self.app.pixel_info.set(f"Mode: None | {pixel_part} | {zoom_part}")
            except Exception:
                self.app.pixel_info.set(f"Mode: None | Zoom: {'ON' if self.app.zoom_box_mode else 'OFF'}")


    def toggle_artery_mode(self):
        """Toggles the artery/distance measurement mode."""
        self.app.save_state()
        new_state = not self.app.artery_mode
        self._set_mode("artery_mode", new_state)
        if new_state:
            # Model doesn't need explicit reset here, adding dots handles pairs
            pass

    def toggle_calibration_mode(self):
        """Toggles the calibration mode."""
        self.app.save_state()
        new_state = not self.app.calibration_mode
        
        if new_state:
            # Always reset current calibration dots in the model when entering mode
            self.app.measurement_model.calibration_dots = [] 
            if self.app.measurement_model.calibration_done:
                if not messagebox.askyesno("Recalibrate?", 
                                           "Calibration already exists. Reset and recalibrate?", 
                                           parent=self.app.root):
                    return # User cancelled, don't enter mode
                else:
                    # User wants to recalibrate, reset model's calibration fully
                    self.app.measurement_model.reset_calibration() 
            self._set_mode("calibration_mode", True)
            self.app.utils.update_dot_coords_display() # Show empty calib dots
        else: # Turning calibration mode OFF
            self._set_mode("calibration_mode", False)
        self.app.display_image()


    def toggle_angle_mode(self):
        """Toggles the angle measurement mode."""
        self.app.save_state()
        new_state = not self.app.angle_mode
        self._set_mode("angle_mode", new_state)
        if new_state:
            # Model's add_angle_point will handle clearing if 3 points already exist
            # and user clicks again.
            self.app.measurement.set(self.mode_config["angle_mode"]["status_msg"])


    def toggle_line_mode(self):
        """Toggles the line measurement mode."""
        self.app.save_state()
        new_state = not self.app.line_mode
        self._set_mode("line_mode", new_state)
        if new_state:
            # Model's add_line_mode_point will handle clearing if 4 points exist.
            self.app.measurement.set(self.mode_config["line_mode"]["status_msg"])

    # Note: Methods like reset_artery_mode, reset_calibration, reset_angle_mode, reset_lines
    # are now primarily handled in ImageAnalyzer (main.py) as they involve calls to
    # MeasurementModel for data reset and then UI updates (display_image, update_tables, etc.).
    # ModeHandler focuses on toggling the *state* of the mode.

    # Canny ROI selection mode toggle is handled in ImageAnalyzer (main.py)
    # as it also needs to interact with FilterHandler.
    # This ModeHandler could be extended to manage it if desired, but current split is okay.

    # toggle_keep_dots_fixed is an app-level setting, not a "mode" in the same sense,
    # so it's handled directly in ImageAnalyzer (main.py).
