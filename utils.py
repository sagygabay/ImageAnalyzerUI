import tkinter as tk
from tkinter import filedialog, messagebox, Toplevel
from tkinter.ttk import Treeview, Scrollbar
import math
import json
import datetime
import os
from PIL import Image, ImageDraw, UnidentifiedImageError
import traceback

class AppUtils:
    def __init__(self, app):
        """
        Initializes utility functions for the Image Analyzer application.

        Args:
            app: The main application instance (ImageAnalyzer).
        """
        self.app = app

    def update_dot_coords_display(self):
        """Updates the text box showing coordinates of points in various modes."""
        if not hasattr(self.app, 'dot_coords_text') or not self.app.dot_coords_text.winfo_exists():
            return

        model = self.app.measurement_model
        text = ""

        if model.calibration_dots:
            text += "--- Calibration ---\n"
            for i, (x, y) in enumerate(model.calibration_dots):
                text += f"  Dot {i+1}: ({x:.1f}, {y:.1f})\n"
            if len(model.calibration_dots) >= 2 and model.calibration_done:
                dist_px = math.sqrt((model.calibration_dots[1][0] - model.calibration_dots[0][0])**2 +
                                    (model.calibration_dots[1][1] - model.calibration_dots[0][1])**2)
                real_dist = model.calibration_reference_distance_mm if model.calibration_reference_distance_mm is not None else (dist_px / model.calibration_factor)
                text += f"  -> Dist: {dist_px:.2f}px = {real_dist:.2f}mm (Factor: {model.calibration_factor:.4f} px/mm)\n"
            elif len(model.calibration_dots) >= 2:
                dist_px = math.sqrt((model.calibration_dots[1][0] - model.calibration_dots[0][0])**2 +
                                    (model.calibration_dots[1][1] - model.calibration_dots[0][1])**2)
                text += f"  -> Dist: {dist_px:.2f}px (Pending Calibration)\n"

        if model.artery_dots:
            text += "\n--- Dots Mode (Artery/Distance) ---\n"
            pair_count = 1
            i = 0
            while i < len(model.artery_dots):
                if i + 1 < len(model.artery_dots):
                    x1, y1 = model.artery_dots[i]
                    x2, y2 = model.artery_dots[i+1]
                    # Find the corresponding measurement in the log for accurate display
                    logged_meas = next((m for m in reversed(model.measurements_log) 
                                        if m.get("type") == "artery_distance" and m.get("points") == [(x1,y1), (x2,y2)]), None)
                    if logged_meas:
                        dist_px = logged_meas.get("distance_px", 0)
                        angle = logged_meas.get("angle_deg", 0)
                        text += f"  Pair {pair_count}: ({x1:.1f},{y1:.1f}) -> ({x2:.1f},{y2:.1f})\n"
                        if model.calibration_done and "distance_mm" in logged_meas:
                            text += f"    Dist: {dist_px:.2f}px = {logged_meas['distance_mm']:.3f}mm | Angle: {angle:.1f}°\n"
                        else:
                            text += f"    Dist: {dist_px:.2f}px | Angle: {angle:.1f}° (Uncalibrated)\n"
                    else: # Should not happen if model logs correctly
                        text += f"  Pair {pair_count} (Processing): ({x1:.1f},{y1:.1f}) -> ({x2:.1f},{y2:.1f})\n"
                    i += 2
                else:
                    x, y = model.artery_dots[i]
                    text += f"  Pair {pair_count} (Pending): ({x:.1f}, {y:.1f})\n"
                    i += 1
                pair_count += 1
        
        if model.line_points:
            text += "\n--- Line Mode Points ---\n"
            for idx, (x, y) in enumerate(model.line_points):
                 text += f"  Point {idx+1}: ({x:.1f}, {y:.1f})\n"
            if len(model.line_points) == 4:
                 text += "  (Measurement logged. 'Reset Lines' for new.)\n"
            elif len(model.line_points) > 0:
                 text += f"  ({4 - len(model.line_points)} more points needed)\n"


        if model.angle_points:
            text += "\n--- Angle Mode Points ---\n"
            for idx, (x, y) in enumerate(model.angle_points):
                 text += f"  Point {idx+1}: ({x:.1f}, {y:.1f})\n"
            if len(model.angle_points) == 3:
                text += "  (Angle measured. Click to start new.)\n"
            elif len(model.angle_points) > 0:
                text += f"  ({3 - len(model.angle_points)} more points needed)\n"

        self.app.dot_coords_text.config(state=tk.NORMAL)
        self.app.dot_coords_text.delete("1.0", tk.END)
        self.app.dot_coords_text.insert(tk.END, text if text else "No points placed yet.")
        self.app.dot_coords_text.config(state=tk.DISABLED)
        self.app.dot_coords_text.yview_moveto(1.0)

    def update_tables(self):
        """Updates the measurement summary table using data from MeasurementModel."""
        if not hasattr(self.app, 'measurement_table') or not self.app.measurement_table.winfo_exists():
            return

        try:
            for item in self.app.measurement_table.get_children():
                self.app.measurement_table.delete(item)
        except tk.TclError:
            return

        model = self.app.measurement_model
        for i, meas in enumerate(model.get_all_measurements()): # Use getter for logged measurements
            if not isinstance(meas, dict) or 'type' not in meas:
                print(f"Warning: Invalid measurement data found at index {i}: {meas}")
                try:
                    self.app.measurement_table.insert("", tk.END, iid=f"error_{i}", values=("Error", "Invalid Data", "", ""))
                except tk.TclError: pass
                continue

            m_type_orig = meas.get("type", "unknown")
            m_type_display = m_type_orig.replace("_", " ").capitalize()
            px_dist_str, mm_dist_str, angle_str = "N/A", "N/A", "N/A"

            try:
                if m_type_orig == "artery_distance":
                    px_dist = meas.get('distance_px')
                    angle = meas.get('angle_deg')
                    px_dist_str = f"{px_dist:.2f}" if px_dist is not None else "N/A"
                    angle_str = f"{angle:.1f}" if angle is not None else "N/A"
                    if model.calibration_done and "distance_mm" in meas:
                         mm_dist_str = f"{meas['distance_mm']:.3f}"
                    elif model.calibration_done: mm_dist_str = "Error" # Calibrated but no mm value
                    else: mm_dist_str = "Uncalib."

                elif m_type_orig == "angle":
                    angle = meas.get('angle_deg')
                    angle_str = f"{angle:.2f}" if angle is not None else "N/A"

                elif m_type_orig == "line_measurement":
                    l1px = meas.get('length1_px'); l2px = meas.get('length2_px')
                    angle_dev = meas.get('angle_deviation_deg')
                    avg_px = meas.get('avg_distance_px')
                    px_dist_str = f"Avg:{avg_px:.2f} (L1:{l1px:.1f}, L2:{l2px:.1f})" if all(v is not None for v in [avg_px, l1px, l2px]) else "N/A"
                    angle_str = f"{angle_dev:.1f}" if angle_dev is not None else "N/A"
                    if model.calibration_done:
                        avg_mm = meas.get('avg_distance_mm'); l1mm = meas.get('length1_mm'); l2mm = meas.get('length2_mm')
                        if all(v is not None for v in [avg_mm, l1mm, l2mm]):
                            mm_dist_str = f"Avg:{avg_mm:.3f} (L1:{l1mm:.2f}, L2:{l2mm:.2f})"
                        else: mm_dist_str = "N/A" # Some mm values missing
                    else: mm_dist_str = "Uncalib."
                
                elif m_type_orig == "calibration_set":
                    px_dist = meas.get('distance_px')
                    mm_val = meas.get('real_value_mm')
                    factor = meas.get('calibration_factor')
                    px_dist_str = f"{px_dist:.2f}" if px_dist is not None else "N/A"
                    mm_dist_str = f"{mm_val:.3f}" if mm_val is not None else "N/A"
                    angle_str = f"{factor:.4f} px/mm" if factor is not None else "N/A"
                    m_type_display = "Calibration Set"

                else:
                    m_type_display = f"{m_type_orig.replace('_', ' ').capitalize()} (?)"

                try:
                    self.app.measurement_table.insert("", tk.END, iid=str(i), values=(m_type_display, px_dist_str, mm_dist_str, angle_str))
                except tk.TclError: pass
            except Exception as e:
                print(f"Error processing measurement for table display (idx {i}, type {m_type_orig}): {e}")
                traceback.print_exc()
                try:
                    self.app.measurement_table.insert("", tk.END, iid=f"proc_error_{i}", values=(m_type_display, "Proc Error", str(e)[:20], ""))
                except tk.TclError: pass

    def export_annotated_image(self):
        """Exports the currently displayed image with overlays from MeasurementModel."""
        if not self.app.img_original:
            messagebox.showerror("Export Error", "No image loaded to export.", parent=self.app.root)
            return

        img_to_export = (self.app.img_filtered if self.app.img_filtered is not None else self.app.img_original).copy()
        if img_to_export.mode not in ('RGB', 'RGBA'): img_to_export = img_to_export.convert('RGBA')
        elif img_to_export.mode == 'RGB': img_to_export = img_to_export.convert('RGBA')
        draw = ImageDraw.Draw(img_to_export)
        dot_radius, line_width = 3, 2
        model = self.app.measurement_model

        for x, y in model.calibration_dots:
            draw.ellipse((x - dot_radius, y - dot_radius, x + dot_radius, y + dot_radius), fill="cyan", outline="black")
        for i in range(len(model.artery_dots)):
            x, y = model.artery_dots[i]
            draw.ellipse((x - dot_radius, y - dot_radius, x + dot_radius, y + dot_radius), fill="yellow", outline="black")
            if i % 2 == 1: draw.line([(model.artery_dots[i-1]), (x,y)], fill="yellow", width=line_width)
        for i in range(len(model.line_points)):
            x, y = model.line_points[i]
            draw.ellipse((x - dot_radius, y - dot_radius, x + dot_radius, y + dot_radius), fill="magenta", outline="black")
            if i % 2 == 1: draw.line([(model.line_points[i-1]), (x,y)], fill="magenta", width=line_width)
        if model.angle_points:
            for p in model.angle_points:
                draw.ellipse((p[0]-dot_radius, p[1]-dot_radius, p[0]+dot_radius, p[1]+dot_radius), fill="lime green", outline="black")
            if len(model.angle_points) > 1: draw.line(model.angle_points, fill="lime green", width=line_width, joint="curve")
        
        tick_radius = 2 # For line mode visualization ticks
        for p1, p2 in model.line_measurement_visualization_points: # Get from model
             draw.ellipse((p1[0]-tick_radius, p1[1]-tick_radius, p1[0]+tick_radius, p1[1]+tick_radius), fill="red", outline="red")
             draw.line([p1, p2], fill="red", width=1)

        base_name = os.path.splitext(os.path.basename(self.app.file_path))[0] if self.app.file_path else "image"
        save_path = filedialog.asksaveasfilename(
            title="Save Annotated Image As", initialfile=f"{base_name}_annotated.png", defaultextension=".png",
            filetypes=[("PNG", "*.png"), ("JPEG", "*.jpg"), ("BMP", "*.bmp"), ("TIFF", "*.tif"), ("All", "*.*")],
            parent=self.app.root
        )
        if save_path:
            try:
                save_format = os.path.splitext(save_path)[1].lower()
                final_image_to_save = img_to_export
                if save_format in [".jpg", ".jpeg"] and final_image_to_save.mode == 'RGBA':
                     bg = Image.new("RGB", final_image_to_save.size, (255,255,255))
                     bg.paste(final_image_to_save, mask=final_image_to_save.split()[3])
                     final_image_to_save = bg
                elif save_format in [".jpg", ".jpeg"] and final_image_to_save.mode != 'RGB':
                     final_image_to_save = final_image_to_save.convert('RGB')
                final_image_to_save.save(save_path)
                messagebox.showinfo("Export Successful", f"Annotated image saved to:\n{save_path}", parent=self.app.root)
            except Exception as e:
                messagebox.showerror("Export Error", f"Failed to save image:\n{e}", parent=self.app.root)
                print(traceback.format_exc())

    def save_measurements_to_json(self):
        """Saves data from MeasurementModel and app metadata to a JSON file."""
        if not self.app.file_path:
            messagebox.showerror("Save Error", "No image loaded.", parent=self.app.root); return
        
        model = self.app.measurement_model
        if not model.get_all_measurements() and not model.calibration_done: # Check if any data exists
             if not messagebox.askokcancel("Save?", "No measurements or calibration recorded. Save empty file?", parent=self.app.root):
                 return

        name = self.app.name_var.get().strip()
        diameter_str = self.app.diameter_var.get().strip()
        if not name: messagebox.showerror("Input Error", "Please enter a 'Name'.", parent=self.app.root); self.app.name_entry.focus_set(); return
        if not diameter_str: messagebox.showerror("Input Error", "Please enter 'Real Ø (mm)'.", parent=self.app.root); self.app.diameter_entry.focus_set(); return
        try:
            real_diameter = float(diameter_str)
            if real_diameter <= 0: raise ValueError("Diameter must be positive")
        except ValueError:
            messagebox.showerror("Input Error", "Valid positive number for 'Real Ø (mm)'.", parent=self.app.root); self.app.diameter_entry.focus_set(); return

        current_time = datetime.datetime.now()
        time_str = current_time.strftime("%Y-%m-%d_%H-%M-%S")
        
        # Get model state, which includes calibration and measurements log
        model_data_dict = model.to_dict()

        data_to_save = {
            "metadata": {
                "source_image_path": self.app.file_path,
                "source_image_name": os.path.basename(self.app.file_path),
                "analysis_name": name,
                "analysis_timestamp_iso": current_time.isoformat(),
                "expected_real_diameter_mm": real_diameter,
                "software_version": "ImageAnalyzer_Refactored_v1" 
            },
            "calibration_details": { # Extracted from model_data_dict for clarity
                "calibrated": model_data_dict.get("calibration_done", False),
                "pixels_per_mm": model_data_dict.get("calibration_factor"),
                "calibration_points_original_px": model_data_dict.get("calibration_dots"),
                "calibration_reference_distance_mm": model_data_dict.get("calibration_reference_distance_mm")
            },
            "measurements_log": model_data_dict.get("measurements_log", []) # The actual log
        }
        
        default_filename = f"analysis_{name}_{time_str}.json"
        filename = filedialog.asksaveasfilename(
            title="Save Analysis Data As", initialfile=default_filename, defaultextension=".json",
            filetypes=[("JSON Files", "*.json"), ("All Files", "*.*")], parent=self.app.root
        )
        if filename:
            try:
                with open(filename, 'w') as f:
                    json.dump(data_to_save, f, indent=4)
                messagebox.showinfo("Save Successful", f"Analysis data saved to:\n{filename}", parent=self.app.root)
            except Exception as e:
                messagebox.showerror("Save Error", f"Failed to save JSON file:\n{e}", parent=self.app.root)
                print(traceback.format_exc())

    def show_line_measurements(self):
        """Displays detailed results of the last Line Mode measurement from MeasurementModel."""
        model = self.app.measurement_model
        line_measurement = None
        for meas in reversed(model.get_all_measurements()):
            if meas.get("type") == "line_measurement": # Updated type
                line_measurement = meas
                break
        if not line_measurement:
            messagebox.showinfo("No Line Measurement", "No Line Mode measurement found.", parent=self.app.root); return

        window = Toplevel(self.app.root); window.title("Line Mode - Details"); window.geometry("450x400"); window.transient(self.app.root); window.grab_set()
        text_frame = tk.Frame(window); text_frame.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)
        text_widget = tk.Text(text_frame, height=15, width=50, wrap=tk.WORD, font=("Courier New", 10))
        scrollbar = Scrollbar(text_frame, orient=tk.VERTICAL, command=text_widget.yview); text_widget.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y); text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        text_widget.insert(tk.END, "--- Line Mode Summary ---\n\n")
        l1px = line_measurement.get('length1_px'); l2px = line_measurement.get('length2_px'); angle_dev = line_measurement.get('angle_deviation_deg')
        text_widget.insert(tk.END, f"Line 1 Len: {l1px:>8.2f} px" if l1px is not None else "N/A")
        if model.calibration_done and 'length1_mm' in line_measurement: text_widget.insert(tk.END, f"  ({line_measurement['length1_mm']:.3f} mm)\n")
        else: text_widget.insert(tk.END, "\n")
        text_widget.insert(tk.END, f"Line 2 Len: {l2px:>8.2f} px" if l2px is not None else "N/A")
        if model.calibration_done and 'length2_mm' in line_measurement: text_widget.insert(tk.END, f"  ({line_measurement['length2_mm']:.3f} mm)\n")
        else: text_widget.insert(tk.END, "\n")
        text_widget.insert(tk.END, f"Angle Dev: {angle_dev:>7.2f}°\n\n" if angle_dev is not None else "N/A\n\n")

        dists_px = line_measurement.get('distances_px', []) # Should be a list from model
        dists_mm = line_measurement.get('distances_mm', [])
        text_widget.insert(tk.END, "--- Sampled Distances ---\n # |   Pixels  |    mm\n---|-----------|-----------\n")
        for i, d_px in enumerate(dists_px):
             line = f"{i+1:>2} | {d_px:>9.2f} |"
             if model.calibration_done and i < len(dists_mm): line += f" {dists_mm[i]:>9.3f}\n"
             else: line += "   N/A\n"
             text_widget.insert(tk.END, line)
        
        avg_px = line_measurement.get('avg_distance_px', 0)
        min_px = min(dists_px) if dists_px else 0
        max_px = max(dists_px) if dists_px else 0
        text_widget.insert(tk.END, "\n--- Stats (Sampled) ---\n")
        text_widget.insert(tk.END, f"Avg Dist: {avg_px:>8.2f} px")
        if model.calibration_done and 'avg_distance_mm' in line_measurement: text_widget.insert(tk.END, f"  ({line_measurement['avg_distance_mm']:.3f} mm)\n")
        else: text_widget.insert(tk.END, "\n")
        # ... (min/max stats similarly) ...

        text_widget.config(state=tk.DISABLED)
        tk.Button(window, text="Close", command=window.destroy).pack(pady=5)
        window.wait_window()
