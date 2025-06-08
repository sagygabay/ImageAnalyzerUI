import math

class MeasurementModel:
    """
    Manages all measurement and annotation data for the Image Analyzer.
    This includes calibration data, points for various measurement modes,
    and the list of completed measurements.
    """
    def __init__(self):
        """
        Initializes the data model with empty structures.
        """
        self.reset_all_data()

    def reset_all_data(self, keep_calibration=False):
        """
        Resets all measurement and point data.

        Args:
            keep_calibration (bool): If True, calibration data is preserved.
                                     Otherwise, it's also reset.
        """
        if not keep_calibration:
            self.calibration_dots = []  # List of (x, y) tuples for calibration
            self.calibration_factor = 1.0  # Pixels per unit (e.g., mm)
            self.calibration_done = False
            self.calibration_reference_distance_mm = None # The real distance used for calibration

        self.artery_dots = []  # List of (x, y) tuples for artery/distance pairs
        self.angle_points = []  # List of (x, y) tuples for angle measurement (max 3)
        self.line_points = []   # List of (x, y) tuples for line mode (max 4)
        
        # Stores finalized measurements. Each entry is a dictionary.
        # Example: {"type": "artery", "points": [...], "distance_px": ..., "distance_mm": ...}
        # Example: {"type": "angle", "points": [...], "angle_deg": ...}
        # Example: {"type": "line", "points": [...], "distances_px": [...], ...}
        self.measurements_log = []

        # Temporary storage for visualization, e.g., ticks for line mode
        self.line_measurement_visualization_points = []

    def add_calibration_dot(self, point_orig):
        """
        Adds a point for calibration. Max 2 points.

        Args:
            point_orig (tuple): (x, y) coordinate in original image space.
        
        Returns:
            int: Number of calibration dots currently stored.
        """
        if len(self.calibration_dots) < 2:
            self.calibration_dots.append(point_orig)
        return len(self.calibration_dots)

    def set_calibration(self, distance_px, real_distance_mm):
        """
        Sets the calibration factor based on pixel distance and real-world distance.

        Args:
            distance_px (float): The measured distance in pixels.
            real_distance_mm (float): The corresponding real-world distance in mm.
        
        Returns:
            bool: True if calibration was set successfully, False otherwise.
        """
        if real_distance_mm > 0 and distance_px > 0:
            self.calibration_factor = distance_px / real_distance_mm
            self.calibration_done = True
            self.calibration_reference_distance_mm = real_distance_mm
            # Log calibration event
            self.log_measurement({
                "type": "calibration_set",
                "points": self.calibration_dots[:], # Store a copy
                "distance_px": round(distance_px, 2),
                "real_value_mm": round(real_distance_mm, 3),
                "calibration_factor": round(self.calibration_factor, 6)
            })
            return True
        self.calibration_done = False
        return False

    def reset_calibration(self):
        """Resets only the calibration data."""
        self.calibration_dots = []
        self.calibration_factor = 1.0
        self.calibration_done = False
        self.calibration_reference_distance_mm = None
        # Optionally remove calibration_set from log, or keep for history
        self.measurements_log = [m for m in self.measurements_log if m.get("type") != "calibration_set"]


    def add_artery_dot(self, point_orig):
        """
        Adds a point for artery/distance measurement. Pairs of points form a measurement.

        Args:
            point_orig (tuple): (x, y) coordinate in original image space.
        """
        self.artery_dots.append(point_orig)
        if len(self.artery_dots) % 2 == 0:
            p1 = self.artery_dots[-2]
            p2 = self.artery_dots[-1]
            dist_px = math.sqrt((p2[0] - p1[0])**2 + (p2[1] - p1[1])**2)
            angle_rad = math.atan2(-(p2[1] - p1[1]), p2[0] - p1[0]) # Y is often inverted in graphics
            angle_deg = math.degrees(angle_rad)
            if angle_deg < 0: angle_deg += 360

            measurement_data = {
                "type": "artery_distance",
                "points": [p1, p2],
                "distance_px": round(dist_px, 2),
                "angle_deg": round(angle_deg, 1)
            }
            if self.calibration_done:
                measurement_data["distance_mm"] = round(dist_px / self.calibration_factor, 3)
            self.log_measurement(measurement_data)

    def delete_last_artery_pair(self):
        """Deletes the last pair of artery dots and their corresponding measurement."""
        if len(self.artery_dots) >= 2:
            removed_p2 = self.artery_dots.pop()
            removed_p1 = self.artery_dots.pop()
            # Remove the corresponding measurement from the log
            self.measurements_log = [
                m for m in self.measurements_log 
                if not (m.get("type") == "artery_distance" and 
                        m.get("points") == [removed_p1, removed_p2])
            ]
            return True
        elif len(self.artery_dots) == 1: # Single pending dot
            self.artery_dots.pop()
            return True
        return False

    def reset_artery_dots_and_measurements(self):
        """Clears all artery dots and their logged measurements."""
        self.artery_dots = []
        self.measurements_log = [m for m in self.measurements_log if m.get("type") != "artery_distance"]

    def add_angle_point(self, point_orig):
        """
        Adds a point for angle measurement. Three points (P1, Vertex, P2) form an angle.

        Args:
            point_orig (tuple): (x, y) coordinate in original image space.
        
        Returns:
            int: Number of angle points currently stored for the active angle.
        """
        if len(self.angle_points) >= 3: # Start new angle if 3 points already exist
            self.angle_points = []
        
        self.angle_points.append(point_orig)

        if len(self.angle_points) == 3:
            p1, vertex, p2 = self.angle_points
            v1 = (p1[0] - vertex[0], p1[1] - vertex[1])
            v2 = (p2[0] - vertex[0], p2[1] - vertex[1])
            
            dot_product = v1[0] * v2[0] + v1[1] * v2[1]
            mag_v1 = math.sqrt(v1[0]**2 + v1[1]**2)
            mag_v2 = math.sqrt(v2[0]**2 + v2[1]**2)

            if mag_v1 == 0 or mag_v2 == 0:
                angle_deg = 0.0 # Or handle as error
            else:
                cos_angle = dot_product / (mag_v1 * mag_v2)
                cos_angle = max(-1.0, min(1.0, cos_angle)) # Clamp for precision issues
                angle_deg = math.degrees(math.acos(cos_angle))
            
            self.log_measurement({
                "type": "angle",
                "points": self.angle_points[:], # Store a copy
                "angle_deg": round(angle_deg, 2)
            })
        return len(self.angle_points)

    def reset_angle_points_and_measurements(self):
        """Clears current angle points and all logged angle measurements."""
        self.angle_points = []
        self.measurements_log = [m for m in self.measurements_log if m.get("type") != "angle"]

    def add_line_mode_point(self, point_orig):
        """
        Adds a point for line mode. Four points define two lines for parallel distance measurement.

        Args:
            point_orig (tuple): (x, y) coordinate in original image space.
        
        Returns:
            int: Number of line mode points currently stored.
        """
        if len(self.line_points) >= 4: # Start new line set
            self.line_points = []
            self.line_measurement_visualization_points = []

        self.line_points.append(point_orig)

        if len(self.line_points) == 4:
            self._calculate_and_log_line_measurement()
        
        return len(self.line_points)

    def _calculate_and_log_line_measurement(self):
        """Internal: Calculates and logs measurements for the current 4 line points."""
        if len(self.line_points) != 4:
            return

        p1, p2 = self.line_points[0], self.line_points[1] # First line
        p3, p4 = self.line_points[2], self.line_points[3] # Second line

        len1_px = math.sqrt((p2[0] - p1[0])**2 + (p2[1] - p1[1])**2)
        len2_px = math.sqrt((p4[0] - p3[0])**2 + (p4[1] - p3[1])**2)

        # Angle between lines
        v_line1 = (p2[0] - p1[0], p2[1] - p1[1])
        v_line2 = (p4[0] - p3[0], p4[1] - p3[1])
        angle_dev_deg = 0
        if (v_line1[0]**2 + v_line1[1]**2 > 0) and (v_line2[0]**2 + v_line2[1]**2 > 0):
            unit_v1 = (v_line1[0]/len1_px, v_line1[1]/len1_px) if len1_px > 0 else (0,0)
            unit_v2 = (v_line2[0]/len2_px, v_line2[1]/len2_px) if len2_px > 0 else (0,0)
            dot = unit_v1[0]*unit_v2[0] + unit_v1[1]*unit_v2[1]
            dot = max(-1.0, min(1.0, dot)) # Clamp
            angle_dev_deg = math.degrees(math.acos(dot))
            if angle_dev_deg > 90: angle_dev_deg = 180 - angle_dev_deg # Smallest angle

        # Perpendicular distances (simplified: average distance for now)
        # For a more robust solution, consider multiple sample points as in original app
        # Here, we'll take distance from midpoint of line1 to line2 for simplicity
        mid_p1 = ((p1[0]+p2[0])/2, (p1[1]+p2[1])/2)
        
        # Distance from point mid_p1 to line (p3, p4)
        # Formula: |(x0-x1)(y2-y1) - (y0-y1)(x2-x1)| / sqrt((x2-x1)^2 + (y2-y1)^2)
        # where (x0,y0) is mid_p1, and line is (x1,y1)-(x2,y2) i.e. p3-p4
        num = abs((mid_p1[0]-p3[0])*(p4[1]-p3[1]) - (mid_p1[1]-p3[1])*(p4[0]-p3[0]))
        den = math.sqrt((p4[0]-p3[0])**2 + (p4[1]-p3[1])**2)
        avg_dist_px = num / den if den > 0 else 0
        
        # For visualization, we'd need the projection point on line 2
        # This is a simplified version for the model. The app can enhance this.
        self.line_measurement_visualization_points = [] # Clear old ones
        # Example: self.line_measurement_visualization_points.append((mid_p1, projected_point_on_line2))

        measurement_data = {
            "type": "line_measurement",
            "points": self.line_points[:],
            "length1_px": round(len1_px, 2),
            "length2_px": round(len2_px, 2),
            "angle_deviation_deg": round(angle_dev_deg, 1),
            "avg_distance_px": round(avg_dist_px, 2),
            "distances_px": [round(avg_dist_px, 2)] # Placeholder for multiple distances
        }
        if self.calibration_done:
            measurement_data["length1_mm"] = round(len1_px / self.calibration_factor, 3)
            measurement_data["length2_mm"] = round(len2_px / self.calibration_factor, 3)
            measurement_data["avg_distance_mm"] = round(avg_dist_px / self.calibration_factor, 3)
            measurement_data["distances_mm"] = [round(avg_dist_px / self.calibration_factor, 3)]

        self.log_measurement(measurement_data)

    def reset_line_points_and_measurements(self):
        """Clears current line points and all logged line measurements."""
        self.line_points = []
        self.line_measurement_visualization_points = []
        self.measurements_log = [m for m in self.measurements_log if m.get("type") != "line_measurement"]

    def log_measurement(self, measurement_data):
        """
        Adds a finalized measurement to the log.

        Args:
            measurement_data (dict): The dictionary containing measurement details.
        """
        self.measurements_log.append(measurement_data)

    def get_all_measurements(self):
        """
        Returns a copy of all logged measurements.

        Returns:
            list: A list of measurement dictionaries.
        """
        return [m.copy() for m in self.measurements_log]

    def clear_all_measurements(self):
        """Clears all logged measurements, but keeps point lists for active modes."""
        self.measurements_log = []

    # --- Persistence (Save/Load) ---
    # These are placeholders. The main app currently handles JSON saving.
    # This model could be extended to handle its own state persistence.

    def to_dict(self):
        """
        Serializes the current model state to a dictionary.

        Returns:
            dict: A dictionary representing the model's state.
        """
        return {
            "calibration_dots": self.calibration_dots,
            "calibration_factor": self.calibration_factor,
            "calibration_done": self.calibration_done,
            "calibration_reference_distance_mm": self.calibration_reference_distance_mm,
            "artery_dots": self.artery_dots, # Current, possibly incomplete
            "angle_points": self.angle_points, # Current, possibly incomplete
            "line_points": self.line_points,   # Current, possibly incomplete
            "measurements_log": self.measurements_log
        }

    def from_dict(self, data_dict):
        """
        Restores the model state from a dictionary.

        Args:
            data_dict (dict): A dictionary representing the model's state.
        """
        self.calibration_dots = data_dict.get("calibration_dots", [])
        self.calibration_factor = data_dict.get("calibration_factor", 1.0)
        self.calibration_done = data_dict.get("calibration_done", False)
        self.calibration_reference_distance_mm = data_dict.get("calibration_reference_distance_mm")
        
        self.artery_dots = data_dict.get("artery_dots", [])
        self.angle_points = data_dict.get("angle_points", [])
        self.line_points = data_dict.get("line_points", [])
        
        self.measurements_log = data_dict.get("measurements_log", [])
        
        # Reset visualization points as they are usually transient
        self.line_measurement_visualization_points = []

# Example usage (for testing or if run directly)
if __name__ == "__main__":
    model = MeasurementModel()
    
    # Simulate calibration
    model.add_calibration_dot((10, 10))
    model.add_calibration_dot((110, 10)) # 100 pixels
    model.set_calibration(100.0, 10.0) # 100px = 10mm => factor = 10 px/mm
    print(f"Calibration done: {model.calibration_done}, Factor: {model.calibration_factor}")

    # Simulate artery measurement
    model.add_artery_dot((20, 20))
    model.add_artery_dot((20, 70)) # 50 pixels
    
    # Simulate angle measurement
    model.add_angle_point((0,0))
    model.add_angle_point((50,0)) # Vertex
    model.add_angle_point((50,50)) # Should be 90 degrees

    print("\nAll Measurements Logged:")
    for meas in model.get_all_measurements():
        print(meas)

    print("\nModel state as dict:")
    print(model.to_dict())
