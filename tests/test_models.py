import unittest
import math

# Adjust import path based on how tests are run
from gui_for_image_ana.models import MeasurementModel

class TestMeasurementModel(unittest.TestCase):

    def setUp(self):
        """Set up for each test."""
        self.model = MeasurementModel()

    def test_initialization(self):
        """Test model initializes with empty data structures."""
        self.assertEqual(self.model.calibration_dots, [])
        self.assertEqual(self.model.calibration_factor, 1.0)
        self.assertFalse(self.model.calibration_done)
        self.assertIsNone(self.model.calibration_reference_distance_mm)
        self.assertEqual(self.model.artery_dots, [])
        self.assertEqual(self.model.angle_points, [])
        self.assertEqual(self.model.line_points, [])
        self.assertEqual(self.model.measurements_log, [])
        self.assertEqual(self.model.line_measurement_visualization_points, [])

    def test_reset_all_data(self):
        """Test resetting all data."""
        self.model.add_calibration_dot((1,1))
        self.model.set_calibration(10, 1) # factor = 10
        self.model.add_artery_dot((2,2))
        self.model.log_measurement({"type": "test"})
        
        self.model.reset_all_data()
        self.test_initialization() # Should be back to initial state

    def test_reset_all_data_keep_calibration(self):
        """Test resetting data while keeping calibration."""
        self.model.add_calibration_dot((10,10))
        self.model.add_calibration_dot((110,10)) # 100px
        self.model.set_calibration(100.0, 10.0) # factor = 10
        
        calib_dots_before = self.model.calibration_dots[:]
        calib_factor_before = self.model.calibration_factor
        calib_done_before = self.model.calibration_done
        calib_ref_dist_before = self.model.calibration_reference_distance_mm
        
        self.model.add_artery_dot((5,5)) # Add some other data
        
        self.model.reset_all_data(keep_calibration=True)
        
        self.assertEqual(self.model.calibration_dots, calib_dots_before)
        self.assertEqual(self.model.calibration_factor, calib_factor_before)
        self.assertEqual(self.model.calibration_done, calib_done_before)
        self.assertEqual(self.model.calibration_reference_distance_mm, calib_ref_dist_before)
        self.assertEqual(self.model.artery_dots, []) # Other data should be reset
        # Check measurements_log for calibration_set persistence if desired
        self.assertTrue(any(m['type'] == 'calibration_set' for m in self.model.measurements_log))


    def test_add_calibration_dot(self):
        """Test adding calibration dots."""
        self.assertEqual(self.model.add_calibration_dot((10, 10)), 1)
        self.assertEqual(self.model.calibration_dots, [(10, 10)])
        self.assertEqual(self.model.add_calibration_dot((20, 20)), 2)
        self.assertEqual(self.model.calibration_dots, [(10, 10), (20, 20)])
        # Test adding more than 2 (should not add)
        self.assertEqual(self.model.add_calibration_dot((30, 30)), 2)
        self.assertEqual(self.model.calibration_dots, [(10, 10), (20, 20)])

    def test_set_calibration(self):
        """Test setting calibration factor."""
        self.model.add_calibration_dot((0,0))
        self.model.add_calibration_dot((100,0)) # 100px distance
        
        self.assertTrue(self.model.set_calibration(100.0, 10.0)) # 10mm real distance
        self.assertTrue(self.model.calibration_done)
        self.assertEqual(self.model.calibration_factor, 10.0) # 100px / 10mm = 10 px/mm
        self.assertEqual(self.model.calibration_reference_distance_mm, 10.0)
        self.assertEqual(len(self.model.measurements_log), 1)
        self.assertEqual(self.model.measurements_log[0]['type'], 'calibration_set')

        # Test invalid calibration
        self.assertFalse(self.model.set_calibration(100.0, 0)) # Zero real distance
        self.assertFalse(self.model.calibration_done) # Should reset done flag

    def test_reset_calibration(self):
        """Test resetting calibration."""
        self.model.add_calibration_dot((0,0)); self.model.add_calibration_dot((10,0))
        self.model.set_calibration(10,1)
        self.assertTrue(self.model.calibration_done)
        
        self.model.reset_calibration()
        self.assertEqual(self.model.calibration_dots, [])
        self.assertEqual(self.model.calibration_factor, 1.0)
        self.assertFalse(self.model.calibration_done)
        self.assertIsNone(self.model.calibration_reference_distance_mm)
        self.assertFalse(any(m['type'] == 'calibration_set' for m in self.model.measurements_log))


    def test_add_artery_dot_and_measurement(self):
        """Test adding artery dots and automatic measurement logging."""
        self.model.add_artery_dot((0, 0))
        self.assertEqual(len(self.model.artery_dots), 1)
        self.assertEqual(len(self.model.measurements_log), 0) # No measurement yet

        self.model.add_artery_dot((0, 50)) # 50px distance
        self.assertEqual(len(self.model.artery_dots), 2)
        self.assertEqual(len(self.model.measurements_log), 1)
        
        last_meas = self.model.measurements_log[0]
        self.assertEqual(last_meas['type'], 'artery_distance')
        self.assertEqual(last_meas['points'], [(0,0), (0,50)])
        self.assertAlmostEqual(last_meas['distance_px'], 50.0)
        self.assertAlmostEqual(last_meas['angle_deg'], 270.0) # Straight down

        # Test with calibration
        self.model.set_calibration(10.0, 1.0) # 10 px/mm
        self.model.add_artery_dot((10,0))
        self.model.add_artery_dot((10,20)) # 20px distance
        
        self.assertEqual(len(self.model.measurements_log), 3) # 1 calib_set + 2 artery_distance
        last_artery_meas = self.model.measurements_log[2]
        self.assertAlmostEqual(last_artery_meas['distance_px'], 20.0)
        self.assertAlmostEqual(last_artery_meas['distance_mm'], 2.0) # 20px / 10px/mm

    def test_delete_last_artery_pair(self):
        """Test deleting the last artery pair."""
        self.model.add_artery_dot((0,0)); self.model.add_artery_dot((0,10)) # Pair 1
        self.model.add_artery_dot((5,5)); self.model.add_artery_dot((5,15)) # Pair 2
        self.assertEqual(len(self.model.artery_dots), 4)
        self.assertEqual(len(self.model.measurements_log), 2)

        self.assertTrue(self.model.delete_last_artery_pair())
        self.assertEqual(len(self.model.artery_dots), 2)
        self.assertEqual(len(self.model.measurements_log), 1)
        self.assertEqual(self.model.measurements_log[0]['points'], [(0,0), (0,10)])

        self.assertTrue(self.model.delete_last_artery_pair())
        self.assertEqual(len(self.model.artery_dots), 0)
        self.assertEqual(len(self.model.measurements_log), 0)

        self.assertFalse(self.model.delete_last_artery_pair()) # Nothing to delete

    def test_add_angle_point_and_measurement(self):
        """Test adding angle points and measurement logging."""
        self.assertEqual(self.model.add_angle_point((0,0)), 1)
        self.assertEqual(self.model.add_angle_point((50,0)), 2) # Vertex
        self.assertEqual(len(self.model.measurements_log), 0)

        self.assertEqual(self.model.add_angle_point((50,50)), 3) # 90 degree angle
        self.assertEqual(len(self.model.measurements_log), 1)
        last_meas = self.model.measurements_log[0]
        self.assertEqual(last_meas['type'], 'angle')
        self.assertEqual(last_meas['points'], [(0,0), (50,0), (50,50)])
        self.assertAlmostEqual(last_meas['angle_deg'], 90.0)

        # Test starting a new angle
        self.assertEqual(self.model.add_angle_point((100,100)), 1) # Clears old points, starts new
        self.assertEqual(self.model.angle_points, [(100,100)])
        self.assertEqual(len(self.model.measurements_log), 1) # Old measurement remains

    def test_add_line_mode_point_and_measurement(self):
        """Test adding line mode points and measurement logging."""
        self.assertEqual(self.model.add_line_mode_point((0,10)), 1)
        self.assertEqual(self.model.add_line_mode_point((100,10)), 2) # Line 1
        self.assertEqual(self.model.add_line_mode_point((0,30)), 3)
        self.assertEqual(len(self.model.measurements_log), 0)

        self.assertEqual(self.model.add_line_mode_point((100,30)), 4) # Line 2 (parallel to line 1)
        self.assertEqual(len(self.model.measurements_log), 1)
        
        last_meas = self.model.measurements_log[0]
        self.assertEqual(last_meas['type'], 'line_measurement')
        self.assertEqual(len(last_meas['points']), 4)
        self.assertAlmostEqual(last_meas['length1_px'], 100.0)
        self.assertAlmostEqual(last_meas['length2_px'], 100.0)
        self.assertAlmostEqual(last_meas['angle_deviation_deg'], 0.0) # Parallel
        self.assertAlmostEqual(last_meas['avg_distance_px'], 20.0) # Distance between y=10 and y=30

    def test_to_dict_and_from_dict_serialization(self):
        """Test serialization to and from dictionary."""
        self.model.add_calibration_dot((1,1)); self.model.add_calibration_dot((11,1))
        self.model.set_calibration(10, 1) # factor 10
        self.model.add_artery_dot((0,0)); self.model.add_artery_dot((0,5)) # 5px, 0.5mm
        
        data_dict = self.model.to_dict()
        
        new_model = MeasurementModel()
        new_model.from_dict(data_dict)
        
        self.assertEqual(new_model.calibration_dots, self.model.calibration_dots)
        self.assertEqual(new_model.calibration_factor, self.model.calibration_factor)
        self.assertEqual(new_model.calibration_done, self.model.calibration_done)
        self.assertEqual(new_model.artery_dots, self.model.artery_dots)
        self.assertEqual(len(new_model.measurements_log), len(self.model.measurements_log))
        # Could do a deeper comparison of measurements_log if needed
        self.assertEqual(new_model.measurements_log[0]['type'], 'calibration_set')
        self.assertEqual(new_model.measurements_log[1]['type'], 'artery_distance')


if __name__ == '__main__':
    unittest.main()
