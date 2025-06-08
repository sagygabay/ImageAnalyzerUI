import unittest
from PIL import Image, ImageChops
import numpy as np

# Adjust import path based on how tests are run
# If tests are run from the project root (e.g., `python -m unittest discover` or `pytest`)
from gui_for_image_ana.filters import FilterHandler
from gui_for_image_ana.main import ImageAnalyzer # Needed for app instance context

# If tests are run from within the tests directory, you might need:
# import sys
# sys.path.append('../') # Add parent directory to path
# from filters import FilterHandler
# from main import ImageAnalyzer


class MockApp:
    """A minimal mock of the ImageAnalyzer app for FilterHandler testing."""
    def __init__(self):
        self.img_original = None
        self.img_filtered = None
        self.canny_selection_mode = False # For _apply_canny_filter_internal logic
        self.zoom_box_mode = False
        self.zoom_box = None # Mock zoom_box canvas if needed for some updates
        self.measurement = MockStringVar() # Mock Tkinter StringVar
        self.canny_low = MockIntVar(100)
        self.canny_high = MockIntVar(200)
        self.buttons = {} # Mock buttons dict

    def display_image(self):
        # print("MockApp.display_image called")
        pass

    def update_zoom_box_content(self, event):
        # print("MockApp.update_zoom_box_content called")
        pass

class MockStringVar:
    def __init__(self, value=""):
        self._value = value
    def get(self):
        return self._value
    def set(self, value):
        self._value = value

class MockIntVar:
    def __init__(self, value=0):
        self._value = value
    def get(self):
        return self._value
    def set(self, value):
        self._value = value


class TestFilterHandler(unittest.TestCase):

    def setUp(self):
        """Set up for each test."""
        self.mock_app = MockApp()
        self.filter_handler = FilterHandler(self.mock_app)
        # Create a simple dummy RGBA image for testing
        self.test_image_pil = Image.new("RGBA", (100, 100), (128, 128, 128, 255))
        self.mock_app.img_original = self.test_image_pil.copy()

    def test_initialization(self):
        """Test FilterHandler initializes with default Canny thresholds."""
        self.assertEqual(self.filter_handler.canny_low_thresh, 100)
        self.assertEqual(self.filter_handler.canny_high_thresh, 200)
        self.assertIsNone(self.filter_handler.canny_roi_start_orig)
        self.assertFalse(self.filter_handler.is_global_canny_active)

    def test_update_canny_thresholds(self):
        """Test updating Canny thresholds also triggers apply_filters_and_update_display."""
        self.mock_app.display_image = unittest.mock.MagicMock() # Mock display_image
        self.filter_handler.update_canny_thresholds(50, 150)
        self.assertEqual(self.filter_handler.canny_low_thresh, 50)
        self.assertEqual(self.filter_handler.canny_high_thresh, 150)
        self.mock_app.display_image.assert_called_once()

    def test_set_canny_roi_original_coords(self):
        """Test setting Canny ROI coordinates."""
        self.filter_handler.set_canny_roi_original_coords((10, 20), (50, 60))
        self.assertEqual(self.filter_handler.canny_roi_start_orig, (10, 20))
        self.assertEqual(self.filter_handler.canny_roi_end_orig, (50, 60))

        # Test with inverted coordinates
        self.filter_handler.set_canny_roi_original_coords((50, 60), (10, 20))
        self.assertEqual(self.filter_handler.canny_roi_start_orig, (10, 20))
        self.assertEqual(self.filter_handler.canny_roi_end_orig, (50, 60))

        self.filter_handler.set_canny_roi_original_coords(None, None)
        self.assertIsNone(self.filter_handler.canny_roi_start_orig)

    def test_apply_global_canny_filter(self):
        """Test applying global Canny filter."""
        self.filter_handler.is_global_canny_active = True
        self.filter_handler.update_canny_thresholds(50,100) # Ensure specific thresholds
        
        # Create a simple image with a clear edge for Canny
        img_data = np.zeros((100, 100, 4), dtype=np.uint8)
        img_data[:, :50, :] = (0, 0, 0, 255)  # Black left half
        img_data[:, 50:, :] = (255, 255, 255, 255) # White right half
        self.mock_app.img_original = Image.fromarray(img_data, 'RGBA')

        self.filter_handler.apply_filters_and_update_display()
        
        self.assertIsNotNone(self.mock_app.img_filtered)
        # Further checks: verify that img_filtered contains Canny edges.
        # This might involve converting to numpy array and checking for non-zero pixels
        # where edges are expected.
        # For simplicity, we'll just check it's different from original.
        # A more robust test would compare against a pre-computed Canny output.
        self.assertFalse(ImageChops.difference(self.mock_app.img_original.convert("RGB"), self.mock_app.img_filtered.convert("RGB")).getbbox() is None, 
                         "Filtered image should be different from original after Canny.")

    def test_apply_roi_canny_filter(self):
        """Test applying Canny filter to an ROI."""
        self.filter_handler.set_canny_roi_original_coords((10, 10), (60, 60)) # ROI in original image space
        self.mock_app.canny_selection_mode = False # ROI selection is complete
        self.filter_handler.update_canny_thresholds(50,100)

        # Create an image where Canny should only find edges within the ROI
        img_data = np.zeros((100, 100, 4), dtype=np.uint8)
        img_data[20:50, 20:50, :] = (200, 200, 200, 255) # A square within ROI
        self.mock_app.img_original = Image.fromarray(img_data, 'RGBA')
        
        self.filter_handler.apply_filters_and_update_display()
        self.assertIsNotNone(self.mock_app.img_filtered)
        
        # More detailed checks would be needed here:
        # 1. Ensure the area outside the ROI in img_filtered is the same as img_original.
        # 2. Ensure the area inside the ROI in img_filtered has Canny edges.

    def test_reset_filter_states(self):
        """Test resetting filter states."""
        self.filter_handler.is_global_canny_active = True
        self.filter_handler.set_canny_roi_original_coords((10,10), (20,20))
        self.filter_handler.update_canny_thresholds(10,20)
        self.mock_app.img_filtered = self.test_image_pil.copy() # Simulate a filtered image

        self.filter_handler.reset_filter_states()

        self.assertFalse(self.filter_handler.is_global_canny_active)
        self.assertIsNone(self.filter_handler.canny_roi_start_orig)
        self.assertEqual(self.filter_handler.canny_low_thresh, 100) # Resets to default
        self.assertEqual(self.filter_handler.canny_high_thresh, 200) # Resets to default
        self.assertIsNone(self.mock_app.img_filtered) # Should be cleared in app

    # --- Placeholder tests for other filters (denoise, threshold, contrast) ---
    # def test_apply_denoise_gaussian(self):
    #     self.filter_handler.is_denoise_active = True
    #     self.filter_handler.denoise_params = {"type": "gaussian", "radius": 2}
    #     self.filter_handler.apply_filters_and_update_display()
    #     self.assertIsNotNone(self.mock_app.img_filtered)
    #     # Add assertions to check if GaussianBlur was applied

    # def test_apply_threshold_binary(self):
    #     self.filter_handler.is_threshold_active = True
    #     self.filter_handler.threshold_params = {"value": 100, "method": "binary"}
    #     self.filter_handler.apply_filters_and_update_display()
    #     self.assertIsNotNone(self.mock_app.img_filtered)
    #     # Add assertions for thresholding

    # def test_simple_invert_colors(self):
    #     from gui_for_image_ana.filters import simple_invert_colors # Import directly
    #     inverted_img = simple_invert_colors(self.test_image_pil)
    #     self.assertIsNotNone(inverted_img)
    #     # Check a few pixel values to confirm inversion
    #     original_pixel = self.test_image_pil.getpixel((0,0))
    #     inverted_pixel = inverted_img.getpixel((0,0))
    #     self.assertEqual(inverted_pixel[0], 255 - original_pixel[0])
    #     self.assertEqual(inverted_pixel[1], 255 - original_pixel[1])
    #     self.assertEqual(inverted_pixel[2], 255 - original_pixel[2])
    #     self.assertEqual(inverted_pixel[3], original_pixel[3]) # Alpha should be same


if __name__ == '__main__':
    unittest.main()
