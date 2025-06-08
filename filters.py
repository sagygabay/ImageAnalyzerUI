import cv2
import numpy as np
from PIL import Image, ImageFilter # ImageFilter can be used for other filters

class FilterHandler:
    def __init__(self, app_instance):
        """
        Manages image filtering operations and their states.

        Args:
            app_instance: The main application instance (e.g., ImageAnalyzer),
                          used for accessing the image and updating the display.
        """
        self.app = app_instance

        # Canny filter parameters and state
        self.canny_low_thresh = 100
        self.canny_high_thresh = 200
        self.canny_roi_start_orig = None  # (x1, y1) in original image coordinates
        self.canny_roi_end_orig = None    # (x2, y2) in original image coordinates
        self.is_global_canny_active = False
        # Note: is_canny_roi_selection_mode is managed by the main app's ModeHandler

        # --- Placeholder for other filter states and parameters ---
        # Example: Denoise
        # self.is_denoise_active = False
        # self.denoise_params = {"type": "gaussian", "radius": 1}

        # Example: Threshold
        # self.is_threshold_active = False
        # self.threshold_params = {"value": 128, "method": "binary"}

        # Example: Contrast
        # self.is_contrast_active = False
        # self.contrast_params = {"factor": 1.5}

    def update_canny_thresholds(self, low, high):
        """
        Updates the Canny edge detection thresholds. Called by UI elements.

        Args:
            low (int): The lower threshold.
            high (int): The higher threshold.
        """
        self.canny_low_thresh = low
        self.canny_high_thresh = high
        # Automatically re-apply filters when thresholds change
        self.apply_filters_and_update_display()

    def set_canny_roi_original_coords(self, start_coord_orig, end_coord_orig):
        """
        Sets the Region of Interest (ROI) for the Canny filter using original image coordinates.

        Args:
            start_coord_orig (tuple): (x1, y1) of the ROI in original image space.
            end_coord_orig (tuple): (x2, y2) of the ROI in original image space.
        """
        if start_coord_orig and end_coord_orig:
            self.canny_roi_start_orig = (
                min(start_coord_orig[0], end_coord_orig[0]),
                min(start_coord_orig[1], end_coord_orig[1])
            )
            self.canny_roi_end_orig = (
                max(start_coord_orig[0], end_coord_orig[0]),
                max(start_coord_orig[1], end_coord_orig[1])
            )
        else:
            self.canny_roi_start_orig = None
            self.canny_roi_end_orig = None
        # self.apply_filters_and_update_display() # Apply when ROI is finalized by app

    def _apply_canny_filter_internal(self, image_pil_rgba):
        """
        Internal method to apply Canny filter based on current state.

        Args:
            image_pil_rgba (PIL.Image.Image): Input RGBA image.

        Returns:
            PIL.Image.Image: Edges as an RGBA image (e.g., blue edges on transparent background),
                             or None if Canny is not applicable.
        """
        if not image_pil_rgba:
            return None

        img_cv_gray = np.array(image_pil_rgba.convert('L')) # Convert to grayscale for Canny
        edges_cv = np.zeros_like(img_cv_gray) # Initialize empty edge map

        apply_to_full_image = self.is_global_canny_active
        apply_to_roi = (self.canny_roi_start_orig and self.canny_roi_end_orig and
                        not self.app.canny_selection_mode) # Apply if ROI is set and not actively selecting

        if apply_to_full_image:
            edges_cv = cv2.Canny(img_cv_gray, self.canny_low_thresh, self.canny_high_thresh)
        elif apply_to_roi:
            x1, y1 = int(self.canny_roi_start_orig[0]), int(self.canny_roi_start_orig[1])
            x2, y2 = int(self.canny_roi_end_orig[0]), int(self.canny_roi_end_orig[1])

            h, w = img_cv_gray.shape
            x1, y1 = max(0, x1), max(0, y1)
            x2, y2 = min(w, x2), min(h, y2)

            if x1 >= x2 or y1 >= y2: # Invalid ROI
                return None # No Canny to apply

            roi_gray = img_cv_gray[y1:y2, x1:x2]
            if roi_gray.size == 0: return None # Empty ROI

            edges_roi_cv = cv2.Canny(roi_gray, self.canny_low_thresh, self.canny_high_thresh)
            edges_cv[y1:y2, x1:x2] = edges_roi_cv
        else:
            return None # Canny not active or ROI not finalized

        # Convert edges (white on black) to a colored overlay (e.g., blue on transparent)
        # Create a blue image for edges
        blue_edges_rgba = np.zeros((h, w, 4), dtype=np.uint8)
        blue_edges_rgba[edges_cv == 255] = [0, 0, 255, 255] # Blue color for edges, fully opaque

        return Image.fromarray(blue_edges_rgba, 'RGBA')

    def apply_filters_and_update_display(self, *args):
        """
        Applies all active filters to the original image and tells the main app to update its display.
        This method is typically called when a filter parameter changes or mode toggles.
        The *args parameter allows it to be used as a callback for Tkinter variable traces.
        """
        if not self.app.img_original:
            self.app.img_filtered = None
            self.app.display_image()
            return

        # Start with a fresh copy of the original image (ensure it's RGBA for compositing)
        current_image_pil = self.app.img_original.convert("RGBA")

        # 1. Apply Canny Filter (if active)
        # Thresholds are updated via update_canny_thresholds from app's IntVars
        canny_edge_overlay_pil = self._apply_canny_filter_internal(current_image_pil)
        if canny_edge_overlay_pil:
            # Alpha composite the Canny edges onto the current image
            current_image_pil = Image.alpha_composite(current_image_pil, canny_edge_overlay_pil)

        # 2. Apply Denoise Filter (Placeholder)
        # if self.is_denoise_active:
        #    current_image_pil = self.apply_denoise(current_image_pil, self.denoise_params)

        # 3. Apply Threshold Filter (Placeholder)
        # if self.is_threshold_active:
        #    current_image_pil = self.apply_threshold(current_image_pil, self.threshold_params)
        
        # 4. Apply Contrast Adjustment (Placeholder)
        # if self.is_contrast_active:
        #    current_image_pil = self.apply_contrast_adjustment(current_image_pil, self.contrast_params)

        self.app.img_filtered = current_image_pil
        
        # Update status message in the main app
        status_parts = []
        if self.is_global_canny_active:
            status_parts.append(f"GlobalCanny({self.canny_low_thresh},{self.canny_high_thresh})")
        elif self.canny_roi_start_orig and self.canny_roi_end_orig and not self.app.canny_selection_mode:
            status_parts.append(f"CannyROI({self.canny_low_thresh},{self.canny_high_thresh})")
        # Add other active filters to status_parts here

        if not status_parts:
            self.app.measurement.set("Status: Filters reset or no filter active.")
        else:
            self.app.measurement.set(f"Status: Active Filters: {', '.join(status_parts)}")

        self.app.display_image()
        if self.app.zoom_box_mode and self.app.zoom_box and self.app.zoom_box.winfo_exists():
            self.app.update_zoom_box_content(None) # Event can be None

    def reset_filter_states(self):
        """Resets all filter states and parameters to their defaults and updates display."""
        self.canny_low_thresh = 100
        self.canny_high_thresh = 200
        self.canny_roi_start_orig = None
        self.canny_roi_end_orig = None
        self.is_global_canny_active = False
        # self.is_denoise_active = False
        # etc. for other filters

        self.app.img_filtered = None # Clear the filtered image in the main app
        
        # Update UI elements in the main app (e.g., button states, slider values)
        if hasattr(self.app, 'canny_low') and hasattr(self.app, 'canny_high'):
            self.app.canny_low.set(self.canny_low_thresh)
            self.app.canny_high.set(self.canny_high_thresh)
        if "Global Canny" in self.app.buttons:
            try: self.app.buttons["Global Canny"].config(relief="raised")
            except Exception: pass # Ignore if button not ready
        if hasattr(self.app, 'image_canvas') and self.app.image_canvas.winfo_exists():
            self.app.image_canvas.delete("canny_rect") # Remove visual ROI

        # Trigger a re-display with no filters
        self.apply_filters_and_update_display()
        self.app.measurement.set("Status: Filters reset.")

    # --- Placeholder methods for other filters ---

    def apply_denoise(self, image_pil, params):
        """
        Applies a denoising filter to the image.

        Args:
            image_pil (PIL.Image.Image): The input image.
            params (dict): Parameters for denoising (e.g., {"type": "median", "kernel_size": 3}).
                           Supported types: "gaussian", "median".

        Returns:
            PIL.Image.Image: The denoised image.
        """
        filter_type = params.get("type", "gaussian")
        if filter_type == "gaussian":
            radius = params.get("radius", 1)
            return image_pil.filter(ImageFilter.GaussianBlur(radius=radius))
        elif filter_type == "median":
            kernel_size = params.get("kernel_size", 3)
            if kernel_size % 2 == 0: kernel_size += 1 # Must be odd
            return image_pil.filter(ImageFilter.MedianFilter(size=kernel_size))
        # Add more denoising types as needed
        return image_pil

    def apply_threshold(self, image_pil, params):
        """
        Applies thresholding to the image.

        Args:
            image_pil (PIL.Image.Image): The input image.
            params (dict): Parameters for thresholding (e.g., {"value": 128, "method": "binary"}).
                           Supported methods: "binary", "binary_inv".

        Returns:
            PIL.Image.Image: The thresholded image.
        """
        gray_image = image_pil.convert('L')
        threshold_value = params.get("value", 128)
        method = params.get("method", "binary")

        if method == "binary":
            return gray_image.point(lambda x: 255 if x > threshold_value else 0, mode='L').convert(image_pil.mode)
        elif method == "binary_inv":
            return gray_image.point(lambda x: 0 if x > threshold_value else 255, mode='L').convert(image_pil.mode)
        # Add more thresholding types (Otsu, adaptive) as needed
        return image_pil

    def apply_contrast_adjustment(self, image_pil, params):
        """
        Adjusts the contrast of the image.

        Args:
            image_pil (PIL.Image.Image): The input image.
            params (dict): Parameters for contrast adjustment (e.g., {"factor": 1.5}).

        Returns:
            PIL.Image.Image: The contrast-adjusted image.
        """
        from PIL import ImageEnhance
        factor = params.get("factor", 1.0) # 1.0 means no change
        enhancer = ImageEnhance.Contrast(image_pil)
        return enhancer.enhance(factor)

# Example of a standalone filter function (could also be in this file or a sub-module)
def simple_invert_colors(image_pil):
    """
    Inverts the colors of a PIL image.

    Args:
        image_pil (PIL.Image.Image): The input image.

    Returns:
        PIL.Image.Image: The inverted image.
    """
    if image_pil.mode == 'RGBA':
        r,g,b,a = image_pil.split()
        r = r.point(lambda i: 255 - i)
        g = g.point(lambda i: 255 - i)
        b = b.point(lambda i: 255 - i)
        return Image.merge('RGBA', (r,g,b,a))
    elif image_pil.mode == 'RGB':
        # For RGB, we can use PIL.ImageOps.invert
        from PIL import ImageOps
        return ImageOps.invert(image_pil.convert('RGB'))
    elif image_pil.mode == 'L':
        from PIL import ImageOps
        return ImageOps.invert(image_pil.convert('L'))
    else:
        # For other modes, attempt conversion to RGB, invert, then try to convert back
        # This might not always be ideal.
        try:
            original_mode = image_pil.mode
            rgb_image = image_pil.convert('RGB')
            from PIL import ImageOps
            inverted_rgb = ImageOps.invert(rgb_image)
            return inverted_rgb.convert(original_mode)
        except Exception:
            # Fallback if conversion path is tricky, just return original
            return image_pil
