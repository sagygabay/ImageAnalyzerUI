# Image Analyzer GUI

## Project Overview

The Image Analyzer GUI is a Python application built with Tkinter for performing various image analysis tasks, including measurements (distance, angle, parallel lines), calibration, and image filtering (e.g., Canny edge detection). The application allows users to load images, apply operations, and save both annotated images and measurement data.

The project is structured into several modules to separate concerns:

-   `main.py`: Contains the main `ImageAnalyzer` class, GUI setup, event loop, and orchestration logic.
-   `filters.py`: Manages image filtering operations through the `FilterHandler` class.
-   `models.py`: Defines the `MeasurementModel` class for storing and managing all annotation, measurement, and calibration data.
-   `modes.py`: Contains the `ModeHandler` class for managing different user interaction modes (e.g., calibration mode, artery mode).
-   `utils.py`: Provides utility functions through the `AppUtils` class for tasks like UI updates, data saving, and image export.

## Folder Structure

```
gui_for_image_ana/
├── main.py             # Main application runner and GUI logic
├── filters.py          # Image filtering handlers
├── models.py           # Data models for measurements and annotations
├── modes.py            # User interaction mode management
├── utils.py            # Utility functions (UI updates, saving, export)
├── README.md           # This file
├── requirements.txt    # Project dependencies (to be created)
├── config.json         # Example configuration file (to be created)
├── .gitignore          # Git ignore file (to be created)
└── tests/              # Unit tests (to be created)
    ├── __init__.py
    ├── test_filters.py
    └── test_models.py
```

## Installation

1.  **Clone the repository (if applicable) or ensure all project files are in a directory.**

2.  **Create a virtual environment (recommended):**
    ```bash
    python -m venv venv
    source venv/bin/activate  # On Windows: venv\Scripts\activate
    ```

3.  **Install dependencies:**
    A `requirements.txt` file should be created listing all necessary packages.
    Example content for `requirements.txt`:
    ```
    Pillow>=9.0.0
    numpy>=1.20.0
    opencv-python>=4.5.0
    ttkthemes>=3.2.0  # Optional, for themed GUI
    ```
    Install using:
    ```bash
    pip install -r requirements.txt
    ```

## Usage

1.  **Navigate to the project directory:**
    ```bash
    cd path/to/gui_for_image_ana_project_root 
    ```
    (Note: If `gui_for_image_ana` is a sub-directory, you might run from its parent or adjust Python's module search path if running `main.py` directly from within `gui_for_image_ana`.)

2.  **Run the application:**
    Assuming you are in the directory containing the `gui_for_image_ana` package:
    ```bash
    python -m gui_for_image_ana.main
    ```
    Alternatively, if you are inside the `gui_for_image_ana` directory itself, you might need to adjust imports or run as a module from the parent directory. For simplicity, the command above is preferred if `gui_for_image_ana` is treated as a package.

3.  **Using the Application:**
    *   **Load Image:** Click "Load Image" to open an image file. Navigation (Next/Previous) is available via arrow keys.
    *   **Calibration:**
        *   Click "Calibrate" to enter calibration mode.
        *   Click two points on the image.
        *   Enter the known real-world distance between these points (e.g., in mm).
        *   "Reset Calibration" clears the current calibration.
        *   "Keep Dots Fixed" will preserve calibration dots when changing images if toggled ON.
    *   **Dots Mode (Distance/Angle Measurement):**
        *   Click "Dots Mode".
        *   Click pairs of points to measure distance and angle. Results are logged.
        *   "Reset Dots" clears all dots and measurements in this mode.
        *   "Delete Last Pair" removes the last placed pair of dots.
    *   **Angle Mode:**
        *   Click "Angle Mode".
        *   Click three points (P1, Vertex, P2) to define and measure an angle.
        *   "Reset Angle" clears current angle points and logged angle measurements.
    *   **Line Mode (Parallel Lines):**
        *   Click "Line Mode".
        *   Click four points to define two lines. The application will calculate distances between them.
        *   "Reset Lines" clears points and measurements for this mode.
        *   "Show Line Measurements" displays detailed sampled distances.
    *   **Filters (Canny Edge Detection):**
        *   **Global Canny Filter:** Applies Canny edge detection to the entire image. Thresholds can be adjusted with sliders.
        *   **Canny ROI Selection:** Allows drawing a rectangle (Region of Interest) to apply Canny filter only to that area.
        *   Adjust "Low" and "High" Canny thresholds using the sliders. Changes apply live.
        *   "Reset Filters" removes all filter effects and ROI selections.
    *   **Zoom:**
        *   Use mouse wheel over the image to zoom in/out centered on the cursor.
        *   Use "+" and "-" keys to zoom in/out centered on the image view.
        *   "Zoom In Box": Toggles a magnified view of the area around the mouse cursor.
    *   **Export Image:** Saves the currently displayed image (with annotations and filter effects) to a file.
    *   **Save Measurements:**
        *   Enter a "Name" for the analysis and the "Real Ø (mm)" (expected diameter, for context).
        *   Click "Save Measurements" to save all calibration data and logged measurements to a JSON file.
    *   **Undo/Redo:** Use Ctrl+Z / Ctrl+Y or buttons to undo/redo actions.

## Configuration

A `config.json` (or `.yaml`) file can be used for parameter defaults. (This file needs to be created and integrated).
Example `config.json`:
```json
{
  "default_canny_low_threshold": 50,
  "default_canny_high_threshold": 150,
  "zoom_box_factor": 4,
  "default_theme": "arc"
}
```
The application would need to be updated to load and use these settings.

## Development & Testing

-   **Unit Tests:** A `tests/` directory should contain unit tests for:
    -   `test_filters.py`: Testing filter logic (e.g., Canny application, ROI handling).
    -   `test_models.py`: Testing data management in `MeasurementModel` (e.g., adding points, calibration logic, measurement logging).
    Frameworks like `unittest` or `pytest` can be used.

    Example `pytest` command (run from project root):
    ```bash
    pytest
    ```

---

This README provides a starting point and should be expanded as the project evolves.
