> [!WARNING]
> This program is AI generated. Sometimes the OCR may make mistakes, such as capturing the number 6 as 5 or other numbers incorrectly. 

For better results:
- Adjust the match threshold slider a bit higher.
- Try replacing the template number images in the `img` folder with clearer samples.

Results may vary depending on image quality and template accuracy.

# Number Sum Calculator

A Python GUI application that captures a screen region, extracts numbers using OCR, and highlights pairs of numbers that sum to 10.
![Total Downloads](https://img.shields.io/github/downloads/thealtersky/nikke_azx/total)

![](/assets/capture.png)

## Features

- **Window Capture**: Press `Win + Shift + Z` or click the "Capture" button to start
- **Resizable Selection**: Drag to select any area of your screen
- **OCR Processing**: Automatically extracts numbers from the captured image
- **Smart Detection**: Finds all pairs of numbers that sum to 10
- **Visual Highlighting**: Highlights matching pairs with colored rectangles
- **Save Results**: Save the highlighted image to disk

## Installation

1. Install Python 3.8 or higher

2. Install the required packages:
```bash
pip install -r requirements.txt
```

3. Install Tesseract OCR:
   - **Windows**: Download and install from https://github.com/UB-Mannheim/tesseract/wiki
   - Add Tesseract to your PATH or update the code with the installation path
   - **Linux**: `sudo apt-get install tesseract-ocr`
   - **Mac**: `brew install tesseract`

## Usage

1. Run the application:
```bash
python main.py
```

2. Click the "Capture" button or press `Win + Shift + Z`

3. Drag to select the area containing numbers

4. Release to capture - the app will:
   - Extract numbers using OCR
   - Find pairs that sum to 10
   - Highlight them with colored rectangles
   - Display the result

5. Click "Save Result" to save the highlighted image

## Requirements

- Python 3.8+
- Tesseract OCR engine
- See `requirements.txt` for Python packages

## How It Works

1. **Capture**: Uses PIL ImageGrab to capture the selected screen region
2. **OCR**: Uses pytesseract to extract numbers and their positions
3. **Calculation**: Finds pairs of numbers that sum to 10
4. **Visualization**: Draws colored rectangles around matching pairs
5. **Display**: Shows the result in the GUI window

## Notes

- The application requires administrator privileges on Windows to capture the screen
- For best OCR results, ensure numbers are clear and well-lit
- The app automatically filters low-confidence OCR results
