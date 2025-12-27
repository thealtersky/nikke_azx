> [!IMPORTANT]
> **Disclaimer**: This software is AI generated. This software is provided "as is", without warranty of any kind. The author is not responsible for any issues or damage caused by the use of this tool. Use at your own risk. 

> [!NOTE]
> This project is based on code from the original repository https://github.com/Miguensio/AZX-Nikke-Script/
> 
> All core logic, features, and much of the implementation are derived from or adapted from the above source.

# Sum10 Puzzle Solver

A Python GUI tool for automating and solving the "Sum 10" puzzle in the NIKKE game. The app uses OpenCV template matching to recognize numbers on the game grid, and provides both manual and automatic solving features with a modern PyQt5 interface and visual overlays.

![Demo]()

## Features

- **Game Window Detection**: Automatically finds and focuses the NIKKE game window.
- **Grid Scanning**: Captures the puzzle grid and recognizes numbers using OpenCV template matching.
- **Manual Controls**: Scan, clean, and highlight right, down, and square sums manually.
- **Auto Solve**: Automatically finds and executes all valid sum-10 solutions, with visual highlights for each step.
- **Overlay Visualization**: See real-time highlights of detected sums directly over the game window.
- **Dark Mode UI**: Modern, dark-themed PyQt5 interface.
- **Hotkeys**: F1–F6 for quick actions, F12 to cancel auto-solve, ESC to exit.
- **Activity Log**: View all actions and results in a scrollable log.

## Installation

1. Install Python 3.8 or higher (Python 3.8–3.11 recommended).
2. Install the required packages:
   ```bash
   pip install -r requirements.txt
   ```
3. (Optional) Build as a standalone EXE using auto-py-to-exe or PyInstaller:
   ```bash
   pyinstaller --noconfirm --onefile --windowed --add-data "templates;templates/" main.py
   ```

## Usage

1. Start the NIKKE game and open the puzzle.
2. Run the application:
   ```bash
   python main.py
   ```
3. Use the GUI or hotkeys to:
   - Scan the grid (F5)
   - Find right sums (F2), down sums (F3), or square sums (F4)
   - Auto solve the puzzle (F6)
   - Cancel auto solve (F12)
   - Clean matrix (F1)
   - Exit (ESC)
4. Watch the overlay for visual feedback as the solver works.

You need to run the program as an administrator to use Auto Solve.

## Requirements

- Python 3.8+
- Windows OS (required for window focus and overlays)
- See `requirements.txt` for Python dependencies (PyQt5, OpenCV, mss, pyautogui, etc.)
- NIKKE game running and visible on your desktop

## How It Works

1. **Grid Capture**: Uses mss to capture the puzzle grid from the game window.
2. **Number Recognition**: Uses OpenCV template matching with digit images in the `templates/` folder to recognize numbers in each cell.
3. **Solution Search**: Finds all valid right, down, and square sum-10 solutions in the grid.
4. **Automation**: Simulates mouse drags to solve the puzzle automatically, with visual overlay highlights for each step.
5. **Overlay**: Draws colored rectangles over the game window to show detected sums and actions in real time.