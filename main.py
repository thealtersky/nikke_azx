import keyboard
import mss
import numpy as np
import cv2
import threading
import sys
import time
import pyautogui
import psutil
import win32gui
import win32con
import win32process
from PyQt5 import QtCore, QtGui, QtWidgets
import sys
import os
from threading import Lock

# Helper for resource loading (works for dev and PyInstaller EXE)
def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)
import ctypes
from PyQt5.QtWidgets import (QApplication, QMainWindow, QPushButton, QVBoxLayout, 
                             QHBoxLayout, QWidget, QLabel, QSpinBox, QGroupBox, 
                             QTextEdit, QCheckBox)

class Overlay(QtWidgets.QWidget):

    def __init__(self, rows, cols, left_start, top_start, offset_x, offset_y, w, h):
        super().__init__()

        self.rows = rows
        self.cols = cols

        self.left_start = left_start
        self.top_start = top_start
        self.offset_x = offset_x
        self.offset_y = offset_y

        self.cell_w = w
        self.cell_h = h

        self.cells = []

        self.setWindowFlags(
            QtCore.Qt.FramelessWindowHint |
            QtCore.Qt.WindowStaysOnTopHint |
            QtCore.Qt.Tool
        )
        self.setAttribute(QtCore.Qt.WA_TranslucentBackground)

        self.setGeometry(0, 0, 1920, 1080)

        self.show()

        #Click-through
        hwnd = self.winId().__int__()
        ctypes.windll.user32.SetWindowLongW(
            hwnd, -20,
            ctypes.windll.user32.GetWindowLongW(hwnd, -20)
            | 0x80000  #WS_EX_LAYERED
            | 0x20     #WS_EX_TRANSPARENT
        )

    def set_cells(self, cell_list):
        self.cells = cell_list
        self.update()

    def paintEvent(self, event):
        painter = QtGui.QPainter(self)

        for (r, c, color) in self.cells:
            r_px = self.top_start + r * self.offset_y
            c_px = self.left_start + c * self.offset_x

            brush = QtGui.QColor(*color)
            painter.setBrush(brush)
            painter.setPen(QtGui.QPen(QtGui.QColor(255,255,255,180), 2))

            painter.drawRect(c_px, r_px, self.cell_w, self.cell_h)


class ControlGUI(QMainWindow):
    log_signal = QtCore.pyqtSignal(str)
    
    def __init__(self):
        super().__init__()
        self.solver = None
        self.log_signal.connect(self.append_log)
        self.init_ui()

    def set_solver(self, solver):
        """Set the solver instance and connect signals"""
        self.solver = solver
        
        # Connect button signals
        self.scan_btn.clicked.connect(self.solver.get_matrix_numbers)
        self.clean_btn.clicked.connect(self.solver.clean_matrix)
        self.right_btn.clicked.connect(self.solver.sums_right)
        self.down_btn.clicked.connect(self.solver.sums_down)
        self.square_btn.clicked.connect(self.solver.sums_square)
        self.cancel_btn.clicked.connect(self.solver.cancel_auto_solve)


    def init_ui(self):
        self.setWindowTitle("Sum10 Puzzle Solver")
        self.setGeometry(100, 100, 500, 700)
        self.setWindowFlags(QtCore.Qt.WindowStaysOnTopHint)

        # Apply dark mode stylesheet
        dark_stylesheet = """
        QWidget {
            background-color: #232629;
            color: #e0e0e0;
        }
        QGroupBox {
            border: 1px solid #444;
            border-radius: 6px;
            margin-top: 10px;
            background-color: #282c34;
        }
        QGroupBox:title {
            subcontrol-origin: margin;
            subcontrol-position: top left;
            padding: 0 3px;
            color: #ffb300;
        }
        QLabel {
            color: #e0e0e0;
        }
        QPushButton {
            background-color: #333;
            color: #fff;
            border: 1px solid #555;
            border-radius: 4px;
            padding: 6px 12px;
        }
        QPushButton:hover {
            background-color: #444;
        }
        QPushButton:pressed {
            background-color: #222;
        }
        QSpinBox, QTextEdit, QLineEdit {
            background-color: #181a1b;
            color: #e0e0e0;
            border: 1px solid #444;
            border-radius: 4px;
        }
        QCheckBox {
            color: #e0e0e0;
        }
        """
        self.setStyleSheet(dark_stylesheet)

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        # Status label
        self.status_label = QLabel("Status: Ready")
        self.status_label.setStyleSheet("font-weight: bold; padding: 1px; background-color: #232629; color: #ffb300; border-radius: 5px;")
        layout.addWidget(self.status_label)

        # Process Detection Group
        process_group = QGroupBox("Game Detection")
        process_layout = QVBoxLayout()
        
        self.process_status = QLabel("❌ NIKKE not detected")
        self.process_status.setStyleSheet("padding: 5px; color: red; font-weight: bold;")
        process_layout.addWidget(self.process_status)
        
        detect_btn = QPushButton("🔍 Detect NIKKE Process")
        detect_btn.clicked.connect(self.detect_game)
        process_layout.addWidget(detect_btn)
        
        self.auto_detect_check = QCheckBox("Auto-detect on solve")
        self.auto_detect_check.setChecked(True)
        process_layout.addWidget(self.auto_detect_check)
        
        process_group.setLayout(process_layout)
        layout.addWidget(process_group)

        # Grid settings group
        grid_group = QGroupBox("Grid Settings")
        grid_layout = QVBoxLayout()
        
        row_layout = QHBoxLayout()
        row_layout.addWidget(QLabel("Rows:"))
        self.rows_spin = QSpinBox()
        self.rows_spin.setValue(16)
        self.rows_spin.setRange(1, 50)
        row_layout.addWidget(self.rows_spin)
        grid_layout.addLayout(row_layout)
        
        col_layout = QHBoxLayout()
        col_layout.addWidget(QLabel("Columns:"))
        self.cols_spin = QSpinBox()
        self.cols_spin.setValue(10)
        self.cols_spin.setRange(1, 50)
        col_layout.addWidget(self.cols_spin)
        grid_layout.addLayout(col_layout)
        
        grid_group.setLayout(grid_layout)
        layout.addWidget(grid_group)

        # Manual control buttons
        manual_group = QGroupBox("Manual Controls")
        manual_layout = QVBoxLayout()

        self.scan_btn = QPushButton("📷 Scan Matrix (F5)")
        manual_layout.addWidget(self.scan_btn)

        self.clean_btn = QPushButton("🧹 Clean Matrix (F1)")
        manual_layout.addWidget(self.clean_btn)

        self.right_btn = QPushButton("➡️ Find Right Sums (F2)")
        manual_layout.addWidget(self.right_btn)

        self.down_btn = QPushButton("⬇️ Find Down Sums (F3)")
        manual_layout.addWidget(self.down_btn)

        self.square_btn = QPushButton("⬛ Find Square Sums (F4)")
        manual_layout.addWidget(self.square_btn)

        manual_group.setLayout(manual_layout)
        layout.addWidget(manual_group)

        # Automation controls
        auto_group = QGroupBox("Automation")
        auto_layout = QVBoxLayout()

        self.auto_solve_btn = QPushButton("🤖 AUTO SOLVE (F6)")
        self.auto_solve_btn.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold; padding: 10px; font-size: 14px;")
        self.auto_solve_btn.clicked.connect(self.start_auto_solve)
        auto_layout.addWidget(self.auto_solve_btn)

        self.cancel_btn = QPushButton("⛔ CANCEL AUTO SOLVE (F12)")
        self.cancel_btn.setStyleSheet("background-color: #FF9800; color: white; font-weight: bold; padding: 10px; font-size: 14px;")
        self.cancel_btn.setEnabled(False)
        auto_layout.addWidget(self.cancel_btn)

        delay_layout = QHBoxLayout()
        delay_layout.addWidget(QLabel("Drag Delay (ms):"))
        self.delay_spin = QSpinBox()
        self.delay_spin.setValue(100)
        self.delay_spin.setRange(50, 1000)
        delay_layout.addWidget(self.delay_spin)
        auto_layout.addLayout(delay_layout)

        auto_group.setLayout(auto_layout)
        layout.addWidget(auto_group)

        # Log display
        log_group = QGroupBox("Activity Log")
        log_layout = QVBoxLayout()
        
        self.log_display = QTextEdit()
        self.log_display.setReadOnly(True)
        self.log_display.setMaximumHeight(150)
        self.log_display.setStyleSheet("background-color: #1e1e1e; color: #00ff00; font-family: 'Consolas', monospace; font-size: 10px;")
        log_layout.addWidget(self.log_display)
        
        clear_log_btn = QPushButton("Clear Log")
        clear_log_btn.clicked.connect(self.log_display.clear)
        log_layout.addWidget(clear_log_btn)
        
        log_group.setLayout(log_layout)
        layout.addWidget(log_group)

        # Exit button
        self.exit_btn = QPushButton("❌ Exit (ESC)")
        self.exit_btn.setStyleSheet("background-color: #f44336; color: white; font-weight: bold; padding: 8px;")
        self.exit_btn.clicked.connect(self.close_app)
        layout.addWidget(self.exit_btn)

    def update_status(self, message):
        self.status_label.setText(f"Status: {message}")
        
    def log(self, message):
        """Thread-safe logging"""
        self.log_signal.emit(message)
    
    def append_log(self, message):
        """Append log message to display"""
        timestamp = time.strftime("%H:%M:%S")
        self.log_display.append(f"[{timestamp}] {message}")
        # Auto-scroll to bottom
        self.log_display.verticalScrollBar().setValue(
            self.log_display.verticalScrollBar().maximum()
        )

    def detect_game(self):
        """Detect NIKKE game process"""
        if self.solver:
            found = self.solver.find_nikke_process()
            if found:
                self.process_status.setText("✅ NIKKE detected and focused")
                self.process_status.setStyleSheet("padding: 5px; color: green; font-weight: bold;")
                self.log("Game window found and brought to foreground")
            else:
                self.process_status.setText("❌ NIKKE not found")
                self.process_status.setStyleSheet("padding: 5px; color: red; font-weight: bold;")
                self.log("NIKKE process not found. Please start the game.")

    def start_auto_solve(self):
        self.update_status("Starting auto-solve...")
        self.log("=== AUTO SOLVE INITIATED ===")
        self.auto_solve_btn.setEnabled(False)
        self.cancel_btn.setEnabled(True)
        threading.Thread(target=self.solver.auto_solve, daemon=True).start()

    def close_app(self):
        self.log("Application closing...")
        QApplication.quit()


class PuzzleSolver:
    def __init__(self, overlay, gui):
        self.overlay = overlay
        self.gui = gui
        
        self.rows = 16
        self.columns = 10
        
        self.offset_x = 51
        self.offset_y = 52
        self.top_start = 221
        self.left_start = 708
        self.capture_area_w = 44
        self.capture_area_h = 45
        
        self.numbers = []
        self.matrix = []
        self.solutions = []
        
        self.nikke_hwnd = None
        self.cancel_flag = False
        self.cancel_lock = Lock()
        self.is_auto_solving = False
        
        self.start_area = {
            "top": self.top_start,
            "left": self.left_start,
            "width": self.capture_area_w,
            "height": self.capture_area_h
        }
        
        # Start keyboard monitor thread
        self.start_keyboard_monitor()

    def log(self, message):
        """Send log message to GUI"""
        if self.gui:
            self.gui.log(message)
        else:
            print(message)
    
    def start_keyboard_monitor(self):
        """Start a background thread to continuously monitor F12 key"""
        def monitor_f12():
            while True:
                try:
                    if keyboard.is_pressed('f12'):
                        if self.is_auto_solving:
                            print("F12 DETECTED WHILE AUTO-SOLVING!")
                            self.cancel_auto_solve()
                            time.sleep(0.5)  # Debounce
                    time.sleep(0.1)  # Check every 100ms
                except Exception as e:
                    print(f"Keyboard monitor error: {e}")
                    time.sleep(1)
        
        monitor_thread = threading.Thread(target=monitor_f12, daemon=True)
        monitor_thread.start()
        self.log("Keyboard monitor started for F12 cancel")

    def cancel_auto_solve(self):
        """Cancel the ongoing auto-solve operation"""
        with self.cancel_lock:
            self.cancel_flag = True
            print("CANCEL FLAG SET TO TRUE")  # Debug print
        self.log("⛔⛔⛔ CANCEL REQUESTED - STOPPING ASAP ⛔⛔⛔")
    
    def is_cancelled(self):
        """Thread-safe check if cancellation was requested"""
        with self.cancel_lock:
            return self.cancel_flag

    def find_nikke_process(self):
        """Find and focus NIKKE game window"""
        self.log("Searching for NIKKE process...")
        
        # Search for common NIKKE process names
        process_names = ['nikke.exe', 'NIKKE.exe', 'Nikke.exe']
        
        for proc in psutil.process_iter(['pid', 'name']):
            try:
                if proc.info['name'] in process_names:
                    self.log(f"Found process: {proc.info['name']} (PID: {proc.info['pid']})")
                    
                    # Find the window associated with this process
                    def callback(hwnd, hwnds):
                        if win32gui.IsWindowVisible(hwnd):
                            _, pid = win32process.GetWindowThreadProcessId(hwnd)
                            if pid == proc.info['pid']:
                                hwnds.append(hwnd)
                        return True
                    
                    hwnds = []
                    win32gui.EnumWindows(callback, hwnds)
                    
                    if hwnds:
                        self.nikke_hwnd = hwnds[0]
                        self.focus_game_window()
                        return True
                        
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                pass
        
        self.log("NIKKE process not found!")
        return False

    def focus_game_window(self):
        """Bring NIKKE window to foreground"""
        if self.nikke_hwnd:
            try:
                # Restore if minimized
                if win32gui.IsIconic(self.nikke_hwnd):
                    win32gui.ShowWindow(self.nikke_hwnd, win32con.SW_RESTORE)
                
                # Bring to foreground
                win32gui.SetForegroundWindow(self.nikke_hwnd)
                self.log("Game window brought to foreground")
                time.sleep(0.5)  # Wait for window to be ready
                return True
            except Exception as e:
                self.log(f"Error focusing window: {e}")
                return False
        return False

    def get_cell_center(self, row, col):
        """Get the center pixel coordinates of a cell"""
        x = self.left_start + col * self.offset_x + self.cell_w // 2
        y = self.top_start + row * self.offset_y + self.cell_h // 2
        return x, y

    def perform_drag(self, start_row, start_col, end_row, end_col):
        """Perform a drag operation from start cell to end cell"""
        # Check cancellation before drag
        if self.is_cancelled():
            self.log("Drag cancelled before starting")
            return False
            
        start_x, start_y = self.get_cell_center(start_row, start_col)
        end_x, end_y = self.get_cell_center(end_row, end_col)
        
        delay = self.gui.delay_spin.value() / 1000.0
        
        self.log(f"Dragging from ({start_row},{start_col}) to ({end_row},{end_col})")
        
        pyautogui.moveTo(start_x, start_y, duration=0.1)
        
        if self.is_cancelled():
            return False
            
        time.sleep(0.05)
        pyautogui.mouseDown()
        
        if self.is_cancelled():
            pyautogui.mouseUp()
            return False
            
        time.sleep(delay)
        pyautogui.moveTo(end_x, end_y, duration=0.2)
        
        if self.is_cancelled():
            pyautogui.mouseUp()
            return False
            
        time.sleep(0.05)
        pyautogui.mouseUp()
        time.sleep(0.2)
        
        return True

    def auto_solve(self):
        """Automatically solve the puzzle by scanning and executing drags, with visual highlights."""
        # Reset cancel flag and set auto-solving state
        with self.cancel_lock:
            self.cancel_flag = False
            self.is_auto_solving = True
        
        print("AUTO-SOLVE STARTED - is_auto_solving =", self.is_auto_solving)
        
        try:
            # Step 1: Detect and focus game if enabled
            if self.gui.auto_detect_check.isChecked():
                self.gui.update_status("Detecting game...")
                if not self.find_nikke_process():
                    self.gui.update_status("Game not found!")
                    self.log("ERROR: Cannot proceed without game window")
                    self.gui.auto_solve_btn.setEnabled(True)
                    self.gui.cancel_btn.setEnabled(False)
                    return
                time.sleep(0.5)

            # Check for cancellation
            if self.is_cancelled():
                self.gui.update_status("Cancelled before scan")
                self.log("Auto-solve cancelled by user")
                self.gui.auto_solve_btn.setEnabled(True)
                self.gui.cancel_btn.setEnabled(False)
                return

            # SCAN ONCE at the beginning
            self.gui.update_status("Scanning matrix...")
            self.log("=== INITIAL SCAN: Reading grid ===")
            time.sleep(0.3)

            self.get_matrix_numbers()

            if not self.matrix:
                self.gui.update_status("Failed to scan matrix!")
                self.log("ERROR: Matrix scan failed")
                self.gui.auto_solve_btn.setEnabled(True)
                self.gui.cancel_btn.setEnabled(False)
                return

            total_solutions_executed = 0
            iteration = 1
            
            # MAIN LOOP: Keep finding and executing solutions using memory
            while True:
                # Check for cancellation
                if self.is_cancelled():
                    self.gui.update_status(f"Cancelled ({total_solutions_executed} completed)")
                    self.log(f"Auto-solve cancelled by user after {total_solutions_executed} solutions")
                    self.gui.auto_solve_btn.setEnabled(True)
                    self.gui.cancel_btn.setEnabled(False)
                    return

                # Check if there are any numbers left in our memory matrix
                has_numbers = False
                for row in self.matrix:
                    for cell in row:
                        if cell not in [' ', "\u2192", "\u2193", "\u25A0", "\u25A1", "\u25BA", "\u25BC"]:
                            if isinstance(cell, int) and 1 <= cell <= 9:
                                has_numbers = True
                                break
                    if has_numbers:
                        break
                
                if not has_numbers:
                    self.gui.update_status(f"Puzzle Complete! ({total_solutions_executed} total solutions)")
                    self.log(f"=== PUZZLE COMPLETE: No more numbers in memory ===")
                    self.log(f"Total solutions executed: {total_solutions_executed}")
                    self.gui.auto_solve_btn.setEnabled(True)
                    self.gui.cancel_btn.setEnabled(False)
                    return

                # Find all solutions in current memory state
                self.gui.update_status(f"Analyzing solutions (iteration {iteration})...")
                self.log(f"=== ITERATION {iteration}: Finding solutions ===")
                self.find_all_solutions()

                if not self.solutions:
                    self.gui.update_status(f"No more solutions - Complete! ({total_solutions_executed} total)")
                    self.log(f"No valid solutions found in iteration {iteration}")
                    self.log(f"=== AUTO SOLVE COMPLETE: {total_solutions_executed} total solutions executed ===")
                    self.gui.auto_solve_btn.setEnabled(True)
                    self.gui.cancel_btn.setEnabled(False)
                    return

                self.log(f"Found {len(self.solutions)} valid solutions in iteration {iteration}")

                # Create a set to track which cells have been used in this iteration
                used_cells = set()

                # Execute ALL solutions from current scan, tracking used cells
                solutions_this_iteration = 0
                for solution in self.solutions:
                    # Check for cancellation
                    if self.is_cancelled():
                        self.gui.update_status(f"Cancelled ({total_solutions_executed} completed)")
                        self.log(f"⛔ Auto-solve CANCELLED after {total_solutions_executed} solutions")
                        self.gui.auto_solve_btn.setEnabled(True)
                        self.gui.cancel_btn.setEnabled(False)
                        return

                    sol_type, start_r, start_c, end_r, end_c = solution

                    # Get all cells involved in this solution
                    solution_cells = set()
                    if sol_type == 'right':
                        for c in range(start_c, end_c + 1):
                            solution_cells.add((start_r, c))
                    elif sol_type == 'down':
                        for r in range(start_r, end_r + 1):
                            solution_cells.add((r, start_c))
                    elif sol_type == 'square':
                        for r in range(min(start_r, end_r), max(start_r, end_r) + 1):
                            for c in range(start_c, end_c + 1):
                                solution_cells.add((r, c))

                    # Check if any cells in this solution have already been used
                    if solution_cells & used_cells:
                        self.log(f"Skipping solution at ({start_r},{start_c}) - cells already used this iteration")
                        continue

                    # Mark these cells as used
                    used_cells.update(solution_cells)

                    total_solutions_executed += 1
                    solutions_this_iteration += 1
                    self.gui.update_status(f"Solution #{total_solutions_executed} (iter {iteration}: {solutions_this_iteration}/{len(self.solutions)}) - {sol_type}")
                    self.log(f"Executing solution #{total_solutions_executed}: {sol_type} at ({start_r},{start_c})")

                    # Highlight the solution visually before dragging
                    if sol_type == 'right':
                        positions = end_c - start_c
                        for x in range(positions + 1):
                            if self.matrix[start_r][start_c + x] not in [' ', "\u2192", "\u2193", "\u25A0", "\u25A1", "\u25BA", "\u25BC"]:
                                if x == 0:
                                    self.matrix[start_r][start_c] = "\u25BA"
                                else:
                                    self.matrix[start_r][start_c + x] = "\u2192"
                    elif sol_type == 'down':
                        positions = end_r - start_r
                        for x in range(positions + 1):
                            if self.matrix[start_r + x][start_c] not in [' ', "\u2192", "\u2193", "\u25A0", "\u25A1", "\u25BA", "\u25BC"]:
                                if x == 0:
                                    self.matrix[start_r][start_c] = "\u25BC"
                                else:
                                    self.matrix[start_r + x][start_c] = "\u2193"
                    elif sol_type == 'square':
                        r0, c0, r1, c1 = start_r, start_c, end_r, end_c
                        for j in range(min(r0, r1), max(r0, r1) + 1):
                            for k in range(c0, c1 + 1):
                                if self.matrix[j][k] not in [' ', "\u2192", "\u2193", "\u25A0", "\u25A1", "\u25BA", "\u25BC"]:
                                    if j == start_r and k == start_c:
                                        self.matrix[j][k] = "\u25A1"
                                    else:
                                        self.matrix[j][k] = "\u25A0"

                    self.update_overlay()
                    time.sleep(0.2)  # Show highlight before drag

                    try:
                        drag_success = self.perform_drag(start_r, start_c, end_r, end_c)
                        if not drag_success:
                            self.gui.update_status(f"Cancelled ({total_solutions_executed - 1} completed)")
                            self.log(f"⛔ Auto-solve CANCELLED after {total_solutions_executed - 1} solutions")
                            self.gui.auto_solve_btn.setEnabled(True)
                            self.gui.cancel_btn.setEnabled(False)
                            return
                        
                        # Mark the cells as empty in our matrix memory after successful drag
                        for cell in solution_cells:
                            r, c = cell
                            self.matrix[r][c] = ' '
                        
                        time.sleep(0.3)  # Wait for animation
                    except Exception as e:
                        self.log(f"ERROR during drag: {e}")

                    # Clear overlay after drag
                    self.update_overlay()

                self.log(f"Iteration {iteration} complete: {solutions_this_iteration} solutions executed")
                iteration += 1
                # No rescan - continue with memory matrix
        
        finally:
            # Always reset auto-solving state when done
            with self.cancel_lock:
                self.is_auto_solving = False
            print("AUTO-SOLVE ENDED - is_auto_solving =", self.is_auto_solving)

    def find_all_solutions(self):
        """Find all valid sum=10 solutions in the matrix"""
        self.solutions = []
        
        # Find right sums
        for r in range(self.rows):
            for c in range(self.columns):
                if self.has_special_char(r, c):
                    continue
                
                valid, positions = self.checkRight(self.columns, r, c, self.matrix)
                if valid:
                    self.solutions.append(('right', r, c, r, c + positions))
        
        # Find down sums
        for r in range(self.rows):
            for c in range(self.columns):
                if self.has_special_char(r, c):
                    continue
                
                valid, positions = self.checkDown(self.rows, r, c, self.matrix)
                if valid:
                    self.solutions.append(('down', r, c, r + positions, c))
        
        # Find square sums
        for r in range(self.rows):
            for c in range(self.columns):
                if self.has_special_char(r, c):
                    continue
                
                valid, max_r, max_c = self.checkSquareDown(self.rows, self.columns, r, c, self.matrix)
                if valid:
                    self.solutions.append(('square', r, c, max_r, max_c))
                else:
                    valid, max_r, max_c = self.checkSquareUp(self.rows, self.columns, r, c, self.matrix)
                    if valid:
                        self.solutions.append(('square', r, c, max_r, max_c))

    def get_matrix_numbers(self):
        self.numbers = []
        self.matrix = []
        
        counter = 0
        self.start_area["top"] = self.top_start
        self.start_area["left"] = self.left_start

        self.log(f"Scanning {self.rows}x{self.columns} grid...")

        with mss.mss() as sct:
            for r in range(self.rows):
                for c in range(self.columns):
                    counter += 1

                    image = sct.grab(self.start_area)
                    img_np = np.array(image)
                    img_gray = cv2.cvtColor(img_np, cv2.COLOR_BGR2GRAY)

                    best_score = -10 
                    best_match_digit = -1


                    for j in range(1, 10):
                        template_path = resource_path(f"templates/T{j}.png")
                        template = cv2.imread(template_path)

                        if template is None:
                            continue

                        temp_gray = cv2.cvtColor(template, cv2.COLOR_BGR2GRAY)
                        res = cv2.matchTemplate(img_gray, temp_gray, cv2.TM_CCOEFF_NORMED)
                        min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(res)

                        if max_val > best_score:
                            best_score = max_val
                            best_match_digit = j
                    
                    if best_match_digit != -1:
                        self.numbers.append(best_match_digit)
                    else:
                        self.numbers.append(" ")

                    self.start_area["left"] += self.offset_x

                self.start_area["left"] = self.left_start
                self.start_area["top"] += self.offset_y

        self.createMatrix()
        self.gui.update_status("Matrix scanned successfully")
        self.log(f"OCR scan complete: {counter} cells analyzed")

    def createMatrix(self):
        self.matrix = []
        pos = 0
        for i in range(self.rows):
            row = []
            for j in range(self.columns):
                row.append(self.numbers[pos])
                pos += 1
            self.matrix.append(row)
        self.printMatrix()

    def printMatrix(self):
        self.log("Matrix preview:")
        for i in range(min(3, self.rows)):  # Show first 3 rows only
            row_str = ""
            for j in range(self.columns):
                row_str += f"{self.matrix[i][j]}  "
            self.log(row_str)
        if self.rows > 3:
            self.log("...")

    def checkRight(self, columns, r, c, matrix):
        sum_val = 0
        for j in range(columns - c):
            if c + j >= columns:
                return False, 0
            if self.has_special_char(r, c + j):
                continue
            sum_val += matrix[r][c + j]
            if sum_val > 10:
                return False, 0
            elif sum_val == 10:
                return True, j
        return False, 0

    def sums_right(self):
        count = 0
        for r in range(self.rows):
            for c in range(self.columns):
                if self.has_special_char(r, c):
                    continue
                validSum, positions = self.checkRight(self.columns, r, c, self.matrix)
                if validSum:
                    count += 1
                    for x in range(1, positions + 1):
                        self.matrix[r][c] = "\u25BA"
                        self.matrix[r][c + x] = "\u2192"
        self.printMatrix()
        self.update_overlay()
        self.gui.update_status(f"Found {count} right sums")
        self.log(f"Right sums found: {count}")

    def checkDown(self, rows, r, c, matrix):
        sum_val = 0
        for j in range(rows - r):
            if r + j >= rows:
                return False, 0
            if self.has_special_char(r + j, c):
                continue
            sum_val += matrix[r + j][c]
            if sum_val > 10:
                return False, 0
            elif sum_val == 10:
                return True, j
        return False, 0

    def sums_down(self):
        count = 0
        for c in range(self.columns):
            for r in range(self.rows):
                if self.has_special_char(r, c):
                    continue
                validSum, positions = self.checkDown(self.rows, r, c, self.matrix)
                if validSum:
                    count += 1
                    for x in range(1, positions + 1):
                        self.matrix[r][c] = "\u25BC"
                        self.matrix[x + r][c] = "\u2193"
        self.printMatrix()
        self.update_overlay()
        self.gui.update_status(f"Found {count} down sums")
        self.log(f"Down sums found: {count}")

    def checkSquareDown(self, rows, columns, start_r, start_c, matrix):
        if start_r >= rows or start_c >= columns:
            return False, 0, 0
        if self.has_special_char(start_r, start_c):
            return False, 0, 0

        edge_rows = start_r + 1
        edge_columns = start_c + 1

        while edge_rows < rows and edge_columns < columns:
            sum_val = 0
            for r in range(start_r, edge_rows + 1):
                for c in range(start_c, edge_columns + 1):
                    if not self.has_special_char(r, c):
                        sum_val += int(matrix[r][c])
            
            if sum_val > 10:
                return False, 0, 0
            if sum_val == 10:
                return True, edge_rows, edge_columns
            
            edge_rows += 1
            edge_columns += 1

        return False, 0, 0

    def checkSquareUp(self, rows, columns, start_r, start_c, matrix):
        if start_r <= 0 or start_c >= columns:
            return False, 0, 0
        if self.has_special_char(start_r, start_c):
            return False, 0, 0

        edge_rows = start_r - 1
        edge_columns = start_c + 1

        while edge_rows >= 0 and edge_columns < columns:
            sum_val = 0
            for r in range(edge_rows, start_r + 1):
                for c in range(start_c, edge_columns + 1):
                    if not self.has_special_char(r, c):
                        sum_val += int(matrix[r][c])
            
            if sum_val > 10:
                return False, 0, 0
            if sum_val == 10:
                return True, edge_rows, edge_columns
            
            edge_rows -= 1
            edge_columns += 1

        return False, 0, 0

    def sums_square(self):
        count = 0
        for r in range(self.rows):
            for c in range(self.columns):
                if self.has_special_char(r, c):
                    continue
                
                validSumDown, max_r, max_c = self.checkSquareDown(self.rows, self.columns, r, c, self.matrix)
                if validSumDown:
                    count += 1
                    for j in range(r, max_r + 1):
                        for k in range(c, max_c + 1):
                            self.matrix[j][k] = "\u25A0"
                    self.matrix[r][c] = "\u25A1"
                else:
                    validSumUp, max_r, max_c = self.checkSquareUp(self.rows, self.columns, r, c, self.matrix)
                    if validSumUp:
                        count += 1
                        for j in range(r, max_r - 1, -1):
                            for k in range(c, max_c + 1):
                                self.matrix[j][k] = "\u25A0"
                        self.matrix[r][c] = "\u25A1"
        
        self.printMatrix()
        self.update_overlay()
        self.gui.update_status(f"Found {count} square sums")
        self.log(f"Square sums found: {count}")

    def has_special_char(self, r, c):
        return self.matrix[r][c] in ["\u2192", "\u2193", "\u25A0", "\u25A1", "\u25BA", "\u25BC", " "]

    def clean_matrix(self):
        for r in range(self.rows):
            for c in range(self.columns):
                if self.has_special_char(r, c):
                    self.matrix[r][c] = " "
        self.printMatrix()
        self.update_overlay()
        self.gui.update_status("Matrix cleaned")
        self.log("Special characters cleaned from matrix")

    def update_overlay(self):
        cell_list = []
        for r in range(self.rows):
            for c in range(self.columns):
                value = self.matrix[r][c]
                if value == "\u25A0":
                    cell_list.append((r, c, (255, 0, 0, 70)))
                elif value == "\u25A1":
                    cell_list.append((r, c, (255, 0, 0, 200)))
                elif value == "\u2192":
                    cell_list.append((r, c, (0, 255, 0, 70)))
                elif value == "\u2193":
                    cell_list.append((r, c, (0, 0, 255, 70)))
                elif value == "\u25BA":
                    cell_list.append((r, c, (0, 255, 0, 200)))
                elif value == "\u25BC":
                    cell_list.append((r, c, (0, 0, 255, 200)))
                elif value == " ":
                    cell_list.append((r, c, (0, 0, 0, 0)))
        self.overlay.set_cells(cell_list)

    @property
    def cell_w(self):
        return self.capture_area_w

    @property
    def cell_h(self):
        return self.capture_area_h


def setup_hotkeys(solver):
    keyboard.add_hotkey('f5', solver.get_matrix_numbers)
    keyboard.add_hotkey('f1', solver.clean_matrix)
    keyboard.add_hotkey('f2', solver.sums_right)
    keyboard.add_hotkey('f3', solver.sums_down)
    keyboard.add_hotkey('f4', solver.sums_square)
    keyboard.add_hotkey('f6', solver.auto_solve)
    
    # F12 for cancel - with debug
    def cancel_handler():
        print("F12 PRESSED - CANCEL HANDLER CALLED")
        solver.cancel_auto_solve()
    
    keyboard.add_hotkey('f12', cancel_handler)
    keyboard.add_hotkey("esc", lambda: QApplication.quit())
    
    print("Hotkeys registered successfully!")


if __name__ == "__main__":
    app = QApplication(sys.argv)

    # Create overlay
    overlay = Overlay(16, 10, 708, 221, 51, 52, 44, 45)

    # Create GUI and solver
    gui = ControlGUI()
    solver = PuzzleSolver(overlay, gui)
    gui.set_solver(solver)

    # Setup hotkeys in a separate thread
    threading.Thread(target=lambda: setup_hotkeys(solver), daemon=True).start()
    threading.Thread(target=keyboard.wait, daemon=True).start()

    gui.show()
    sys.exit(app.exec_())