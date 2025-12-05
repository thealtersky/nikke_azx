import tkinter as tk
from tkinter import messagebox
from PIL import Image, ImageGrab, ImageDraw, ImageFont, ImageTk
import keyboard
import numpy as np
import os
import sys
import cv2
from collections import defaultdict

# Handle both development and PyInstaller bundled execution
def resource_path(relative_path):
    """Get the correct path for resources in both dev and compiled modes"""
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

# Template matching - load digit templates
digit_templates = {}
TEMPLATE_DIR = resource_path("img")

def load_templates():
    """Load digit templates from img folder"""
    global digit_templates
    digit_templates = {}
    
    if not os.path.exists(TEMPLATE_DIR):
        os.makedirs(TEMPLATE_DIR)
        return False
    
    for i in range(1, 10):  # 1 to 9 only
        template_path = os.path.join(TEMPLATE_DIR, f"{i}.png")
        if os.path.exists(template_path):
            # Load and preprocess template the same way as the main image
            template_img = cv2.imread(template_path)
            if template_img is not None:
                # Convert to grayscale
                template_gray = cv2.cvtColor(template_img, cv2.COLOR_BGR2GRAY)
                # Apply same enhancement
                clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8,8))
                template_enhanced = clahe.apply(template_gray)
                digit_templates[i] = template_enhanced
                print(f"Loaded template for digit {i}: {template_enhanced.shape}")
    
    return len(digit_templates) > 0

class NumberSumCapture:
    def __init__(self, root):
        self.root = root
        self.root.title("Number Sum Calculator")
        self.root.attributes('-topmost', True)
        
        # Load templates
        if not load_templates():
            messagebox.showwarning("Templates Missing", 
                                 "No digit templates found in 'img' folder!\n\n"
                                 "Please add template images:\n"
                                 "img/1.png, img/2.png, ... img/9.png\n\n"
                                 "Capture clean digit images from your game.")
        
        # Variables for capture window
        self.capture_window = None
        self.start_x = None
        self.start_y = None
        self.current_x = None
        self.current_y = None
        self.rect = None
        self.is_selecting = False
        
        # Captured image
        self.captured_image = None
        self.result_image = None
        
        # Store original window position to restore later
        self.original_window_x = None
        self.original_window_y = None
        
        # Setup UI
        self.setup_ui()
        
        # Register global hotkey Win+Shift+Z
        # NOTE: The 'keyboard' library hotkey may only work when the Python window is focused unless run as administrator.
        try:
            keyboard.add_hotkey('win+shift+z', self.start_capture)
        except Exception as e:
            print(f"Hotkey registration failed: {e}")
        
    def setup_ui(self):
        # --- Dark mode colors ---
        bg_main = "#23272e"
        bg_panel = "#282c34"
        fg_text = "#f8f8f2"
        fg_subtle = "#bbbbbb"
        accent_green = "#4CAF50"
        accent_blue = "#2196F3"
        accent_grey = "#444a56"
        accent_slider = "#444a56"
        accent_slider_trough = "#353940"
        status_ok = "#7CFC00"
        status_warn = "#ffb86c"
        status_err = "#ff5555"

        self.root.configure(bg=bg_main)

        # Main container (horizontal)
        container = tk.Frame(self.root, bg=bg_main)
        container.pack(fill=tk.BOTH, expand=True)

        # Left frame for controls
        left_frame = tk.Frame(container, padx=20, pady=20, bg=bg_panel)
        left_frame.pack(side=tk.LEFT, fill=tk.Y)

        # Title
        title_label = tk.Label(left_frame, text="Number Sum Calculator (Target: 10)",
                               font=("Arial", 14, "bold"), bg=bg_panel, fg=fg_text)
        title_label.pack(pady=(0, 10))

        # Instructions (with hotkey warning)
        instructions = tk.Label(left_frame,
            text="Press Win + Shift + Z to start capture (may require running as administrator, or may only work when this window is focused)\n"
                 "Drag to select area, release to capture\n"
                 "Or click the button below:",
            justify=tk.LEFT,
            fg=fg_subtle,
            bg=bg_panel)
        instructions.pack(pady=(0, 10))

        # Button row for Capture and Save Result
        button_row = tk.Frame(left_frame, bg=bg_panel)
        button_row.pack(pady=10)

        # Capture button
        self.capture_btn = tk.Button(button_row, text="Capture",
                                     command=self.start_capture,
                                     bg=accent_green, fg=fg_text,
                                     activebackground="#388e3c", activeforeground=fg_text,
                                     font=("Arial", 12, "bold"),
                                     padx=20, pady=10, bd=0, relief=tk.FLAT)
        self.capture_btn.pack(side=tk.LEFT, padx=(0, 10))

        # Save Result button (initially disabled)
        self.save_btn = tk.Button(button_row, text="Save Result",
                                  command=self.save_result,
                                  bg=accent_blue, fg=fg_text,
                                  activebackground="#1565c0", activeforeground=fg_text,
                                  font=("Arial", 12, "bold"),
                                  padx=10, pady=10, bd=0, relief=tk.FLAT,
                                  state=tk.DISABLED)
        self.save_btn.pack(side=tk.LEFT)

        # Threshold control
        threshold_frame = tk.Frame(left_frame, bg=bg_panel)
        threshold_frame.pack(pady=5)
        tk.Label(threshold_frame, text="Match Threshold:", bg=bg_panel, fg=fg_subtle).pack(side=tk.LEFT)
        self.threshold_var = tk.DoubleVar(value=0.90)
        threshold_slider = tk.Scale(threshold_frame, from_=0.5, to=1, resolution=0.05,
                                   orient=tk.HORIZONTAL, variable=self.threshold_var,
                                   bg=accent_slider, fg=fg_text, troughcolor=accent_slider_trough,
                                   highlightthickness=0, bd=0, sliderrelief=tk.FLAT, width=12, length=120)
        threshold_slider.pack(side=tk.LEFT)

        # Status label
        self.status_label = tk.Label(left_frame, text="Ready", fg=accent_blue, bg=bg_panel)
        self.status_label.pack(pady=10)

        # Right frame for result image
        self.result_frame = tk.Frame(container, padx=10, pady=10, bg=bg_main)
        self.result_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
    def start_capture(self):
        """Start the capture window"""
        self.status_label.config(text="Select area to capture...", fg="orange")
        
        # Store original window position and move it far off-screen
        self.original_window_x = self.root.winfo_x()
        self.original_window_y = self.root.winfo_y()
        self.root.geometry(f"+{-5000}+{-5000}")  # Move far off-screen
        self.root.update()
        
        # Delay 300ms to ensure window is fully off-screen before showing capture overlay
        self.root.after(300, self._show_capture_overlay)
    
    def _show_capture_overlay(self):
        """Show the capture overlay (called after delay)"""
        # Create fullscreen capture window
        self.capture_window = tk.Toplevel()
        self.capture_window.attributes('-fullscreen', True)
        self.capture_window.attributes('-alpha', 0.3)
        self.capture_window.attributes('-topmost', True)
        
        # Create canvas for drawing selection
        self.canvas = tk.Canvas(self.capture_window, cursor="cross", 
                               bg='grey', highlightthickness=0)
        self.canvas.pack(fill=tk.BOTH, expand=tk.TRUE)
        
        # Bind mouse events
        self.canvas.bind('<Button-1>', self.on_mouse_down)
        self.canvas.bind('<B1-Motion>', self.on_mouse_drag)
        self.canvas.bind('<ButtonRelease-1>', self.on_mouse_up)
        self.canvas.bind('<Escape>', lambda e: self.cancel_capture())
        
    def on_mouse_down(self, event):
        """Mouse button pressed"""
        self.start_x = event.x
        self.start_y = event.y
        self.is_selecting = True
        
        # Delete previous rectangle if exists
        if self.rect:
            self.canvas.delete(self.rect)
            
    def on_mouse_drag(self, event):
        """Mouse dragging"""
        if self.is_selecting:
            self.current_x = event.x
            self.current_y = event.y
            
            # Update rectangle
            if self.rect:
                self.canvas.delete(self.rect)
            
            self.rect = self.canvas.create_rectangle(
                self.start_x, self.start_y, 
                self.current_x, self.current_y,
                outline='red', width=2
            )
            
    def on_mouse_up(self, event):
        """Mouse button released - capture the region"""
        if self.is_selecting:
            self.current_x = event.x
            self.current_y = event.y
            self.is_selecting = False
            
            # Close capture window
            self.capture_window.destroy()
            
            # DO NOT restore window yet - delay capture to ensure it's off screen
            # Delay capture by 300ms to ensure overlay window is fully gone before screenshot
            self.root.after(300, self.capture_region)
            
    def cancel_capture(self):
        """Cancel the capture"""
        self.capture_window.destroy()
        
        # Restore main window to original position
        if self.original_window_x is not None and self.original_window_y is not None:
            self.root.geometry(f"+{self.original_window_x}+{self.original_window_y}")
        self.root.update()
        
        self.status_label.config(text="Capture cancelled", fg="red")
        
    def capture_region(self):
        """Capture the selected screen region"""
        # Get coordinates (ensure x1 < x2 and y1 < y2)
        x1 = min(self.start_x, self.current_x)
        y1 = min(self.start_y, self.current_y)
        x2 = max(self.start_x, self.current_x)
        y2 = max(self.start_y, self.current_y)

        # Clear previous result image and all widgets in result_frame
        self.result_image = None
        if hasattr(self, '_current_img_label') and self._current_img_label:
            self._current_img_label.destroy()
            self._current_img_label = None
        for widget in self.result_frame.winfo_children():
            widget.destroy()

        # Capture screenshot
        self.status_label.config(text="Capturing...", fg="blue")
        self.root.update()

        try:
            # Capture the region
            screenshot = ImageGrab.grab(bbox=(x1, y1, x2, y2))
            self.captured_image = screenshot

            # Process the image
            self.process_image()

        except Exception as e:
            messagebox.showerror("Error", f"Failed to capture: {str(e)}")
            self.status_label.config(text="Capture failed", fg="red")
        finally:
            # Restore main window to original position AFTER capture is done
            if self.original_window_x is not None and self.original_window_y is not None:
                self.root.geometry(f"+{self.original_window_x}+{self.original_window_y}")
            self.root.update()
    
    def non_max_suppression(self, detections, overlap_thresh=0.5):
        """Apply non-maximum suppression to remove overlapping detections"""
        if len(detections) == 0:
            return []
        
        # Sort by confidence
        detections = sorted(detections, key=lambda x: x['confidence'], reverse=True)
        
        keep = []
        while detections:
            # Take the detection with highest confidence
            best = detections.pop(0)
            keep.append(best)
            
            # Remove detections that overlap significantly
            detections = [d for d in detections if not self.overlaps(best, d, overlap_thresh)]
        
        return keep
    
    def overlaps(self, det1, det2, thresh):
        """Check if two detections overlap more than threshold"""
        x1 = max(det1['x'], det2['x'])
        y1 = max(det1['y'], det2['y'])
        x2 = min(det1['x'] + det1['w'], det2['x'] + det2['w'])
        y2 = min(det1['y'] + det1['h'], det2['y'] + det2['h'])
        
        if x2 < x1 or y2 < y1:
            return False
        
        intersection = (x2 - x1) * (y2 - y1)
        area1 = det1['w'] * det1['h']
        area2 = det2['w'] * det2['h']
        
        overlap = intersection / min(area1, area2)
        return overlap > thresh
            
    def process_image(self):
        """Process the captured image using optimized template matching"""
        self.status_label.config(text="Analyzing image...", fg="blue")
        self.root.update()
        
        if not digit_templates:
            messagebox.showerror("No Templates", 
                               "No digit templates loaded!\n\n"
                               "Please add digit images (1.png to 9.png) in the 'img' folder.")
            self.status_label.config(text="Templates missing", fg="red")
            return
        
        try:
            # Convert to numpy array and grayscale
            img_array = np.array(self.captured_image)
            gray = cv2.cvtColor(img_array, cv2.COLOR_RGB2GRAY)
            
            # Apply same preprocessing as templates
            clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8,8))
            enhanced = clahe.apply(gray)
            
            # Save preprocessed image for debugging
            cv2.imwrite("debug_enhanced.png", enhanced)
            print("Saved debug_enhanced.png for inspection")
            
            all_detections = []
            threshold = self.threshold_var.get()
            
            # Only try 3-4 scales instead of 7
            scales = [0.9, 1.0, 1.1]
            
            total_iterations = len(digit_templates) * len(scales)
            current_iter = 0
            
            # Match each digit template
            for digit in sorted(digit_templates.keys()):
                template = digit_templates[digit]
                
                for scale in scales:
                    current_iter += 1
                    progress = (current_iter / total_iterations) * 100
                    self.status_label.config(text=f"Matching digit {digit} ({progress:.0f}%)...", fg="blue")
                    self.root.update()
                    
                    # Resize template
                    if scale != 1.0:
                        new_w = int(template.shape[1] * scale)
                        new_h = int(template.shape[0] * scale)
                        if new_w < 5 or new_h < 5:
                            continue
                        scaled_template = cv2.resize(template, (new_w, new_h), interpolation=cv2.INTER_AREA)
                    else:
                        scaled_template = template
                    
                    # Skip if template is larger than image
                    if scaled_template.shape[0] > enhanced.shape[0] or scaled_template.shape[1] > enhanced.shape[1]:
                        continue
                    
                    # Template matching - use only the fastest method
                    result = cv2.matchTemplate(enhanced, scaled_template, cv2.TM_CCOEFF_NORMED)
                    
                    # Find matches above threshold
                    locations = np.where(result >= threshold)
                    
                    # Get template dimensions
                    h, w = scaled_template.shape
                    
                    # Store matches
                    for pt in zip(*locations[::-1]):
                        x, y = pt
                        confidence = float(result[y, x])
                        
                        all_detections.append({
                            'number': digit,
                            'x': x,
                            'y': y,
                            'w': w,
                            'h': h,
                            'center_x': x + w // 2,
                            'center_y': y + h // 2,
                            'scale': scale,
                            'confidence': confidence
                        })
            
            print(f"\nFound {len(all_detections)} raw detections")
            
            # Apply non-maximum suppression to remove duplicates
            self.status_label.config(text="Removing duplicates...", fg="blue")
            self.root.update()
            
            numbers_with_positions = self.non_max_suppression(all_detections, overlap_thresh=0.3)
            
            if not numbers_with_positions:
                messagebox.showwarning("No Numbers Found", 
                                      f"Could not match any numbers with threshold {threshold}.\n"
                                      "Try:\n"
                                      "1. Lower the threshold slider\n"
                                      "2. Check if templates match game digits\n"
                                      "3. Recapture templates from the game")
                self.status_label.config(text="No numbers matched", fg="red")
                return
            
            # Debug: Show detected numbers
            detected_nums = [n['number'] for n in numbers_with_positions]
            print(f"\nAfter NMS: {len(detected_nums)} unique detections")
            print(f"First 30 numbers: {detected_nums[:30]}")
            from collections import Counter
            distribution = Counter(detected_nums)
            print(f"Number distribution: {dict(sorted(distribution.items()))}")
            
            # Find pairs that sum to 10
            self.find_and_highlight_pairs(numbers_with_positions)
            
        except Exception as e:
            import traceback
            traceback.print_exc()
            messagebox.showerror("Error", f"Processing failed: {str(e)}")
            self.status_label.config(text="Processing failed", fg="red")
            
    def find_and_highlight_pairs(self, numbers_data):
        """Find pairs that sum to 10 and highlight them (grid-based horizontal/vertical)"""
        self.status_label.config(text="Finding pairs that sum to 10...", fg="blue")
        self.root.update()
        
        # Create a copy of the image for drawing
        result_img = self.captured_image.copy()
        
        # Create overlay for semi-transparent highlighting
        overlay = Image.new('RGBA', result_img.size, (255, 255, 255, 0))
        overlay_draw = ImageDraw.Draw(overlay)
        
        # For regular drawing (text and boxes)
        draw = ImageDraw.Draw(result_img)
        
        # Try to use a nice font
        try:
            font = ImageFont.truetype("arial.ttf", 12)
        except:
            font = ImageFont.load_default()
        
        # Build a grid structure by clustering positions
        if len(numbers_data) < 2:
            self.result_image = result_img
            self.display_result(0, len(numbers_data))
            return
            
        # Sort all unique x and y positions
        all_x = sorted([d['center_x'] for d in numbers_data])
        all_y = sorted([d['center_y'] for d in numbers_data])
        
        # Cluster nearby positions to find grid columns and rows
        def cluster_positions(positions, min_gap):
            """Group positions into clusters (columns or rows)"""
            if not positions:
                return []
            clusters = [[positions[0]]]
            for pos in positions[1:]:
                if pos - clusters[-1][-1] < min_gap:
                    clusters[-1].append(pos)
                else:
                    clusters.append([pos])
            # Return average position for each cluster
            return [sum(c) / len(c) for c in clusters]
        
        # Estimate cell size
        avg_w = sum(d['w'] for d in numbers_data) / len(numbers_data)
        avg_h = sum(d['h'] for d in numbers_data) / len(numbers_data)
        
        # Cluster to find grid lines (use half cell size as minimum gap)
        grid_cols = cluster_positions(all_x, avg_w * 0.6)
        grid_rows = cluster_positions(all_y, avg_h * 0.6)
        
        print(f"\nDetected grid: {len(grid_cols)} columns x {len(grid_rows)} rows")
        print(f"Cell size estimate: {avg_w:.1f} x {avg_h:.1f}")
        
        # Assign each number to a grid cell
        for data in numbers_data:
            # Find nearest column
            col_idx = min(range(len(grid_cols)), key=lambda i: abs(grid_cols[i] - data['center_x']))
            # Find nearest row
            row_idx = min(range(len(grid_rows)), key=lambda i: abs(grid_rows[i] - data['center_y']))
            data['grid_col'] = col_idx
            data['grid_row'] = row_idx
        
        # Create grid lookup: (col, row) -> data
        grid_lookup = {}
        for data in numbers_data:
            key = (data['grid_col'], data['grid_row'])
            if key not in grid_lookup:
                grid_lookup[key] = []
            grid_lookup[key].append(data)
        
        # For each cell, keep only the highest confidence match
        for key in grid_lookup:
            if len(grid_lookup[key]) > 1:
                grid_lookup[key] = [max(grid_lookup[key], key=lambda x: x.get('confidence', 0))]
        
        print(f"Grid has {len(grid_lookup)} unique cells populated")
        
        # Draw all detected numbers for debugging
        for data_list in grid_lookup.values():
            for data in data_list:
                x, y, w, h = data['x'], data['y'], data['w'], data['h']
                num = data['number']
                # Draw thin box around detected area
                draw.rectangle([x-1, y-1, x+w+1, y+h+1], outline='lightblue', width=1)
                # Draw detected number in GREEN at bottom-right corner of the cell
                text = str(num)
                # Position the green number at bottom-right of detected area
                text_x = x + w - 8
                text_y = y + h - 12
                # Draw with green color and slight shadow for visibility
                draw.text((text_x+1, text_y+1), text, fill='black', font=font)  # shadow
                draw.text((text_x, text_y), text, fill='lime', font=font)  # green number
        
        # Find all pairs that sum to 10 (horizontal or vertical adjacent)
        # Important: Once a cell is used in a pair, it cannot be used again
        pairs_found = []
        used_cells = set()
        
        # Sort by position (top-to-bottom, left-to-right) for consistent pairing
        sorted_cells = sorted(grid_lookup.items(), key=lambda x: (x[0][1], x[0][0]))
        
        for key, data_list in sorted_cells:
            # Skip if this cell is already used in a pair
            if key in used_cells or not data_list:
                continue
            
            data1 = data_list[0]
            num1 = data1['number']
            target = 10 - num1
            col, row = data1['grid_col'], data1['grid_row']
            
            # Check 4 adjacent cells: right, down, left, up
            # Prioritize right and down to avoid double-checking
            neighbors = [
                ((col + 1, row), 'horizontal', 'right'),
                ((col, row + 1), 'vertical', 'down'),
                ((col - 1, row), 'horizontal', 'left'),
                ((col, row - 1), 'vertical', 'up'),
            ]
            
            pair_found = False
            for neighbor_key, direction, dir_name in neighbors:
                # Check neighbor exists, is not used, and has the target number
                if neighbor_key in grid_lookup and neighbor_key not in used_cells:
                    data2_list = grid_lookup[neighbor_key]
                    if data2_list:
                        num2 = data2_list[0]['number']
                        # DEBUG: Check what we're comparing
                        if num1 + num2 == 10:
                            data2 = data2_list[0]
                            pairs_found.append((data1, data2, direction))
                            # Mark BOTH cells as used immediately
                            used_cells.add(key)
                            used_cells.add(neighbor_key)
                            print(f"Found {direction} pair ({dir_name}): {num1} + {num2} = 10 at grid ({col},{row}) and {neighbor_key}")
                            pair_found = True
                            break
                        else:
                            # Debug: log near-misses
                            if abs((num1 + num2) - 10) <= 2:
                                print(f"  Near miss at ({col},{row}): {num1} + {num2} = {num1+num2} (not 10)")
            
            # If we found a pair, this cell is now used
            if pair_found:
                continue
        
        # Highlight the pairs with semi-transparent red marker
        for data1, data2, is_horizontal in pairs_found:
            # Calculate bounding box for both numbers
            x1 = min(data1['x'], data2['x'])
            y1 = min(data1['y'], data2['y'])
            x2 = max(data1['x'] + data1['w'], data2['x'] + data2['w'])
            y2 = max(data1['y'] + data1['h'], data2['y'] + data2['h'])
            
            # Add padding
            padding = 3
            x1 -= padding
            y1 -= padding
            x2 += padding
            y2 += padding
            
            # Draw semi-transparent red marker
            overlay_draw.rectangle([x1, y1, x2, y2], fill=(255, 0, 0, 128), outline=(255, 0, 0, 255), width=2)
        
        # Composite the overlay onto the result image
        result_img = result_img.convert('RGBA')
        result_img = Image.alpha_composite(result_img, overlay)
        result_img = result_img.convert('RGB')
        
        self.result_image = result_img
        
        # Display the result
        self.display_result(len(pairs_found), len(grid_lookup))
        
    def display_result(self, pairs_count, total_numbers):
        """Display the result image"""
        # (No need to clear result_frame here; already cleared in capture_region)
        self._current_img_label = None

        # Update status
        self.status_label.config(
            text=f"Found {pairs_count} pairs (from {total_numbers} unique cells)",
            fg="green" if pairs_count > 0 else "orange"
        )

        # Enable the Save Result button
        if hasattr(self, 'save_btn'):
            self.save_btn.config(state=tk.NORMAL, bg="#2196F3")

        # Display the result image
        if self.result_image:
            # Resize if too large
            max_width = 900
            max_height = 700
            img = self.result_image.copy()

            if img.width > max_width or img.height > max_height:
                img.thumbnail((max_width, max_height), Image.Resampling.LANCZOS)

            # Convert to PhotoImage
            photo = ImageTk.PhotoImage(img)

            # Create label to display image
            img_label = tk.Label(self.result_frame, image=photo, bg=self.result_frame['bg'])
            img_label.image = photo  # Keep a reference
            img_label.pack()
            self._current_img_label = img_label
            
    def save_result(self):
        """Save the result image"""
        if self.result_image:
            filename = "result_sum_10.png"
            self.result_image.save(filename)
            messagebox.showinfo("Saved", f"Result saved as {filename}")
        else:
            # Should not be possible, but keep button disabled if no image
            if hasattr(self, 'save_btn'):
                self.save_btn.config(state=tk.DISABLED, bg="#888888")
            
    def run(self):
        """Run the application"""
        self.root.mainloop()

if __name__ == "__main__":
    root = tk.Tk()
    app = NumberSumCapture(root)
    app.run()