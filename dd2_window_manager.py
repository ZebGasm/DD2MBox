import time
import psutil
import keyboard
from pynput.mouse import Button, Controller as MouseController
from pynput import mouse
import win32gui
import win32con
import win32api
import win32process
import pywintypes # Added for win32api types
from ctypes import windll
import tkinter as tk
from tkinter import ttk, messagebox
import threading
import json
import os
import sys

class DraggableSquare(tk.Frame):
    """A simple draggable frame widget."""
    def __init__(self, master, box_number, bg_color='grey', text_color='white', **kwargs): # Added colors
        # Extract width and height from kwargs, provide defaults if not present
        width = kwargs.pop('width', 50)
        height = kwargs.pop('height', 50)
        
        super().__init__(master, **kwargs) # Pass remaining kwargs to super
        self.configure(bg=bg_color, width=width, height=height, cursor="fleur") # Use extracted width/height
        self._drag_data = {"x": 0, "y": 0}
        self.bind("<ButtonPress-1>", self._on_press)
        self.bind("<B1-Motion>", self._on_drag)

        # Add a label to display the number
        self.number_label = tk.Label(self, text=str(box_number), bg=bg_color, fg=text_color, font=("Arial", 12, "bold")) # Use bg_color, text_color
        self.number_label.place(relx=0.5, rely=0.5, anchor=tk.CENTER)
        # Bind events to the label as well, so dragging works even if clicking on the number
        self.number_label.bind("<ButtonPress-1>", self._on_press)
        self.number_label.bind("<B1-Motion>", self._on_drag)
        # Make sure the label doesn't "eat" mouse events, passing them to the parent frame
        self.number_label.bind("<ButtonRelease-1>", lambda event: self.event_generate("<ButtonRelease-1>", x=event.x, y=event.y))

    def _on_press(self, event):
        """Records the initial click position."""
        self._drag_data["x"] = event.x
        self._drag_data["y"] = event.y

    def _on_drag(self, event):
        """Moves the widget with the mouse."""
        x = self.winfo_pointerx() - self._drag_data["x"]
        y = self.winfo_pointery() - self._drag_data["y"]
        self.place(x=x, y=y)

class ShoppingOverlay(tk.Toplevel):
    """
    An overlay window for shopping mode with draggable, semi-transparent squares.
    """
    def __init__(self, master, initial_positions=None):
        super().__init__(master)
        self.master = master
        self.overrideredirect(True)
        self.attributes('-topmost', True)
        # self.attributes('-alpha', 0.6) # Removed to avoid conflict with transparentcolor

        # Use a transparent color to make the window background invisible
        # Using magenta, a color unlikely to be in game content
        self.transparent_color = 'magenta'
        self.attributes('-transparentcolor', self.transparent_color)
        self.config(bg=self.transparent_color)

        # Fullscreen geometry
        self.geometry(f"{self.winfo_screenwidth()}x{self.winfo_screenheight()}+0+0")

        self.shopping_squares = []
        self.utility_squares = []

        # Default positions for shopping boxes if not provided or malformed
        default_shopping_positions = [{'x': 50 + (i * 60), 'y': 50} for i in range(8)]
        shopping_positions = initial_positions.get('shopping_boxes', default_shopping_positions) if isinstance(initial_positions, dict) else default_shopping_positions

        # Default positions for utility boxes (3 of them, different starting point)
        default_utility_positions = [{'x': 50 + (i * 60), 'y': 150} for i in range(3)] # Changed range to 3
        utility_positions = initial_positions.get('utility_boxes', default_utility_positions) if isinstance(initial_positions, dict) else default_utility_positions

        for i, pos in enumerate(shopping_positions):
            square = DraggableSquare(self, box_number=i+1, bg_color='grey', text_color='white', width=50, height=50) # Explicitly set size
            square.place(x=pos.get('x', 50 + (i * 60)), y=pos.get('y', 50))
            self.shopping_squares.append(square)

        for i, pos in enumerate(utility_positions):
            square = DraggableSquare(self, box_number=i+1, bg_color='purple', text_color='white', width=25, height=25) # Half size
            square.place(x=pos.get('x', 50 + (i * 60)), y=pos.get('y', 150))
            self.utility_squares.append(square)

        self.is_click_through = False
        self.master.update_status("Shopping overlay created in SETUP mode.")

    def get_box_positions(self):
        """Returns the current x, y coordinates of all squares."""
        shopping_positions = []
        for square in self.shopping_squares:
            shopping_positions.append({'x': square.winfo_x(), 'y': square.winfo_y()})
        
        utility_positions = []
        for square in self.utility_squares:
            utility_positions.append({'x': square.winfo_x(), 'y': square.winfo_y()})
            
        return {'shopping_boxes': shopping_positions, 'utility_boxes': utility_positions}

    def set_click_through(self, enable: bool):
        """Toggles the click-through property of the overlay window."""
        hwnd = self.winfo_id()
        try:
            current_style = windll.user32.GetWindowLongW(hwnd, win32con.GWL_EXSTYLE)
            if enable:
                new_style = current_style | win32con.WS_EX_LAYERED | win32con.WS_EX_TRANSPARENT
                windll.user32.SetWindowLongW(hwnd, win32con.GWL_EXSTYLE, new_style)
                
                # Make the magenta background transparent
                win32gui.SetLayeredWindowAttributes(hwnd, win32api.RGB(255, 0, 255), 0, win32con.LWA_COLORKEY)
                
                win32gui.SetWindowPos(hwnd, 0, 0, 0, 0, 0, win32con.SWP_NOMOVE | win32con.SWP_NOSIZE | win32con.SWP_NOZORDER | win32con.SWP_FRAMECHANGED)
                self.is_click_through = True
                self.master.update_status("Click-through enabled.")
            else:
                # To make it interactive again, remove the transparent style
                new_style = current_style & ~win32con.WS_EX_TRANSPARENT
                windll.user32.SetWindowLongW(hwnd, win32con.GWL_EXSTYLE, new_style)
                win32gui.SetWindowPos(hwnd, 0, 0, 0, 0, 0, win32con.SWP_NOMOVE | win32con.SWP_NOSIZE | win32con.SWP_NOZORDER | win32con.SWP_FRAMECHANGED)
                self.is_click_through = False
                self.master.update_status("Click-through disabled.")
        except Exception as e:
            self.master.update_status(f"Error setting window style: {e}")

class WindowManager(tk.Tk): # Inherit from tk.Tk
    """
    Manages game window layout and input broadcasting based on an AHK script.
    Includes a GUI for control.
    """
    def __init__(self):
        super().__init__() # Call parent constructor
        self.title("Zeb DD2 Script")
        self.geometry("900x400") # Adjust size as needed

        # Determine the application path for PyInstaller compatibility
        if getattr(sys, 'frozen', False):
            self.application_path = os.path.dirname(sys.executable)
        else:
            self.application_path = os.path.dirname(os.path.abspath(__file__))

        self.target_exe = "DunDefGame.exe"
        self.main_w = 1720
        self.main_h = 900
        self.padding = 0
        
        self.dd2_windows = []  # List of HWNDs
        self.main_window_index = 0
        self.last_main_hwnd = 0
        
        self.select_mode = False
        self.mouse_listener = None
        self.rotation_hook_ids = {} # hotkey hook references

        self.rotation_hotkeys_enabled = True # New state variable

        self.key_map = {
            'g': ord('G'),
            'esc': win32con.VK_ESCAPE,
            'pgup': win32con.VK_PRIOR,
            'pgdn': win32con.VK_NEXT,
            'm': ord('M'),
            'y': ord('Y'),
        }
        self.key_delay_ms = 20 # From AHK script

        self.g_presser_enabled = False
        self.inactive_sender_enabled = False

        self.intervalG = 1200 # Time between 'G' presses (in milliseconds)
        self.keyGSend = "g"   # The key to send for G-presser
        self.g_presser_timer = None # For F7 functionality
        self.inactive_sender_timer = None # For F10 functionality
        self.mouse_controller = MouseController()
        self.shopping_cycle_count = 0 # Track completed shopping box cycles
        self.utility_box_cycle_index = 0 # New: Track which utility box to interact with (0 for box 2, 1 for box 3)
        self.shopping_loop_id = None # Store the ID of the scheduled shopping loop

        # Shopping Mode state
        self.shopping_mode_state = "OFF" # OFF, SETUP, AUTO-RUN
        self.shopping_overlay = None
        self.shopping_config_file = 'shopping_overlay_config.json'
        self.status_text = None # Initialize status_text to None
        self.box_positions = self._load_box_positions()
        self.original_cursor_pos = None
        self.esc_hook_id = None # Initialize esc hotkey hook id

        self._apply_terminal_theme() # Apply the new theme
        self.refresh_monitor_work_area()
        self.create_widgets() # New method to set up GUI
        
        self._register_ahk_hotkeys()
        # Initialize rotation hotkeys based on the rotation_hotkeys_enabled state
        if self.rotation_hotkeys_enabled:
            self._enable_rotation_hotkeys()

        # Find and apply initial window layout
        self.find_dd2_windows()
        self.apply_layout()

        self.update_status("DD2 Window Manager GUI Initialized.")
        self.update_status("Hotkeys: UP/DOWN to rotate main.")

        # Ensure the mainloop is running to process GUI events
        self.protocol("WM_DELETE_WINDOW", self._on_closing)

    def _apply_terminal_theme(self):
        """Creates and applies a custom 'cool terminal' theme using Catppuccin Mocha colors."""
        self.theme = {
            'bg': '#1E1E2E',  # Catppuccin Mocha: Base
            'fg': '#CDD6F4',  # Catppuccin Mocha: Text (main foreground)
            'bg_alt': '#181825', # Catppuccin Mocha: Mantle (for text area background)
            'fg_alt': '#A6E3A1', # Catppuccin Mocha: Green (for accents/status labels)
            'font': ('Consolas', 10),
            'font_bold': ('Consolas', 10, 'bold')
        }

        # Apply to root window
        self.configure(bg=self.theme['bg'])

        # Create and configure ttk style
        style = ttk.Style(self)
        style.theme_use('clam')

        # General widget styling
        style.configure('.',
                        background=self.theme['bg'],
                        foreground=self.theme['fg'],
                        font=self.theme['font'],
                        borderwidth=1)

        # Frame and LabelFrame
        style.configure('TFrame', background=self.theme['bg'])
        style.configure('TLabelframe',
                        background=self.theme['bg'],
                        relief='solid',
                        bordercolor=self.theme['fg_alt']) # Use accent for border
        style.configure('TLabelframe.Label',
                        background=self.theme['bg'],
                        foreground=self.theme['fg_alt'], # Use accent for label text
                        font=self.theme['font_bold'])

        # Label
        style.configure('TLabel', foreground=self.theme['fg']) # Default labels use main foreground
        style.configure('Status.TLabel', foreground=self.theme['fg_alt']) # Status labels use accent foreground

        # Button
        style.configure('TButton',
                        font=self.theme['font_bold'],
                        relief='solid',
                        bordercolor=self.theme['fg_alt'], # Use accent for border
                        background=self.theme['bg'], # Button background
                        foreground=self.theme['fg'], # Button text foreground
                        padding=5)
        style.map('TButton',
                  background=[('pressed', self.theme['fg_alt']), ('active', self.theme['bg_alt'])],
                  foreground=[('pressed', self.theme['bg']), ('active', self.theme['fg'])])
        
        # Scrollbar
        style.configure("TScrollbar",
                background=self.theme['bg'],
                troughcolor=self.theme['bg_alt'],
                bordercolor=self.theme['bg'],
                arrowcolor=self.theme['fg_alt']) # Use accent for arrow color
        style.map("TScrollbar",
                background=[('active', self.theme['fg_alt'])])


    def create_widgets(self):
        # 1. Status Log Area (top section of the window)
        info_frame = ttk.LabelFrame(self, text="[ STATUS LOG ]", padding=5)
        info_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=10, pady=(10, 5))

        self.status_text = tk.Text(info_frame, height=5, state=tk.DISABLED, wrap=tk.WORD,
                                   font=self.theme['font'],
                                   bg=self.theme['bg_alt'],
                                   fg=self.theme['fg'],
                                   insertbackground=self.theme['fg'], # Cursor color
                                   relief='solid',
                                   borderwidth=0)
        self.status_text.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)
        
        self.status_scrollbar = ttk.Scrollbar(info_frame, command=self.status_text.yview, style="TScrollbar")
        self.status_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.status_text.config(yscrollcommand=self.status_scrollbar.set)

        # 2. Control Panels Area (bottom section of the window, holds horizontal frames)
        control_panels_container = ttk.Frame(self) # This frame will hold the three horizontal frames
        control_panels_container.pack(side=tk.BOTTOM, fill=tk.BOTH, expand=True, padx=10, pady=(5, 10))

        # --- Hotkey Controls (leftmost) ---
        hotkey_frame = ttk.LabelFrame(control_panels_container, text="[ HOTKEY STATUS ]", padding=10)
        hotkey_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))

        self.rotation_status_label = ttk.Label(hotkey_frame, text="Rotation (Up/Down): ENABLED", style='Status.TLabel')
        self.rotation_status_label.pack(anchor=tk.W, pady=2)
        
        self.g_presser_status_label = ttk.Label(hotkey_frame, text="G-Presser (F7): OFF", style='Status.TLabel')
        self.g_presser_status_label.pack(anchor=tk.W, pady=2)

        self.inactive_sender_status_label = ttk.Label(hotkey_frame, text="Inactive Sender (F10): OFF", style='Status.TLabel')
        self.inactive_sender_status_label.pack(anchor=tk.W, pady=2)
        
        ttk.Label(hotkey_frame, text="Select Window (F8)").pack(anchor=tk.W, pady=2)
        ttk.Label(hotkey_frame, text="Emergency Kill (F9)").pack(anchor=tk.W, pady=2)

        # --- Actions (middle) ---
        actions_frame = ttk.LabelFrame(control_panels_container, text="[ ACTIONS ]", padding=10)
        actions_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)

        self.rotation_toggle_button = ttk.Button(actions_frame, text="Toggle Rotation", command=self._toggle_rotation_hotkeys)
        self.rotation_toggle_button.pack(fill=tk.X, pady=4)
        
        self.g_presser_toggle_button = ttk.Button(actions_frame, text="Toggle G-Presser", command=self._toggle_g_presser_gui)
        self.g_presser_toggle_button.pack(fill=tk.X, pady=4)
        
        self.inactive_sender_toggle_button = ttk.Button(actions_frame, text="Toggle Inactive Sender", command=self._toggle_inactive_sender_gui)
        self.inactive_sender_toggle_button.pack(fill=tk.X, pady=4)

        ttk.Button(actions_frame, text="Refresh DD2 Windows", command=self._refresh_dd2_windows).pack(fill=tk.X, pady=4)


        # --- Shopping Mode (rightmost) ---
        shopping_frame = ttk.LabelFrame(control_panels_container, text="[ SHOPPING MODE ]", padding=10)
        shopping_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(5, 0))

        self.shopping_status_label = ttk.Label(shopping_frame, text="Status: OFF", style='Status.TLabel')
        self.shopping_status_label.pack(anchor=tk.W, pady=2)

        self.shopping_toggle_button = ttk.Button(shopping_frame, text="Enable Shopping", command=self._toggle_shopping_mode)
        self.shopping_toggle_button.pack(fill=tk.X, pady=4)

    def update_status(self, message):
        """Updates the log area in the GUI or prints to console if GUI not ready."""
        timestamp = time.strftime("[%H:%M:%S]")
        formatted_message = f"{timestamp} {message}"
        if self.status_text: # Check if the widget has been created
            self.status_text.config(state=tk.NORMAL)
            self.status_text.insert(tk.END, formatted_message + "\n")
            self.status_text.see(tk.END) # Scroll to the end
            self.status_text.config(state=tk.DISABLED)
        else:
            print(formatted_message) # Print to console during early initialization

    def _load_box_positions(self):
        """Loads box positions from the config file."""
        default_shopping_positions = [{'x': 50 + (i * 60), 'y': 50} for i in range(8)]
        default_utility_positions = [{'x': 50 + (i * 60), 'y': 150} for i in range(3)] # Changed range to 3
        default_all_positions = {
            'shopping_boxes': default_shopping_positions,
            'utility_boxes': default_utility_positions
        }

        config_full_path = os.path.join(self.application_path, self.shopping_config_file)
        self.update_status(f"Attempting to load box configurations from: {config_full_path}")

        if not os.path.exists(config_full_path):
            self.update_status(f"Config file not found at {config_full_path}. Using defaults for all boxes.")
            return default_all_positions
        try:
            with open(config_full_path, 'r') as f:
                loaded_positions = json.load(f)
            # Basic validation: ensure it's a dict and contains expected keys/lists
            if (isinstance(loaded_positions, dict) and
                'shopping_boxes' in loaded_positions and isinstance(loaded_positions['shopping_boxes'], list) and len(loaded_positions['shopping_boxes']) == 8 and
                'utility_boxes' in loaded_positions and isinstance(loaded_positions['utility_boxes'], list) and len(loaded_positions['utility_boxes']) == 3): # Changed length to 3
                self.update_status("Loaded all box positions from config.")
                return loaded_positions
            else:
                self.update_status("Config file for box positions is malformed or incomplete. Using defaults for all boxes.")
                return default_all_positions
        except (json.JSONDecodeError, IOError) as e:
            self.update_status(f"Error loading box configurations from {config_full_path}: {e}. Using defaults for all boxes.")
            return default_all_positions

    def _save_box_positions(self, positions_dict): # Renamed argument for clarity
        """Saves box positions to the config file."""
        config_full_path = os.path.join(self.application_path, self.shopping_config_file)
        try:
            with open(config_full_path, 'w') as f:
                json.dump(positions_dict, f, indent=4)
            self.update_status(f"Saved all box positions to: {config_full_path}")
        except IOError as e:
            self.update_status(f"Error saving box configurations to {config_full_path}: {e}")

    def _shopping_loop(self, step_index=0, box_index=0):
        """The main non-blocking loop for auto-shopping, broken into granular steps."""
        if self.shopping_mode_state != "AUTO-RUN":
            self.update_status("Stopping shopping loop (early exit).")
            if self.original_cursor_pos:
                win32api.SetCursorPos(self.original_cursor_pos)
            self.shopping_loop_id = None # Clear ID as we are stopping
            return

        if step_index == 0: # Start of a new box interaction (move mouse)
            if box_index < len(self.box_positions['shopping_boxes']):
                pos = self.box_positions['shopping_boxes'][box_index]
                self.update_status(f"Auto-shopping at box {box_index + 1}: ({pos['x']}, {pos['y']})")
                self.mouse_controller.position = (pos['x'], pos['y'])
                self.shopping_loop_id = self.after(100, self._shopping_loop, 1, box_index) # Move to step 1 (first Enter) after 100ms
            else:
                # All shopping boxes processed for this cycle, proceed to utility boxes
                self.update_status("All shopping boxes processed in this cycle.")
                self.shopping_loop_id = self.after(100, self._shopping_loop, 5) # Move to utility box interaction (step 5) after 100ms
        
        elif step_index == 1: # First Enter
            keyboard.press_and_release('enter')
            self.shopping_loop_id = self.after(1000, self._shopping_loop, 2, box_index) # Move to step 2 (second Enter) after 1s
            
        elif step_index == 2: # Second Enter
            keyboard.press_and_release('enter')
            self.shopping_loop_id = self.after(1000, self._shopping_loop, 3, box_index) # Move to step 3 (third Enter) after 1s
            
        elif step_index == 3: # Third Enter
            keyboard.press_and_release('enter')
            self.update_status(f"Finished box {box_index+1}. Waiting before next box (15 seconds).")
            self.shopping_loop_id = self.after(15000, self._shopping_loop, 0, box_index + 1) # Move to next box (step 0 for next box) after 15s
            
        elif step_index == 5: # Utility box interaction
            # --- New logic for utility boxes ---
            if self.utility_box_cycle_index == 0: # Current cycle needs utility box 2
                self.update_status("Interacting with Utility Box 2.")
                self._interact_with_utility_box(2) # Interact with utility box 2
                self.utility_box_cycle_index = 1 # Set next cycle to utility box 3
            else: # Current cycle needs utility box 3
                self.update_status("Interacting with Utility Box 3.")
                self._interact_with_utility_box(3) # Interact with utility box 3
                self.utility_box_cycle_index = 0 # Set next cycle to utility box 2

            self.shopping_cycle_count += 1 # Increment cycle count after utility box interaction

            if self.shopping_cycle_count >= 3: # Stop after 3 full cycles (Utility Box 3 interacted with, and 3rd shopping pass initiated)
                self.update_status("Completed all shopping and utility box interactions, and final shopping pass. Stopping auto-shopping.")
                if self.shopping_loop_id:
                    self.after_cancel(self.shopping_loop_id) # Cancel any pending after call
                self.shopping_loop_id = None # Clear ID
                self._toggle_shopping_mode() # Transition to OFF state
            else:
                self.update_status("Restarting shopping sequence from box 1.")
                self.shopping_loop_id = self.after(1000, self._shopping_loop, 0, 0) # Restart shopping loop (step 0, box 0) after 1s

    def _toggle_shopping_mode(self):
        """Cycles through the shopping mode states: OFF -> SETUP -> AUTO-RUN -> OFF."""
        if self.shopping_mode_state == "OFF":
            self.shopping_overlay = ShoppingOverlay(self, initial_positions=self.box_positions)
            self.shopping_mode_state = "SETUP"
            self.shopping_toggle_button.config(text="Start Auto-Shop")
            self.shopping_status_label.config(text="Status: SETUP")
            self.update_status("Shopping overlay enabled. Drag squares to position them.")

        elif self.shopping_mode_state == "SETUP":
            if self.shopping_overlay:
                self.box_positions = self.shopping_overlay.get_box_positions()
                self._save_box_positions(self.box_positions)
                self.shopping_overlay.set_click_through(True)
            
            for square in self.shopping_overlay.utility_squares:
                square.place_forget()
            self.update_status("Utility boxes hidden.")
            
            self.shopping_mode_state = "AUTO-RUN"
            self.shopping_toggle_button.config(text="Stop Auto-Shop")
            self.shopping_status_label.config(text="Status: AUTO-RUN")
            self.update_status("Auto-shopping started.")
            
            if self.dd2_windows and len(self.dd2_windows) > self.main_window_index:
                self._activate_window(self.dd2_windows[self.main_window_index])
                time.sleep(0.1)

            self.original_cursor_pos = win32api.GetCursorPos()
            self.esc_hook_id = keyboard.add_hotkey('ctrl+alt+s', self._handle_esc_press)
            self.update_status("Hotkeys: ctrl+alt+s to stop auto-shopping.")
            self._shopping_loop(0, 0)

        elif self.shopping_mode_state == "AUTO-RUN":
            if self.shopping_overlay:
                self.shopping_overlay.destroy()
                self.shopping_overlay = None

            if self.shopping_loop_id:
                self.after_cancel(self.shopping_loop_id)
                self.shopping_loop_id = None

            self.shopping_mode_state = "OFF"
            self.shopping_toggle_button.config(text="Enable Shopping")
            self.shopping_status_label.config(text="Status: OFF")
            self.update_status("Auto-shopping stopped.")
            self.shopping_cycle_count = 0
            
            if self.esc_hook_id:
                keyboard.remove_hotkey(self.esc_hook_id)
                self.esc_hook_id = None

    def _handle_esc_press(self):
        """Handler for Esc key press to stop auto-shopping."""
        if self.shopping_mode_state == "AUTO-RUN":
            self.update_status("ESC pressed. Stopping auto-shopping.")
            self._toggle_shopping_mode()

    def _interact_with_utility_box(self, box_number):
        """Temporarily hides the entire overlay, clicks a utility box position, and shows overlay again."""
        if not self.shopping_overlay or not self.shopping_overlay.utility_squares:
            self.update_status(f"Error: Utility boxes not available to interact with box #{box_number}.")
            return

        if 1 <= box_number <= len(self.shopping_overlay.utility_squares):
            utility_square = self.shopping_overlay.utility_squares[box_number - 1]
            x = utility_square.winfo_x() + utility_square.winfo_width() // 2 # Center of the box
            y = utility_square.winfo_y() + utility_square.winfo_height() // 2 # Center of the box

            original_pos = win32api.GetCursorPos()

            if self.dd2_windows and len(self.dd2_windows) > self.main_window_index:
                self._activate_window(self.dd2_windows[self.main_window_index])
            
            self.update_status(f"Temporarily hiding overlay to click utility box #{box_number}.")
            self.shopping_overlay.withdraw()
            time.sleep(0.05)

            self.mouse_controller.position = (x, y)
            time.sleep(0.1)
            self.mouse_controller.click(Button.left, 1)
            time.sleep(0.5)

            self.shopping_overlay.deiconify()
            self.update_status(f"Restored overlay after clicking utility box #{box_number}.")
            time.sleep(0.05)

            win32api.SetCursorPos(original_pos)

        else:
            self.update_status(f"Error: Utility box #{box_number} is out of range.")
            
    def _toggle_rotation_hotkeys(self):
        """Toggles the state of UP/DOWN hotkeys."""
        self.rotation_hotkeys_enabled = not self.rotation_hotkeys_enabled
        if self.rotation_hotkeys_enabled:
            self._enable_rotation_hotkeys()
            self.rotation_status_label.config(text="Rotation (Up/Down): ENABLED")
            self.update_status("Rotation hotkeys ENABLED.")
        else:
            self._disable_rotation_hotkeys()
            self.rotation_status_label.config(text="Rotation (Up/Down): DISABLED")
            self.update_status("Rotation hotkeys DISABLED.")

    def _enable_rotation_hotkeys(self):
        """Registers the UP and DOWN hotkeys to be active."""
        self.rotation_hook_ids['up'] = keyboard.add_hotkey('up', lambda: self.rotate_main_window('up'))
        self.rotation_hook_ids['down'] = keyboard.add_hotkey('down', lambda: self.rotate_main_window('down'))
        self.update_status("UP/DOWN hotkeys are ENABLED and registered.")

    def _disable_rotation_hotkeys(self):
        """Disables the UP and DOWN hotkeys."""
        if 'up' in self.rotation_hook_ids:
            self.rotation_hook_ids['up']() # Call the stored unhook function
            del self.rotation_hook_ids['up']
        if 'down' in self.rotation_hook_ids:
            self.rotation_hook_ids['down']() # Call the stored unhook function
            del self.rotation_hook_ids['down']
        self.update_status("UP/DOWN hotkeys unhooked.")



    def refresh_monitor_work_area(self):
        """Gets the primary monitor's work area, excluding the taskbar."""
        monitor_info = win32api.GetMonitorInfo(win32api.MonitorFromPoint((0, 0)))
        work_area = monitor_info['Work']
        self.m_left, self.m_top, self.m_right, self.m_bottom = work_area

    def find_dd2_windows(self):
        """
        Finds all active window handles (HWND) for the target executable.
        """
        new_window_list = []
        pids = [p.info['pid'] for p in psutil.process_iter(['pid', 'name']) if p.info['name'] == self.target_exe]

        def callback(hwnd, hwnds):
            if not win32gui.IsWindowVisible(hwnd) or not win32gui.IsWindowEnabled(hwnd):
                return True
            _, found_pid = win32process.GetWindowThreadProcessId(hwnd)
            if found_pid in pids:
                hwnds.append(hwnd)
            return True

        win32gui.EnumWindows(callback, new_window_list)
        self.dd2_windows = new_window_list
        return len(self.dd2_windows)

    def rotate_main_window(self, direction='up'):
        """
        Rotates the main window selection.
        """
        if not self.rotation_hotkeys_enabled:
            self.update_status("Rotation hotkeys are disabled. Cannot rotate main window.")
            return
        if self.find_dd2_windows() < 1:
            self.update_status("No DD2 windows found to rotate.")
            return

        start_index = 0
        if self.last_main_hwnd in self.dd2_windows:
            try:
                start_index = self.dd2_windows.index(self.last_main_hwnd)
            except ValueError:
                start_index = self.main_window_index # Fallback
        else:
             start_index = self.main_window_index

        if direction == 'up':
            self.main_window_index = start_index + 1
        else: # 'down'
            self.main_window_index = start_index - 1

        if self.main_window_index >= len(self.dd2_windows):
            self.main_window_index = 0
        if self.main_window_index < 0:
            self.main_window_index = len(self.dd2_windows) - 1
            
        self.update_status(f"Rotating main window. New main index: {self.main_window_index}")
        self.apply_layout()

    def apply_layout(self):
        """
        Applies the window layout based on the current main window.
        """
        if not self.dd2_windows:
            self.update_status("Cannot apply layout, no windows found.")
            return

        main_hwnd = self.dd2_windows[self.main_window_index]
        self.last_main_hwnd = main_hwnd

        secondary_windows = [hwnd for hwnd in self.dd2_windows if hwnd != main_hwnd]
        
        # 1. Restore all windows first to ensure they can be moved
        for hwnd in self.dd2_windows:
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        time.sleep(0.05)

        # 2. Position MAIN window
        x_main, y_main = self.m_left, self.m_top
        win32gui.SetWindowPos(main_hwnd, win32con.HWND_TOP, x_main, y_main, self.main_w, self.main_h, win32con.SWP_SHOWWINDOW)

        # 3. Position SECONDARY windows
        sec1_hwnd = secondary_windows[0] if len(secondary_windows) > 0 else None
        sec2_hwnd = secondary_windows[1] if len(secondary_windows) > 1 else None
        sec3_hwnd = secondary_windows[2] if len(secondary_windows) > 2 else None

        # SIDE window (sec1)
        if sec1_hwnd:
            side_w = self.m_right - (x_main + self.main_w) - self.padding
            side_h = self.main_h
            x = x_main + self.main_w + self.padding
            y = y_main
            win32gui.SetWindowPos(sec1_hwnd, win32con.HWND_NOTOPMOST, x, y, side_w, side_h, win32con.SWP_SHOWWINDOW | win32con.SWP_NOACTIVATE)
            
        # BOTTOM windows (sec2 & sec3)
        bottom_h = self.m_bottom - (y_main + self.main_h) - self.padding * 2
        sec_w = (self.main_w - self.padding) // 2
        
        if sec2_hwnd:
            x = x_main
            y = y_main + self.main_h + self.padding
            win32gui.SetWindowPos(sec2_hwnd, win32con.HWND_NOTOPMOST, x, y, sec_w, bottom_h, win32con.SWP_SHOWWINDOW | win32con.SWP_NOACTIVATE)

        if sec3_hwnd:
            x = x_main + sec_w + self.padding
            y = y_main + self.main_h + self.padding
            win32gui.SetWindowPos(sec3_hwnd, win32con.HWND_NOTOPMOST, x, y, sec_w, bottom_h, win32con.SWP_SHOWWINDOW | win32con.SWP_NOACTIVATE)
            
        # 4. Activate the main window
        self._activate_window(main_hwnd)
        self.update_status("Layout applied.")

    def _activate_window(self, hwnd):
        """Robustly activate a window with retries."""
        try:
            current_foreground = win32gui.GetForegroundWindow()
            if current_foreground == hwnd:
                return

            # Store the current foreground window's thread ID
            current_foreground_thread_id = windll.user32.GetWindowThreadProcessId(current_foreground, None)
            target_thread_id = windll.user32.GetWindowThreadProcessId(hwnd, None)

            # Attach thread input to steal focus if different threads
            if current_foreground_thread_id != target_thread_id:
                windll.user32.AttachThreadInput(current_foreground_thread_id, target_thread_id, True)
                time.sleep(0.01) # Small delay for attachment to take effect

            # Attempt to set foreground window multiple times
            max_retries = 5
            for i in range(max_retries):
                win32gui.BringWindowToTop(hwnd)
                win32gui.SetForegroundWindow(hwnd)
                # Check if it's foreground
                if win32gui.GetForegroundWindow() == hwnd:
                    self.update_status(f"Window {hwnd} activated successfully after {i+1} attempts.")
                    break
                time.sleep(0.05) # Wait a bit before retrying
            else:
                self.update_status(f"Warning: Window {hwnd} might not be foreground after {max_retries} attempts.")

            # Detach thread input
            if current_foreground_thread_id != target_thread_id:
                windll.user32.AttachThreadInput(current_foreground_thread_id, target_thread_id, False)

        except Exception as e:
            self.update_status(f"Error activating window {hwnd}: {e}")

    def toggle_select_mode(self):
        """Toggles mouse selection mode for setting the main window."""
        self.select_mode = not self.select_mode
        if self.select_mode:
            self.update_status("SELECT MODE ACTIVATED: Click on a DD2 window to make it main.")
            self.update_status("Press F8 again to cancel.")
            if self.mouse_listener is None or not self.mouse_listener.running:
                self.mouse_listener = mouse.Listener(on_click=self._on_global_mouse_click)
                self.mouse_listener.start()
        else:
            self.update_status("SELECT MODE DEACTIVATED.")
            if self.mouse_listener and self.mouse_listener.running:
                self.mouse_listener.stop()
                self.mouse_listener = None # Clear listener after stopping

    def _on_global_mouse_click(self, x, y, button, pressed):
        """Global mouse click handler for select mode."""
        if not self.select_mode or not pressed or button != mouse.Button.left:
            return True # Continue listening

        clicked_hwnd = win32gui.WindowFromPoint((x, y))
        
        # Check if the clicked window is one of the DD2 windows
        if clicked_hwnd in self.dd2_windows:
            self.last_main_hwnd = clicked_hwnd
            self.main_window_index = self.dd2_windows.index(clicked_hwnd)
            self.update_status(f"Window {clicked_hwnd} selected as new main.")
            
            # Deactivate select mode and apply layout
            self.after(0, self.toggle_select_mode) # Use after() to call from GUI thread
            self.after(0, self.apply_layout)     # Use after() to call from GUI thread
            
            return False # Stop the listener
        else:
            self.update_status("Clicked window is not a detected DD2 window.")
            return True # Continue listening if not a DD2 window

    def _refresh_dd2_windows(self):
        """Refreshes the list of DD2 windows and updates status."""
        count = self.find_dd2_windows()
        self.update_status(f"Refreshed: Found {count} DD2 windows.")
        if count > 0:
            # Re-apply layout to ensure the 'main' window is visually correct
            self.apply_layout() 
        else:
            self.dd2_windows = [] # Clear the list if no windows are found
            self.main_window_index = 0
            self.last_main_hwnd = 0

    def _on_closing(self):
        """Handles proper shutdown when the GUI window is closed."""
        self.update_status("GUI window closed. Exiting application.")
        if self.shopping_overlay:
            # Save positions before closing if overlay is open
            all_box_positions = self.shopping_overlay.get_box_positions()
            self._save_box_positions(all_box_positions)
            self.shopping_overlay.destroy()
        if self.mouse_listener and self.mouse_listener.running:
            self.mouse_listener.stop()
        if self.esc_hook_id: # Ensure esc hotkey is removed on exit
            keyboard.remove_hotkey(self.esc_hook_id)
        if self.original_cursor_pos: # Restore mouse cursor if it was moved by auto-run
            win32api.SetCursorPos(self.original_cursor_pos)
            self.original_cursor_pos = None
        keyboard.unhook_all()
        self.destroy() # Destroy the Tkinter window
        sys.exit(0) # Ensure the entire script exits

    def _send_key_to_all_dd2_windows(self, key_name):
        """Sends a specified key to all detected DD2 windows."""
        vk_code = self.key_map.get(key_name.lower())
        if not vk_code:
            self.update_status(f"Error: Unknown key '{key_name}' for sending.")
            return

        self.find_dd2_windows() # Refresh window list
        if not self.dd2_windows:
            self.update_status("No DD2 windows found to send key to.")
            return

        for hwnd in self.dd2_windows:
            self._send_key_to_window(hwnd, vk_code, self.key_delay_ms)
        self.update_status(f"Sent '{key_name}' to all DD2 windows.")

    def _send_key_to_inactive_dd2_windows(self, key_name):
        """Sends a specified key to all detected DD2 windows, excluding the foreground window."""
        vk_code = self.key_map.get(key_name.lower())
        if not vk_code:
            self.update_status(f"Error: Unknown key '{key_name}' for sending.")
            return

        self.find_dd2_windows() # Refresh window list
        if not self.dd2_windows:
            self.update_status("No DD2 windows found to send key to.")
            return

        active_hwnd = win32gui.GetForegroundWindow()
        for hwnd in self.dd2_windows:
            if hwnd != active_hwnd:
                self._send_key_to_window(hwnd, vk_code, self.key_delay_ms)
        self.update_status(f"Sent '{key_name}' to inactive DD2 windows.")

    def _g_presser_loop(self):
        if self.g_presser_enabled:
            self._send_key_to_all_dd2_windows(self.keyGSend)
            self.g_presser_timer = self.after(self.intervalG, self._g_presser_loop)

    def _toggle_g_presser(self):
        self.g_presser_enabled = not self.g_presser_enabled
        if self.g_presser_enabled:
            self.update_status(f"G-Presser ON: Sending '{self.keyGSend}' every {self.intervalG}ms.")
            self._g_presser_loop() # Start the loop
        else:
            if self.g_presser_timer:
                self.after_cancel(self.g_presser_timer)
                self.g_presser_timer = None
            self.update_status("G-Presser OFF.")

    def _inactive_sender_loop(self):
        if self.inactive_sender_enabled:
            self._send_key_to_inactive_dd2_windows(self.keyGSend)
            self._send_key_to_inactive_dd2_windows('esc')
            self.inactive_sender_timer = self.after(100, self._inactive_sender_loop) # 100ms as per AHK script

    def _toggle_inactive_sender(self):
        self.inactive_sender_enabled = not self.inactive_sender_enabled
        if self.inactive_sender_enabled:
            self.update_status(f"Inactive Sender ON: Sending '{self.keyGSend}' and 'Esc' to inactive windows every 100ms.")
            self._inactive_sender_loop() # Start the loop
        else:
            if self.inactive_sender_timer:
                self.after_cancel(self.inactive_sender_timer)
                self.inactive_sender_timer = None
            self.update_status("Inactive Sender OFF.")

    def _toggle_g_presser_gui(self):
        self._toggle_g_presser()
        if self.g_presser_enabled:
            self.g_presser_status_label.config(text="G-Presser (F7): ON")
        else:
            self.g_presser_status_label.config(text="G-Presser (F7): OFF")

    def _toggle_inactive_sender_gui(self):
        self._toggle_inactive_sender()
        if self.inactive_sender_enabled:
            self.inactive_sender_status_label.config(text="Inactive Sender (F10): ON")
        else:
            self.inactive_sender_status_label.config(text="Inactive Sender (F10): OFF")

    def _register_ahk_hotkeys(self):
        """Registers hotkeys from the AHK script functionality."""
        keyboard.add_hotkey('f7', self._toggle_g_presser_gui)
        keyboard.add_hotkey('f10', self._toggle_inactive_sender_gui)
        keyboard.add_hotkey('p', lambda: self._send_key_to_all_dd2_windows('esc'))
        keyboard.add_hotkey('n', lambda: self._send_key_to_all_dd2_windows('m'))
        keyboard.add_hotkey('y', lambda: self._send_key_to_all_dd2_windows('y'))
        keyboard.add_hotkey('g', lambda: self._send_key_to_all_dd2_windows('g'))
        keyboard.add_hotkey('page up', lambda: self._send_key_to_all_dd2_windows('pgup'))
        keyboard.add_hotkey('page down', lambda: self._send_key_to_all_dd2_windows('pgdn'))
        keyboard.add_hotkey('b', lambda: self._send_key_to_inactive_dd2_windows('m'))
        
        # F9 (Emergency Kill Switch)
        keyboard.add_hotkey('f9', self._on_closing)
        self.update_status("AHK hotkeys registered.")

    def _send_key_to_window(self, hwnd, vk_code, key_delay=20):
        """Sends a key press (down and up) to a specific window."""
        try:
            win32api.PostMessage(hwnd, win32con.WM_KEYDOWN, vk_code, 0)
            time.sleep(key_delay / 1000.0) # Convert ms to seconds
            win32api.PostMessage(hwnd, win32con.WM_KEYUP, vk_code, 0)
        except pywintypes.error as e:
            self.update_status(f"Error sending key to window {hwnd}: {e}")

def main():
    app = WindowManager()
    app.mainloop()

if __name__ == "__main__":
    main()
