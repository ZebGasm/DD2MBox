import time
import psutil
import keyboard
from pynput import mouse
import win32gui
import win32con
import win32api
import win32process
from ctypes import windll
import tkinter as tk
from tkinter import ttk, messagebox
import threading
import sv_ttk # Import sv_ttk

class WindowManager(tk.Tk): # Inherit from tk.Tk
    """
    Manages game window layout and input broadcasting based on an AHK script.
    Includes a GUI for control.
    """
    def __init__(self):
        super().__init__() # Call parent constructor
        sv_ttk.use_dark_theme() # Apply the dark theme
        self.title("DD2 Window Manager")
        self.geometry("400x350") # Adjust size as needed

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

        self.refresh_monitor_work_area()
        self.create_widgets() # New method to set up GUI
        
        # Register hotkeys for F8
        # F8 hotkey removed
        # Register rotation hotkeys unconditionally
        self.rotation_hook_ids['up'] = keyboard.add_hotkey('up', lambda: self.rotate_main_window('up'))
        self.rotation_hook_ids['down'] = keyboard.add_hotkey('down', lambda: self.rotate_main_window('down'))
        self.update_status("UP/DOWN hotkeys registered unconditionally.")

        # Find and apply initial window layout
        self.find_dd2_windows()
        self.apply_layout()

        self.update_status("DD2 Window Manager GUI Initialized.")
        self.update_status("Hotkeys: UP/DOWN to rotate main.")

        # Ensure the mainloop is running to process GUI events
        self.protocol("WM_DELETE_WINDOW", self._on_closing)

    def create_widgets(self):
        # Frame for hotkey controls
        hotkey_frame = ttk.LabelFrame(self, text="Hotkey Controls", padding=10)
        hotkey_frame.pack(padx=10, pady=10, fill=tk.X)

        # Rotation hotkeys toggle
        self.rotation_status_label = ttk.Label(hotkey_frame, text="Rotation Hotkeys: Enabled", foreground="green")
        self.rotation_status_label.pack(anchor=tk.W, pady=5)
        self.rotation_toggle_button = ttk.Button(hotkey_frame, text="Toggle Rotation Hotkeys (Up/Down)", command=self._toggle_rotation_hotkeys)
        self.rotation_toggle_button.pack(fill=tk.X, pady=2)

        # Select mode hotkey (F8 is always registered, but this button could initiate a different flow if needed,
        # or just be informational)
        ttk.Label(hotkey_frame, text="").pack(anchor=tk.W, pady=5)
        
        # Add a refresh button for windows
        ttk.Button(self, text="Refresh DD2 Windows", command=self._refresh_dd2_windows).pack(fill=tk.X, padx=10, pady=5)
        
        # Status/Info Area
        info_frame = ttk.LabelFrame(self, text="Status/Log", padding=5)
        info_frame.pack(padx=10, pady=5, fill=tk.BOTH, expand=True)

        self.status_text = tk.Text(info_frame, height=5, state=tk.DISABLED, wrap=tk.WORD, font=("Consolas", 9))
        self.status_text.pack(fill=tk.BOTH, expand=True)
        self.status_scrollbar = ttk.Scrollbar(info_frame, command=self.status_text.yview)
        self.status_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.status_text.config(yscrollcommand=self.status_scrollbar.set)

    def update_status(self, message):
        """Updates the log area in the GUI."""
        timestamp = time.strftime("[%H:%M:%S]")
        self.status_text.config(state=tk.NORMAL)
        self.status_text.insert(tk.END, f"{timestamp} {message}\n")
        self.status_text.see(tk.END) # Scroll to the end
        self.status_text.config(state=tk.DISABLED)

    def _toggle_rotation_hotkeys(self):
        """Toggles the state of UP/DOWN hotkeys."""
        self.update_status(f"Toggle button pressed. Current rotation_hotkeys_enabled: {self.rotation_hotkeys_enabled}")
        self.rotation_hotkeys_enabled = not self.rotation_hotkeys_enabled
        self.update_status(f"After toggle, rotation_hotkeys_enabled: {self.rotation_hotkeys_enabled}")
        if self.rotation_hotkeys_enabled:
            self._enable_rotation_hotkeys()
            self.rotation_status_label.config(text="Rotation Hotkeys: Enabled", foreground="green")
            self.update_status("Rotation hotkeys ENABLED.")
        else:
            self._disable_rotation_hotkeys()
            self.rotation_status_label.config(text="Rotation Hotkeys: Disabled", foreground="red")
            self.update_status("Rotation hotkeys DISABLED.")

    def _enable_rotation_hotkeys(self):
        """Prepares for UP and DOWN hotkeys to be active."""
        # Hotkeys are always registered in __init__, so no need to add them here.
        # This function primarily updates the logical state and GUI.
        self.update_status("UP/DOWN hotkeys are logically ENABLED.")

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
        if self.mouse_listener and self.mouse_listener.running:
            self.mouse_listener.stop()
        keyboard.unhook_all()
        self.destroy() # Destroy the Tkinter window

def main():
    app = WindowManager()
    app.mainloop()

if __name__ == "__main__":
    main()