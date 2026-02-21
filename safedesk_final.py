import cv2
import customtkinter as ctk
from ultralytics import YOLO
import ctypes
import threading
import sqlite3
import os
import pandas as pd
from datetime import datetime
from tkinter import messagebox
import time
import sys

# TRAY imports
import pystray
from PIL import Image as PILImage


def resource_path(relative_path: str) -> str:
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


def get_app_data_path():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.abspath(".")


# SET PATHS
APP_DATA_PATH = get_app_data_path()
DB_PATH = os.path.join(APP_DATA_PATH, "compliance_logs.db")
ALERTS_DIR = os.path.join(APP_DATA_PATH, "Alerts")
MODEL_PATH = resource_path("yolov8n.pt")
TRAY_ICON_PATH = resource_path("SafedeskAI.ico")

os.makedirs(ALERTS_DIR, exist_ok=True)

# Database
conn = sqlite3.connect(DB_PATH, check_same_thread=False)
cursor = conn.cursor()
cursor.execute('''CREATE TABLE IF NOT EXISTS violations 
                 (id INTEGER PRIMARY KEY AUTOINCREMENT, timestamp TEXT, 
                  employee_name TEXT, object_detected TEXT, image_path TEXT)''')
conn.commit()


class SafeDeskApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        # 🔥 SHOW UI IMMEDIATELY (before anything else)
        self.title("🛡️ SafeDesk AI - Professional Monitoring")
        self.geometry("1100x650")
        ctk.set_appearance_mode("dark")

        # Show loading state first
        self.loading_label = ctk.CTkLabel(self, text="⏳ Loading SafeDesk AI...",
                                          font=ctk.CTkFont(size=24, weight="bold"))
        self.loading_label.place(relx=0.5, rely=0.5, anchor="center")
        self.update()  # Force UI update

        # Initialize variables
        self.monitoring = False
        self.model = None
        self.model_loading = False  # Track if model is being loaded
        self.cap = None
        self.conf_threshold = 0.40
        self.alert_cooldown_seconds = 10
        self.required_streak = 3
        self.detect_streak = 0
        self.action_mode = "Log Only"
        self.last_alert_time = 0.0
        self.tray_icon = None
        self.tray_running = False
        self.open_photo_windows = {}

        # Setup UI immediately
        self.setup_ui()
        self.loading_label.destroy()  # Remove loading message

        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.start_tray()

        # 🔥 Load model in background AFTER UI is shown
        self.after(100, self.load_model_background)

    def load_model_background(self):
        """Load YOLO model in background thread"""
        if not self.model_loading and self.model is None:
            self.model_loading = True
            self.status_text.configure(text="⏳ Loading AI Model...", text_color="orange")
            threading.Thread(target=self._load_model_thread, daemon=True).start()

    def _load_model_thread(self):
        """Background thread to load model"""
        try:
            print("🔄 Loading YOLO model in background...")
            self.model = YOLO(MODEL_PATH)
            print("✅ Model loaded!")
            # Update UI from main thread
            self.after(0, self._model_loaded_callback)
        except Exception as e:
            print(f"❌ Model load failed: {e}")
            self.after(0, lambda: self._model_load_failed(str(e)))

    def _model_loaded_callback(self):
        """Called when model is loaded"""
        self.model_loading = False
        self.status_text.configure(text="🔴 SYSTEM READY (Click Start)", text_color="#2ecc71")

    def _model_load_failed(self, error):
        """Called when model fails to load"""
        self.model_loading = False
        self.status_text.configure(text="❌ Model Load Failed", text_color="red")
        messagebox.showerror("Error", f"Failed to load AI model:\n{error}")

    def get_model(self):
        """Get model, load if not ready"""
        if self.model is None:
            if not self.model_loading:
                self.load_model_background()
            raise RuntimeError("Model still loading, please wait...")
        return self.model

    def setup_ui(self):
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # Sidebar
        self.sidebar = ctk.CTkFrame(self, width=240, corner_radius=0)
        self.sidebar.grid(row=0, column=0, sticky="nsew")

        self.logo_label = ctk.CTkLabel(self.sidebar, text="🛡️ SafeDesk AI",
                                       font=ctk.CTkFont(size=22, weight="bold"))
        self.logo_label.pack(pady=22)

        self.start_btn = ctk.CTkButton(self.sidebar, text="▶️ Start Monitoring",
                                       fg_color="#2ecc71", hover_color="#27ae60",
                                       command=self.start_monitoring, height=40)
        self.start_btn.pack(pady=8, padx=18, fill="x")

        self.stop_btn = ctk.CTkButton(self.sidebar, text="⏹️ Stop Monitoring",
                                      fg_color="#e74c3c", hover_color="#c0392b",
                                      command=self.stop_monitoring, height=40)
        self.stop_btn.pack(pady=8, padx=18, fill="x")

        self.export_btn = ctk.CTkButton(self.sidebar, text="📊 Export Report",
                                        fg_color="#3498db", hover_color="#2980b9",
                                        command=self.export_to_excel_auto_open, height=40)
        self.export_btn.pack(pady=8, padx=18, fill="x")

        self.manager_btn = ctk.CTkButton(self.sidebar, text="🔐 Manager Dashboard (PIN)",
                                         fg_color="#f39c12", hover_color="#e67e22",
                                         command=self.manager_pin_prompt, height=40)
        self.manager_btn.pack(pady=8, padx=18, fill="x")

        # Settings
        settings_title = ctk.CTkLabel(self.sidebar, text="⚙️ Settings",
                                      font=ctk.CTkFont(size=14, weight="bold"))
        settings_title.pack(pady=(18, 6))

        self.mode_label = ctk.CTkLabel(self.sidebar, text="Action Mode",
                                       font=ctk.CTkFont(size=12))
        self.mode_label.pack(pady=(6, 2), padx=18, anchor="w")
        self.mode_menu = ctk.CTkOptionMenu(self.sidebar, values=["Log Only", "Warn", "Lock"],
                                           command=self.set_action_mode)
        self.mode_menu.set(self.action_mode)
        self.mode_menu.pack(pady=(0, 8), padx=18, fill="x")

        self.conf_label = ctk.CTkLabel(self.sidebar, text=f"Confidence: {self.conf_threshold:.2f}",
                                       font=ctk.CTkFont(size=12))
        self.conf_label.pack(pady=(6, 2), padx=18, anchor="w")
        self.conf_slider = ctk.CTkSlider(self.sidebar, from_=0.25, to=0.75,
                                         number_of_steps=50, command=self.on_conf_change)
        self.conf_slider.set(self.conf_threshold)
        self.conf_slider.pack(pady=(0, 10), padx=18, fill="x")

        # Logs
        self.log_label = ctk.CTkLabel(self.sidebar, text="📋 Recent Violations",
                                      font=ctk.CTkFont(size=14, weight="bold"))
        self.log_label.pack(pady=(10, 5))
        self.log_box = ctk.CTkTextbox(self.sidebar, width=200, height=180, font=("Arial", 11))
        self.log_box.pack(pady=8, padx=12, fill="both")
        self.refresh_logs()

        # Status
        self.status_container = ctk.CTkFrame(self, corner_radius=15, fg_color="#1a1a1a")
        self.status_container.grid(row=0, column=1, padx=40, pady=40, sticky="nsew")

        self.shield_icon = ctk.CTkLabel(self.status_container, text="🛡️",
                                        font=ctk.CTkFont(size=120))
        self.shield_icon.pack(expand=True, pady=(50, 10))

        self.status_text = ctk.CTkLabel(self.status_container, text="⏳ Initializing...",
                                        font=ctk.CTkFont(size=28, weight="bold"), text_color="orange")
        self.status_text.pack(expand=True, pady=(10, 20))

        self.info_text = ctk.CTkLabel(self.status_container,
                                      text="Camera feed hidden for privacy.\nAI monitoring mobile phones in real-time.",
                                      font=ctk.CTkFont(size=14), text_color="#888")
        self.info_text.pack(expand=True, pady=(0, 50))

    def start_monitoring(self):
        # 🔥 Check if model is loaded
        if self.model is None:
            if self.model_loading:
                messagebox.showinfo("Please Wait", "AI model is still loading...\nPlease try again in a few seconds.")
                return
            else:
                messagebox.showerror("Error", "AI model failed to load.\nPlease restart the application.")
                return

        if not self.monitoring:
            self.monitoring = True
            self.detect_streak = 0

            self.cap = cv2.VideoCapture(0, cv2.CAP_DSHOW)
            self.cap.set(cv2.CAP_PROP_FRAME_WIDTH, 640)
            self.cap.set(cv2.CAP_PROP_FRAME_HEIGHT, 480)
            self.cap.set(cv2.CAP_PROP_FPS, 30)

            if self.cap.isOpened():
                self.status_text.configure(text="🟢 MONITORING LIVE", text_color="#2ecc71")
                threading.Thread(target=self.update_frame, daemon=True).start()
            else:
                messagebox.showerror("🚫 Camera Error", "Check Privacy Settings!")
                self.monitoring = False

    def set_action_mode(self, value: str):
        self.action_mode = value

    def on_conf_change(self, v):
        self.conf_threshold = float(v)
        self.conf_label.configure(text=f"Confidence: {self.conf_threshold:.2f}")

    def stop_monitoring(self):
        self.monitoring = False
        self.detect_streak = 0
        if self.cap:
            self.cap.release()
            self.cap = None
        # Show ready status instead of offline
        if self.model is not None:
            self.status_text.configure(text="🔴 SYSTEM READY (Click Start)", text_color="#2ecc71")
        else:
            self.status_text.configure(text="🔴 SYSTEM OFFLINE", text_color="gray")

    def update_frame(self):
        while self.monitoring:
            try:
                ret, frame = self.cap.read()
                if not ret or not self.monitoring:
                    time.sleep(0.1)
                    continue

                # Model already loaded, just use it
                results = self.model(frame, conf=self.conf_threshold, verbose=False)
                phone_detected = False

                for r in results:
                    if r.boxes is not None:
                        for box in r.boxes:
                            if int(box.cls[0]) == 67:
                                phone_detected = True
                                break
                        if phone_detected: break

                if phone_detected:
                    self.detect_streak += 1
                else:
                    self.detect_streak = 0

                if self.detect_streak >= self.required_streak:
                    now = time.time()
                    if now - self.last_alert_time >= self.alert_cooldown_seconds:
                        self.last_alert_time = now
                        self.after(0, lambda f=frame.copy(): self.save_violation(f))
                        self.detect_streak = 0

                time.sleep(0.03)
            except Exception as e:
                print(f"Frame error: {e}")
                time.sleep(0.1)

    def save_violation(self, frame):
        now = datetime.now()
        timestamp_str = now.strftime("%Y-%m-%d %H:%M:%S")
        img_name = os.path.join(ALERTS_DIR, f"violation_{now.strftime('%Y%m%d_%H%M%S')}.jpg")

        try:
            employee_id = os.getlogin()
        except:
            employee_id = "Unknown"

        cv2.imwrite(img_name, frame)

        if os.path.exists(img_name):
            cursor.execute(
                'INSERT INTO violations (timestamp, employee_name, object_detected, image_path) VALUES (?, ?, ?, ?)',
                (timestamp_str, employee_id, "Mobile Phone", img_name))
            conn.commit()
            self.after(0, self.refresh_logs)

            if self.action_mode == "Warn":
                self.after(0, lambda: messagebox.showwarning("🚨 VIOLATION", "Mobile phone detected!"))
            elif self.action_mode == "Lock":
                ctypes.windll.user32.LockWorkStation()

    def refresh_logs(self):
        try:
            self.log_box.configure(state="normal")
            self.log_box.delete("1.0", "end")
            cursor.execute("SELECT timestamp, employee_name, object_detected FROM violations ORDER BY id DESC LIMIT 10")
            for ts, emp, obj in cursor.fetchall():
                self.log_box.insert("end", f"🚨 {ts} | {emp} | {obj}\n")
            self.log_box.configure(state="disabled")
        except:
            pass

    def export_to_excel_auto_open(self):
        try:
            df = pd.read_sql_query("SELECT * FROM violations ORDER BY id DESC", conn)
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            filename = f"SafeDesk_Report_{timestamp}.xlsx"

            desktop_path = os.path.join(os.path.expanduser("~"), "Desktop", filename)
            df.to_excel(desktop_path, index=False)

            os.startfile(desktop_path)

            self.status_text.configure(text=f"📊 Report opened!", text_color="#3498db")
            self.after(3000, lambda: self.status_text.configure(
                text="🔴 SYSTEM READY (Click Start)" if self.model else "🔴 SYSTEM OFFLINE",
                text_color="#2ecc71" if self.model else "gray"))

        except Exception as e:
            messagebox.showerror("❌ Error", f"Report export failed:\n{str(e)}")

    def manager_pin_prompt(self):
        self.manager_pin_window = ctk.CTkToplevel(self)
        self.manager_pin_window.title("🔐 Manager Access Required")
        self.manager_pin_window.geometry("350x220")
        self.manager_pin_window.resizable(False, False)
        self.manager_pin_window.transient(self)

        ctk.CTkLabel(self.manager_pin_window, text="👨‍💼 MANAGER AUTHENTICATION",
                     font=ctk.CTkFont(size=20, weight="bold")).pack(pady=20)
        ctk.CTkLabel(self.manager_pin_window, text="Enter 4-digit PIN:",
                     font=ctk.CTkFont(size=14)).pack(pady=(0, 10))

        self.pin_entry = ctk.CTkEntry(self.manager_pin_window, show="*", width=200,
                                      height=40, font=ctk.CTkFont(size=16))
        self.pin_entry.pack(pady=10)
        self.pin_entry.focus()

        ctk.CTkButton(self.manager_pin_window, text="✅ VERIFY PIN", fg_color="#27ae60",
                      height=40, font=ctk.CTkFont(size=14, weight="bold"),
                      command=self.verify_manager_pin).pack(pady=15)

    def verify_manager_pin(self):
        pin = self.pin_entry.get()
        if pin == "2613":
            self.manager_pin_window.destroy()
            self.show_manager_dashboard()
        else:
            self.pin_entry.delete(0, "end")
            messagebox.showerror("❌ ACCESS DENIED", "Wrong PIN! Try again.")

    def show_manager_dashboard(self):
        dashboard = ctk.CTkToplevel(self)
        dashboard.title("📊 SafeDesk AI - MANAGER DASHBOARD")
        dashboard.geometry("1300x750")
        dashboard.minsize(1000, 600)

        stats_frame = ctk.CTkFrame(dashboard, fg_color="#1a1a1a")
        stats_frame.pack(fill="x", padx=20, pady=20)

        total_violations = cursor.execute("SELECT COUNT(*) FROM violations").fetchone()[0]
        today_violations = cursor.execute("""
            SELECT COUNT(*) FROM violations WHERE DATE(timestamp)=DATE('now')
        """).fetchone()[0]
        week_violations = cursor.execute("""
            SELECT COUNT(*) FROM violations WHERE DATE(timestamp) >= DATE('now', '-7 days')
        """).fetchone()[0]

        ctk.CTkLabel(stats_frame, text=f"🔴 TOTAL VIOLATIONS: {total_violations}",
                     font=ctk.CTkFont(size=32, weight="bold"), text_color="#e74c3c").pack(pady=20)
        ctk.CTkLabel(stats_frame, text=f"🟢 TODAY: {today_violations} | 📅 THIS WEEK: {week_violations}",
                     font=ctk.CTkFont(size=20)).pack(pady=(0, 20))

        table_frame = ctk.CTkScrollableFrame(dashboard, height=400)
        table_frame.pack(fill="both", expand=True, padx=20, pady=10)

        headers = ["Date & Time", "👤 Employee", "📱 Device", "📸 Photo", "✅ Status"]
        for col, header in enumerate(headers):
            lbl = ctk.CTkLabel(table_frame, text=header, font=ctk.CTkFont(size=14, weight="bold"),
                               width=200 if col < 3 else 100, fg_color="#2b2b2b", height=35)
            lbl.grid(row=0, column=col, padx=2, pady=10, sticky="ew")

        cursor.execute("SELECT * FROM violations ORDER BY id DESC LIMIT 50")
        violations = cursor.fetchall()

        for row_idx, row in enumerate(violations, 1):
            timestamp = row[1]
            emp_name = row[2]
            device = row[3]
            img_path = str(row[4]) if row[4] else ""

            if device and ("C:\\" in str(device) or "violation_" in str(device)):
                img_path = str(device)
                device = str(emp_name)
                try:
                    emp_name = os.getlogin()
                except:
                    emp_name = "User"

            status = "✅ Photo Saved" if os.path.exists(img_path) else "❌ Missing"

            ctk.CTkLabel(table_frame, text=timestamp, width=200, anchor="w").grid(
                row=row_idx, column=0, padx=2, pady=4, sticky="ew")

            emp_label = ctk.CTkLabel(table_frame, text=emp_name, width=150, text_color="black",
                                     fg_color="#ffebee", corner_radius=8, height=28)
            emp_label.grid(row=row_idx, column=1, padx=2, pady=4)

            ctk.CTkLabel(table_frame, text=device, width=120).grid(
                row=row_idx, column=2, padx=2, pady=4)

            ctk.CTkButton(table_frame, text="👁️ VIEW", width=90, height=28,
                          fg_color="#3498db", hover_color="#2980b9",
                          command=lambda p=img_path: self.open_violation_photo(p)).grid(
                row=row_idx, column=3, padx=2, pady=4)

            status_color = "#2ecc71" if "Saved" in status else "#e74c3c"
            ctk.CTkLabel(table_frame, text=status, width=100, fg_color=status_color).grid(
                row=row_idx, column=4, padx=2, pady=4)

        controls_frame = ctk.CTkFrame(dashboard)
        controls_frame.pack(fill="x", padx=20, pady=20)

        ctk.CTkButton(controls_frame, text="📊 Export FULL Manager Report", fg_color="#27ae60",
                      width=220, height=45, font=ctk.CTkFont(size=16),
                      command=self.export_manager_report_auto_open).pack(side="left", padx=15, pady=15)

        ctk.CTkButton(controls_frame, text="🗑️ Clear ALL Data", fg_color="#e74c3c",
                      width=180, height=45, font=ctk.CTkFont(size=16),
                      command=self.clear_all_violations).pack(side="right", padx=15, pady=15)

    def open_violation_photo(self, img_path):
        if not img_path or not os.path.exists(img_path):
            messagebox.showerror("❌ Error", "Photo file not found!")
            return

        if img_path in self.open_photo_windows:
            self.open_photo_windows[img_path].lift()
            self.open_photo_windows[img_path].focus_force()
            return

        try:
            photo_win = ctk.CTkToplevel(self)
            photo_win.title("📸 VIOLATION EVIDENCE")
            photo_win.geometry("700x550")
            photo_win.resizable(True, True)
            photo_win.transient(self)
            photo_win.grab_set()

            self.open_photo_windows[img_path] = photo_win

            def on_window_close():
                del self.open_photo_windows[img_path]
                photo_win.destroy()

            photo_win.protocol("WM_DELETE_WINDOW", on_window_close)

            img = PILImage.open(img_path)
            max_size = (650, 480)
            img.thumbnail(max_size)
            ctk_img = ctk.CTkImage(light_image=img, dark_image=img, size=img.size)

            img_label = ctk.CTkLabel(photo_win, image=ctk_img, text="")
            img_label.pack(expand=True, pady=(20, 10))

            close_btn = ctk.CTkButton(photo_win, text="❌ Close", command=on_window_close,
                                      fg_color="#e74c3c", hover_color="#c0392b", height=35)
            close_btn.pack(pady=(0, 20), fill="x", padx=20)

        except Exception as e:
            if img_path in self.open_photo_windows:
                del self.open_photo_windows[img_path]

    def export_manager_report_auto_open(self):
        try:
            df = pd.read_sql_query("""
                SELECT 
                    timestamp as 'Date & Time',
                    employee_name as 'Employee Name', 
                    object_detected as 'Device Detected',
                    CASE 
                        WHEN DATE(timestamp)=DATE('now') THEN '🟢 TODAY'
                        WHEN DATE(timestamp)=DATE('now','-1 day') THEN '🟡 YESTERDAY'
                        ELSE '🔵 PREVIOUS'
                    END as 'Period',
                    image_path as 'Photo Evidence'
                FROM violations ORDER BY timestamp DESC
            """, conn)

            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            filename = f"SafeDesk_Manager_Report_{timestamp}.xlsx"

            desktop_path = os.path.join(os.path.expanduser("~"), "Desktop", filename)
            df.to_excel(desktop_path, index=False)

            os.startfile(desktop_path)

        except Exception as e:
            messagebox.showerror("❌ Error", f"Export failed:\n{str(e)}")

    def clear_all_violations(self):
        if messagebox.askyesno("⚠️ CONFIRM DELETE", "Delete ALL violation records?\nThis cannot be undone!"):
            cursor.execute("DELETE FROM violations")
            conn.commit()
            self.refresh_logs()
            messagebox.showinfo("✅ DATABASE CLEARED", "All violation data deleted!")

    def start_tray(self):
        if self.tray_running:
            return
        self.tray_running = True

        try:
            if TRAY_ICON_PATH and os.path.exists(TRAY_ICON_PATH):
                icon_img = PILImage.open(TRAY_ICON_PATH)
            else:
                icon_img = PILImage.new("RGB", (64, 64), (46, 204, 113))
        except Exception:
            icon_img = PILImage.new("RGB", (64, 64), (46, 204, 113))

        def _show(icon, item):
            self.after(0, self.show_window)

        def _hide(icon, item):
            self.after(0, self.hide_window)

        def _exit(icon, item):
            self.after(0, self.exit_app)
            icon.stop()

        menu = pystray.Menu(
            pystray.MenuItem("👁️ Show", _show),
            pystray.MenuItem("🙈 Hide", _hide),
            pystray.MenuItem("❌ Exit", _exit)
        )
        self.tray_icon = pystray.Icon("SafeDeskAI", icon_img, "SafeDesk AI", menu)
        threading.Thread(target=self.tray_icon.run, daemon=True).start()

    def show_window(self):
        self.deiconify()
        self.lift()
        self.focus_force()

    def hide_window(self):
        self.withdraw()

    def on_close(self):
        self.hide_window()

    def exit_app(self):
        for win in self.open_photo_windows.values():
            win.destroy()
        self.open_photo_windows.clear()

        self.stop_monitoring()
        conn.close()
        self.destroy()


if __name__ == "__main__":
    app = SafeDeskApp()
    app.mainloop()