import cv2
import threading
import time
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from PIL import Image, ImageTk
from ultralytics import YOLO
import pyttsx3
import csv
import os
from collections import defaultdict, deque
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from email.message import EmailMessage

# Email config (replace with your values)
EMAIL_ADDRESS = '220701108@rajalakshmi.edu.in'
EMAIL_PASSWORD = 'ypee zjxv qztc xkfa'
TO_EMAIL = 'vijayakumarjitheeswaran@gmail.com'

# Dummy user database (for demo purposes)
USER_DB = {
    'admin': {'password': 'admin123', 'role': 'admin'},
    'user': {'password': 'user123', 'role': 'user'}
}

class QualityControlApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Automated Quality Control System")
        self.model = YOLO("runs/detect/package_defect_custom2/weights/best.pt")
        self.cap = None
        self.running = False
        self.engine = pyttsx3.init()

        self.total = 0
        self.defects = 0
        self.defect_types = defaultdict(int)
        self.last_defects = deque(maxlen=5)
        self.frame_skip = 0

        self.setup_login()

    def setup_login(self):
        # Login screen setup
        self.login_frame = ttk.Frame(self.root)
        self.login_frame.pack()

        self.username_label = ttk.Label(self.login_frame, text="Username:")
        self.username_label.grid(row=0, column=0, padx=10, pady=10)

        self.username_entry = ttk.Entry(self.login_frame)
        self.username_entry.grid(row=0, column=1, padx=10, pady=10)

        self.password_label = ttk.Label(self.login_frame, text="Password:")
        self.password_label.grid(row=1, column=0, padx=10, pady=10)

        self.password_entry = ttk.Entry(self.login_frame, show="*")
        self.password_entry.grid(row=1, column=1, padx=10, pady=10)

        self.login_button = ttk.Button(self.login_frame, text="Login", command=self.login)
        self.login_button.grid(row=2, column=0, columnspan=2, pady=10)

    def login(self):
        username = self.username_entry.get()
        password = self.password_entry.get()

        if username in USER_DB and USER_DB[username]['password'] == password:
            self.role = USER_DB[username]['role']
            self.username = username
            self.login_frame.destroy()  # Remove login screen
            self.setup_dashboard()  # Go to dashboard
        else:
            messagebox.showerror("Login Error", "Invalid username or password")

    def setup_dashboard(self):
        # Main dashboard UI setup
        self.dashboard_frame = ttk.Frame(self.root)
        self.dashboard_frame.pack()

        # Stats Section
        stats_frame = ttk.Frame(self.dashboard_frame)
        stats_frame.pack()

        self.total_label = ttk.Label(stats_frame, text="Total: 0")
        self.total_label.grid(row=0, column=0, padx=10)

        self.defect_label = ttk.Label(stats_frame, text="Defects: 0")
        self.defect_label.grid(row=0, column=1, padx=10)

        self.quality_label = ttk.Label(stats_frame, text="Quality: 100%")
        self.quality_label.grid(row=0, column=2, padx=10)

        self.start_btn = ttk.Button(self.dashboard_frame, text="Start Camera", command=self.toggle_camera)
        self.start_btn.pack(pady=10)

        # Additional dashboard features
        if self.role == 'admin':
            self.admin_dashboard()

    def admin_dashboard(self):
        # Admin features: Defect History and Analytics
        self.defect_history_btn = ttk.Button(self.dashboard_frame, text="View Defect History", command=self.view_defect_history)
        self.defect_history_btn.pack(pady=10)

        self.analysis_btn = ttk.Button(self.dashboard_frame, text="View Analytics", command=self.view_analytics)
        self.analysis_btn.pack(pady=10)

    def view_defect_history(self):
        # Display defect history in a new window
        history_window = tk.Toplevel(self.root)
        history_window.title("Defect History")
        tree = ttk.Treeview(history_window, columns=("Timestamp", "Defect", "Confidence"), show="headings")
        tree.heading("Timestamp", text="Timestamp")
        tree.heading("Defect", text="Defect")
        tree.heading("Confidence", text="Confidence")

        with open("predictions_log.csv", 'r') as f:
            reader = csv.reader(f)
            for row in reader:
                tree.insert("", "end", values=row)

        tree.pack()

    def view_analytics(self):
        # Display defect trends (bar chart) in a new window
        analytics_window = tk.Toplevel(self.root)
        analytics_window.title("Defect Analytics")
        fig, ax = plt.subplots(figsize=(5, 3))
        ax.bar(self.defect_types.keys(), self.defect_types.values(), color='orange')
        ax.set_title("Defects by Type")
        canvas = FigureCanvasTkAgg(fig, master=analytics_window)
        canvas.get_tk_widget().pack()
        canvas.draw()

    def toggle_camera(self):
        if not self.running:
            self.cap = cv2.VideoCapture(0, cv2.CAP_DSHOW)  # faster on Windows
            self.running = True
            self.start_btn.config(text="Stop Camera")
            self.update_frame()
        else:
            self.running = False
            self.cap.release()
            self.start_btn.config(text="Start Camera")

    def update_frame(self):
        if not self.running:
            return

        ret, frame = self.cap.read()
        if not ret:
            self.root.after(10, self.update_frame)
            return

        # Skip every other frame to speed up
        self.frame_skip = (self.frame_skip + 1) % 2
        if self.frame_skip == 0:
            self.process_and_display_frame(frame)

        # Continue the loop
        self.root.after(10, self.update_frame)

    def process_and_display_frame(self, frame):
        results = self.model.predict(frame, verbose=False)[0]

        for box in results.boxes:
            x1, y1, x2, y2 = map(int, box.xyxy[0])
            cls = int(box.cls[0])
            conf = float(box.conf[0])
            label = results.names[cls]
            cv2.rectangle(frame, (x1, y1), (x2, y2), (0, 0, 255), 2)
            cv2.putText(frame, f'{label} {conf:.2f}', (x1, y1 - 10),
                        cv2.FONT_HERSHEY_SIMPLEX, 0.7, (0, 0, 255), 2)

            # Record defect once per object
            self.record_defect(label, conf)

        # Show annotated frame in GUI
        img = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        img = ImageTk.PhotoImage(Image.fromarray(img))
        self.video_label.imgtk = img
        self.video_label.configure(image=img)

    def record_defect(self, label, conf):
        self.total += 1
        self.defects += 1
        self.defect_types[label] += 1
        self.last_defects.append(f"{label} - {conf:.2f}")
        self.update_stats()

        threading.Thread(target=self.voice_alert).start()
        threading.Thread(target=self.save_csv, args=(label, conf)).start()

    def update_stats(self):
        self.total_label.config(text=f"Total: {self.total}")
        self.defect_label.config(text=f"Defects: {self.defects}")
        quality = 100 * (1 - self.defects / max(1, self.total))
        self.quality_label.config(text=f"Quality: {quality:.2f}%")

    def voice_alert(self):
        self.engine.say("Defect detected!")
        self.engine.runAndWait()

    def save_csv(self, defect, conf):
        with open("predictions_log.csv", 'a', newline='') as f:
            writer = csv.writer(f)
            writer.writerow([time.strftime("%Y-%m-%d %H:%M:%S"), defect, f"{conf:.2f}"])

    def on_close(self):
        self.running = False
        if self.cap:
            self.cap.release()
        self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = QualityControlApp(root)
    root.protocol("WM_DELETE_WINDOW", app.on_close)
    root.mainloop()
