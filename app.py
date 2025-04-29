import cv2
import threading
import time
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
from PIL import Image, ImageTk
from ultralytics import YOLO
import pyttsx3
import csv
import os
from collections import defaultdict, deque
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import smtplib
from email.message import EmailMessage
from email.utils import formatdate
from datetime import datetime

# Email config (replace with your values)
EMAIL_ADDRESS = '220701108@rajalakshmi.edu.in'
EMAIL_PASSWORD = 'ypee zjxv qztc xkfa'

# Company details
COMPANY_NAME = "Stark technologies"
COMPANY_ADDRESS = "sriperumbudur, Chennai, India"
COMPANY_PHONE = "+91 9876543210"
COMPANY_LOGO = None  # Can be replaced with actual logo path

# Dummy user database (for demo purposes)
USER_DB = {
    'admin': {'password': 'admin123', 'role': 'admin', 'name': 'Admin User', 'email': 'admin@company.com'},
    'user': {'password': 'user123', 'role': 'user', 'name': 'Operator 1', 'email': 'operator@company.com'}
}

class QualityControlApp:
    def __init__(self, root):
        self.root = root
        self.root.title(f"{COMPANY_NAME} - Automated Quality Control System")
        self.root.geometry("1200x800")
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        
        # Initialize components
        self.model = None
        self.cap = None
        self.running = False
        self.engine = None
        self.current_frame = None
        self.video_label = None
        self.video_panel_created = False

        # Statistics
        self.total = 0
        self.defects = 0
        self.defect_types = defaultdict(int)
        self.last_defects = deque(maxlen=5)
        self.frame_skip = 0
        self.shift_start_time = datetime.now()
        self.current_employee = None

        # Initialize TTS engine
        try:
            self.engine = pyttsx3.init()
        except Exception as e:
            print(f"Warning: Could not initialize TTS engine - {str(e)}")

        # Load YOLO model
        self.load_model()

        # Setup UI
        self.setup_login()

    def load_model(self):
        """Load the YOLO model with error handling"""
        try:
            self.model = YOLO("runs/detect/package_defect_custom2/weights/best.pt")
        except Exception as e:
            messagebox.showerror("Model Error", f"Failed to load YOLO model: {str(e)}")
            self.model = None

    def setup_login(self):
        """Setup the login screen"""
        # Clear any existing frames
        for widget in self.root.winfo_children():
            widget.destroy()

        # Login screen setup
        self.login_frame = ttk.Frame(self.root, padding="30")
        self.login_frame.pack(expand=True)

        # Style configuration
        style = ttk.Style()
        style.configure("TLabel", font=('Helvetica', 12))
        style.configure("TButton", font=('Helvetica', 12), padding=6)
        style.configure("TEntry", font=('Helvetica', 12), padding=6)

        # Company header
        header_frame = ttk.Frame(self.login_frame)
        header_frame.grid(row=0, column=0, columnspan=2, pady=(0, 20))
        
        company_label = ttk.Label(header_frame, text=COMPANY_NAME, font=('Helvetica', 18, 'bold'))
        company_label.pack()
        
        system_label = ttk.Label(header_frame, text="Quality Control System", font=('Helvetica', 14))
        system_label.pack()

        # Username
        self.username_label = ttk.Label(self.login_frame, text="Employee ID:")
        self.username_label.grid(row=1, column=0, padx=10, pady=10, sticky="e")

        self.username_entry = ttk.Entry(self.login_frame, width=30)
        self.username_entry.grid(row=1, column=1, padx=10, pady=10)

        # Password
        self.password_label = ttk.Label(self.login_frame, text="Password:")
        self.password_label.grid(row=2, column=0, padx=10, pady=10, sticky="e")

        self.password_entry = ttk.Entry(self.login_frame, show="*", width=30)
        self.password_entry.grid(row=2, column=1, padx=10, pady=10)

        # Login button
        self.login_button = ttk.Button(self.login_frame, text="Login", command=self.login)
        self.login_button.grid(row=3, column=0, columnspan=2, pady=20)

        # Center the login form
        self.login_frame.place(relx=0.5, rely=0.5, anchor="center")

    def login(self):
        """Handle login authentication"""
        username = self.username_entry.get()
        password = self.password_entry.get()

        if username in USER_DB and USER_DB[username]['password'] == password:
            self.role = USER_DB[username]['role']
            self.username = username
            self.current_employee = USER_DB[username]['name']
            self.login_frame.destroy()
            self.setup_dashboard()
        else:
            messagebox.showerror("Login Error", "Invalid Employee ID or password")

    def setup_dashboard(self):
        """Setup the main dashboard after successful login"""
        # Clear any existing frames
        for widget in self.root.winfo_children():
            widget.destroy()

        # Main frame
        self.main_frame = ttk.Frame(self.root)
        self.main_frame.pack(fill=tk.BOTH, expand=True)

        # Left panel for controls (30% width)
        self.control_frame = ttk.LabelFrame(self.main_frame, text="Controls", padding="10")
        self.control_frame.pack(side=tk.LEFT, fill=tk.Y, padx=10, pady=10)

        # Right panel for camera feed (70% width)
        self.camera_frame = ttk.LabelFrame(self.main_frame, text="Package Inspection Feed", padding="10")
        self.camera_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Create video label
        self.video_label = ttk.Label(self.camera_frame)
        self.video_label.pack(fill=tk.BOTH, expand=True)
        self.video_panel_created = True

        # Employee info section
        self.setup_employee_section()

        # Stats Section
        self.setup_stats_section()

        # Control buttons
        self.setup_control_buttons()

        # Additional dashboard features for admin
        if self.role == 'admin':
            self.admin_dashboard()

        # Logout button
        self.setup_logout_button()

    def setup_employee_section(self):
        """Setup the employee information section"""
        self.employee_frame = ttk.LabelFrame(self.control_frame, text="Employee Information", padding="10")
        self.employee_frame.pack(fill=tk.X, pady=10)

        name_label = ttk.Label(self.employee_frame, text=f"Name: {self.current_employee}")
        name_label.pack(anchor="w", pady=5)

        role_label = ttk.Label(self.employee_frame, text=f"Role: {self.role.capitalize()}")
        role_label.pack(anchor="w", pady=5)

        shift_label = ttk.Label(self.employee_frame, text=f"Shift Start: {self.shift_start_time.strftime('%Y-%m-%d %H:%M:%S')}")
        shift_label.pack(anchor="w", pady=5)

    def setup_stats_section(self):
        """Setup the statistics display section"""
        self.stats_frame = ttk.LabelFrame(self.control_frame, text="Package Statistics", padding="10")
        self.stats_frame.pack(fill=tk.X, pady=10)

        self.total_label = ttk.Label(self.stats_frame, text="Total Packages: 0")
        self.total_label.pack(anchor="w", pady=5)

        self.defect_label = ttk.Label(self.stats_frame, text="Defective Packages: 0")
        self.defect_label.pack(anchor="w", pady=5)

        self.quality_label = ttk.Label(self.stats_frame, text="Quality Rate: 100%")
        self.quality_label.pack(anchor="w", pady=5)

    def update_stats(self):
        """Update the statistics display"""
        quality_rate = ((self.total - self.defects) / self.total) * 100 if self.total > 0 else 100
        
        self.total_label.config(text=f"Total Packages: {self.total}")
        self.defect_label.config(text=f"Defective Packages: {self.defects}")
        self.quality_label.config(text=f"Quality Rate: {quality_rate:.1f}%")

    def setup_control_buttons(self):
        """Setup the control buttons"""
        self.btn_frame = ttk.Frame(self.control_frame)
        self.btn_frame.pack(fill=tk.X, pady=10)

        self.start_btn = ttk.Button(self.btn_frame, text="Start Inspection", command=self.toggle_camera)
        self.start_btn.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)

        self.capture_btn = ttk.Button(self.btn_frame, text="Capture Image", command=self.capture_image, state=tk.DISABLED)
        self.capture_btn.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)

    def admin_dashboard(self):
        """Setup admin-specific features"""
        admin_frame = ttk.LabelFrame(self.control_frame, text="Quality Control Tools", padding="10")
        admin_frame.pack(fill=tk.X, pady=10)

        self.defect_history_btn = ttk.Button(admin_frame, text="Defect History", command=self.view_defect_history)
        self.defect_history_btn.pack(fill=tk.X, pady=5)

        self.analysis_btn = ttk.Button(admin_frame, text="Quality Dashboard", command=self.view_analytics)
        self.analysis_btn.pack(fill=tk.X, pady=5)

        self.export_btn = ttk.Button(admin_frame, text="Export Report", command=self.export_data)
        self.export_btn.pack(fill=tk.X, pady=5)

        self.email_btn = ttk.Button(admin_frame, text="Send Quality Report", command=self.send_email_report)
        self.email_btn.pack(fill=tk.X, pady=5)

    def setup_logout_button(self):
        """Setup the logout button"""
        logout_btn = ttk.Button(self.control_frame, text="Logout", command=self.setup_login)
        logout_btn.pack(side=tk.BOTTOM, pady=20, fill=tk.X)

    def toggle_camera(self):
        """Toggle camera on/off"""
        if not self.running:
            self.start_camera()
        else:
            self.stop_camera()

    def start_camera(self):
        """Start the camera capture"""
        try:
            self.cap = cv2.VideoCapture(0)
            if not self.cap.isOpened():
                raise Exception("Could not open video device")
            
            # Set camera resolution
            self.cap.set(cv2.CAP_PROP_FRAME_WIDTH, 640)
            self.cap.set(cv2.CAP_PROP_FRAME_HEIGHT, 480)
            
            self.running = True
            self.start_btn.config(text="Stop Inspection")
            self.capture_btn.config(state=tk.NORMAL)
            self.update_frame()
        except Exception as e:
            messagebox.showerror("Camera Error", f"Failed to start camera: {str(e)}")
            if self.cap:
                self.cap.release()
            self.running = False
            self.start_btn.config(text="Start Inspection")

    def stop_camera(self):
        """Stop the camera capture"""
        self.running = False
        if self.cap:
            self.cap.release()
        self.start_btn.config(text="Start Inspection")
        self.capture_btn.config(state=tk.DISABLED)

    def update_frame(self):
        """Update the camera frame continuously"""
        if not self.running or not self.video_panel_created:
            return

        ret, frame = self.cap.read()
        if not ret:
            self.root.after(10, self.update_frame)
            return

        # Store the current frame for potential capture
        self.current_frame = frame.copy()

        # Process the frame
        self.process_and_display_frame(frame)

        # Continue the loop
        self.root.after(10, self.update_frame)

    def process_and_display_frame(self, frame):
        """Process frame with YOLO, detect packages and their sizes, and display results"""
        if self.model is None or not self.video_panel_created:
            return

        try:
            # Convert frame to RGB (YOLO expects RGB)
            frame_rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)

            # Run inference - only detect packages (class 0 if that's your package class)
            results = self.model.predict(frame_rgb, verbose=False)[0]

            # Draw bounding boxes and process each detected package
            for box in results.boxes:
                x1, y1, x2, y2 = map(int, box.xyxy[0])
                cls = int(box.cls[0])
                conf = float(box.conf[0])
                label = results.names[cls]

                # Calculate box dimensions
                width = x2 - x1
                height = y2 - y1

                # Draw rectangle and label with confidence and size
                color = (0, 255, 0) if conf > 0.8 else (0, 0, 255)  # Green for good, Red for potentially defective
                cv2.rectangle(frame_rgb, (x1, y1), (x2, y2), color, 2)
                cv2.putText(frame_rgb, f'{label} {conf:.2f} ({width}x{height})', (x1, y1 - 10),
                            cv2.FONT_HERSHEY_SIMPLEX, 0.6, color, 2)

                # Record package information including size
                self.record_package(conf, width, height, label)

            # Display the annotated frame
            img = Image.fromarray(frame_rgb)
            imgtk = ImageTk.PhotoImage(image=img)

            # Update the label
            self.video_label.imgtk = imgtk
            self.video_label.configure(image=imgtk)

        except Exception as e:
            print(f"Error processing frame: {str(e)}")

    def record_package(self, conf, width, height, label):
        """Record package statistics including size"""
        self.total += 1
        if conf < 0.8:  # Consider as potentially defective if confidence is low
            self.defects += 1
            self.defect_types[label] += 1
            self.last_defects.append(f"{label} - {conf:.2f} (Size: {width}x{height})")
            self.update_stats()

            # Run voice alert in background
            if self.engine:
                threading.Thread(target=self.voice_alert, args=(label,), daemon=True).start()

        # Save to CSV with size information
        self.save_to_csv(conf, width, height, label)

    def voice_alert(self, defect_type):
        """Voice alert for defective packages"""
        try:
            self.engine.say(f"Warning! Potential {defect_type} detected")
            self.engine.runAndWait()
        except Exception as e:
            print(f"Error in voice alert: {str(e)}")

    def save_to_csv(self, conf, width, height, label):
        """Save package information including size to CSV file"""
        try:
            file_exists = os.path.isfile("packages_log.csv")

            with open("packages_log.csv", 'a', newline='') as f:
                writer = csv.writer(f)
                if not file_exists:
                    writer.writerow(["Timestamp", "Employee", "Label", "Status", "Confidence", "Width", "Height"])

                status = "Potentially Defective" if conf < 0.8 else "Good"
                writer.writerow([
                    time.strftime("%Y-%m-%d %H:%M:%S"),
                    self.current_employee,
                    label,
                    status,
                    f"{conf:.2f}",
                    width,
                    height
                ])
        except Exception as e:
            print(f"Error saving to CSV: {str(e)}")

    def capture_image(self):
        """Capture and save the current frame"""
        if self.current_frame is not None:
            try:
                # Ask user for save location
                file_path = filedialog.asksaveasfilename(
                    defaultextension=".jpg",
                    filetypes=[("JPEG files", "*.jpg"), ("PNG files", "*.png"), ("All files", "*.*")],
                    title="Save package image"
                )
                
                if file_path:
                    # Convert to RGB before saving
                    img_rgb = cv2.cvtColor(self.current_frame, cv2.COLOR_BGR2RGB)
                    cv2.imwrite(file_path, img_rgb)
                    messagebox.showinfo("Success", f"Package image saved to {file_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save image: {str(e)}")

    def view_defect_history(self):
        """Display package history in a new window"""
        try:
            if not os.path.exists("packages_log.csv"):
                messagebox.showinfo("Info", "No package history found")
                return

            history_window = tk.Toplevel(self.root)
            history_window.title("Package Inspection History")
            history_window.geometry("1200x600")

            # Create a frame for the treeview and scrollbars
            tree_frame = ttk.Frame(history_window)
            tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

            # Add scrollbars
            v_scroll = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL)
            h_scroll = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL)

            # Create the treeview
            tree = ttk.Treeview(tree_frame,
                                            columns=("Timestamp", "Employee", "Label", "Status", "Confidence", "Width", "Height"),
                                            show="headings",
                                            yscrollcommand=v_scroll.set,
                                            xscrollcommand=h_scroll.set)

            # Configure scrollbars
            v_scroll.config(command=tree.yview)
            h_scroll.config(command=tree.xview)

            # Define headings
            tree.heading("Timestamp", text="Timestamp")
            tree.heading("Employee", text="Employee")
            tree.heading("Label", text="Label")
            tree.heading("Status", text="Status")
            tree.heading("Confidence", text="Confidence Level")
            tree.heading("Width", text="Width")
            tree.heading("Height", text="Height")

            # Set column widths
            tree.column("Timestamp", width=200, anchor="center")
            tree.column("Employee", width=150, anchor="center")
            tree.column("Label", width=150, anchor="center")
            tree.column("Status", width=150, anchor="center")
            tree.column("Confidence", width=120, anchor="center")
            tree.column("Width", width=100, anchor="center")
            tree.column("Height", width=100, anchor="center")

            # Add data from CSV
            with open("packages_log.csv", 'r') as f:
                reader = csv.reader(f)
                next(reader, None)  # Skip header if exists
                for row in reader:
                    if len(row) >= 7:  # Ensure we have all columns
                        tree.insert("", "end", values=row[:7])

            # Pack the treeview and scrollbars
            tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            v_scroll.pack(side=tk.RIGHT, fill=tk.Y)
            h_scroll.pack(side=tk.BOTTOM, fill=tk.X)

        except Exception as e:
            messagebox.showerror("Error", f"Failed to load package history: {str(e)}")

    def view_analytics(self):
        """Display package analytics in a new window"""
        try:
            if not os.path.exists("packages_log.csv"):
                messagebox.showinfo("Info", "No package data available for analytics")
                return

            # Read data from CSV
            timestamps = []
            employees = []
            labels = []
            statuses = []
            confidences = []
            widths = []  # To store width
            heights = [] # To store height
            
            with open("packages_log.csv", 'r') as f:
                reader = csv.reader(f)
                next(reader, None)  # Skip header
                for row in reader:
                    if len(row) >= 7:
                        timestamps.append(row[0])
                        employees.append(row[1])
                        labels.append(row[2])
                        statuses.append(row[3])
                        confidences.append(float(row[4]))
                        widths.append(int(row[5]))  # Extract width
                        heights.append(int(row[6])) # Extract height

            if not timestamps:
                messagebox.showinfo("Info", "No data available for analytics")
                return

            analytics_window = tk.Toplevel(self.root)
            analytics_window.title("Package Quality Analytics")
            analytics_window.geometry("1000x800")

            # Create notebook for multiple tabs
            notebook = ttk.Notebook(analytics_window)
            notebook.pack(fill=tk.BOTH, expand=True)

            # Tab 1: Quality Overview
            overview_frame = ttk.Frame(notebook)
            notebook.add(overview_frame, text="Quality Overview")

            # Calculate statistics
            total = len(statuses)
            defective = statuses.count("Potentially Defective")
            good = total - defective
            quality_rate = (good / total) * 100 if total > 0 else 100

            # Create figure with subplots
            fig1, (ax1, ax2) = plt.subplots(2, 1, figsize=(10, 10))
            
            # Bar chart for quality overview
            ax1.bar(["Good", "Defective"], [good, defective], color=['green', 'red'])
            ax1.set_title("Package Quality Overview")
            ax1.set_ylabel("Count")
            
            # Text information
            ax2.axis('off')
            ax2.text(0.1, 0.8, f"Total Packages: {total}", fontsize=12)
            ax2.text(0.1, 0.6, f"Good Packages: {good} ({quality_rate:.2f}%)", fontsize=12)
            ax2.text(0.1, 0.4, f"Potentially Defective Packages: {defective} ({100-quality_rate:.2f}%)", fontsize=12)
            ax2.text(0.1, 0.2, f"Shift Start Time: {self.shift_start_time.strftime('%Y-%m-%d %H:%M:%S')}", fontsize=12)

            # Add canvas to tab
            canvas1 = FigureCanvasTkAgg(fig1, master=overview_frame)
            canvas1.draw()
            canvas1.get_tk_widget().pack(fill=tk.BOTH, expand=True)

            # Tab 2: Defect Type Analysis
            defect_frame = ttk.Frame(notebook)
            notebook.add(defect_frame, text="Defect Types")

            # Calculate defect type distribution
            defect_counts = defaultdict(int)
            for label, status in zip(labels, statuses):
                if status == "Potentially Defective":
                    defect_counts[label] += 1

            # Create figure for defect types
            fig2, ax = plt.subplots(figsize=(10, 6))
            
            if defect_counts:
                # Pie chart for defect types
                labels_pie = list(defect_counts.keys())
                sizes = list(defect_counts.values())
                
                ax.pie(sizes, labels=labels_pie, autopct='%1.1f%%', startangle=90)
                ax.axis('equal')  # Equal aspect ratio ensures pie is drawn as a circle
                ax.set_title("Defect Type Distribution")
            else:
                ax.text(0.5, 0.5, "No defects found in the data", 
                       ha='center', va='center', fontsize=12)
                ax.axis('off')

            # Add canvas to tab
            canvas2 = FigureCanvasTkAgg(fig2, master=defect_frame)
            canvas2.draw()
            canvas2.get_tk_widget().pack(fill=tk.BOTH, expand=True)

            # Tab 3: Size Analysis
            size_analysis_frame = ttk.Frame(notebook)
            notebook.add(size_analysis_frame, text="Size Analysis")
            
            # Create figure for size analysis
            fig3, (ax1, ax2) = plt.subplots(1, 2, figsize=(12, 6))  # 1 row, 2 columns
            
            # Scatter plot of width vs height with status coloring
            colors = ['green' if s == "Good" else 'red' for s in statuses]
            ax1.scatter(widths, heights, c=colors, alpha=0.5)
            ax1.set_title("Package Size Distribution")
            ax1.set_xlabel("Width (pixels)")
            ax1.set_ylabel("Height (pixels)")

            # Box plot of width and height
            ax2.boxplot([widths, heights], labels=['Width', 'Height'])
            ax2.set_title("Box Plot of Package Dimensions")
            ax2.set_ylabel("Size (pixels)")

            # Add canvas
            canvas3 = FigureCanvasTkAgg(fig3, master=size_analysis_frame)
            canvas3.draw()
            canvas3.get_tk_widget().pack(fill=tk.BOTH, expand=True)

            # Tab 4: Time Analysis
            time_frame = ttk.Frame(notebook)
            notebook.add(time_frame, text="Time Analysis")

            # Convert timestamps to datetime objects
            times = [datetime.strptime(ts, "%Y-%m-%d %H:%M:%S") for ts in timestamps]
            hours = [t.hour + t.minute/60 for t in times]
            
            # Create figure for time analysis
            fig4, ax = plt.subplots(figsize=(10, 6))
            
            # Scatter plot of defects over time
            colors = ['green' if s == "Good" else 'red' for s in statuses]
            ax.scatter(hours, confidences, c=colors, alpha=0.5)
            ax.set_title("Package Quality Over Time")
            ax.set_xlabel("Hour of Day")
            ax.set_ylabel("Confidence Level")
            ax.set_ylim(0, 1.1)
            
            # Add mean line
            mean_conf = sum(confidences)/len(confidences)
            ax.axhline(y=mean_conf, color='blue', linestyle='--', label=f'Mean Confidence: {mean_conf:.2f}')
            ax.legend()

            # Add canvas to tab
            canvas4 = FigureCanvasTkAgg(fig4, master=time_frame)
            canvas4.draw()
            canvas4.get_tk_widget().pack(fill=tk.BOTH, expand=True)

            # Add tight layout to all figures
            plt.tight_layout()

        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate analytics: {str(e)}")

    def export_data(self):
        """Export package data to a file with company header"""
        try:
            if not os.path.exists("packages_log.csv"):
                messagebox.showinfo("Info", "No data available to export")
                return

            file_path = filedialog.asksaveasfilename(
                defaultextension=".csv",
                filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
                title="Export package quality data"
            )
            
            if file_path:
                # Create a new file with company header
                with open(file_path, 'w', newline='') as outfile:
                    writer = csv.writer(outfile)
                    
                    # Write company header
                    writer.writerow([COMPANY_NAME])
                    writer.writerow([COMPANY_ADDRESS])
                    writer.writerow([f"Phone: {COMPANY_PHONE}"])
                    writer.writerow(["Package Quality Report"])
                    writer.writerow([f"Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"])
                    writer.writerow([f"Generated by: {self.current_employee}"])
                    writer.writerow([])  # Empty line
                    
                    # Write the actual data
                    with open("packages_log.csv", 'r') as infile:
                        reader = csv.reader(infile)
                        for row in reader:
                            writer.writerow(row)
                    
                messagebox.showinfo("Success", f"Report exported to {file_path}")
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export report: {str(e)}")

    def send_email_report(self):
        """Send quality report via email with professional formatting"""
        try:
            if not os.path.exists("packages_log.csv"):
                messagebox.showinfo("Info", "No data available for report")
                return

            # Get recipient email from user
            to_email = simpledialog.askstring("Recipient Email", 
                                            "Enter the recipient email address:",
                                            parent=self.root)
            if not to_email:
                return  # User cancelled

            # Validate email format
            if "@" not in to_email or "." not in to_email.split("@")[-1]:
                messagebox.showerror("Error", "Please enter a valid email address")
                return

            # Calculate statistics
            with open("packages_log.csv", 'r') as f:
                reader = csv.reader(f)
                next(reader, None)  # Skip header
                data = list(reader)
                
            total = len(data)
            defective = sum(1 for row in data if len(row) >= 4 and row[3] == "Potentially Defective")
            good = total - defective
            quality_rate = (good / total) * 100 if total > 0 else 100

            # Create email message
            msg = EmailMessage()
            msg['Subject'] = f'{COMPANY_NAME} - Package Quality Report'
            msg['From'] = EMAIL_ADDRESS
            msg['To'] = to_email
            msg['Date'] = formatdate(localtime=True)
            
            # Email body with HTML formatting
            html = f"""
            <html>
                <body>
                    <div style="font-family: Arial, sans-serif; max-width: 800px; margin: 0 auto;">
                        <div style="background-color: #f8f9fa; padding: 20px; border-bottom: 3px solid #007bff;">
                            <h1 style="color: #007bff; margin: 0;">{COMPANY_NAME}</h1>
                            <p style="margin: 5px 0 0; color: #6c757d;">{COMPANY_ADDRESS}</p>
                        </div>
                        
                        <div style="padding: 20px;">
                            <h2 style="color: #343a40;">Package Quality Report</h2>
                            <p>Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
                            <p>Generated by: {self.current_employee}</p>
                            
                            <h3 style="margin-top: 20px;">Quality Summary</h3>
                            <table style="width: 100%; border-collapse: collapse; margin-bottom: 20px;">
                                <tr style="background-color: #f8f9fa;">
                                    <th style="padding: 10px; text-align: left; border: 1px solid #dee2e6;">Metric</th>
                                    <th style="padding: 10px; text-align: left; border: 1px solid #dee2e6;">Value</th>
                                </tr>
                                <tr>
                                    <td style="padding: 10px; border: 1px solid #dee2e6;">Total Packages</td>
                                    <td style="padding: 10px; border: 1pxsolid #dee2e6;">{total}</td>
                                </tr>
                                <tr style="background-color: #f8f9fa;">
                                    <td style="padding: 10px; border: 1px solid #dee2e6;">Good Packages</td>
                                    <td style="padding: 10px; border: 1px solid #dee2e6;">{good} ({quality_rate:.2f}%)</td>
                                </tr>
                                <tr>
                                    <td style="padding: 10px; border: 1px solid #dee2e6;">Potentially Defective Packages</td>
                                    <td style="padding: 10px; border: 1px solid #dee2e6;">{defective} ({100-quality_rate:.2f}%)</td>
                                </tr>
                            </table>
                            
                            <p style="margin-top: 20px;">This report was automatically generated by the Quality Control System.</p>
                        </div>
                    </div>
                </body>
            </html>
            """
            
            msg.add_alternative(html, subtype='html')

            # Attach CSV file
            with open("packages_log.csv", 'rb') as f:
                file_data = f.read()
                msg.add_attachment(file_data,
                                            maintype="text",
                                            subtype="csv",
                                            filename=f"package_quality_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv")

            # Send email
            with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
                smtp.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
                smtp.send_message(msg)
            
            messagebox.showinfo("Success", f"Quality report email sent to {to_email}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to send email: {str(e)}")

    def on_close(self):
        """Clean up resources before closing"""
        self.running = False
        if self.cap:
            self.cap.release()
        if self.engine:
            self.engine.stop()
        self.root.destroy()

if __name__ == "__main__":
    try:
        root = tk.Tk()
        app = QualityControlApp(root)
        root.mainloop()
    except Exception as e:
        messagebox.showerror("Fatal Error", f"Application crashed: {str(e)}")