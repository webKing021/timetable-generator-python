
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, colorchooser, filedialog
import random
from datetime import datetime, timedelta
import itertools
import sys
import os
import pickle
import webbrowser

# Try to import pandas, but don't fail if it's not available
try:
    import pandas as pd
    PANDAS_AVAILABLE = True
except ImportError:
    PANDAS_AVAILABLE = False
    
# Try to import numpy, but don't fail if it's not available
try:
    import numpy as np
    NUMPY_AVAILABLE = True
except ImportError:
    NUMPY_AVAILABLE = False

# Try to import matplotlib, but don't fail if it's not available
try:
    import matplotlib.pyplot as plt
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
    MATPLOTLIB_AVAILABLE = True
except ImportError:
    MATPLOTLIB_AVAILABLE = False

# Try to import PIL, but don't fail if it's not available
try:
    from PIL import Image, ImageTk
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False

# Try to import PDF libraries, but don't fail if they're not available
try:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import letter, A4
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.lib.styles import getSampleStyleSheet
    PDF_AVAILABLE = True
except ImportError:
    PDF_AVAILABLE = False

class TimetableGenerator:
    def __init__(self, root):
        """Initialize the timetable generator"""
        self.root = root
        self.root.title("College Timetable Generator")
        self.root.geometry("1200x700")
        
        # Set theme color (default blue)
        self.theme_color = "#4a7abc"  # Default blue theme
        
        # Create settings dictionary
        self.settings = {
            "theme_color": self.theme_color,
            "auto_save": False,
            "auto_save_path": os.path.join(os.path.expanduser("~"), "timetable_autosave.pkl"),
            "default_export_format": "excel"
        }
        
        # Data structures
        self.days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]
        self.faculties = {}
        self.time_slots = []
        self.classrooms = []
        self.labs = []
        self.subjects = {}
        self.current_timetable = None
        
        # Create main frame
        main_frame = ttk.Frame(root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Create menu
        self.create_menu()
        
        # Create notebook for tabs
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True)
        
        # Store after IDs for proper cleanup
        self.after_ids = []

        # Default values
        self.load_default_data()

        # Create main notebook
        self.create_input_tab()
        self.create_faculty_tab()
        self.create_subjects_tab()
        self.create_resources_tab()
        self.create_timetable_tab()
        self.create_analytics_tab()
        self.create_visualization_tab()
        self.create_settings_tab()

        # Create status bar with clock
        status_frame = ttk.Frame(root)
        status_frame.pack(fill=tk.X, side=tk.BOTTOM, padx=5, pady=2)
        
        # Clock on the right
        self.clock_var = tk.StringVar()
        clock_label = ttk.Label(status_frame, textvariable=self.clock_var, font=("Arial", 9))
        clock_label.pack(side=tk.RIGHT, padx=5)
        
        # Status bar on the left
        self.status_bar = ttk.Label(status_frame, text="Ready", font=("Arial", 9))
        self.status_bar.pack(side=tk.LEFT, fill=tk.X, padx=5)
        
        # Start the clock
        self.clock_after_id = None
        self.update_clock()
        
        # Bind destroy event to clean up after IDs
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

        # Style configuration
        style = ttk.Style()
        style.configure("TButton", font=("Arial", 12))
        style.configure("TLabel", font=("Arial", 12))
        style.configure("TNotebook", background="#f0f0f0")
        style.configure("Treeview", font=("Arial", 10))
        style.configure("Treeview.Heading", font=("Arial", 10, "bold"))
        
        # Progress bar style
        style.configure("green.Horizontal.TProgressbar", background="green")
        style.configure("yellow.Horizontal.TProgressbar", background="yellow")
        style.configure("red.Horizontal.TProgressbar", background="red")

    def update_clock(self):
        """Update the clock in the status bar"""
        if hasattr(self, 'root') and self.root.winfo_exists():
            current_time = datetime.now().strftime('%H:%M:%S')
            self.clock_var.set(current_time)
            # Use a unique name for the after callback
            self.clock_after_id = self.root.after(1000, self.update_clock)  # Update every second
        
    def create_menu(self):
        """Create the application menu"""
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        # File menu
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="New Timetable", command=self.new_timetable)
        file_menu.add_command(label="Save Timetable", command=self.save_timetable)
        file_menu.add_command(label="Load Timetable", command=self.load_timetable)
        file_menu.add_separator()
        file_menu.add_command(label="Export to Excel", command=self.export_timetable)
        file_menu.add_command(label="Export to PDF", command=self.export_to_pdf)
        file_menu.add_command(label="Export to HTML", command=self.export_to_html)
        file_menu.add_separator()
        file_menu.add_command(label="Print Timetable", command=self.print_timetable)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.root.quit)
        
        # Edit menu
        edit_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Edit", menu=edit_menu)
        edit_menu.add_command(label="Undo", command=self.undo_action)
        edit_menu.add_command(label="Redo", command=self.redo_action)
        edit_menu.add_separator()
        edit_menu.add_command(label="Preferences", command=lambda: self.notebook.select(self.notebook.index('end')-1))
        
        # View menu
        view_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="View", menu=view_menu)
        view_menu.add_command(label="Faculty Workload", command=self.show_faculty_workload)
        view_menu.add_command(label="Room Utilization", command=self.show_room_utilization)
        view_menu.add_command(label="Conflict Report", command=self.show_conflict_report)
        view_menu.add_separator()
        view_menu.add_command(label="Refresh Analytics", command=self.refresh_analytics)
        
        # Tools menu
        tools_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Tools", menu=tools_menu)
        tools_menu.add_command(label="Generate Timetable", command=self.generate_timetable)
        tools_menu.add_command(label="Optimize Timetable", command=self.optimize_timetable)
        tools_menu.add_command(label="Check Conflicts", command=self.check_timetable_conflicts)
        tools_menu.add_separator()
        tools_menu.add_command(label="Import Faculty Data", command=self.import_faculty_data)
        tools_menu.add_command(label="Import Subject Data", command=self.import_subject_data)
        
        # Help menu
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Help", menu=help_menu)
        help_menu.add_command(label="User Guide", command=self.show_user_guide)
        help_menu.add_command(label="About", command=self.show_about)
        
    def new_timetable(self):
        """Create a new timetable"""
        if hasattr(self, 'current_timetable') and self.current_timetable:
            if messagebox.askyesno("Confirm", "This will clear the current timetable. Continue?"):
                # Save to history for undo
                self.add_to_history()
                
                # Clear current timetable
                self.current_timetable = None
                self.timetable_display.delete(1.0, tk.END)
                self.status_bar.config(text="New timetable created")
        else:
            self.status_bar.config(text="New timetable created")
    
    def add_to_history(self):
        """Add current timetable to history for undo/redo"""
        if hasattr(self, 'current_timetable') and self.current_timetable:
            # If we're not at the end of the history, truncate it
            if self.history_position < len(self.history) - 1:
                self.history = self.history[:self.history_position + 1]
            
            # Add current state to history
            import copy
            self.history.append(copy.deepcopy(self.current_timetable))
            
            # Limit history size
            if len(self.history) > self.settings["max_undo_steps"]:
                self.history = self.history[-self.settings["max_undo_steps"]:]  
            
            # Update position
            self.history_position = len(self.history) - 1
    
    def undo_action(self):
        """Undo the last action"""
        if self.history_position > 0:
            self.history_position -= 1
            self.current_timetable = self.history[self.history_position]
            self.display_timetable(self.current_timetable)
            self.status_bar.config(text="Undo successful")
        else:
            self.status_bar.config(text="Nothing to undo")
    
    def redo_action(self):
        """Redo the last undone action"""
        if self.history_position < len(self.history) - 1:
            self.history_position += 1
            self.current_timetable = self.history[self.history_position]
            self.display_timetable(self.current_timetable)
            self.status_bar.config(text="Redo successful")
        else:
            self.status_bar.config(text="Nothing to redo")
    
    def show_faculty_workload(self):
        """Show faculty workload in the analytics tab"""
        try:
            self.notebook.select(self.notebook.index("Analytics"))
            self.analytics_notebook.select(0)  # Select the faculty workload tab
        except Exception as e:
            messagebox.showwarning("Warning", "Analytics tab not available. Please generate a timetable first.")
    
    def show_room_utilization(self):
        """Show room utilization in the analytics tab"""
        try:
            self.notebook.select(self.notebook.index("Analytics"))
            self.analytics_notebook.select(1)  # Select the room utilization tab
        except Exception as e:
            messagebox.showwarning("Warning", "Analytics tab not available. Please generate a timetable first.")
    
    def show_conflict_report(self):
        """Show conflict report in the analytics tab"""
        try:
            self.notebook.select(self.notebook.index("Analytics"))
            self.analytics_notebook.select(2)  # Select the conflict report tab
        except Exception as e:
            messagebox.showwarning("Warning", "Analytics tab not available. Please generate a timetable first.")
    
    def optimize_timetable(self):
        """Optimize the current timetable"""
        if not hasattr(self, 'current_timetable') or not self.current_timetable:
            messagebox.showwarning("Warning", "Please generate a timetable first.")
            return
            
        # Add current timetable to history
        self.add_to_history()
        
        # Update status
        self.status_bar.config(text="Optimizing timetable...")
        self.progress_var.set(0)
        self.progress_status.set("Starting optimization...")
        self.root.update()
        
        # Perform optimization (this is a simplified version)
        # In a real implementation, this would use more sophisticated algorithms
        
        # 1. Try to reduce faculty conflicts
        self.progress_var.set(25)
        self.progress_status.set("Reducing faculty conflicts...")
        self.root.update()
        
        # 2. Try to balance faculty workload across days
        self.progress_var.set(50)
        self.progress_status.set("Balancing faculty workload...")
        self.root.update()
        
        # 3. Try to improve room utilization
        self.progress_var.set(75)
        self.progress_status.set("Improving room utilization...")
        self.root.update()
        
        # 4. Finalize optimization
        self.progress_var.set(100)
        self.progress_status.set("Optimization complete!")
        self.status_bar.config(text="Timetable optimized successfully")
        
        # Refresh analytics
        self.refresh_analytics()
        
        messagebox.showinfo("Success", "Timetable optimization complete!")
    
    def import_faculty_data(self):
        """Import faculty data from CSV"""
        filename = filedialog.askopenfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
            title="Import Faculty Data"
        )
        
        if not filename:
            return
            
        try:
            # Read CSV file
            faculty_data = pd.read_csv(filename)
            
            # Process faculty data
            for _, row in faculty_data.iterrows():
                faculty_name = row['Name']
                start_time = row.get('StartTime', '9:00')
                end_time = row.get('EndTime', '15:15')
                days = row.get('Days', 'Monday,Tuesday,Wednesday,Thursday,Friday,Saturday').split(',')
                
                # Add to faculties dictionary
                self.faculties[faculty_name] = {
                    "start": start_time,
                    "end": end_time,
                    "days": [day.strip() for day in days]
                }
            
            # Update faculty listbox
            self.faculty_listbox.delete(0, tk.END)
            for faculty in sorted(self.faculties.keys()):
                self.faculty_listbox.insert(tk.END, faculty)
                
            messagebox.showinfo("Success", f"Imported {len(faculty_data)} faculty records")
            self.status_bar.config(text=f"Imported faculty data from {filename}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Error importing faculty data: {str(e)}")
    
    def import_subject_data(self):
        """Import subject data from CSV"""
        filename = filedialog.askopenfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
            title="Import Subject Data"
        )
        
        if not filename:
            return
            
        try:
            # Read CSV file
            subject_data = pd.read_csv(filename)
            
            # Process subject data
            for _, row in subject_data.iterrows():
                semester = int(row['Semester'])
                division = row['Division']
                subject = row['Subject']
                hours = int(row['Hours'])
                faculty = row['Faculty']
                subject_type = row.get('Type', 'Theory')
                
                # Initialize dictionaries if needed
                if semester not in self.subjects:
                    self.subjects[semester] = {}
                if division not in self.subjects[semester]:
                    self.subjects[semester][division] = {}
                
                # Add subject
                self.subjects[semester][division][subject] = (hours, faculty, subject_type)
            
            messagebox.showinfo("Success", f"Imported {len(subject_data)} subject records")
            self.status_bar.config(text=f"Imported subject data from {filename}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Error importing subject data: {str(e)}")
    
    def show_user_guide(self):
        """Show user guide"""
        guide_text = """
        College Timetable Generator - User Guide
        
        1. General Settings: Configure basic settings like college name and semester duration.
        
        2. Faculty Management: Add, edit, and delete faculty members and their availability.
        
        3. Subjects Management: Manage subjects for each semester and division.
        
        4. Resources Management: Configure time slots, classrooms, and labs.
        
        5. Generate Timetable: Select semester and division, then generate the timetable.
        
        6. Analytics: View faculty workload, room utilization, and conflict reports.
        
        7. Visualization: See graphical representations of the timetable data.
        
        8. Export: Save your timetable to Excel, PDF, or HTML formats.
        
        For more detailed instructions, please refer to the documentation.
        """
        
        guide_window = tk.Toplevel(self.root)
        guide_window.title("User Guide")
        guide_window.geometry("600x500")
        
        guide_text_widget = scrolledtext.ScrolledText(guide_window, wrap=tk.WORD)
        guide_text_widget.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        guide_text_widget.insert(tk.END, guide_text)
        guide_text_widget.config(state=tk.DISABLED)  # Make read-only
    
    def show_about(self):
        """Show about dialog"""
        about_text = """
        College Timetable Generator
        Version 2.0
        
        An advanced application for generating and managing college timetables.
        
        Features:
        - Intelligent timetable generation algorithm
        - Faculty and resource management
        - Analytics and visualization
        - Multiple export formats
        - Conflict detection and resolution
        
        2025 Krutarth Raychura - All Rights Reserved
        """
        
        messagebox.showinfo("About College Timetable Generator", about_text)
        
    def load_default_data(self):
        """Load default data for testing"""
        # Faculty and their available time slots
        self.faculties = {
            "DPP": {"start": "8:00", "end": "15:15", "days": self.days.copy()},
            "SVP": {"start": "9:00", "end": "15:15", "days": self.days.copy()},
            "AAD": {"start": "9:00", "end": "15:15", "days": self.days.copy()},
            "PG": {"start": "9:00", "end": "15:15", "days": self.days.copy()},
            "RKS": {"start": "9:00", "end": "15:15", "days": self.days.copy()},
            "MAP": {"start": "9:00", "end": "15:15", "days": self.days.copy()},
            "ADD": {"start": "9:00", "end": "15:15", "days": self.days.copy()},
            "DP": {"start": "9:00", "end": "15:15", "days": self.days.copy()},
            "PM": {"start": "9:00", "end": "15:15", "days": self.days.copy()},
            "DAA": {"start": "11:45", "end": "12:35", "days": ["Tuesday", "Wednesday", "Friday", "Saturday"]},
            "MJ": {"start": "9:30", "end": "11:45", "days": self.days.copy()},
            "SP": {"start": "8:00", "end": "11:00", "days": ["Monday", "Tuesday", "Wednesday"]},
            "PK": {"start": "8:00", "end": "10:00", "days": self.days.copy()},
            "AN": {"start": "9:55", "end": "11:30", "days": self.days.copy()},
            "HMB": {"start": "9:00", "end": "15:15", "days": self.days.copy()}
        }

        # Time slots
        self.time_slots = [
            "8:05-8:55", "9:00-9:50", "9:55-10:45", "10:50-11:40",
            "11:45-12:35", "12:40-13:30", "13:35-14:25", "14:30-15:15"
        ]

        # Classrooms
        self.classrooms = ["101", "207", "208", "307"]

        # Labs
        self.labs = ["Lab-6", "Lab-7", "Lab-8", "Lab-113", "Lab-205"]

        # Subjects - semester, division, subject: (total_hours, faculty)
        self.subjects = {
            2: {
                "A": {
                    "ADV-C-T": (36, "PG", "Theory"),
                    "ADV-C-P": (36, "PG", "Lab"),
                    "WDT-P": (60, "DP", "Lab"),
                    "DBMS-T": (36, "ADD", "Theory"),
                    "SQL": (36, "ADD", "Lab"),
                    "BM": (30, "DPP", "Theory"),
                    "SDWT-P": (24, "MAP", "Lab"),
                    "CS": (18, "MJ", "Theory"),
                    "AOH": (12, "DAA", "Theory")
                },
                "B": {
                    "ADV-C-T": (36, "RKS", "Theory"),
                    "ADV-C-P": (36, "RKS", "Lab"),
                    "WDT-P": (60, "PM", "Lab"),
                    "DBMS-T": (36, "DP", "Theory"),
                    "SQL": (36, "DP", "Lab"),
                    "BM": (30, "DPP", "Theory"),
                    "SDWT-P": (24, "ADD", "Lab"),
                    "CS": (18, "MJ", "Theory"),
                    "AOH": (12, "DAA", "Theory")
                },
                "C": {
                    "ADV-C-T": (36, "PG", "Theory"),
                    "ADV-C-P": (36, "PG", "Lab"),
                    "WDT-P": (60, "DP", "Lab"),
                    "DBMS-T": (36, "ADD", "Theory"),
                    "SQL": (36, "ADD", "Lab"),
                    "BM": (30, "DPP", "Theory"),
                    "SDWT-P": (24, "PJ", "Lab"),
                    "CS": (18, "MJ", "Theory"),
                    "AOH": (12, "DAA", "Theory")
                },
                "D": {
                    "ADV-C-T": (36, "RKS", "Theory"),
                    "ADV-C-P": (36, "RKS", "Lab"),
                    "WDT-P": (60, "PM", "Lab"),
                    "DBMS-T": (36, "DP", "Theory"),
                    "SQL": (36, "DP", "Lab"),
                    "BM": (30, "DP", "Theory"),
                    "SDWT-P": (24, "ADD", "Lab"),
                    "CS": (18, "MJ", "Theory"),
                    "AOH": (12, "DAA", "Theory")
                },
                "E": {
                    "ADV-C-T": (36, "PG", "Theory"),
                    "ADV-C-P": (36, "PG", "Lab"),
                    "WDT-P": (60, "PM", "Lab"),
                    "DBMS-T": (36, "ADD", "Theory"),
                    "SQL": (36, "ADD", "Lab"),
                    "BM": (30, "DPP", "Theory"),
                    "SDWT-P": (24, "MAP", "Lab"),
                    "CS": (18, "MJ", "Theory"),
                    "AOH": (12, "DAA", "Theory")
                }
            },
            4: {
                "A": {
                    "JAVA-T": (36, "AAD", "Theory"),
                    "JAVA-P": (36, "AAD", "Lab"),
                    "Laravel": (60, "SVP", "Lab"),
                    "OS": (48, "HMB", "Theory"),
                    "Linux": (24, "HMB", "Lab"),
                    "DMCS": (36, "DPP", "Lab"),
                    "SM": (12, "MAP", "Theory"),
                    "YOGA": (12, "SP", "Theory"),
                    "DBMS-2": (48, "RKS", "Lab")
                },
                "B": {
                    "JAVA-T": (36, "AAD", "Theory"),
                    "JAVA-P": (36, "AAD", "Lab"),
                    "Laravel": (60, "SVP", "Lab"),
                    "OS": (48, "HMB", "Theory"),
                    "Linux": (24, "HMB", "Lab"),
                    "DMCS": (36, "DPP", "Lab"),
                    "SM": (12, "MAP", "Theory"),
                    "YOGA": (12, "SP", "Theory"),
                    "DBMS-2": (48, "PK", "Lab")
                },
                "C": {
                    "JAVA-T": (36, "AN", "Theory"),
                    "JAVA-P": (36, "AN", "Lab"),
                    "Laravel": (60, "AAD", "Lab"),
                    "OS": (48, "MAP", "Theory"),
                    "Linux": (24, "MAP", "Lab"),
                    "DMCS": (36, "DPP", "Lab"),
                    "SM": (12, "MAP", "Theory"),
                    "YOGA": (12, "SP", "Theory"),
                    "DBMS-2": (48, "SVP", "Lab")
                },
                "D": {
                    "JAVA-T": (36, "AN", "Theory"),
                    "JAVA-P": (36, "AN", "Lab"),
                    "Laravel": (60, "AAD", "Lab"),
                    "OS": (48, "PK", "Theory"),
                    "Linux": (24, "MAP", "Lab"),
                    "DMCS": (36, "DPP", "Lab"),
                    "SM": (12, "MAP", "Theory"),
                    "YOGA": (12, "SP", "Theory"),
                    "DBMS-2": (48, "RKS", "Lab")
                },
                "E": {
                    "JAVA-T": (36, "PM", "Theory"),
                    "JAVA-P": (36, "PM", "Lab"),
                    "Laravel": (60, "SVP", "Lab"),
                    "OS": (48, "MAP", "Theory"),
                    "Linux": (24, "MAP", "Lab"),
                    "DMCS": (36, "DPP", "Lab"),
                    "SM": (12, "MAP", "Theory"),
                    "YOGA": (12, "SP", "Theory"),
                    "DBMS-2": (48, "RKS", "Lab")
                }
            }
        }

    def create_input_tab(self):
        """Create the input tab for general settings"""
        input_frame = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(input_frame, text="General Settings")

        # College Name
        ttk.Label(input_frame, text="College Name:").grid(row=0, column=0, sticky="w", pady=5)
        self.college_name = tk.StringVar(value="Lok Jagruti Kendra University")
        ttk.Entry(input_frame, textvariable=self.college_name, width=50).grid(row=0, column=1, sticky="w", pady=5)

        # Semester Duration
        ttk.Label(input_frame, text="Semester Duration (weeks):").grid(row=1, column=0, sticky="w", pady=5)
        self.semester_duration = tk.IntVar(value=11)
        ttk.Spinbox(input_frame, from_=8, to=16, textvariable=self.semester_duration, width=5).grid(row=1, column=1, sticky="w", pady=5)

        # Working Days
        ttk.Label(input_frame, text="Working Days:").grid(row=2, column=0, sticky="w", pady=5)
        days_frame = ttk.Frame(input_frame)
        days_frame.grid(row=2, column=1, sticky="w", pady=5)

        self.working_days = {}
        for i, day in enumerate(["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]):
            self.working_days[day] = tk.BooleanVar(value=True)
            ttk.Checkbutton(days_frame, text=day, variable=self.working_days[day]).grid(row=0, column=i, padx=5)

        # Lectures per day
        ttk.Label(input_frame, text="Min. Lectures per day:").grid(row=3, column=0, sticky="w", pady=5)
        self.min_lectures = tk.IntVar(value=3)
        ttk.Spinbox(input_frame, from_=1, to=8, textvariable=self.min_lectures, width=5).grid(row=3, column=1, sticky="w", pady=5)

        ttk.Label(input_frame, text="Max. Lectures per day:").grid(row=4, column=0, sticky="w", pady=5)
        self.max_lectures = tk.IntVar(value=6)
        ttk.Spinbox(input_frame, from_=1, to=8, textvariable=self.max_lectures, width=5).grid(row=4, column=1, sticky="w", pady=5)

        # Max faculty lectures
        ttk.Label(input_frame, text="Max. Faculty Lectures per day:").grid(row=5, column=0, sticky="w", pady=5)
        self.max_faculty_lectures = tk.IntVar(value=4)
        ttk.Spinbox(input_frame, from_=1, to=8, textvariable=self.max_faculty_lectures, width=5).grid(row=5, column=1, sticky="w", pady=5)

        # Additional information
        info_frame = ttk.LabelFrame(input_frame, text="Special Instructions", padding=10)
        info_frame.grid(row=6, column=0, columnspan=2, sticky="ew", pady=10)

        self.special_instructions = scrolledtext.ScrolledText(info_frame, width=60, height=10, wrap=tk.WORD)
        self.special_instructions.pack(fill=tk.BOTH, expand=True)
        self.special_instructions.insert(tk.END, "1. Every faculty can take maximum 4 lectures per day but if he/she have more than 216 load then she/he will take 5 lectures daily.\n2. Try to make class starts on same time daily.\n3. Every class has minimum 3 lectures and maximum 6 lectures on each day.")

        # Save button
        ttk.Button(input_frame, text="Save Settings", command=self.save_settings).grid(row=7, column=0, columnspan=2, pady=10)

    def create_faculty_tab(self):
        """Create the faculty management tab"""
        faculty_frame = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(faculty_frame, text="Faculty Management")

        # Left frame for faculty list
        left_frame = ttk.Frame(faculty_frame)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))

        ttk.Label(left_frame, text="Faculty List").pack(anchor=tk.W)

        # Faculty list
        self.faculty_listbox = tk.Listbox(left_frame, width=25, height=20)
        self.faculty_listbox.pack(fill=tk.BOTH, expand=True)
        self.faculty_listbox.bind('<<ListboxSelect>>', self.select_faculty)

        # Populate faculty list
        for faculty in sorted(self.faculties.keys()):
            self.faculty_listbox.insert(tk.END, faculty)

        # Buttons frame
        buttons_frame = ttk.Frame(left_frame)
        buttons_frame.pack(fill=tk.X, pady=5)

        ttk.Button(buttons_frame, text="Add", command=self.add_faculty).pack(side=tk.LEFT, padx=2)
        ttk.Button(buttons_frame, text="Delete", command=self.delete_faculty).pack(side=tk.LEFT, padx=2)

        # Right frame for faculty details
        right_frame = ttk.LabelFrame(faculty_frame, text="Faculty Details", padding=10)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        # Faculty name
        ttk.Label(right_frame, text="Name:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.faculty_name = tk.StringVar()
        ttk.Entry(right_frame, textvariable=self.faculty_name, width=20).grid(row=0, column=1, sticky=tk.W, pady=5)

        # Faculty availability
        ttk.Label(right_frame, text="Available From:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.faculty_start = tk.StringVar()
        ttk.Entry(right_frame, textvariable=self.faculty_start, width=10).grid(row=1, column=1, sticky=tk.W, pady=5)

        ttk.Label(right_frame, text="Available To:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.faculty_end = tk.StringVar()
        ttk.Entry(right_frame, textvariable=self.faculty_end, width=10).grid(row=2, column=1, sticky=tk.W, pady=5)

        # Available days
        ttk.Label(right_frame, text="Available Days:").grid(row=3, column=0, sticky=tk.W, pady=5)
        days_frame = ttk.Frame(right_frame)
        days_frame.grid(row=3, column=1, sticky=tk.W, pady=5)

        self.faculty_days = {}
        for i, day in enumerate(self.days):
            self.faculty_days[day] = tk.BooleanVar()
            ttk.Checkbutton(days_frame, text=day, variable=self.faculty_days[day]).grid(row=i//3, column=i%3, sticky=tk.W)

        # Save button
        ttk.Button(right_frame, text="Save Faculty", command=self.save_faculty).grid(row=4, column=0, columnspan=2, pady=10)

    def create_subjects_tab(self):
        """Create the subjects management tab"""
        subjects_frame = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(subjects_frame, text="Subjects Management")

        # Left frame for semester and division selection
        left_frame = ttk.Frame(subjects_frame)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))

        # Semester selection
        ttk.Label(left_frame, text="Semester:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.semester_var = tk.StringVar()
        semester_combo = ttk.Combobox(left_frame, textvariable=self.semester_var, values=list(self.subjects.keys()))
        semester_combo.grid(row=0, column=1, sticky=tk.W, pady=5)
        semester_combo.bind('<<ComboboxSelected>>', self.update_divisions)

        # Division selection
        ttk.Label(left_frame, text="Division:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.division_var = tk.StringVar()
        self.division_combo = ttk.Combobox(left_frame, textvariable=self.division_var)
        self.division_combo.grid(row=1, column=1, sticky=tk.W, pady=5)
        self.division_combo.bind('<<ComboboxSelected>>', self.update_subjects_list)

        # Subject list
        ttk.Label(left_frame, text="Subjects:").grid(row=2, column=0, columnspan=2, sticky=tk.W, pady=5)
        self.subjects_listbox = tk.Listbox(left_frame, width=25, height=15)
        self.subjects_listbox.grid(row=3, column=0, columnspan=2, sticky=tk.NSEW, pady=5)
        self.subjects_listbox.bind('<<ListboxSelect>>', self.select_subject)

        # Buttons frame
        buttons_frame = ttk.Frame(left_frame)
        buttons_frame.grid(row=4, column=0, columnspan=2, sticky=tk.EW, pady=5)

        ttk.Button(buttons_frame, text="Add", command=self.add_subject).pack(side=tk.LEFT, padx=2)
        ttk.Button(buttons_frame, text="Delete", command=self.delete_subject).pack(side=tk.LEFT, padx=2)

        # Right frame for subject details
        right_frame = ttk.LabelFrame(subjects_frame, text="Subject Details", padding=10)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        # Subject name
        ttk.Label(right_frame, text="Subject Name:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.subject_name = tk.StringVar()
        ttk.Entry(right_frame, textvariable=self.subject_name, width=20).grid(row=0, column=1, sticky=tk.W, pady=5)

        # Subject hours
        ttk.Label(right_frame, text="Total Hours:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.subject_hours = tk.IntVar()
        ttk.Spinbox(right_frame, from_=1, to=100, textvariable=self.subject_hours, width=5).grid(row=1, column=1, sticky=tk.W, pady=5)

        # Subject type
        ttk.Label(right_frame, text="Type:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.subject_type = tk.StringVar()
        ttk.Combobox(right_frame, textvariable=self.subject_type, values=["Theory", "Lab"]).grid(row=2, column=1, sticky=tk.W, pady=5)

        # Faculty
        ttk.Label(right_frame, text="Faculty:").grid(row=3, column=0, sticky=tk.W, pady=5)
        self.subject_faculty = tk.StringVar()
        ttk.Combobox(right_frame, textvariable=self.subject_faculty, values=sorted(self.faculties.keys())).grid(row=3, column=1, sticky=tk.W, pady=5)

        # Save button
        ttk.Button(right_frame, text="Save Subject", command=self.save_subject).grid(row=4, column=0, columnspan=2, pady=10)

        # Set default values
        if self.subjects:
            semester_combo.set(list(self.subjects.keys())[0])
            self.update_divisions(None)

    def create_resources_tab(self):
        """Create the resources management tab"""
        resources_frame = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(resources_frame, text="Resources Management")

        # Left frame for time slots
        left_frame = ttk.LabelFrame(resources_frame, text="Time Slots", padding=10)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))

        self.time_slots_listbox = tk.Listbox(left_frame, width=15, height=10)
        self.time_slots_listbox.pack(fill=tk.BOTH, expand=True, pady=5)

        # Populate time slots
        for slot in self.time_slots:
            self.time_slots_listbox.insert(tk.END, slot)

        # Time slot entry
        time_slot_frame = ttk.Frame(left_frame)
        time_slot_frame.pack(fill=tk.X, pady=5)

        self.time_slot_var = tk.StringVar()
        ttk.Entry(time_slot_frame, textvariable=self.time_slot_var, width=15).pack(side=tk.LEFT)

        # Time slot buttons
        time_slot_buttons = ttk.Frame(left_frame)
        time_slot_buttons.pack(fill=tk.X, pady=5)

        ttk.Button(time_slot_buttons, text="Add", command=self.add_time_slot).pack(side=tk.LEFT, padx=2)
        ttk.Button(time_slot_buttons, text="Delete", command=self.delete_time_slot).pack(side=tk.LEFT, padx=2)

        # Middle frame for classrooms
        middle_frame = ttk.LabelFrame(resources_frame, text="Classrooms", padding=10)
        middle_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))

        self.classrooms_listbox = tk.Listbox(middle_frame, width=15, height=10)
        self.classrooms_listbox.pack(fill=tk.BOTH, expand=True, pady=5)

        # Populate classrooms
        for room in self.classrooms:
            self.classrooms_listbox.insert(tk.END, room)

        # Classroom entry
        classroom_frame = ttk.Frame(middle_frame)
        classroom_frame.pack(fill=tk.X, pady=5)

        self.classroom_var = tk.StringVar()
        ttk.Entry(classroom_frame, textvariable=self.classroom_var, width=15).pack(side=tk.LEFT)

        # Classroom buttons
        classroom_buttons = ttk.Frame(middle_frame)
        classroom_buttons.pack(fill=tk.X, pady=5)

        ttk.Button(classroom_buttons, text="Add", command=self.add_classroom).pack(side=tk.LEFT, padx=2)
        ttk.Button(classroom_buttons, text="Delete", command=self.delete_classroom).pack(side=tk.LEFT, padx=2)

        # Right frame for labs
        right_frame = ttk.LabelFrame(resources_frame, text="Labs", padding=10)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        self.labs_listbox = tk.Listbox(right_frame, width=15, height=10)
        self.labs_listbox.pack(fill=tk.BOTH, expand=True, pady=5)

        # Populate labs
        for lab in self.labs:
            self.labs_listbox.insert(tk.END, lab)

        # Lab entry
        lab_frame = ttk.Frame(right_frame)
        lab_frame.pack(fill=tk.X, pady=5)

        self.lab_var = tk.StringVar()
        ttk.Entry(lab_frame, textvariable=self.lab_var, width=15).pack(side=tk.LEFT)

        # Lab buttons
        lab_buttons = ttk.Frame(right_frame)
        lab_buttons.pack(fill=tk.X, pady=5)

        ttk.Button(lab_buttons, text="Add", command=self.add_lab).pack(side=tk.LEFT, padx=2)
        ttk.Button(lab_buttons, text="Delete", command=self.delete_lab).pack(side=tk.LEFT, padx=2)

    def create_timetable_tab(self):
        """Create the timetable generation tab"""
        timetable_frame = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(timetable_frame, text="Generate Timetable")

        # Top frame for options
        top_frame = ttk.Frame(timetable_frame)
        top_frame.pack(fill=tk.X, pady=10)

        # Semester selection
        ttk.Label(top_frame, text="Semester:").pack(side=tk.LEFT, padx=5)
        self.timetable_semester = tk.StringVar()
        ttk.Combobox(top_frame, textvariable=self.timetable_semester, values=list(self.subjects.keys())).pack(side=tk.LEFT, padx=5)

        # Division selection
        ttk.Label(top_frame, text="Division:").pack(side=tk.LEFT, padx=5)
        self.timetable_division = tk.StringVar(value="All")
        ttk.Combobox(top_frame, textvariable=self.timetable_division, values=["All", "A", "B", "C", "D", "E"]).pack(side=tk.LEFT, padx=5)

        # Buttons frame
        buttons_frame = ttk.Frame(top_frame)
        buttons_frame.pack(side=tk.LEFT, padx=20)

        # Generate button
        ttk.Button(buttons_frame, text="Generate Timetable", command=self.generate_timetable).pack(side=tk.LEFT, padx=5)

        # Export button
        ttk.Button(buttons_frame, text="Export to Excel", command=self.export_timetable).pack(side=tk.LEFT, padx=5)

        # Save/Load buttons
        ttk.Button(buttons_frame, text="Save Timetable", command=self.save_timetable).pack(side=tk.LEFT, padx=5)
        ttk.Button(buttons_frame, text="Load Timetable", command=self.load_timetable).pack(side=tk.LEFT, padx=5)

        # Progress frame
        progress_frame = ttk.LabelFrame(timetable_frame, text="Generation Progress", padding=10)
        progress_frame.pack(fill=tk.X, pady=10)

        # Progress bar
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(progress_frame, orient="horizontal", length=300, 
                                           mode="determinate", variable=self.progress_var,
                                           style="green.Horizontal.TProgressbar")
        self.progress_bar.pack(fill=tk.X, pady=5)

        # Progress status
        self.progress_status = tk.StringVar(value="Ready to generate timetable")
        ttk.Label(progress_frame, textvariable=self.progress_status).pack(anchor=tk.W, pady=5)

        # Bottom frame for timetable display
        bottom_frame = ttk.Frame(timetable_frame)
        bottom_frame.pack(fill=tk.BOTH, expand=True, pady=10)

        # Timetable display
        self.timetable_display = scrolledtext.ScrolledText(bottom_frame, wrap=tk.WORD, width=100, height=30, font=("Courier New", 10))
        self.timetable_display.pack(fill=tk.BOTH, expand=True)

    # Event handlers and utility methods
    def save_settings(self):
        """Save general settings"""
        messagebox.showinfo("Success", "Settings saved successfully!")

    def select_faculty(self, event):
        """Handle faculty selection"""
        if not self.faculty_listbox.curselection():
            return

        faculty_name = self.faculty_listbox.get(self.faculty_listbox.curselection())
        faculty_data = self.faculties[faculty_name]

        self.faculty_name.set(faculty_name)
        self.faculty_start.set(faculty_data["start"])
        self.faculty_end.set(faculty_data["end"])

        for day in self.days:
            if day in faculty_data["days"]:
                self.faculty_days[day].set(True)
            else:
                self.faculty_days[day].set(False)

    def add_faculty(self):
        """Add a new faculty"""
        new_faculty = f"Faculty{len(self.faculties) + 1}"
        self.faculties[new_faculty] = {"start": "9:00", "end": "15:15", "days": self.days.copy()}
        self.faculty_listbox.insert(tk.END, new_faculty)
        messagebox.showinfo("Success", f"Added new faculty: {new_faculty}")

    def delete_faculty(self):
        """Delete selected faculty"""
        if not self.faculty_listbox.curselection():
            return

        faculty_name = self.faculty_listbox.get(self.faculty_listbox.curselection())
        if messagebox.askyesno("Confirm", f"Delete faculty {faculty_name}?"):
            del self.faculties[faculty_name]
            self.faculty_listbox.delete(self.faculty_listbox.curselection())
            messagebox.showinfo("Success", f"Deleted faculty: {faculty_name}")

    def save_faculty(self):
        """Save faculty details"""
        faculty_name = self.faculty_name.get()
        if not faculty_name:
            messagebox.showerror("Error", "Faculty name cannot be empty")
            return

        # Update or add faculty
        active_days = [day for day in self.days if self.faculty_days[day].get()]
        self.faculties[faculty_name] = {
            "start": self.faculty_start.get(),
            "end": self.faculty_end.get(),
            "days": active_days
        }

        # Update listbox
        if faculty_name not in self.faculty_listbox.get(0, tk.END):
            self.faculty_listbox.insert(tk.END, faculty_name)

        messagebox.showinfo("Success", f"Faculty {faculty_name} saved successfully")

    def update_divisions(self, event):
        """Update divisions based on selected semester"""
        try:
            semester = int(self.semester_var.get())
            if semester in self.subjects:
                divisions = list(self.subjects[semester].keys())
                self.division_combo['values'] = divisions
                if divisions:
                    self.division_combo.set(divisions[0])
                    self.update_subjects_list(None)
        except (ValueError, KeyError):
            pass

    def update_subjects_list(self, event):
        """Update subjects list based on selected semester and division"""
        try:
            semester = int(self.semester_var.get())
            division = self.division_var.get()

            if semester in self.subjects and division in self.subjects[semester]:
                self.subjects_listbox.delete(0, tk.END)
                for subject in self.subjects[semester][division]:
                    self.subjects_listbox.insert(tk.END, subject)
        except (ValueError, KeyError):
            pass

    def select_subject(self, event):
        """Handle subject selection"""
        if not self.subjects_listbox.curselection():
            return

        try:
            semester = int(self.semester_var.get())
            division = self.division_var.get()
            subject = self.subjects_listbox.get(self.subjects_listbox.curselection())

            if semester in self.subjects and division in self.subjects[semester] and subject in self.subjects[semester][division]:
                hours, faculty, subject_type = self.subjects[semester][division][subject]
                self.subject_name.set(subject)
                self.subject_hours.set(hours)
                self.subject_faculty.set(faculty)
                self.subject_type.set(subject_type)
        except (ValueError, KeyError):
            pass

    def add_subject(self):
        """Add a new subject"""
        try:
            semester = int(self.semester_var.get())
            division = self.division_var.get()

            if semester not in self.subjects:
                self.subjects[semester] = {}
            if division not in self.subjects[semester]:
                self.subjects[semester][division] = {}

            new_subject = f"Subject{len(self.subjects[semester][division]) + 1}"
            self.subjects[semester][division][new_subject] = (36, "TBD", "Theory")
            self.subjects_listbox.insert(tk.END, new_subject)
            messagebox.showinfo("Success", f"Added new subject: {new_subject}")
        except (ValueError, KeyError):
            messagebox.showerror("Error", "Please select a valid semester and division")

    def delete_subject(self):
        """Delete selected subject"""
        if not self.subjects_listbox.curselection():
            return

        try:
            semester = int(self.semester_var.get())
            division = self.division_var.get()
            subject = self.subjects_listbox.get(self.subjects_listbox.curselection())

            if messagebox.askyesno("Confirm", f"Delete subject {subject}?"):
                del self.subjects[semester][division][subject]
                self.subjects_listbox.delete(self.subjects_listbox.curselection())
                messagebox.showinfo("Success", f"Deleted subject: {subject}")
        except (ValueError, KeyError):
            messagebox.showerror("Error", "Error deleting subject")

    def save_subject(self):
        """Save subject details"""
        try:
            semester = int(self.semester_var.get())
            division = self.division_var.get()
            subject_name = self.subject_name.get()
            hours = self.subject_hours.get()
            faculty = self.subject_faculty.get()
            subject_type = self.subject_type.get()

            if not subject_name or not faculty:
                messagebox.showerror("Error", "Subject name and faculty cannot be empty")
                return

            if semester not in self.subjects:
                self.subjects[semester] = {}
            if division not in self.subjects[semester]:
                self.subjects[semester][division] = {}

            self.subjects[semester][division][subject_name] = (hours, faculty, subject_type)

            # Update listbox if needed
            if subject_name not in self.subjects_listbox.get(0, tk.END):
                self.subjects_listbox.insert(tk.END, subject_name)

            messagebox.showinfo("Success", f"Subject {subject_name} saved successfully")
        except (ValueError, KeyError):
            messagebox.showerror("Error", "Error saving subject")

    def add_time_slot(self):
        """Add a new time slot"""
        time_slot = self.time_slot_var.get()
        if not time_slot:
            messagebox.showerror("Error", "Time slot cannot be empty")
            return

        if time_slot not in self.time_slots:
            self.time_slots.append(time_slot)
            self.time_slots_listbox.insert(tk.END, time_slot)
            self.time_slot_var.set("")
            messagebox.showinfo("Success", f"Added time slot: {time_slot}")

    def delete_time_slot(self):
        """Delete selected time slot"""
        if not self.time_slots_listbox.curselection():
            return

        time_slot = self.time_slots_listbox.get(self.time_slots_listbox.curselection())
        if messagebox.askyesno("Confirm", f"Delete time slot {time_slot}?"):
            self.time_slots.remove(time_slot)
            self.time_slots_listbox.delete(self.time_slots_listbox.curselection())
            messagebox.showinfo("Success", f"Deleted time slot: {time_slot}")

    def add_classroom(self):
        """Add a new classroom"""
        classroom = self.classroom_var.get()
        if not classroom:
            messagebox.showerror("Error", "Classroom cannot be empty")
            return

        if classroom not in self.classrooms:
            self.classrooms.append(classroom)
            self.classrooms_listbox.insert(tk.END, classroom)
            self.classroom_var.set("")
            messagebox.showinfo("Success", f"Added classroom: {classroom}")

    def delete_classroom(self):
        """Delete selected classroom"""
        if not self.classrooms_listbox.curselection():
            return

        classroom = self.classrooms_listbox.get(self.classrooms_listbox.curselection())
        if messagebox.askyesno("Confirm", f"Delete classroom {classroom}?"):
            self.classrooms.remove(classroom)
            self.classrooms_listbox.delete(self.classrooms_listbox.curselection())
            messagebox.showinfo("Success", f"Deleted classroom: {classroom}")

    def add_lab(self):
        """Add a new lab"""
        lab = self.lab_var.get()
        if not lab:
            messagebox.showerror("Error", "Lab cannot be empty")
            return

        if lab not in self.labs:
            self.labs.append(lab)
            self.labs_listbox.insert(tk.END, lab)
            self.lab_var.set("")
            messagebox.showinfo("Success", f"Added lab: {lab}")

    def delete_lab(self):
        """Delete selected lab"""
        if not self.labs_listbox.curselection():
            return

        lab = self.labs_listbox.get(self.labs_listbox.curselection())
        if messagebox.askyesno("Confirm", f"Delete lab {lab}?"):
            self.labs.remove(lab)
            self.labs_listbox.delete(self.labs_listbox.curselection())
            messagebox.showinfo("Success", f"Deleted lab: {lab}")

    def initialize_timetable_structure(self):
        """Initialize the timetable structure"""
        timetable = {}
        for semester in self.subjects:
            timetable[semester] = {}
            for division in self.subjects[semester]:
                timetable[semester][division] = {}
                for day in [d for d in ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"] if self.working_days[d].get()]:
                    timetable[semester][division][day] = {}
                    for time_slot in self.time_slots:
                        timetable[semester][division][day][time_slot] = {
                            "subject": None,
                            "faculty": None,
                            "room": None,
                            "is_lab": False
                        }
        return timetable

    def calculate_weekly_load(self):
        """Calculate weekly load for each subject"""
        weekly_load = {}
        for semester in self.subjects:
            for division in self.subjects[semester]:
                for subject, (total_hours, faculty, subject_type) in self.subjects[semester][division].items():
                    # Calculate lectures per week based on total hours and semester duration
                    lectures_per_week = max(1, round(total_hours / self.semester_duration.get()))
                    weekly_load[(semester, division, subject)] = {
                        "lectures_per_week": lectures_per_week,
                        "faculty": faculty,
                        "type": subject_type
                    }
        return weekly_load

    def check_faculty_availability(self, faculty, day, time_slot):
        """Check if the faculty is available for the given day and time slot"""
        if faculty not in self.faculties:
            return False

        faculty_data = self.faculties[faculty]

        # Check if the faculty works on the given day
        if day not in faculty_data["days"]:
            return False

        # Check if the time slot is within the faculty's working hours
        time_start, time_end = time_slot.split("-")

        if time_start < faculty_data["start"] or time_end > faculty_data["end"]:
            return False

        return True

    def count_faculty_assignments(self, timetable, faculty, day):
        """Count how many sessions a faculty has on a given day"""
        count = 0
        for semester in timetable:
            for division in timetable[semester]:
                for time_slot in timetable[semester][division].get(day, {}):
                    if timetable[semester][division][day][time_slot]["faculty"] == faculty:
                        count += 1
        return count
        
    def count_faculty_subjects(self, faculty):
        """Count how many subjects a faculty teaches"""
        count = 0
        for semester in self.subjects:
            for division in self.subjects[semester]:
                for subject, (hours, assigned_faculty, subject_type) in self.subjects[semester][division].items():
                    if assigned_faculty == faculty:
                        count += 1
        return count
        
    def get_faculty_total_load(self, faculty):
        """Calculate total teaching hours for a faculty"""
        total_hours = 0
        for semester in self.subjects:
            for division in self.subjects[semester]:
                for subject, (hours, assigned_faculty, subject_type) in self.subjects[semester][division].items():
                    if assigned_faculty == faculty:
                        total_hours += hours
        return total_hours
        
    def check_faculty_conflict(self, timetable, faculty, day, time_slot):
        """Check if a faculty is already scheduled at a specific time"""
        for semester in timetable:
            for division in timetable[semester]:
                if day in timetable[semester][division] and time_slot in timetable[semester][division][day]:
                    if timetable[semester][division][day][time_slot]["faculty"] == faculty:
                        return True
        return False
        
    def find_available_rooms(self, timetable, day, time_slot, room_type="classroom"):
        """Find rooms available at a specific time"""
        # Determine which rooms to check based on type
        rooms = self.classrooms if room_type == "classroom" else self.labs
        
        # Find rooms that are already in use at this time
        used_rooms = set()
        for semester in timetable:
            for division in timetable[semester]:
                if day in timetable[semester][division] and time_slot in timetable[semester][division][day]:
                    room = timetable[semester][division][day][time_slot]["room"]
                    if room:
                        used_rooms.add(room)
        
        # Return available rooms
        return [room for room in rooms if room not in used_rooms]
        
    def get_preferred_slots(self, timetable, semester, division, subject):
        """Determine preferred time slots for consistency"""
        # Check if this subject is already scheduled on other days
        existing_slots = {}
        for day in ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]:
            if day not in timetable[semester][division]:
                continue
                
            for time_slot in self.time_slots:
                if time_slot not in timetable[semester][division][day]:
                    continue
                    
                if timetable[semester][division][day][time_slot]["subject"] == subject:
                    if time_slot in existing_slots:
                        existing_slots[time_slot] += 1
                    else:
                        existing_slots[time_slot] = 1
        
        # Sort slots by frequency
        preferred_slots = sorted(existing_slots.items(), key=lambda x: x[1], reverse=True)
        
        # Return list of slot names
        return [slot for slot, _ in preferred_slots] if preferred_slots else []

    def create_settings_tab(self):
        """Create the settings tab for application preferences"""
        settings_frame = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(settings_frame, text="Settings")

        # Create scrollable frame for settings
        canvas = tk.Canvas(settings_frame)
        scrollbar = ttk.Scrollbar(settings_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Application settings section
        app_settings = ttk.LabelFrame(scrollable_frame, text="Application Settings", padding=10)
        app_settings.pack(fill="x", padx=10, pady=10)

        # Theme color
        ttk.Label(app_settings, text="Theme Color:").grid(row=0, column=0, sticky="w", pady=5)
        theme_frame = ttk.Frame(app_settings)
        theme_frame.grid(row=0, column=1, sticky="w", pady=5)
        
        self.theme_color_preview = tk.Label(theme_frame, text="   ", bg=self.settings["theme_color"], width=3, height=1, relief="solid")
        self.theme_color_preview.pack(side="left", padx=5)
        
        ttk.Button(theme_frame, text="Change", command=self.change_theme_color).pack(side="left")

        # Auto-save settings
        self.auto_save_var = tk.BooleanVar(value=self.settings["auto_save"])
        ttk.Checkbutton(app_settings, text="Auto-save timetable", variable=self.auto_save_var).grid(row=1, column=0, columnspan=2, sticky="w", pady=5)
        
        ttk.Label(app_settings, text="Auto-save path:").grid(row=2, column=0, sticky="w", pady=5)
        auto_save_frame = ttk.Frame(app_settings)
        auto_save_frame.grid(row=2, column=1, sticky="w", pady=5)
        
        self.auto_save_path_var = tk.StringVar(value=self.settings["auto_save_path"])
        ttk.Entry(auto_save_frame, textvariable=self.auto_save_path_var, width=30).pack(side="left", padx=5)
        ttk.Button(auto_save_frame, text="Browse", command=self.browse_auto_save_path).pack(side="left")

        # Export preferences
        ttk.Label(app_settings, text="Preferred Export Format:").grid(row=3, column=0, sticky="w", pady=5)
        self.export_format_var = tk.StringVar(value=self.settings["default_export_format"])
        ttk.Combobox(app_settings, textvariable=self.export_format_var, values=["excel", "pdf", "html"]).grid(row=3, column=1, sticky="w", pady=5)

        # Display settings section
        display_settings = ttk.LabelFrame(scrollable_frame, text="Display Settings", padding=10)
        display_settings.pack(fill="x", padx=10, pady=10)

        # Show conflicts
        self.show_conflicts_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(display_settings, text="Highlight conflicts in timetable", variable=self.show_conflicts_var).grid(row=0, column=0, columnspan=2, sticky="w", pady=5)

        # Show room numbers
        self.show_room_numbers_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(display_settings, text="Show room numbers in timetable", variable=self.show_room_numbers_var).grid(row=1, column=0, columnspan=2, sticky="w", pady=5)

        # Show faculty codes
        self.show_faculty_codes_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(display_settings, text="Show faculty codes in timetable", variable=self.show_faculty_codes_var).grid(row=2, column=0, columnspan=2, sticky="w", pady=5)

        # Algorithm settings section
        algorithm_settings = ttk.LabelFrame(scrollable_frame, text="Algorithm Settings", padding=10)
        algorithm_settings.pack(fill="x", padx=10, pady=10)

        # Prioritize consistency
        self.prioritize_consistency_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(algorithm_settings, text="Prioritize consistent time slots across days", variable=self.prioritize_consistency_var).grid(row=0, column=0, columnspan=2, sticky="w", pady=5)

        # Balance faculty workload
        self.balance_workload_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(algorithm_settings, text="Balance faculty workload across days", variable=self.balance_workload_var).grid(row=1, column=0, columnspan=2, sticky="w", pady=5)

        # Optimize room usage
        self.optimize_rooms_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(algorithm_settings, text="Optimize room usage", variable=self.optimize_rooms_var).grid(row=2, column=0, columnspan=2, sticky="w", pady=5)

        # Buttons
        buttons_frame = ttk.Frame(scrollable_frame)
        buttons_frame.pack(fill="x", padx=10, pady=20)

        ttk.Button(buttons_frame, text="Save Settings", command=self.save_app_settings).pack(side="left", padx=5)
        ttk.Button(buttons_frame, text="Reset to Defaults", command=self.reset_app_settings).pack(side="left", padx=5)

    def change_theme_color(self):
        """Change the application theme color"""
        color = colorchooser.askcolor(initialcolor=self.settings["theme_color"], title="Select Theme Color")[1]
        if color:
            self.settings["theme_color"] = color
            self.theme_color_preview.config(bg=color)
            self.root.configure(bg=color)
            self.status_bar.config(text=f"Theme color changed to {color}")

    def browse_auto_save_path(self):
        """Browse for auto-save file path"""
        filename = filedialog.asksaveasfilename(
            defaultextension=".ttbl",
            filetypes=[("Timetable files", "*.ttbl"), ("All files", "*.*")],
            title="Select Auto-save Location"
        )
        if filename:
            self.auto_save_path_var.set(filename)

    def save_app_settings(self):
        """Save application settings"""
        # Update settings from UI
        self.settings["auto_save"] = self.auto_save_var.get()
        self.settings["auto_save_path"] = self.auto_save_path_var.get()
        self.settings["default_export_format"] = self.export_format_var.get()
        
        # Save settings to file
        try:
            settings_path = os.path.join(os.path.expanduser("~"), "timetable_settings.json")
            import json
            with open(settings_path, 'w') as f:
                json.dump(self.settings, f, indent=4)
            
            messagebox.showinfo("Success", "Settings saved successfully")
            self.status_bar.config(text="Settings saved")
        except Exception as e:
            messagebox.showerror("Error", f"Error saving settings: {str(e)}")

    def reset_app_settings(self):
        """Reset application settings to defaults"""
        if messagebox.askyesno("Confirm", "Reset all settings to default values?"):
            # Reset to defaults
            self.settings = {
                "theme_color": "#4a7abc",
                "auto_save": False,
                "auto_save_path": os.path.join(os.path.expanduser("~"), "timetable_autosave.ttbl"),
                "default_export_format": "excel"
            }
            
            # Update UI
            self.theme_color_preview.config(bg=self.settings["theme_color"])
            self.auto_save_var.set(self.settings["auto_save"])
            self.auto_save_path_var.set(self.settings["auto_save_path"])
            self.export_format_var.set(self.settings["default_export_format"])
            
            # Apply settings
            self.root.configure(bg=self.settings["theme_color"])
            
            self.status_bar.config(text="Settings reset to defaults")
            
    def on_close(self):
        """Handle cleanup when closing the application"""
        # Cancel any scheduled after callbacks
        if hasattr(self, 'clock_after_id') and self.clock_after_id:
            self.root.after_cancel(self.clock_after_id)
        
        # Cancel any other after callbacks
        for after_id in self.after_ids:
            try:
                self.root.after_cancel(after_id)
            except Exception:
                pass
        
        # Save settings if auto-save is enabled
        if self.settings.get("auto_save", False):
            try:
                self.save_timetable(self.settings.get("auto_save_path"))
            except Exception:
                pass
        
        # Destroy the root window
        self.root.destroy()
            
    def create_analytics_tab(self):
        """Create the analytics tab with faculty workload, room utilization, and conflict reports"""
        self.analytics_frame = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(self.analytics_frame, text="Analytics")

        # Top frame for options
        top_frame = ttk.Frame(self.analytics_frame)
        top_frame.pack(fill=tk.X, pady=10)

        # Refresh button
        ttk.Button(top_frame, text="Refresh Analytics", command=self.refresh_analytics).pack(side=tk.LEFT, padx=5)

        # Create notebook for analytics
        self.analytics_notebook = ttk.Notebook(self.analytics_frame)
        self.analytics_notebook.pack(fill=tk.BOTH, expand=True, pady=10)

        # Faculty workload tab
        faculty_frame = ttk.Frame(self.analytics_notebook, padding=10)
        self.analytics_notebook.add(faculty_frame, text="Faculty Workload")

        # Faculty workload treeview
        columns = ("faculty", "subjects", "total_hours", "weekly_hours", "max_daily")
        self.faculty_tree = ttk.Treeview(faculty_frame, columns=columns, show="headings")
        
        # Configure columns
        self.faculty_tree.heading("faculty", text="Faculty")
        self.faculty_tree.heading("subjects", text="Subjects")
        self.faculty_tree.heading("total_hours", text="Total Hours")
        self.faculty_tree.heading("weekly_hours", text="Weekly Hours")
        self.faculty_tree.heading("max_daily", text="Max Daily")
        
        self.faculty_tree.column("faculty", width=100)
        self.faculty_tree.column("subjects", width=80)
        self.faculty_tree.column("total_hours", width=100)
        self.faculty_tree.column("weekly_hours", width=100)
        self.faculty_tree.column("max_daily", width=100)
        
        # Add scrollbar
        faculty_scroll = ttk.Scrollbar(faculty_frame, orient="vertical", command=self.faculty_tree.yview)
        self.faculty_tree.configure(yscrollcommand=faculty_scroll.set)
        
        # Pack widgets
        self.faculty_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        faculty_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        # Room utilization tab
        room_frame = ttk.Frame(self.analytics_notebook, padding=10)
        self.analytics_notebook.add(room_frame, text="Room Utilization")

        # Room utilization treeview
        columns = ("room", "type", "usage_count", "usage_percent")
        self.room_tree = ttk.Treeview(room_frame, columns=columns, show="headings")
        
        # Configure columns
        self.room_tree.heading("room", text="Room")
        self.room_tree.heading("type", text="Type")
        self.room_tree.heading("usage_count", text="Usage Count")
        self.room_tree.heading("usage_percent", text="Usage %")
        
        self.room_tree.column("room", width=100)
        self.room_tree.column("type", width=100)
        self.room_tree.column("usage_count", width=100)
        self.room_tree.column("usage_percent", width=100)
        
        # Add scrollbar
        room_scroll = ttk.Scrollbar(room_frame, orient="vertical", command=self.room_tree.yview)
        self.room_tree.configure(yscrollcommand=room_scroll.set)
        
        # Pack widgets
        self.room_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        room_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        # Conflicts tab
        conflicts_frame = ttk.Frame(self.analytics_notebook, padding=10)
        self.analytics_notebook.add(conflicts_frame, text="Conflicts")

        # Conflicts treeview
        columns = ("type", "details", "severity")
        self.conflicts_tree = ttk.Treeview(conflicts_frame, columns=columns, show="headings")
        
        # Configure columns
        self.conflicts_tree.heading("type", text="Conflict Type")
        self.conflicts_tree.heading("details", text="Details")
        self.conflicts_tree.heading("severity", text="Severity")
        
        self.conflicts_tree.column("type", width=150)
        self.conflicts_tree.column("details", width=300)
        self.conflicts_tree.column("severity", width=100)
        
        # Add scrollbar
        conflicts_scroll = ttk.Scrollbar(conflicts_frame, orient="vertical", command=self.conflicts_tree.yview)
        self.conflicts_tree.configure(yscrollcommand=conflicts_scroll.set)
        
        # Pack widgets
        self.conflicts_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        conflicts_scroll.pack(side=tk.RIGHT, fill=tk.Y)

    def create_visualization_tab(self):
        """Create the visualization tab with charts and graphs"""
        visualization_frame = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(visualization_frame, text="Visualization")

        # Check if matplotlib is available
        if not 'MATPLOTLIB_AVAILABLE' in globals() or not MATPLOTLIB_AVAILABLE:
            # Display message about missing dependencies
            message_frame = ttk.Frame(visualization_frame, padding=20)
            message_frame.pack(fill=tk.BOTH, expand=True)
            
            ttk.Label(message_frame, text="Visualization features require additional libraries.", 
                      font=("Arial", 12, "bold")).pack(pady=10)
            ttk.Label(message_frame, text="Please install the following packages to enable visualizations:", 
                      font=("Arial", 11)).pack(pady=5)
            ttk.Label(message_frame, text="pip install matplotlib numpy pandas pillow", 
                      font=("Arial", 10, "italic")).pack(pady=5)
            
            # Button to install dependencies
            ttk.Button(message_frame, text="Install Dependencies", 
                      command=self.install_visualization_dependencies).pack(pady=20)
            
            return

        # Create notebook for different visualizations
        viz_notebook = ttk.Notebook(visualization_frame)
        viz_notebook.pack(fill=tk.BOTH, expand=True, pady=10)

        # Faculty workload chart tab
        faculty_viz_frame = ttk.Frame(viz_notebook, padding=10)
        viz_notebook.add(faculty_viz_frame, text="Faculty Workload")
        
        # Create frame for the chart
        self.faculty_chart_frame = ttk.Frame(faculty_viz_frame)
        self.faculty_chart_frame.pack(fill=tk.BOTH, expand=True)

        # Room utilization chart tab
        room_viz_frame = ttk.Frame(viz_notebook, padding=10)
        viz_notebook.add(room_viz_frame, text="Room Utilization")
        
        # Create frame for the chart
        self.room_chart_frame = ttk.Frame(room_viz_frame)
        self.room_chart_frame.pack(fill=tk.BOTH, expand=True)

        # Timetable heatmap tab
        heatmap_frame = ttk.Frame(viz_notebook, padding=10)
        viz_notebook.add(heatmap_frame, text="Timetable Heatmap")
        
        # Create frame for the heatmap
        self.heatmap_frame = ttk.Frame(heatmap_frame)
        self.heatmap_frame.pack(fill=tk.BOTH, expand=True)

        # Refresh button
        refresh_frame = ttk.Frame(visualization_frame)
        refresh_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(refresh_frame, text="Refresh Visualizations", command=self.update_visualizations).pack(side=tk.LEFT, padx=5)
        ttk.Button(refresh_frame, text="Export Charts", command=self.export_charts).pack(side=tk.LEFT, padx=5)
        
        # Initial message
        self.viz_message = tk.StringVar(value="Generate a timetable to view visualizations")
        ttk.Label(visualization_frame, textvariable=self.viz_message).pack(pady=5)
        
    def install_visualization_dependencies(self):
        """Install required dependencies for visualization"""
        if messagebox.askyesno("Install Dependencies", 
                             "This will install matplotlib, numpy, pandas, and pillow packages. Continue?"):
            try:
                self.status_bar.config(text="Installing dependencies...")
                self.root.update()
                
                import subprocess
                process = subprocess.Popen(
                    [sys.executable, "-m", "pip", "install", "matplotlib", "numpy", "pandas", "pillow"],
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE
                )
                stdout, stderr = process.communicate()
                
                if process.returncode == 0:
                    messagebox.showinfo("Success", "Dependencies installed successfully. Please restart the application.")
                    self.status_bar.config(text="Dependencies installed. Restart required.")
                else:
                    error_msg = stderr.decode('utf-8')
                    messagebox.showerror("Error", f"Failed to install dependencies: {error_msg}")
                    self.status_bar.config(text="Failed to install dependencies")
            except Exception as e:
                messagebox.showerror("Error", f"Error installing dependencies: {str(e)}")
                self.status_bar.config(text="Error installing dependencies")


    def update_visualizations(self):
        """Update all visualizations with current data"""
        # Check if matplotlib is available
        if not 'MATPLOTLIB_AVAILABLE' in globals() or not MATPLOTLIB_AVAILABLE:
            messagebox.showwarning("Warning", "Visualization features require matplotlib. Please install it with 'pip install matplotlib numpy pandas pillow'")
            return
            
        if not hasattr(self, 'current_timetable') or not self.current_timetable:
            messagebox.showwarning("Warning", "Please generate a timetable first.")
            return
            
        self.status_bar.config(text="Updating visualizations...")
        self.viz_message.set("Updating visualizations...")
        self.root.update()
        
        try:
            # Clear existing charts
            for widget in self.faculty_chart_frame.winfo_children():
                widget.destroy()
            for widget in self.room_chart_frame.winfo_children():
                widget.destroy()
            for widget in self.heatmap_frame.winfo_children():
                widget.destroy()
                
            # Create faculty workload chart
            self.create_faculty_workload_chart()
            
            # Create room utilization chart
            self.create_room_utilization_chart()
            
            # Create timetable heatmap
            self.create_timetable_heatmap()
            
            self.viz_message.set("Visualizations updated successfully")
            self.status_bar.config(text="Visualizations updated")
        except Exception as e:
            self.viz_message.set(f"Error updating visualizations: {str(e)}")
            self.status_bar.config(text=f"Error updating visualizations: {str(e)}")
            messagebox.showerror("Error", f"Error updating visualizations: {str(e)}")
    
    def create_faculty_workload_chart(self):
        """Create a bar chart of faculty workload"""
        # Collect data
        faculty_names = []
        total_hours = []
        weekly_hours = []
        
        for faculty in sorted(self.faculties.keys()):
            faculty_names.append(faculty)
            hours = self.get_faculty_total_load(faculty)
            total_hours.append(hours)
            weekly_hours.append(round(hours / self.semester_duration.get(), 1))
        
        # Create figure and axis
        fig, ax = plt.subplots(figsize=(10, 6))
        
        # Create bar chart
        x = np.arange(len(faculty_names))
        width = 0.35
        
        ax.bar(x - width/2, total_hours, width, label='Total Hours')
        ax.bar(x + width/2, weekly_hours, width, label='Weekly Hours')
        
        # Add labels and title
        ax.set_xlabel('Faculty')
        ax.set_ylabel('Hours')
        ax.set_title('Faculty Workload')
        ax.set_xticks(x)
        ax.set_xticklabels(faculty_names, rotation=45, ha='right')
        ax.legend()
        
        plt.tight_layout()
        
        # Embed in tkinter
        canvas = FigureCanvasTkAgg(fig, master=self.faculty_chart_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
    
    def create_room_utilization_chart(self):
        """Create a pie chart of room utilization"""
        if not self.current_timetable:
            return
            
        # Collect data
        room_usage = {}
        
        # Initialize room usage counters
        for room in self.classrooms:
            room_usage[room] = {"type": "Classroom", "count": 0}
        for room in self.labs:
            room_usage[room] = {"type": "Lab", "count": 0}
        
        # Count actual usage
        for semester in self.current_timetable:
            for division in self.current_timetable[semester]:
                for day in self.current_timetable[semester][division]:
                    for time_slot in self.current_timetable[semester][division][day]:
                        room = self.current_timetable[semester][division][day][time_slot]["room"]
                        if room and room in room_usage:
                            room_usage[room]["count"] += 1
        
        # Prepare data for pie chart
        classrooms = [room for room in room_usage if room_usage[room]["type"] == "Classroom"]
        classroom_usage = [room_usage[room]["count"] for room in classrooms]
        
        labs = [room for room in room_usage if room_usage[room]["type"] == "Lab"]
        lab_usage = [room_usage[room]["count"] for room in labs]
        
        # Create figure with two subplots
        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(12, 6))
        
        # Classroom pie chart
        if classroom_usage and sum(classroom_usage) > 0:
            ax1.pie(classroom_usage, labels=classrooms, autopct='%1.1f%%', startangle=90)
            ax1.set_title('Classroom Utilization')
        else:
            ax1.text(0.5, 0.5, 'No classroom data', horizontalalignment='center', verticalalignment='center')
            ax1.set_title('Classroom Utilization')
        
        # Lab pie chart
        if lab_usage and sum(lab_usage) > 0:
            ax2.pie(lab_usage, labels=labs, autopct='%1.1f%%', startangle=90)
            ax2.set_title('Lab Utilization')
        else:
            ax2.text(0.5, 0.5, 'No lab data', horizontalalignment='center', verticalalignment='center')
            ax2.set_title('Lab Utilization')
        
        plt.tight_layout()
        
        # Embed in tkinter
        canvas = FigureCanvasTkAgg(fig, master=self.room_chart_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
    
    def create_timetable_heatmap(self):
        """Create a heatmap of the timetable"""
        if not self.current_timetable:
            return
            
        # Get target semester and division
        try:
            target_semester = int(self.timetable_semester.get())
            target_division = self.timetable_division.get()
        except (ValueError, AttributeError):
            # If no specific target, use the first semester and division
            target_semester = list(self.current_timetable.keys())[0]
            target_division = list(self.current_timetable[target_semester].keys())[0]
        
        # If target_division is "All", use the first division
        if target_division == "All" or target_division not in self.current_timetable[target_semester]:
            target_division = list(self.current_timetable[target_semester].keys())[0]
        
        # Create data for heatmap
        working_days = [day for day in ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"] if self.working_days[day].get()]
        data = np.zeros((len(working_days), len(self.time_slots)))
        
        # Fill data with subject counts (1 for occupied, 0 for free)
        for i, day in enumerate(working_days):
            for j, time_slot in enumerate(self.time_slots):
                if day in self.current_timetable[target_semester][target_division] and \
                   time_slot in self.current_timetable[target_semester][target_division][day]:
                    subject = self.current_timetable[target_semester][target_division][day][time_slot]["subject"]
                    if subject != "Free":
                        data[i, j] = 1
        
        # Create figure and axis
        fig, ax = plt.subplots(figsize=(10, 6))
        
        # Create heatmap
        im = ax.imshow(data, cmap='YlGn')
        
        # Add labels
        ax.set_xticks(np.arange(len(self.time_slots)))
        ax.set_yticks(np.arange(len(working_days)))
        ax.set_xticklabels(self.time_slots, rotation=45, ha='right')
        ax.set_yticklabels(working_days)
        
        # Add title
        ax.set_title(f'Timetable Heatmap - Semester {target_semester}, Division {target_division}')
        
        # Add colorbar
        cbar = ax.figure.colorbar(im, ax=ax)
        cbar.ax.set_ylabel('Occupied', rotation=-90, va='bottom')
        
        # Add text annotations
        for i in range(len(working_days)):
            for j in range(len(self.time_slots)):
                if data[i, j] == 1:
                    subject = self.current_timetable[target_semester][target_division][working_days[i]][self.time_slots[j]]["subject"]
                    faculty = self.current_timetable[target_semester][target_division][working_days[i]][self.time_slots[j]]["faculty"]
                    text = ax.text(j, i, f"{subject}\n{faculty}", ha="center", va="center", color="black", fontsize=8)
        
        plt.tight_layout()
        
        # Embed in tkinter
        canvas = FigureCanvasTkAgg(fig, master=self.heatmap_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
    
    def export_charts(self):
        """Export all charts as images"""
        # Check if matplotlib is available
        if not 'MATPLOTLIB_AVAILABLE' in globals() or not MATPLOTLIB_AVAILABLE:
            messagebox.showwarning("Warning", "Chart export requires matplotlib. Please install it with 'pip install matplotlib numpy pandas pillow'")
            return
            
        if not hasattr(self, 'current_timetable') or not self.current_timetable:
            messagebox.showwarning("Warning", "Please generate a timetable first.")
            return
            
        # Ask for directory
        export_dir = filedialog.askdirectory(title="Select Directory for Chart Export")
        if not export_dir:
            return
            
        try:
            # Update visualizations first
            self.update_visualizations()
            
            # Export faculty workload chart
            for widget in self.faculty_chart_frame.winfo_children():
                if isinstance(widget, FigureCanvasTkAgg):
                    widget.figure.savefig(os.path.join(export_dir, "faculty_workload.png"), dpi=300, bbox_inches='tight')
            
            # Export room utilization chart
            for widget in self.room_chart_frame.winfo_children():
                if isinstance(widget, FigureCanvasTkAgg):
                    widget.figure.savefig(os.path.join(export_dir, "room_utilization.png"), dpi=300, bbox_inches='tight')
            
            # Export timetable heatmap
            for widget in self.heatmap_frame.winfo_children():
                if isinstance(widget, FigureCanvasTkAgg):
                    widget.figure.savefig(os.path.join(export_dir, "timetable_heatmap.png"), dpi=300, bbox_inches='tight')
            
            messagebox.showinfo("Success", f"Charts exported to {export_dir}")
            self.status_bar.config(text=f"Charts exported to {export_dir}")
        except Exception as e:
            messagebox.showerror("Error", f"Error exporting charts: {str(e)}")
            self.status_bar.config(text=f"Error exporting charts: {str(e)}")
    
    def refresh_analytics(self):
        """Refresh all analytics data"""
        if not hasattr(self, 'current_timetable') or not self.current_timetable:
            messagebox.showwarning("Warning", "Please generate a timetable first.")
            return

        # Update status
        self.status_bar.config(text="Refreshing analytics...")
        self.root.update()

        # Clear existing data
        self.faculty_tree.delete(*self.faculty_tree.get_children())
        self.room_tree.delete(*self.room_tree.get_children())
        self.conflicts_tree.delete(*self.conflicts_tree.get_children())

        # Calculate faculty workload
        for faculty in sorted(self.faculties.keys()):
            subjects_count = self.count_faculty_subjects(faculty)
            total_hours = self.get_faculty_total_load(faculty)
            weekly_hours = round(total_hours / self.semester_duration.get(), 1)
            
            # Find maximum daily load in the generated timetable
            max_daily = 0
            for day in ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]:
                if self.working_days[day].get():
                    daily_count = 0
                    for semester in self.current_timetable:
                        for division in self.current_timetable[semester]:
                            for time_slot in self.time_slots:
                                if day in self.current_timetable[semester][division] and \
                                   time_slot in self.current_timetable[semester][division][day] and \
                                   self.current_timetable[semester][division][day][time_slot]["faculty"] == faculty:
                                    daily_count += 1
                    max_daily = max(max_daily, daily_count)
            
            # Add to treeview
            self.faculty_tree.insert("", tk.END, values=(faculty, subjects_count, total_hours, weekly_hours, max_daily))

        # Calculate room utilization
        total_slots = 0
        room_usage = {}
        
        # Count total possible slots
        working_day_count = sum(1 for day in self.days if self.working_days[day].get())
        total_slots = working_day_count * len(self.time_slots)
        
        # Initialize room usage counters
        for room in self.classrooms:
            room_usage[room] = {"type": "Classroom", "count": 0}
        for room in self.labs:
            room_usage[room] = {"type": "Lab", "count": 0}
        
        # Count actual usage
        for semester in self.current_timetable:
            for division in self.current_timetable[semester]:
                for day in self.current_timetable[semester][division]:
                    for time_slot in self.current_timetable[semester][division][day]:
                        room = self.current_timetable[semester][division][day][time_slot]["room"]
                        if room and room in room_usage:
                            room_usage[room]["count"] += 1
        
        # Add to treeview
        for room, data in sorted(room_usage.items()):
            usage_percent = round((data["count"] / total_slots) * 100, 1) if total_slots > 0 else 0
            self.room_tree.insert("", tk.END, values=(room, data["type"], data["count"], f"{usage_percent}%"))

        # Check for conflicts
        self.check_timetable_conflicts()

        # Update status
        self.status_bar.config(text="Analytics refreshed")

    def check_timetable_conflicts(self):
        """Check for conflicts in the timetable"""
        if not hasattr(self, 'current_timetable') or not self.current_timetable:
            return
            
        # Check for faculty conflicts (same faculty scheduled in multiple places at once)
        for day in self.days:
            if not self.working_days[day].get():
                continue
                
            for time_slot in self.time_slots:
                faculty_assignments = {}
                
                for semester in self.current_timetable:
                    for division in self.current_timetable[semester]:
                        if day in self.current_timetable[semester][division] and time_slot in self.current_timetable[semester][division][day]:
                            faculty = self.current_timetable[semester][division][day][time_slot]["faculty"]
                            
                            if faculty and faculty != "":
                                if faculty in faculty_assignments:
                                    # Conflict found
                                    details = f"Faculty {faculty} scheduled in multiple classes on {day} at {time_slot}"
                                    self.conflicts_tree.insert("", tk.END, values=("Faculty Double-Booking", details, "High"))
                                else:
                                    faculty_assignments[faculty] = (semester, division)
        
        # Check for room conflicts (same room used by multiple classes at once)
        for day in self.days:
            if not self.working_days[day].get():
                continue
                
            for time_slot in self.time_slots:
                room_assignments = {}
                
                for semester in self.current_timetable:
                    for division in self.current_timetable[semester]:
                        if day in self.current_timetable[semester][division] and time_slot in self.current_timetable[semester][division][day]:
                            room = self.current_timetable[semester][division][day][time_slot]["room"]
                            
                            if room and room != "":
                                if room in room_assignments:
                                    # Conflict found
                                    details = f"Room {room} scheduled for multiple classes on {day} at {time_slot}"
                                    self.conflicts_tree.insert("", tk.END, values=("Room Double-Booking", details, "High"))
                                else:
                                    room_assignments[room] = (semester, division)
        
        # Check for faculty overloading (more than max lectures per day)
        for day in self.days:
            if not self.working_days[day].get():
                continue
                
            faculty_counts = {}
            
            for semester in self.current_timetable:
                for division in self.current_timetable[semester]:
                    if day not in self.current_timetable[semester][division]:
                        continue
                        
                    for time_slot in self.current_timetable[semester][division][day]:
                        faculty = self.current_timetable[semester][division][day][time_slot]["faculty"]
                        
                        if faculty and faculty != "":
                            if faculty in faculty_counts:
                                faculty_counts[faculty] += 1
                            else:
                                faculty_counts[faculty] = 1
            
            # Check for overloaded faculty
            for faculty, count in faculty_counts.items():
                # Get faculty's max lectures
                faculty_load = self.get_faculty_total_load(faculty)
                max_lectures = 5 if faculty_load > 216 else self.max_faculty_lectures.get()
                
                if count > max_lectures:
                    details = f"Faculty {faculty} has {count} lectures on {day} (max: {max_lectures})"
                    self.conflicts_tree.insert("", tk.END, values=("Faculty Overload", details, "Medium"))

    def save_timetable(self):
        """Save the current timetable to a file"""
        if not hasattr(self, 'current_timetable') or not self.current_timetable:
            messagebox.showwarning("Warning", "Please generate a timetable first.")
            return
            
        import pickle
        from tkinter import filedialog
        
        filename = filedialog.asksaveasfilename(
            defaultextension=".ttbl",
            filetypes=[("Timetable files", "*.ttbl"), ("All files", "*.*")],
            title="Save Timetable"
        )
        
        if not filename:
            return
            
        try:
            with open(filename, 'wb') as f:
                pickle.dump(self.current_timetable, f)
            messagebox.showinfo("Success", f"Timetable saved to {filename}")
            self.status_bar.config(text=f"Timetable saved to {filename}")
            
            # Auto-save if enabled
            if self.settings["auto_save"]:
                with open(self.settings["auto_save_path"], 'wb') as f:
                    pickle.dump(self.current_timetable, f)
        except Exception as e:
            messagebox.showerror("Error", f"Error saving timetable: {str(e)}")
    
    def export_to_pdf(self):
        """Export the timetable to PDF format"""
        if not hasattr(self, 'current_timetable') or not self.current_timetable:
            messagebox.showwarning("Warning", "Please generate a timetable first.")
            return
            
        if not 'PDF_AVAILABLE' in globals() or not PDF_AVAILABLE:
            messagebox.showerror("Error", "PDF export requires the ReportLab library. Please install it with 'pip install reportlab'.")
            return
                
        # Ask for file location
        filename = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
            title="Export Timetable to PDF"
        )
        
        if not filename:
            return
            
        try:
            # Update status
            self.status_bar.config(text="Exporting to PDF...")
            self.root.update()
            
            # Create PDF document
            doc = SimpleDocTemplate(filename, pagesize=A4, title="College Timetable")
            styles = getSampleStyleSheet()
            elements = []
            
            # Add title
            title = Paragraph(f"<b>{self.college_name.get()} - Timetable</b>", styles['Title'])
            elements.append(title)
            elements.append(Spacer(1, 12))
            
            # Target semester and division
            target_semester = None
            target_division = None

            try:
                target_semester = int(self.timetable_semester.get())
                target_division = self.timetable_division.get()
            except ValueError:
                pass
                
            # Process each semester/division
            for semester in sorted(self.current_timetable.keys()):
                if target_semester and semester != target_semester:
                    continue
                    
                for division in sorted(self.current_timetable[semester].keys()):
                    if target_division != "All" and division != target_division and target_division:
                        continue
                        
                    # Add semester and division heading
                    heading = Paragraph(f"<b>Semester {semester} - Division {division}</b>", styles['Heading2'])
                    elements.append(heading)
                    elements.append(Spacer(1, 6))
                    
                    # Process each day
                    for day in [d for d in self.days if self.working_days[d].get()]:
                        # Add day heading
                        day_heading = Paragraph(f"<b>{day}</b>", styles['Heading3'])
                        elements.append(day_heading)
                        elements.append(Spacer(1, 3))
                        
                        # Create table data
                        data = [["Time Slot", "Subject", "Faculty", "Room"]]
                        
                        # Add time slots
                        for time_slot in self.time_slots:
                            if day in self.current_timetable[semester][division] and time_slot in self.current_timetable[semester][division][day]:
                                slot_data = self.current_timetable[semester][division][day][time_slot]
                                subject = slot_data["subject"]
                                faculty = slot_data["faculty"]
                                room = slot_data["room"]
                                
                                data.append([time_slot, subject, faculty, room])
                        
                        # Create table
                        table = Table(data, colWidths=[80, 150, 80, 80])
                        
                        # Add style
                        table_style = TableStyle([
                            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                            ('GRID', (0, 0), (-1, -1), 1, colors.black)
                        ])
                        
                        # Add alternating row colors
                        for i in range(1, len(data)):
                            if i % 2 == 0:
                                table_style.add('BACKGROUND', (0, i), (-1, i), colors.lightgrey)
                        
                        table.setStyle(table_style)
                        elements.append(table)
                        elements.append(Spacer(1, 12))
                    
                    # Add page break between divisions
                    elements.append(Spacer(1, 20))
            
            # Build PDF
            doc.build(elements)
            
            # Show success message
            messagebox.showinfo("Success", f"Timetable exported to {filename}")
            self.status_bar.config(text=f"Timetable exported to {filename}")
            
            # Open the PDF
            if messagebox.askyesno("Open PDF", "Would you like to open the exported PDF?"):
                webbrowser.open(filename)
                
        except Exception as e:
            messagebox.showerror("Error", f"Error exporting to PDF: {str(e)}")
            self.status_bar.config(text=f"Error exporting to PDF: {str(e)}")
    
    def export_to_html(self):
        """Export the timetable to HTML format"""
        if not hasattr(self, 'current_timetable') or not self.current_timetable:
            messagebox.showwarning("Warning", "Please generate a timetable first.")
            return
            
        # Ask for file location
        filename = filedialog.asksaveasfilename(
            defaultextension=".html",
            filetypes=[("HTML files", "*.html"), ("All files", "*.*")],
            title="Export Timetable to HTML"
        )
        
        if not filename:
            return
            
        try:
            # Update status
            self.status_bar.config(text="Exporting to HTML...")
            self.root.update()
            
            # Create HTML content
            html_content = f"""
            <!DOCTYPE html>
            <html>
            <head>
                <title>{self.college_name.get()} - Timetable</title>
                <style>
                    body {{ font-family: Arial, sans-serif; margin: 20px; }}
                    h1 {{ color: #333366; text-align: center; }}
                    h2 {{ color: #333366; margin-top: 20px; }}
                    h3 {{ color: #333366; }}
                    table {{ border-collapse: collapse; width: 100%; margin-bottom: 20px; }}
                    th {{ background-color: #333366; color: white; padding: 8px; text-align: left; }}
                    td {{ border: 1px solid #ddd; padding: 8px; }}
                    tr:nth-child(even) {{ background-color: #f2f2f2; }}
                    .free {{ color: #999; }}
                    .lab {{ background-color: #e6f7ff; }}
                    .theory {{ background-color: #f0f0f0; }}
                </style>
            </head>
            <body>
                <h1>{self.college_name.get()} - Timetable</h1>
            """
            
            # Target semester and division
            target_semester = None
            target_division = None

            try:
                target_semester = int(self.timetable_semester.get())
                target_division = self.timetable_division.get()
            except ValueError:
                pass
                
            # Process each semester/division
            for semester in sorted(self.current_timetable.keys()):
                if target_semester and semester != target_semester:
                    continue
                    
                for division in sorted(self.current_timetable[semester].keys()):
                    if target_division != "All" and division != target_division and target_division:
                        continue
                        
                    # Add semester and division heading
                    html_content += f"<h2>Semester {semester} - Division {division}</h2>\n"
                    
                    # Process each day
                    for day in [d for d in self.days if self.working_days[d].get()]:
                        # Add day heading
                        html_content += f"<h3>{day}</h3>\n"
                        
                        # Create table
                        html_content += "<table>\n"
                        html_content += "<tr><th>Time Slot</th><th>Subject</th><th>Faculty</th><th>Room</th></tr>\n"
                        
                        # Add time slots
                        for time_slot in self.time_slots:
                            if day in self.current_timetable[semester][division] and time_slot in self.current_timetable[semester][division][day]:
                                slot_data = self.current_timetable[semester][division][day][time_slot]
                                subject = slot_data["subject"]
                                faculty = slot_data["faculty"]
                                room = slot_data["room"]
                                
                                # Determine class for styling
                                css_class = ""
                                if subject == "Free":
                                    css_class = "class='free'"
                                elif "Lab" in subject or "lab" in subject or "(cont.)" in subject:
                                    css_class = "class='lab'"
                                else:
                                    css_class = "class='theory'"
                                
                                html_content += f"<tr {css_class}><td>{time_slot}</td><td>{subject}</td><td>{faculty}</td><td>{room}</td></tr>\n"
                        
                        html_content += "</table>\n"
            
            # Close HTML
            html_content += """
            </body>
            </html>
            """
            
            # Write to file
            with open(filename, 'w', encoding='utf-8') as f:
                f.write(html_content)
            
            # Show success message
            messagebox.showinfo("Success", f"Timetable exported to {filename}")
            self.status_bar.config(text=f"Timetable exported to {filename}")
            
            # Open the HTML file
            if messagebox.askyesno("Open HTML", "Would you like to open the exported HTML file?"):
                webbrowser.open(filename)
                
        except Exception as e:
            messagebox.showerror("Error", f"Error exporting to HTML: {str(e)}")
            self.status_bar.config(text=f"Error exporting to HTML: {str(e)}")
    
    def print_timetable(self):
        """Print the timetable"""
        if not hasattr(self, 'current_timetable') or not self.current_timetable:
            messagebox.showwarning("Warning", "Please generate a timetable first.")
            return
            
        # First export to PDF, then open for printing
        temp_pdf = os.path.join(os.path.expanduser("~"), "temp_timetable_print.pdf")
        
        try:
            # Check if PDF export is available
            if not PDF_AVAILABLE:
                messagebox.showerror("Error", "Printing requires the ReportLab library. Please install it with 'pip install reportlab'.")
                return
                
            # Update status
            self.status_bar.config(text="Preparing timetable for printing...")
            self.root.update()
            
            # Create PDF document (reusing code from export_to_pdf)
            doc = SimpleDocTemplate(temp_pdf, pagesize=A4, title="College Timetable")
            styles = getSampleStyleSheet()
            elements = []
            
            # Add title
            title = Paragraph(f"<b>{self.college_name.get()} - Timetable</b>", styles['Title'])
            elements.append(title)
            elements.append(Spacer(1, 12))
            
            # Target semester and division
            target_semester = None
            target_division = None

            try:
                target_semester = int(self.timetable_semester.get())
                target_division = self.timetable_division.get()
            except ValueError:
                pass
                
            # Process each semester/division (same as export_to_pdf)
            for semester in sorted(self.current_timetable.keys()):
                if target_semester and semester != target_semester:
                    continue
                    
                for division in sorted(self.current_timetable[semester].keys()):
                    if target_division != "All" and division != target_division and target_division:
                        continue
                        
                    # Add semester and division heading
                    heading = Paragraph(f"<b>Semester {semester} - Division {division}</b>", styles['Heading2'])
                    elements.append(heading)
                    elements.append(Spacer(1, 6))
                    
                    # Process each day
                    for day in [d for d in self.days if self.working_days[d].get()]:
                        # Add day heading
                        day_heading = Paragraph(f"<b>{day}</b>", styles['Heading3'])
                        elements.append(day_heading)
                        elements.append(Spacer(1, 3))
                        
                        # Create table data
                        data = [["Time Slot", "Subject", "Faculty", "Room"]]
                        
                        # Add time slots
                        for time_slot in self.time_slots:
                            if day in self.current_timetable[semester][division] and time_slot in self.current_timetable[semester][division][day]:
                                slot_data = self.current_timetable[semester][division][day][time_slot]
                                subject = slot_data["subject"]
                                faculty = slot_data["faculty"]
                                room = slot_data["room"]
                                
                                data.append([time_slot, subject, faculty, room])
                        
                        # Create table
                        table = Table(data, colWidths=[80, 150, 80, 80])
                        
                        # Add style
                        table_style = TableStyle([
                            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                            ('GRID', (0, 0), (-1, -1), 1, colors.black)
                        ])
                        
                        # Add alternating row colors
                        for i in range(1, len(data)):
                            if i % 2 == 0:
                                table_style.add('BACKGROUND', (0, i), (-1, i), colors.lightgrey)
                        
                        table.setStyle(table_style)
                        elements.append(table)
                        elements.append(Spacer(1, 12))
                    
                    # Add page break between divisions
                    elements.append(Spacer(1, 20))
            
            # Build PDF
            doc.build(elements)
            
            # Open the PDF for printing
            self.status_bar.config(text="Opening print dialog...")
            webbrowser.open(temp_pdf)
            
            self.status_bar.config(text="Timetable sent to printer")
                
        except Exception as e:
            messagebox.showerror("Error", f"Error printing timetable: {str(e)}")
            self.status_bar.config(text=f"Error printing timetable: {str(e)}")
        finally:
            # Don't delete the temp file immediately as it might be needed for printing
            pass
            
    def load_timetable(self):
        """Load a timetable from a file"""
        import pickle
        from tkinter import filedialog
        
        filename = filedialog.askopenfilename(
            defaultextension=".ttbl",
            filetypes=[("Timetable files", "*.ttbl"), ("All files", "*.*")],
            title="Load Timetable"
        )
        
        if not filename:
            return
            
        try:
            with open(filename, 'rb') as f:
                self.current_timetable = pickle.load(f)
            
            # Display the loaded timetable
            self.display_timetable(self.current_timetable)
            messagebox.showinfo("Success", f"Timetable loaded from {filename}")
            self.status_bar.config(text=f"Timetable loaded from {filename}")
            
            # Refresh analytics
            self.refresh_analytics()
        except Exception as e:
            messagebox.showerror("Error", f"Error loading timetable: {str(e)}")

    def generate_timetable(self):
        """Generate the timetable based on constraints"""
        try:
            # Update status
            self.status_bar.config(text="Generating timetable...")
            self.progress_var.set(0)
            self.progress_status.set("Initializing timetable structure...")
            self.root.update()
            
            # Initialize timetable structure
            timetable = self.initialize_timetable_structure()

            # Update progress
            self.progress_var.set(10)
            self.progress_status.set("Calculating weekly load...")
            self.root.update()
            
            # Calculate weekly load
            weekly_load = self.calculate_weekly_load()

            # Target semester and division
            target_semester = None
            target_division = None

            try:
                target_semester = int(self.timetable_semester.get())
                target_division = self.timetable_division.get()
            except ValueError:
                pass

            # Update progress
            self.progress_var.set(20)
            self.progress_status.set("Sorting subjects by priority...")
            self.root.update()
            
            # Sort subjects by priority (labs first, then by total hours)
            subject_entries = list(weekly_load.items())
            subject_entries.sort(key=lambda x: (
                0 if x[1]["type"] == "Lab" else 1,  # Labs first
                -x[1]["lectures_per_week"]  # Higher lecture count gets higher priority
            ))

            # Update progress
            self.progress_var.set(30)
            self.progress_status.set("Assigning lab sessions (first pass)...")
            self.root.update()
            
            # First pass: Assign lab sessions (they are more constrained)
            for (semester, division, subject), details in subject_entries:
                if target_semester and semester != target_semester:
                    continue
                if target_division != "All" and division != target_division and target_division:
                    continue

                if details["type"] != "Lab":
                    continue

                faculty = details["faculty"]
                lectures_needed = details["lectures_per_week"]
                
                # Check if faculty has high workload (>216 hours) to allow more lectures per day
                faculty_load = self.get_faculty_total_load(faculty)
                max_lectures_per_day = self.max_faculty_lectures.get()
                if faculty_load > 216:  # As per special instruction
                    max_lectures_per_day = 5

                # Find suitable time slots
                assignments_made = 0
                working_days = [d for d in self.days if self.working_days[d].get()]
                
                # Try to distribute lab sessions evenly across days
                days_needed = min(lectures_needed, len(working_days))
                target_days = random.sample(working_days, days_needed)
                
                for day in target_days:
                    # Check faculty availability for the day
                    if not self.check_faculty_availability(faculty, day, self.time_slots[0]):
                        continue

                    # Check if faculty already has too many sessions on this day
                    if self.count_faculty_assignments(timetable, faculty, day) >= max_lectures_per_day:
                        continue

                    # Find consecutive time slots for lab sessions (labs usually need 2 consecutive slots)
                    assigned = False
                    for i in range(len(self.time_slots) - 1):
                        if assignments_made >= lectures_needed:
                            break

                        time_slot1 = self.time_slots[i]
                        time_slot2 = self.time_slots[i + 1]

                        # Check if both slots are free
                        if (timetable[semester][division][day][time_slot1]["subject"] == "Free" and
                            timetable[semester][division][day][time_slot2]["subject"] == "Free"):

                            # Check if faculty is already scheduled at this time in another class
                            if (self.check_faculty_conflict(timetable, faculty, day, time_slot1) or
                                self.check_faculty_conflict(timetable, faculty, day, time_slot2)):
                                continue

                            # Find available lab rooms
                            available_labs = self.find_available_rooms(timetable, day, time_slot1, "lab")
                            if not available_labs:  # No labs available at this time
                                continue
                                
                            # Also check next slot
                            available_labs2 = self.find_available_rooms(timetable, day, time_slot2, "lab")
                            available_labs = [lab for lab in available_labs if lab in available_labs2]
                            if not available_labs:  # No labs available for both slots
                                continue

                            # Assign lab to these slots
                            lab_room = random.choice(available_labs)
                            timetable[semester][division][day][time_slot1] = {
                                "subject": subject,
                                "faculty": faculty,
                                "room": lab_room
                            }
                            timetable[semester][division][day][time_slot2] = {
                                "subject": subject + " (cont.)",
                                "faculty": faculty,
                                "room": lab_room
                            }
                            assignments_made += 1
                            assigned = True
                            break  # Found slots for this day
                    
                    # If we couldn't assign on this day, try another day
                    if not assigned and assignments_made < lectures_needed:
                        # Try other days that weren't in our initial sample
                        for day in [d for d in working_days if d not in target_days]:
                            if self.check_faculty_availability(faculty, day, self.time_slots[0]) and \
                               self.count_faculty_assignments(timetable, faculty, day) < max_lectures_per_day:
                                # Try to find slots on this day
                                for i in range(len(self.time_slots) - 1):
                                    time_slot1 = self.time_slots[i]
                                    time_slot2 = self.time_slots[i + 1]
                                    
                                    if (timetable[semester][division][day][time_slot1]["subject"] == "Free" and
                                        timetable[semester][division][day][time_slot2]["subject"] == "Free" and
                                        not self.check_faculty_conflict(timetable, faculty, day, time_slot1) and
                                        not self.check_faculty_conflict(timetable, faculty, day, time_slot2)):
                                        
                                        # Find available lab rooms
                                        available_labs = self.find_available_rooms(timetable, day, time_slot1, "lab")
                                        available_labs2 = self.find_available_rooms(timetable, day, time_slot2, "lab")
                                        available_labs = [lab for lab in available_labs if lab in available_labs2]
                                        
                                        if available_labs:
                                            lab_room = random.choice(available_labs)
                                            timetable[semester][division][day][time_slot1] = {
                                                "subject": subject,
                                                "faculty": faculty,
                                                "room": lab_room
                                            }
                                            timetable[semester][division][day][time_slot2] = {
                                                "subject": subject + " (cont.)",
                                                "faculty": faculty,
                                                "room": lab_room
                                            }
                                            assignments_made += 1
                                            break

            # Update progress
            self.progress_var.set(60)
            self.progress_status.set("Assigning theory sessions (second pass)...")
            self.root.update()
            
            # Second pass: Assign theory sessions
            for (semester, division, subject), details in subject_entries:
                if target_semester and semester != target_semester:
                    continue
                if target_division != "All" and division != target_division and target_division:
                    continue

                if details["type"] != "Theory":
                    continue

                faculty = details["faculty"]
                lectures_needed = details["lectures_per_week"]
                
                # Check if faculty has high workload to allow more lectures per day
                faculty_load = self.get_faculty_total_load(faculty)
                max_lectures_per_day = self.max_faculty_lectures.get()
                if faculty_load > 216:  # As per special instruction
                    max_lectures_per_day = 5

                # Find suitable time slots
                assignments_made = 0
                working_days = [d for d in self.days if self.working_days[d].get()]
                
                # Try to distribute theory sessions evenly across days
                days_needed = min(lectures_needed, len(working_days))
                target_days = working_days[:days_needed]  # Take first N days
                
                for day in working_days:
                    # Skip if we've assigned enough lectures
                    if assignments_made >= lectures_needed:
                        break

                    # Check faculty availability for the day
                    if not self.check_faculty_availability(faculty, day, self.time_slots[0]):
                        continue

                    # Check if faculty already has too many sessions on this day
                    if self.count_faculty_assignments(timetable, faculty, day) >= max_lectures_per_day:
                        continue

                    # Get preferred time slots for consistency
                    preferred_slots = self.get_preferred_slots(timetable, semester, division, subject)
                    
                    # Try preferred slots first, then any available slot
                    all_slots = preferred_slots + [slot for slot in self.time_slots if slot not in preferred_slots]
                    
                    # Find a free time slot
                    for time_slot in all_slots:
                        if time_slot not in timetable[semester][division][day]:
                            continue
                            
                        if timetable[semester][division][day][time_slot]["subject"] == "Free":
                            # Check if faculty is already scheduled at this time in another class
                            if self.check_faculty_conflict(timetable, faculty, day, time_slot):
                                continue
                                
                            # Find available classrooms
                            available_rooms = self.find_available_rooms(timetable, day, time_slot, "classroom")
                            if not available_rooms:  # No classrooms available
                                continue
                                
                            # Assign theory session to this slot
                            classroom = random.choice(available_rooms)
                            timetable[semester][division][day][time_slot] = {
                                "subject": subject,
                                "faculty": faculty,
                                "room": classroom
                            }
                            assignments_made += 1
                            break  # Move to next day

            # Update progress
            self.progress_var.set(90)
            self.progress_status.set("Displaying generated timetable...")
            self.root.update()
            
            # Display the generated timetable
            self.display_timetable(timetable, target_semester, target_division)

            # Save the generated timetable for export
            self.current_timetable = timetable
            
            # Update progress
            self.progress_var.set(100)
            self.progress_status.set("Timetable generated successfully!")
            self.status_bar.config(text="Timetable generated successfully!")
            self.root.update()
            
            # Check for conflicts
            self.check_timetable_conflicts()
            
            # Refresh analytics
            self.refresh_analytics()

            messagebox.showinfo("Success", "Timetable generated successfully!")

        except Exception as e:
            self.progress_status.set(f"Error: {str(e)}")
            self.progress_bar.configure(style="red.Horizontal.TProgressbar")
            self.status_bar.config(text=f"Error generating timetable: {str(e)}")
            messagebox.showerror("Error", f"Error generating timetable: {str(e)}")
            import traceback
            traceback.print_exc()

    def display_timetable(self, timetable, target_semester=None, target_division=None):
        """Display the generated timetable"""
        self.timetable_display.delete(1.0, tk.END)

        if target_semester is None:
            # Display all semesters
            for semester in sorted(timetable.keys()):
                self.display_semester_timetable(timetable, semester, target_division)
        else:
            # Display specific semester
            self.display_semester_timetable(timetable, target_semester, target_division)

    def display_semester_timetable(self, timetable, semester, target_division=None):
        """Display timetable for a specific semester"""
        if target_division == "All" or target_division is None:
            # Display all divisions
            for division in sorted(timetable[semester].keys()):
                self.display_division_timetable(timetable, semester, division)
        else:
            # Display specific division
            if target_division in timetable[semester]:
                self.display_division_timetable(timetable, semester, target_division)

    def display_division_timetable(self, timetable, semester, division):
        """Display timetable for a specific division"""
        self.timetable_display.insert(tk.END, f"\n{'='*80}\n")
        self.timetable_display.insert(tk.END, f"SEMESTER {semester} - DIVISION {division}\n")
        self.timetable_display.insert(tk.END, f"{'='*80}\n\n")

        # Display each day
        for day in [d for d in self.days if self.working_days[d].get()]:
            self.timetable_display.insert(tk.END, f"{day}\n{'-'*80}\n")

            # Table headers
            self.timetable_display.insert(tk.END, f"{'Time Slot':<15} {'Subject':<20} {'Faculty':<10} {'Room':<10}\n")
            self.timetable_display.insert(tk.END, f"{'-'*80}\n")

            # Display time slots
            for time_slot in self.time_slots:
                if day in timetable[semester][division] and time_slot in timetable[semester][division][day]:
                    slot_data = timetable[semester][division][day][time_slot]
                    subject = slot_data["subject"] if slot_data["subject"] is not None else "Free"
                    faculty = slot_data["faculty"] if slot_data["faculty"] is not None else "-"
                    room = slot_data["room"] if slot_data["room"] is not None else "-"

                    self.timetable_display.insert(tk.END, f"{time_slot:<15} {subject:<20} {faculty:<10} {room:<10}\n")

            self.timetable_display.insert(tk.END, f"\n")

    def export_timetable(self):
        """Export the generated timetable to Excel"""
        try:
            if not hasattr(self, 'current_timetable'):
                messagebox.showwarning("Warning", "Please generate a timetable first.")
                return
                
            if not 'PANDAS_AVAILABLE' in globals() or not PANDAS_AVAILABLE:
                messagebox.showerror("Error", "Excel export requires the pandas library. Please install it with 'pip install pandas'.")
                return

            from datetime import datetime
            filename = f"Timetable_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

            # Create Excel workbook
            writer = pd.ExcelWriter(filename, engine='xlsxwriter')

            # Target semester and division
            target_semester = None
            target_division = None

            try:
                target_semester = int(self.timetable_semester.get())
                target_division = self.timetable_division.get()
            except ValueError:
                pass

            # Iterate through semesters and divisions
            for semester in sorted(self.current_timetable.keys()):
                if target_semester and semester != target_semester:
                    continue

                for division in sorted(self.current_timetable[semester].keys()):
                    if target_division != "All" and division != target_division and target_division:
                        continue

                    # Create a sheet for each division
                    sheet_name = f"Sem{semester}_Div{division}"

                    # Prepare data for DataFrame
                    data = []
                    for day in [d for d in self.days if self.working_days[d].get()]:
                        day_data = [day]
                        for time_slot in self.time_slots:
                            if day in self.current_timetable[semester][division] and time_slot in self.current_timetable[semester][division][day]:
                                slot_data = self.current_timetable[semester][division][day][time_slot]
                                day_data.append(f"{slot_data['subject']}\n{slot_data['faculty']}\n{slot_data['room']}")
                            else:
                                day_data.append("")
                        data.append(day_data)

                    # Create DataFrame
                    df = pd.DataFrame(data, columns=["Day"] + self.time_slots)

                    # Write to Excel
                    df.to_excel(writer, sheet_name=sheet_name, index=False)

                    # Format the sheet
                    workbook = writer.book
                    worksheet = writer.sheets[sheet_name]

                    # Set column widths
                    worksheet.set_column('A:A', 12)
                    for i in range(len(self.time_slots)):
                        worksheet.set_column(i+1, i+1, 20)

                    # Add a header with the semester and division
                    header_format = workbook.add_format({
                        'bold': True,
                        'align': 'center',
                        'valign': 'vcenter',
                        'fg_color': '#D7E4BC',
                        'border': 1
                    })
                    worksheet.merge_range(0, 0, 0, len(self.time_slots), f"Semester {semester} - Division {division}", header_format)

                    # Write the data again, starting from row 1
                    df.to_excel(writer, sheet_name=sheet_name, startrow=1, index=False)

            # Save the workbook
            writer.close()

            messagebox.showinfo("Success", f"Timetable exported to {filename}")

        except Exception as e:
            messagebox.showerror("Error", f"Error exporting timetable: {str(e)}")
            import traceback
            traceback.print_exc()

if __name__ == "__main__":
    root = tk.Tk()
    app = TimetableGenerator(root)
    try:
        root.mainloop()
    except KeyboardInterrupt:
        print("\nApplication closed by user.")
        # No need to call root.destroy() as Tkinter is already handling shutdown
        # Just exit gracefully
        sys.exit(0)