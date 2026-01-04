import tkinter as tk
from tkinter import ttk, messagebox
import sys
import os
import shutil
from pathlib import Path
import importlib.util
import customtkinter as ctk
import traceback

class ReportProcessorHomepage:
    def __init__(self):
        # Use CustomTkinter for consistency with modules
        self.root = ctk.CTk()
        self.root.title("Report Processing Suite")
        self.root.geometry("1200x800")  # Increased height from 900 to 1000
        self.root.configure(fg_color="#f8f9fa")
        
        # Set CustomTkinter theme
        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")
        
        # Center the window
        self.center_window()
        
        # Current module reference
        self.current_module = None
        self.current_frame = None
        
        # Create the main container with scrollable frame
        #self.main_container = ctk.CTkScrollableFrame(self.root, fg_color="#f8f9fa", corner_radius=0)
        self.main_container = ctk.CTkFrame(self.root, fg_color="#f8f9fa", corner_radius=0)
        self.main_container.pack(fill="both", expand=True)
        
        # Show homepage initially
        self.show_homepage()
        
    def center_window(self):
        """Center the window on the screen"""
        self.root.update_idletasks()
        width = 1200
        height = 800  # Updated height
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f"{width}x{height}+{x}+{y}")
    
    def clear_main_container(self):
        """Clear all widgets from main container"""
        for widget in self.main_container.winfo_children():
            widget.destroy()
    
    def reset_input_directories(self):
        """Reset all files in InputDIR while preserving folder structure"""
        try:
            input_dir = "InputDIR"
            if not os.path.exists(input_dir):
                messagebox.showwarning("Warning", "InputDIR folder not found!")
                return
            
            deleted_count = 0
            preserved_folders = []
            
            # Walk through all subdirectories
            for root, dirs, files in os.walk(input_dir):
                # Delete all files in current directory
                for file in files:
                    file_path = os.path.join(root, file)
                    try:
                        os.remove(file_path)
                        deleted_count += 1
                    except Exception as e:
                        print(f"Error deleting {file_path}: {e}")
                
                # Keep track of preserved folders
                if dirs:
                    preserved_folders.extend([os.path.join(root, d) for d in dirs])
            
            messagebox.showinfo(
                "Reset Complete", 
                f"Successfully deleted {deleted_count} files from InputDIR.\n"
                f"Folder structure preserved."
            )
            
        except Exception as e:
            messagebox.showerror("Reset Error", f"Error resetting InputDIR: {str(e)}")
    
    def open_folders(self):
        """Open InputDIR and Output folders"""
        try:
            # Create directories if they don't exist
            os.makedirs("InputDIR", exist_ok=True)
            os.makedirs("Output", exist_ok=True)
            
            # Open folders based on OS
            if os.name == 'nt':  # Windows
                os.startfile("InputDIR")
                os.startfile("Output")
            elif os.name == 'posix':  # macOS and Linux
                os.system(f"open {'InputDIR'}")
                os.system(f"open {'Output'}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Could not open folders: {str(e)}")
    
    def show_homepage(self):
        """Show the homepage interface"""
        self.clear_main_container()
        self.current_module = None
        
        # Create homepage frame with proper sizing
        homepage_frame = ctk.CTkFrame(self.main_container, fg_color="#f8f9fa", corner_radius=0)
        homepage_frame.pack(fill="both", expand=True, padx=30, pady=20)
        
        # Title Header Frame - reduced height
        header_frame = ctk.CTkFrame(homepage_frame, fg_color="#ff6b35", corner_radius=0, height=120)
        header_frame.pack(fill="x", pady=(0, 30))
        header_frame.pack_propagate(False)
        
        title_label = ctk.CTkLabel(
            header_frame,
            text="Report Processing Suite",
            font=ctk.CTkFont(size=42, weight="bold"),  # Slightly smaller font
            text_color="white"
        )
        title_label.pack(pady=(20, 5))
        
        subtitle_label = ctk.CTkLabel(
            header_frame,
            text="Streamline your pickup, returns, and cancellation reporting",
            font=ctk.CTkFont(size=16),  # Slightly smaller font
            text_color="white"
        )
        subtitle_label.pack(pady=(0, 20))
        
        # Main Content Frame
        content_frame = ctk.CTkFrame(homepage_frame, fg_color="transparent")
        content_frame.pack(fill="both", expand=True, pady=(0, 20))
        
        # Create module boxes
        self.create_module_boxes(content_frame)
        
        # Bottom Frame with version and buttons - fixed height
        bottom_frame = ctk.CTkFrame(homepage_frame, fg_color="transparent", height=70)
        bottom_frame.pack(fill="x", pady=(20, 10))
        bottom_frame.pack_propagate(False)
        
        # Left side - Version info
        version_label = ctk.CTkLabel(
            bottom_frame,
            text="Version 2.0 | Last Updated: May 2025",
            font=ctk.CTkFont(size=12),
            text_color="#6c757d"
        )
        version_label.pack(side="left", pady=15)
        
        # Right side - Action buttons
        buttons_frame = ctk.CTkFrame(bottom_frame, fg_color="transparent")
        buttons_frame.pack(side="right", pady=15)
        
        help_btn = ctk.CTkButton(
            buttons_frame,
            text="? Help",
            font=ctk.CTkFont(size=14),
            fg_color="#ffc107",
            hover_color="#ffb300",
            text_color="black",
            corner_radius=8,
            height=40,
            width=100,
            command=self.show_help
        )
        help_btn.pack(side="left", padx=(0, 10))
        
        reset_btn = ctk.CTkButton(
            buttons_frame,
            text="🗑 Reset Files",
            font=ctk.CTkFont(size=14),
            fg_color="#dc3545",
            hover_color="#c82333",
            text_color="white",
            corner_radius=8,
            height=40,
            width=120,
            command=self.reset_input_directories
        )
        reset_btn.pack(side="left", padx=(0, 10))
        
        open_folders_btn = ctk.CTkButton(
            buttons_frame,
            text="📁 Open Folders",
            font=ctk.CTkFont(size=14),
            fg_color="#ff6b35",
            hover_color="#e55a2b",
            text_color="white",
            corner_radius=8,
            height=40,
            width=140,
            command=self.open_folders
        )
        open_folders_btn.pack(side="left")
    
    def create_module_boxes(self, parent):
        """Create the three module boxes with better sizing"""
        # Container for the three boxes
        boxes_frame = ctk.CTkFrame(parent, fg_color="transparent")
        boxes_frame.pack(expand=True, fill="both")
        
        # Configure grid with better spacing
        boxes_frame.grid_columnconfigure(0, weight=1)
        boxes_frame.grid_columnconfigure(1, weight=1)
        boxes_frame.grid_columnconfigure(2, weight=1)
        boxes_frame.grid_rowconfigure(0, weight=1)
        
        # Module configurations
        modules = [
            {
                "title": "Pickup Report\nProcessor",
                "description": "Process pickup reports from multiple courier sources including Sellerflex, Flipkart KC/LL, and Meesho manifests.",
                "color": "#e9ecef",
                "text_color": "#495057",
                "icon": "📦",
                "module_file": "Pickupreportexe.py",
                "module_class": "PickupReportModule",
                "status": "Missing 4 files",
                "status_color": "#ffc107",
                "button_color": "#ff6b35"
            },
            {
                "title": "Returns\nReconciliation",
                "description": "Reconcile return shipments and generate comprehensive return reports with tracking analysis.",
                "color": "#e9ecef",
                "text_color": "#495057",
                "icon": "↩️",
                "module_file": "ReturnsReportexe.py",
                "module_class": "ReturnsReportGUI",
                "status": "Ready",
                "status_color": "#28a745",
                "button_color": "#007bff"
            },
            {
                "title": "Cancellation\nReport",
                "description": "Generate detailed cancellation reports and analyze cancellation patterns across platforms.",
                "color": "#e9ecef",
                "text_color": "#495057",
                "icon": "❌",
                "module_file": "Cancellationexe.py",
                "module_class": "CancellationReportModule",
                "status": "Missing TBD files",
                "status_color": "#ffc107",
                "button_color": "#28a745"
            }
        ]
        
        # Create boxes
        for i, module in enumerate(modules):
            self.create_module_box(boxes_frame, module, i)
    
    def create_module_box(self, parent, module_config, column):
        """Create individual module box with optimized sizing"""
        # Main frame for the box - reduced padding
        box_frame = ctk.CTkFrame(
            parent,
            fg_color=module_config["color"],
            corner_radius=12,
            border_width=1,
            border_color="#dee2e6"
        )
        box_frame.grid(row=0, column=column, padx=15, pady=15, sticky="nsew")
        
        # Inner frame for content - reduced padding
        content_frame = ctk.CTkFrame(box_frame, fg_color="transparent")
        content_frame.pack(expand=True, fill="both", padx=25, pady=25)
        
        # Icon - smaller size
        icon_label = ctk.CTkLabel(
            content_frame,
            text=module_config["icon"],
            font=ctk.CTkFont(size=50),  # Reduced from 60
            text_color=module_config["text_color"]
        )
        icon_label.pack(pady=(0, 15))
        
        # Title - smaller font
        title_label = ctk.CTkLabel(
            content_frame,
            text=module_config["title"],
            font=ctk.CTkFont(size=22, weight="bold"),  # Reduced from 24
            text_color=module_config["text_color"],
            justify="center"
        )
        title_label.pack(pady=(0, 15))
        
        # Description - smaller font and reduced wrap length
        desc_label = ctk.CTkLabel(
            content_frame,
            text=module_config["description"],
            font=ctk.CTkFont(size=13),  # Reduced from 14
            text_color="#6c757d",
            justify="center",
            wraplength=260  # Reduced from 280
        )
        desc_label.pack(pady=(0, 20))
        
        # Status indicator
        status_frame = ctk.CTkFrame(content_frame, fg_color="transparent")
        status_frame.pack(pady=(0, 20))
        
        status_dot = ctk.CTkLabel(
            status_frame,
            text="●",
            font=ctk.CTkFont(size=14),  # Reduced from 16
            text_color=module_config["status_color"]
        )
        status_dot.pack(side="left")
        
        status_label = ctk.CTkLabel(
            status_frame,
            text=module_config["status"],
            font=ctk.CTkFont(size=11),  # Reduced from 12
            text_color=module_config["status_color"]
        )
        status_label.pack(side="left", padx=(5, 0))
        
        # Open Module Button - smaller size
        open_btn = ctk.CTkButton(
            content_frame,
            text="Open Module",
            font=ctk.CTkFont(size=15, weight="bold"),  # Reduced from 16
            fg_color=module_config["button_color"],
            hover_color=self.darken_color(module_config["button_color"]),
            text_color="white",
            corner_radius=8,
            height=45,  # Reduced from 50
            width=160,  # Reduced from 180
            command=lambda config=module_config: self.load_module(config)
        )
        open_btn.pack()
    
    def darken_color(self, color):
        """Helper function to darken a color for hover effect"""
        color_map = {
            "#ff6b35": "#e55a2b",
            "#007bff": "#0056b3",
            "#28a745": "#1e7e34"
        }
        return color_map.get(color, color)
    
    def show_help(self):
        """Show help dialog"""
        help_text = """Report Processing Suite Help

    This application is designed to process different types of reports efficiently.

    Pickup Report Processor:
    It processes pickup reports from Sellerflex, Flipkart KC, Flipkart LL, and Meesho. The required input files should be placed in the InputDIR folder. The expected files are:
    - Sellerflex.csv: Contains pickup data from Sellerflex platform.
    - Flipkart KC.csv: Contains pickup data from Flipkart KC platform.
    - Flipkart LL.csv: Contains pickup data from Flipkart LL platform.
    - Manifest.pdf: Contains manifest details for Meesho pickups.

    Returns Reconciliation:
    This module reconciles return shipments and generates detailed reports. It helps analyze return trends and tracking information. The input files needed are:
    - Returns Flipkart KC.csv: Contains return data for Flipkart KC platform.
    - Returns Flipkart LL.csv: Contains return data for Flipkart LL platform.
    - Returns Meesho.csv: Contains return data for Meesho platform.
    - Returns Sellerflex.csv: Contains return data for Sellerflex platform.

    Cancellation Report:
    It generates cancellation reports and analyzes cancellation trends across different platforms. The required input files are:
    - Manifest.pdf: Contains manifest details for cancellations.
    - Vlookup_data.csv: Contains lookup data for cancellations.
    - Flipkart files: Specific files related to Flipkart cancellations.

    To use the application, place the necessary input files in the InputDIR folder. Click "Open Module" to begin processing. The processed reports will be saved automatically in the Output folder. Use the "Reset Files" option to clear all input files, and "Open Folders" to access both InputDIR and Output directories.

    For technical assistance, please contact the development team."""
        
        messagebox.showinfo("Help", help_text)
    
    def load_module(self, module_config):
        """Load and display a module within the same window"""
        try:
            module_file = module_config["module_file"]
            module_name = module_config["title"].replace('\n', ' ')
            
            # Check if file exists
            if not os.path.exists(module_file):
                messagebox.showerror(
                    "Module Not Found", 
                    f"The module file '{module_file}' was not found.\n\n"
                    f"Please ensure the file exists in the application directory."
                )
                return
            
            # Clear the main container
            self.clear_main_container()
            
            # Create module container frame
            module_container = ctk.CTkFrame(self.main_container, fg_color="transparent", corner_radius=0)
            module_container.pack(fill="both", expand=True)
            
            # Add back button
            back_btn = ctk.CTkButton(
                module_container,
                text="← Back to Homepage",
                font=ctk.CTkFont(size=14),
                fg_color="#6c757d",
                hover_color="#545b62",
                corner_radius=8,
                height=40,
                width=200,
                command=self.show_homepage
            )
            back_btn.pack(pady=10, padx=20, anchor="nw")
            
            # Import and initialize the module
            self.import_and_run_module(module_config, module_container)
            
        except Exception as e:
            messagebox.showerror("Module Load Error", f"Error loading {module_name}:\n\n{str(e)}")
            self.show_homepage()
    
    def import_and_run_module(self, module_config, parent_frame):
        """Import and initialize the specific module"""
        try:
            module_file = module_config["module_file"]
            module_name = module_config["title"].replace('\n', ' ')
            expected_class = module_config["module_class"]
            
            # Get the module name without extension
            module_name_clean = os.path.splitext(os.path.basename(module_file))[0]
            
            # Load module dynamically
            spec = importlib.util.spec_from_file_location(module_name_clean, module_file)
            if spec is None:
                raise ImportError(f"Could not load spec for {module_file}")
            
            module = importlib.util.module_from_spec(spec)
            
            # Add the module's directory to sys.path temporarily
            original_path = sys.path.copy()
            module_dir = os.path.dirname(os.path.abspath(module_file))
            if module_dir not in sys.path:
                sys.path.insert(0, module_dir)
            
            try:
                spec.loader.exec_module(module)
            finally:
                # Restore original sys.path
                sys.path = original_path
            
            # Check if the module has the expected class
            if not hasattr(module, expected_class):
                available_classes = [name for name in dir(module) 
                                   if not name.startswith('_') and 
                                   hasattr(getattr(module, name), '__call__')]
                
                messagebox.showerror(
                    "Class Not Found", 
                    f"Class '{expected_class}' not found in {module_file}.\n\n"
                    f"Available classes: {', '.join(available_classes)}"
                )
                self.show_homepage()
                return
            
            gui_class = getattr(module, expected_class)
            
            # Create instance of the GUI class
            try:
                # Try different constructor signatures
                try:
                    gui_instance = gui_class(parent_frame, back_callback=self.show_homepage)
                except TypeError:
                    try:
                        gui_instance = gui_class(parent_frame)
                    except TypeError:
                        gui_instance = gui_class()
                
                self.current_module = gui_instance
                
            except Exception as constructor_error:
                messagebox.showerror(
                    "Module Constructor Error", 
                    f"Error creating {expected_class} instance:\n\n{str(constructor_error)}"
                )
                self.show_homepage()
                return
            
        except Exception as e:
            messagebox.showerror(
                "Module Import Error", 
                f"Failed to load {module_name} module:\n\n{str(e)}"
            )
            self.show_homepage()
    
    def run(self):
        """Start the application"""
        try:
            self.root.mainloop()
        except Exception as e:
            messagebox.showerror("Application Error", f"An error occurred:\n\n{str(e)}")

def main():
    """Main function to run the application"""
    try:
        app = ReportProcessorHomepage()
        app.run()
    except Exception as e:
        messagebox.showerror("Application Error", f"Application startup error:\n\n{str(e)}")

if __name__ == "__main__":
    main()