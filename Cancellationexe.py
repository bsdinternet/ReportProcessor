import os
import pandas as pd
from datetime import datetime
import pdfplumber
from openpyxl import load_workbook, Workbook
import customtkinter as ctk
from tkinter import messagebox, scrolledtext
import threading
import traceback

class CancellationReportModule:
    def __init__(self, parent_frame=None, back_callback=None, root_window=None):
        self.parent_frame = parent_frame
        self.back_callback = back_callback
        self.root_window = root_window  # Reference to main window for threading
        
        # Data storage
        self.combined_df = pd.DataFrame()
        self.processing = False
        
        # Directory setup
        self.setup_directories()
        
        # Create GUI
        self.create_gui()
        
    def setup_directories(self):
        """Setup directory paths"""
        self.base_dir = os.path.dirname(os.path.abspath(__file__))
        self.input_dir = os.path.join(self.base_dir, 'InputDIR')
        self.cancel_dir = os.path.join(self.input_dir, 'CancellationReport')
        self.pickup_dir = os.path.join(self.input_dir, 'PickupReportfiles')
        self.output_file_path = os.path.join(self.base_dir, 'OutputDIR', 'Cancel_product_report.xlsx')
        
        # Create directories
        os.makedirs(self.cancel_dir, exist_ok=True)
        os.makedirs(self.pickup_dir, exist_ok=True)
        os.makedirs(os.path.join(self.base_dir, 'OutputDIR'), exist_ok=True)
    
    def create_gui(self):
        """Create the main GUI interface"""
        if self.parent_frame is None:
            # Standalone mode
            self.root = ctk.CTk()
            self.root.title("Cancellation Report Processor")
            self.root.geometry("1200x800")
            self.root.configure(fg_color="#f8f9fa")
            main_frame = self.root
            self.is_standalone = True
        else:
            # Embedded mode
            main_frame = self.parent_frame
            self.root = self.root_window if self.root_window else self.parent_frame.winfo_toplevel()
            self.is_standalone = False
        
        # Clear the parent frame if in embedded mode
        if not self.is_standalone:
            for widget in main_frame.winfo_children():
                widget.destroy()
        
        # Main container
        self.main_container = ctk.CTkFrame(main_frame, fg_color="#f8f9fa", corner_radius=0)
        self.main_container.pack(fill="both", expand=True, padx=20, pady=20)
        
        # Header (includes back button for embedded mode)
        self.create_header()
        
        # Content area
        self.create_content_area()
        
        # Status area
        self.create_status_area()
    
# Replace the problematic back button creation in create_header method:

    def create_header(self):
        """Create header section with integrated back button"""
        header_frame = ctk.CTkFrame(self.main_container, fg_color="#ff6b35", corner_radius=12, height=120)
        header_frame.pack(fill="x", pady=(0, 30))
        header_frame.pack_propagate(False)
        
        # Add back button for embedded mode - positioned in top-left of header
        if not self.is_standalone and self.back_callback:
            back_btn = ctk.CTkButton(
                header_frame,
                text="← Back to Home",
                font=ctk.CTkFont(size=14, weight="bold"),
                fg_color="#ff8a65",  # Changed from rgba(255,255,255,0.2)
                hover_color="#ff7043",  # Changed from rgba(255,255,255,0.3)
                text_color="white",
                corner_radius=8,
                height=35,
                width=140,
                command=self.go_back
            )
            back_btn.place(x=20, y=15)  # Position in top-left corner
        
        # Rest of the header code remains the same...
        # Title with icon - centered
        title_frame = ctk.CTkFrame(header_frame, fg_color="transparent")
        title_frame.pack(expand=True, fill="both")
        
        icon_label = ctk.CTkLabel(
            title_frame,
            text="❌",
            font=ctk.CTkFont(size=40),
            text_color="white"
        )
        icon_label.pack(pady=(20, 5))
        
        title_label = ctk.CTkLabel(
            title_frame,
            text="Cancellation Report Processor",
            font=ctk.CTkFont(size=32, weight="bold"),
            text_color="white"
        )
        title_label.pack(pady=(0, 5))
        
        subtitle_label = ctk.CTkLabel(
            title_frame,
            text="Process cancellation reports from Meesho and Flipkart platforms",
            font=ctk.CTkFont(size=16),
            text_color="white"
        )
        subtitle_label.pack(pady=(0, 20))
    
    def create_back_button(self):
        """Create back button for embedded mode"""
        back_frame = ctk.CTkFrame(self.main_container, fg_color="transparent", height=50)
        back_frame.pack(fill="x", pady=(0, 10))
        back_frame.pack_propagate(False)
        
        back_btn = ctk.CTkButton(
            back_frame,
            text="← Back to Home",
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color="#6c757d",
            hover_color="#545b62",
            text_color="white",
            corner_radius=8,
            height=40,
            width=150,
            command=self.go_back
        )
        back_btn.pack(side="left", padx=10, pady=5)
    
    def go_back(self):
        """Handle back button click"""
        if self.back_callback:
            self.back_callback()
        
    def create_content_area(self):
        """Create main content area with controls and data display"""
        content_frame = ctk.CTkFrame(self.main_container, fg_color="transparent")
        content_frame.pack(fill="both", expand=True)
        
        # Left panel - Controls
        left_panel = ctk.CTkFrame(content_frame, fg_color="#ffffff", corner_radius=12, width=400)
        left_panel.pack(side="left", fill="y", padx=(0, 15))
        left_panel.pack_propagate(False)
        
        # Controls title
        controls_label = ctk.CTkLabel(
            left_panel,
            text="Processing Controls",
            font=ctk.CTkFont(size=20, weight="bold"),
            text_color="#495057"
        )
        controls_label.pack(pady=(30, 20))
        
        # File status section
        self.create_file_status_section(left_panel)
        
        # Process button
        self.process_btn = ctk.CTkButton(
            left_panel,
            text="🔄 Process Reports",
            font=ctk.CTkFont(size=16, weight="bold"),
            fg_color="#28a745",
            hover_color="#218838",
            text_color="white",
            corner_radius=8,
            height=50,
            width=300,
            command=self.start_processing
        )
        self.process_btn.pack(pady=(30, 20))
        
        # Save button
        self.save_btn = ctk.CTkButton(
            left_panel,
            text="💾 Save Excel Report",
            font=ctk.CTkFont(size=16, weight="bold"),
            fg_color="#007bff",
            hover_color="#0056b3",
            text_color="white",
            corner_radius=8,
            height=50,
            width=300,
            command=self.save_report,
            state="disabled"
        )
        self.save_btn.pack(pady=(0, 20))
        
        # Clear button
        clear_btn = ctk.CTkButton(
            left_panel,
            text="🗑 Clear Data",
            font=ctk.CTkFont(size=14),
            fg_color="#dc3545",
            hover_color="#c82333",
            text_color="white",
            corner_radius=8,
            height=40,
            width=200,
            command=self.clear_data
        )
        clear_btn.pack(pady=(0, 30))
        
        # Right panel - Data display
        right_panel = ctk.CTkFrame(content_frame, fg_color="#ffffff", corner_radius=12)
        right_panel.pack(side="right", fill="both", expand=True)
        
        # Data display title
        data_label = ctk.CTkLabel(
            right_panel,
            text="Cancelled Products Data",
            font=ctk.CTkFont(size=20, weight="bold"),
            text_color="#495057"
        )
        data_label.pack(pady=(30, 20))
        
        # Data text widget
        self.data_text = ctk.CTkTextbox(
            right_panel,
            font=ctk.CTkFont(family="Courier", size=10),
            fg_color="#f8f9fa",
            text_color="#495057",
            corner_radius=8,
            wrap="none"
        )
        self.data_text.pack(fill="both", expand=True, padx=30, pady=(0, 30))
        
        # Initial message
        self.data_text.insert("1.0", "No data processed yet. Click 'Process Reports' to begin.")
    
    def create_file_status_section(self, parent):
        """Create file status indicators"""
        status_frame = ctk.CTkFrame(parent, fg_color="#f8f9fa", corner_radius=8)
        status_frame.pack(fill="x", padx=30, pady=(0, 20))
        
        status_title = ctk.CTkLabel(
            status_frame,
            text="Required Files Status",
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color="#495057"
        )
        status_title.pack(pady=(15, 10))
        
        # File status indicators
        self.file_status_labels = {}
        
        files_to_check = [
            ("Meesho Manifest.pdf", os.path.join(self.pickup_dir, 'Manifest.pdf')),
            ("Meesho Meesho_data.csv", os.path.join(self.cancel_dir, 'Meesho_data.csv')),
            ("Flipkart Cancel Files", self.cancel_dir),
            ("Flipkart Pickup Files", self.pickup_dir)
        ]
        
        for file_name, file_path in files_to_check:
            self.create_file_status_indicator(status_frame, file_name, file_path)
        
        # Refresh button
        refresh_btn = ctk.CTkButton(
            status_frame,
            text="🔄 Refresh Status",
            font=ctk.CTkFont(size=12),
            fg_color="#6c757d",
            hover_color="#545b62",
            text_color="white",
            corner_radius=6,
            height=30,
            width=150,
            command=self.refresh_file_status
        )
        refresh_btn.pack(pady=(10, 15))
        
        # Initial status check
        self.refresh_file_status()
    
    def create_file_status_indicator(self, parent, file_name, file_path):
        """Create individual file status indicator"""
        indicator_frame = ctk.CTkFrame(parent, fg_color="transparent")
        indicator_frame.pack(fill="x", padx=10, pady=2)
        
        status_dot = ctk.CTkLabel(
            indicator_frame,
            text="●",
            font=ctk.CTkFont(size=12),
            text_color="#dc3545"  # Default to red
        )
        status_dot.pack(side="left", padx=(5, 5))
        
        file_label = ctk.CTkLabel(
            indicator_frame,
            text=file_name,
            font=ctk.CTkFont(size=11),
            text_color="#495057"
        )
        file_label.pack(side="left")
        
        self.file_status_labels[file_name] = status_dot
    
    def refresh_file_status(self):
        """Refresh file status indicators"""
        try:
            # Check Meesho files
            meesho_pdf = os.path.exists(os.path.join(self.pickup_dir, 'Manifest.pdf'))
            meesho_csv = os.path.exists(os.path.join(self.cancel_dir, 'Meesho_data.csv'))
            
            # Check Flipkart files
            fk_cancel_files = False
            fk_pickup_files = False
            
            if os.path.exists(self.cancel_dir):
                fk_cancel_files = any(f.endswith('.csv') and 'Flipkart' in f for f in os.listdir(self.cancel_dir) if os.path.isfile(os.path.join(self.cancel_dir, f)))
            
            if os.path.exists(self.pickup_dir):
                fk_pickup_files = any(f.endswith('.csv') and 'Flipkart' in f for f in os.listdir(self.pickup_dir) if os.path.isfile(os.path.join(self.pickup_dir, f)))
            
            # Update status indicators
            statuses = {
                "Meesho Manifest.pdf": meesho_pdf,
                "Meesho Meesho_data.csv": meesho_csv,
                "Flipkart Cancel Files": fk_cancel_files,
                "Flipkart Pickup Files": fk_pickup_files
            }
            
            for file_name, status in statuses.items():
                if file_name in self.file_status_labels:
                    color = "#28a745" if status else "#dc3545"
                    self.file_status_labels[file_name].configure(text_color=color)
        except Exception as e:
            print(f"Error refreshing file status: {e}")
    
    def create_status_area(self):
        """Create status/progress area"""
        self.status_frame = ctk.CTkFrame(self.main_container, fg_color="#e9ecef", corner_radius=12, height=60)
        self.status_frame.pack(fill="x", pady=(20, 0))
        self.status_frame.pack_propagate(False)
        
        self.status_label = ctk.CTkLabel(
            self.status_frame,
            text="Ready to process cancellation reports",
            font=ctk.CTkFont(size=14),
            text_color="#495057"
        )
        self.status_label.pack(pady=20)
    
    def start_processing(self):
        """Start processing in a separate thread"""
        if self.processing:
            return
        
        self.processing = True
        self.process_btn.configure(state="disabled", text="Processing...")
        self.save_btn.configure(state="disabled")
        
        # Start processing thread
        threading.Thread(target=self.process_reports, daemon=True).start()
    
    def process_reports(self):
        """Main processing function (runs in separate thread)"""
        try:
            self.update_status("Starting report processing...")
            
            # Process Meesho data
            meesho_cancelled_df = self.process_meesho_data()
            
            # Process Flipkart data
            flipkart_df_list = self.process_flipkart_data()
            
            # Combine all data
            self.combine_cancelled_data(meesho_cancelled_df, flipkart_df_list)
            
            # Update GUI in main thread
            self.schedule_gui_update(self.processing_complete)
            
        except Exception as e:
            error_msg = f"Processing error: {str(e)}"
            print(f"Error: {error_msg}")
            traceback.print_exc()
            self.schedule_gui_update(lambda: self.processing_error(error_msg))
    
    def schedule_gui_update(self, callback):
        """Schedule GUI update in main thread (works for both standalone and embedded)"""
        if self.root and hasattr(self.root, 'after'):
            self.root.after(0, callback)
        else:
            # Fallback - call directly (may cause threading issues but better than nothing)
            callback()
    
    def process_meesho_data(self):
        """Process Meesho PDF and CSV data"""
        pdf_path = os.path.join(self.pickup_dir, 'Manifest.pdf')
        vlookup_file_path = os.path.join(self.cancel_dir, 'Meesho_data.csv')
        
        if os.path.exists(pdf_path) and os.path.exists(vlookup_file_path):
            self.update_status("Processing Meesho PDF data...")
            df_extracted = self.extract_data_from_pdf(pdf_path)
            
            self.update_status("Performing VLOOKUP for Meesho data...")
            _, meesho_cancelled_df = self.perform_vlookup(df_extracted, vlookup_file_path)
            
            return meesho_cancelled_df
        else:
            self.update_status("Meesho files not found, skipping...")
            return pd.DataFrame()
    
    def process_flipkart_data(self):
        """Process Flipkart cancellation data"""
        flipkart_df_list = []
        
        self.update_status("Processing Flipkart cancellation data...")
        
        if not os.path.exists(self.cancel_dir):
            return flipkart_df_list
        
        for file in os.listdir(self.cancel_dir):
            if file.endswith('.csv') and 'Flipkart' in file:
                cancel_path = os.path.join(self.cancel_dir, file)
                pickup_path = os.path.join(self.pickup_dir, file)
                
                if os.path.exists(pickup_path):
                    sale_channel = 'Flipkart LL' if 'LL' in file else 'Flipkart KC'
                    fk_df = self.flipkart_cancelled_orders(cancel_path, pickup_path, sale_channel)
                    flipkart_df_list.append(fk_df)
                    self.update_status(f"Processed {file}")
                else:
                    self.update_status(f"Warning: Pickup file not found for {file}")
        
        return flipkart_df_list
    
    def combine_cancelled_data(self, meesho_cancelled_df, flipkart_df_list):
        """Combine all cancelled data"""
        self.update_status("Combining all cancellation data...")
        
        all_data = []
        
        # Add Meesho data
        if meesho_cancelled_df is not None and not meesho_cancelled_df.empty:
            meesho_cancelled_df = meesho_cancelled_df.copy()
            meesho_cancelled_df['SaleChannel'] = 'Meesho'
            meesho_cancelled_df.rename(columns={'AWB': 'Tracking ID'}, inplace=True)
            all_data.append(meesho_cancelled_df[['SaleChannel', 'Sub Order Number', 'Tracking ID', 'Status of the product', 'SKU', 'QTY', 'Invoice Amount']])
        
        # Add Flipkart data
        for df in flipkart_df_list:
            if df is not None and not df.empty:
                all_data.append(df)
        
        if all_data:
            self.combined_df = pd.concat(all_data, ignore_index=True)
        else:
            self.combined_df = pd.DataFrame()
    
    def processing_complete(self):
        """Called when processing is complete"""
        self.processing = False
        self.process_btn.configure(state="normal", text="🔄 Process Reports")
        
        if not self.combined_df.empty:
            self.save_btn.configure(state="normal")
            self.display_data()
            self.update_status(f"Processing complete! Found {len(self.combined_df)} cancelled products.")
        else:
            self.update_status("Processing complete but no cancelled products found.")
            self.data_text.delete("1.0", "end")
            self.data_text.insert("1.0", "No cancelled products found in the processed files.")
    
    def processing_error(self, error_msg):
        """Called when processing encounters an error"""
        self.processing = False
        self.process_btn.configure(state="normal", text="🔄 Process Reports")
        self.update_status(f"Error: {error_msg}")
        if messagebox:
            messagebox.showerror("Processing Error", error_msg)
    
    def display_data(self):
        """Display processed data in the text widget"""
        if self.combined_df.empty:
            self.data_text.delete("1.0", "end")
            self.data_text.insert("1.0", "No data to display.")
            return
        
        # Clear existing content
        self.data_text.delete("1.0", "end")
        
        # Create formatted table
        output_lines = []
        
        # Headers
        headers = list(self.combined_df.columns)
        header_line = " | ".join(f"{header:<15}" for header in headers)
        separator_line = "-" * len(header_line)
        
        output_lines.append(header_line)
        output_lines.append(separator_line)
        
        # Data rows (limit to first 100 for display performance)
        display_rows = min(100, len(self.combined_df))
        for _, row in self.combined_df.head(display_rows).iterrows():
            data_line = " | ".join(f"{str(value):<15}" for value in row.values)
            output_lines.append(data_line)
        
        # Add summary
        output_lines.append("")
        output_lines.append(f"Total Records: {len(self.combined_df)}")
        if len(self.combined_df) > display_rows:
            output_lines.append(f"(Showing first {display_rows} records)")
        
        # Insert into text widget
        content = "\n".join(output_lines)
        self.data_text.insert("1.0", content)
    
    def save_report(self):
        """Save the processed data to Excel file"""
        if self.combined_df.empty:
            if messagebox:
                messagebox.showwarning("No Data", "No data to save. Please process reports first.")
            return
        
        try:
            # Ensure output directory exists
            os.makedirs(os.path.dirname(self.output_file_path), exist_ok=True)
            
            # Save to Excel file
            with pd.ExcelWriter(self.output_file_path, engine='openpyxl') as writer:
                self.combined_df.to_excel(writer, index=False, sheet_name='Cancel products')
            
            if messagebox:
                messagebox.showinfo("Success", f"Report saved successfully to:\n{self.output_file_path}")
            self.update_status(f"Report saved: {len(self.combined_df)} records")
            
        except Exception as e:
            error_msg = f"Error saving report: {str(e)}"
            if messagebox:
                messagebox.showerror("Save Error", error_msg)
            self.update_status(error_msg)
    
    def clear_data(self):
        """Clear all processed data"""
        self.combined_df = pd.DataFrame()
        self.data_text.delete("1.0", "end")
        self.data_text.insert("1.0", "Data cleared. Click 'Process Reports' to begin.")
        self.save_btn.configure(state="disabled")
        self.update_status("Data cleared")
    
    def update_status(self, message):
        """Update status label (thread-safe)"""
        def update():
            if hasattr(self, 'status_label') and self.status_label.winfo_exists():
                self.status_label.configure(text=message)
        
        self.schedule_gui_update(update)
    
    # ========== Original Processing Functions ==========
    
    def extract_data_from_pdf(self, pdf_path):
        """Extract data from Meesho PDF"""
        courier_data = {}
        with pdfplumber.open(pdf_path) as reader:
            for i, page in enumerate(reader.pages[1:], start=2):
                tables = page.extract_tables()
                if tables:
                    page_text = page.extract_text()
                    courier_name = "Unknown Courier"
                    if "Courier :" in page_text:
                        courier_name = page_text.split("Courier :")[1].split("\n")[0].strip()

                    if courier_name not in courier_data:
                        courier_data[courier_name] = []

                    for table in tables:
                        sub_order_data = [str(row[1]).replace("\n", "").strip() if len(row) > 1 else "" for row in table[1:]]
                        awb_data = [str(row[2]).strip() if len(row) > 2 else "123" for row in table[1:]]
                        combined_data = list(zip(sub_order_data, awb_data))
                        courier_data[courier_name].extend(combined_data)

        data_list = [{'Courier': courier, 'AWB': awb, 'Sub Order Number': sub_order} for courier, entries in courier_data.items() for sub_order, awb in entries]
        return pd.DataFrame(data_list)

    def perform_vlookup(self, df_extracted, vlookup_file_path):
        """Perform VLOOKUP for Meesho data"""
        vlookup_data = pd.read_csv(vlookup_file_path)

        vlookup_data.rename(columns={
            'Reason for Credit Entry': 'Status of the product',
            'Sub Order No': 'Sub Order Number',
            'SKU': 'SKU',
            'Quantity': 'QTY',
            'Supplier Listed Price (Incl. GST + Commission)': 'Invoice Amount'
        }, inplace=True)

        merged_df = pd.merge(
            df_extracted,
            vlookup_data[['Sub Order Number', 'Status of the product', 'SKU', 'QTY', 'Invoice Amount']],
            on='Sub Order Number',
            how='left'
        )

        filtered_df = merged_df[merged_df['Status of the product'].astype(str).str.strip().str.upper() == 'CANCELLED']
        return merged_df, filtered_df

    def flipkart_cancelled_orders(self, fk_cancel_path, fk_pickup_path, sale_channel):
        """Process Flipkart cancelled orders"""
        df_cancel = pd.read_csv(fk_cancel_path)
        df_pickup = pd.read_csv(fk_pickup_path)

        # Rename Cancel file columns
        df_cancel.rename(columns={
            df_cancel.columns[0]: 'Order Cancellation Date',
            df_cancel.columns[2]: 'OrderID',
            df_cancel.columns[5]: 'Cancellation Type'
        }, inplace=True)

        # Filter today's cancellations
        df_cancel['Order Cancellation Date'] = pd.to_datetime(df_cancel['Order Cancellation Date'], errors='coerce').dt.date
        today = datetime.today().date()
        df_cancel = df_cancel[df_cancel['Order Cancellation Date'] == today]

        # Filter by "Cancelled by buyer"
        df_cancel = df_cancel[df_cancel['Cancellation Type'].str.strip().str.lower() == 'cancelled by buyer']

        # Extract relevant Order IDs
        order_ids = df_cancel['OrderID'].astype(str).str.strip().unique()

        # Rename Pickup file columns
        df_pickup.rename(columns={
            df_pickup.columns[3]: 'OrderID',
            'SKU': 'SKU',
            'Tracking ID': 'Tracking ID',
            'Quantity': 'QTY',
            'Invoice Amount': 'Invoice Amount'
        }, inplace=True)

        # Match Order IDs
        df_pickup['OrderID'] = df_pickup['OrderID'].astype(str).str.strip()
        df_filtered = df_pickup[df_pickup['OrderID'].isin(order_ids)]

        # Merge and return result
        df_result = pd.merge(
            df_filtered[['OrderID', 'Tracking ID', 'SKU', 'QTY', 'Invoice Amount']],
            df_cancel[['OrderID', 'Cancellation Type']],
            on='OrderID', how='left'
        )
        df_result['SaleChannel'] = sale_channel

        return df_result[['SaleChannel', 'OrderID', 'Tracking ID', 'Cancellation Type', 'SKU', 'QTY', 'Invoice Amount']]

    def run(self):
        """Run the standalone application"""
        if hasattr(self, 'root') and self.is_standalone:
            self.root.mainloop()

# For standalone execution
if __name__ == "__main__":
    app = CancellationReportModule()
    app.run()