import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pdfplumber
import pandas as pd
import re
from datetime import datetime
import os

class PDFToExcelConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF to Excel Converter - TC Data")
        self.root.geometry("1000x800")  # Increased size for preview
        self.root.resizable(True, True)
        
        # Variables
        self.pdf_path = tk.StringVar()
        self.excel_path = tk.StringVar()
        self.extracted_data = []  # Store extracted data
        self.records_count_var = tk.StringVar(value="Total Records: 0")
        
        self.setup_ui()
        
    def setup_ui(self):
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(7, weight=1)  # Make preview area expandable
        
        # Title
        title_label = ttk.Label(main_frame, text="PDF to Excel Converter - TC Data", 
                                font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # Info label - CORRECTED column names
        info_label = ttk.Label(main_frame, text="Extracts: Date | Program | Advt_time (HH:MM:SS) | Advt_theme | Dur", 
                            font=("Arial", 10), foreground="blue")
        info_label.grid(row=1, column=0, columnspan=3, pady=(0, 10))
        
        # PDF File Selection
        ttk.Label(main_frame, text="Select PDF File:").grid(row=2, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.pdf_path, width=50).grid(row=2, column=1, sticky=(tk.W, tk.E), pady=5, padx=(5, 5))
        ttk.Button(main_frame, text="Browse", command=self.browse_pdf).grid(row=2, column=2, pady=5)
        
        # Excel File Output
        ttk.Label(main_frame, text="Save Excel As:").grid(row=3, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.excel_path, width=50).grid(row=3, column=1, sticky=(tk.W, tk.E), pady=5, padx=(5, 5))
        ttk.Button(main_frame, text="Browse", command=self.browse_excel).grid(row=3, column=2, pady=5)
        
        # Buttons Frame
        buttons_frame = ttk.Frame(main_frame)
        buttons_frame.grid(row=4, column=0, columnspan=3, pady=20)
        
        # Extract Button (separate from export)
        extract_btn = ttk.Button(buttons_frame, text="📄 Extract from PDF", 
                                command=self.extract_only, style="Accent.TButton")
        extract_btn.pack(side=tk.LEFT, padx=5)
        
        # Export Button (initially disabled)
        self.export_btn = ttk.Button(buttons_frame, text="💾 Export to Excel", 
                                    command=self.export_to_excel, state=tk.DISABLED)
        self.export_btn.pack(side=tk.LEFT, padx=5)
        
        # Clear Button
        clear_btn = ttk.Button(buttons_frame, text="🗑️ Clear Data", 
                            command=self.clear_data)
        clear_btn.pack(side=tk.LEFT, padx=5)
        
        # Progress Bar
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)
        
        # Records Count Label
        count_label = ttk.Label(main_frame, textvariable=self.records_count_var, 
                            font=("Arial", 12, "bold"), foreground="green")
        count_label.grid(row=6, column=0, columnspan=3, pady=5)
        
        # Preview Frame
        preview_frame = ttk.LabelFrame(main_frame, text="📋 Data Preview", padding=10)
        preview_frame.grid(row=7, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10)
        
        # Configure preview frame
        preview_frame.columnconfigure(0, weight=1)
        preview_frame.rowconfigure(0, weight=1)
        
        # Create Treeview for data preview - CORRECTED column names
        self.preview_tree = ttk.Treeview(preview_frame, columns=('Date', 'Program', 'Advt_time', 'Advt_theme', 'Dur'), 
                                        show='headings', height=10)
        
        # Configure column headings and widths - CORRECTED
        self.preview_tree.heading('Date', text='Date')
        self.preview_tree.heading('Program', text='Program')
        self.preview_tree.heading('Advt_time', text='Time')
        self.preview_tree.heading('Advt_theme', text='Advertisement Theme')
        self.preview_tree.heading('Dur', text='Duration')
        
        self.preview_tree.column('Date', width=100, anchor='center')
        self.preview_tree.column('Program', width=150)
        self.preview_tree.column('Advt_time', width=80, anchor='center')
        self.preview_tree.column('Advt_theme', width=300)
        self.preview_tree.column('Dur', width=80, anchor='center')
        
        # Scrollbars for preview
        preview_v_scrollbar = ttk.Scrollbar(preview_frame, orient="vertical", command=self.preview_tree.yview)
        preview_h_scrollbar = ttk.Scrollbar(preview_frame, orient="horizontal", command=self.preview_tree.xview)
        self.preview_tree.configure(yscrollcommand=preview_v_scrollbar.set, xscrollcommand=preview_h_scrollbar.set)
        
        # Grid the preview components
        self.preview_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        preview_v_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        preview_h_scrollbar.grid(row=1, column=0, sticky=(tk.W, tk.E))
        
        # Status Text (smaller now)
        status_frame = ttk.LabelFrame(main_frame, text="📝 Processing Log", padding=5)
        status_frame.grid(row=8, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)
        status_frame.columnconfigure(0, weight=1)
        
        self.status_text = tk.Text(status_frame, height=8, width=80)
        status_scrollbar = ttk.Scrollbar(status_frame, orient="vertical", command=self.status_text.yview)
        self.status_text.configure(yscrollcommand=status_scrollbar.set)
        
        self.status_text.grid(row=0, column=0, sticky=(tk.W, tk.E))
        status_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
    def browse_pdf(self):
        filename = filedialog.askopenfilename(
            title="Select PDF File",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if filename:
            self.pdf_path.set(filename)
            # Auto-suggest excel filename
            excel_name = os.path.splitext(filename)[0] + "_telecast_data.xlsx"
            self.excel_path.set(excel_name)

    def browse_excel(self):
        filename = filedialog.asksaveasfilename(
            title="Save Excel File As",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if filename:
            self.excel_path.set(filename)

    def log_status(self, message):
        """Add message to status text widget"""
        self.status_text.insert(tk.END, f"{datetime.now().strftime('%H:%M:%S')} - {message}\n")
        self.status_text.see(tk.END)
        self.root.update_idletasks()

    def update_preview(self):
        """Update the preview treeview with extracted data"""
        # Clear existing items
        for item in self.preview_tree.get_children():
            self.preview_tree.delete(item)
        
        # Add new items - CORRECTED column access
        for record in self.extracted_data:
            self.preview_tree.insert('', 'end', values=(
                record['Date'],
                record['Program'][:30] + '...' if len(record['Program']) > 30 else record['Program'],
                record['Advt_time'],
                record['Advt_theme'][:50] + '...' if len(record['Advt_theme']) > 50 else record['Advt_theme'],
                record['Dur']
            ))
        
        # Update record count
        self.records_count_var.set(f"Total Records: {len(self.extracted_data)}")
        
        # Enable/disable export button
        if self.extracted_data:
            self.export_btn.config(state=tk.NORMAL)
        else:
            self.export_btn.config(state=tk.DISABLED)

    def clear_data(self):
        """Clear all extracted data and preview"""
        self.extracted_data = []
        self.update_preview()
        self.status_text.delete(1.0, tk.END)
        self.log_status("Data cleared")
        messagebox.showinfo("Cleared", "All extracted data has been cleared.")

    def preprocess_time_format(self, time_value):
        """Extract and standardize time format from various inputs."""
        try:
            if pd.isna(time_value):
                return '00:00:00'
                
            if isinstance(time_value, str):
                # Handle the telecast format like "20:08:35:07" -> "20:08:35"
                # First try to match the full telecast time format (HH:MM:SS:FF)
                time_pattern = re.search(r'(\d{1,2}:\d{2}:\d{2}):\d{2}', str(time_value))
                if time_pattern:
                    return time_pattern.group(1)
                
                # Then try standard time patterns
                time_pattern = re.search(r'(\d{1,2}:\d{2}(?::\d{2})?)', str(time_value))
                if time_pattern:
                    time_str = time_pattern.group(1)
                    if time_str.count(':') == 1:
                        time_str += ':00'
                    return time_str
                return '00:00:00'
            elif hasattr(time_value, 'strftime'):
                return time_value.strftime('%H:%M:%S')
            return '00:00:00'
        except Exception:
            return '00:00:00'

    def extract_telecast_data(self, text):
        """Extract telecast data from the text"""
        data = []
        self.log_status("Extracting telecast data with programme field and time formatting...")
        
        lines = text.split('\n')
        
        # Pattern to match: Date Programme Time Advertisement_Theme Duration
        # Example: 05/02/2025 ROCKY 20:08:35:07 SLT MOBITEL MAX PLUS TVC - 20SEC 20
        
        for line_num, line in enumerate(lines, 1):
            line = line.strip()
            
            # Skip header lines and short lines
            if not line or len(line) < 20:
                continue
                
            # Skip obvious header lines
            if any(keyword in line.upper() for keyword in ['CHANNEL', 'INVOICE', 'BRAND', 'ADDRESS', 'TELECAST CERTIFICATE', 'PROGRAMME', 'CAPTION', 'DATE']):
                continue
            
            # Pattern for the actual data format: Date Programme Time Advertisement Duration
            # More flexible pattern to handle programme names with spaces
            pattern = r'(\d{2}/\d{2}/\d{4})\s+(.+?)\s+(\d{2}:\d{2}:\d{2}:\d{2})\s+(.+?)\s+(\d+)$'
            
            match = re.search(pattern, line)
            if match:
                try:
                    date = match.group(1)
                    programme = match.group(2).strip()
                    raw_advt_time = match.group(3)
                    advt_theme = match.group(4).strip()
                    duration = match.group(5)
                    
                    # Process the time format
                    advt_time = self.preprocess_time_format(raw_advt_time)
                    
                    # Validate the extracted data
                    if (re.match(r'\d{2}/\d{2}/\d{4}', date) and 
                        programme and advt_theme and 
                        duration.isdigit()):
                        
                        # Clean up fields (remove extra spaces)
                        programme = re.sub(r'\s+', ' ', programme).strip()
                        advt_theme = re.sub(r'\s+', ' ', advt_theme).strip()
                        
                        # CORRECTED: Use exact column names
                        data.append({
                            'Date': date,
                            'Program': programme,  
                            'Advt_time': advt_time,  
                            'Advt_theme': advt_theme,
                            'Dur': int(duration)
                        })
                        
                        self.log_status(f"✓ {date} | {programme[:15]}... | {raw_advt_time}→{advt_time} | {advt_theme[:20]}... | {duration}")
                        
                except Exception as e:
                    self.log_status(f"Error parsing line {line_num}: {str(e)}")
                    continue
            else:
                # Try alternative patterns for edge cases
                alt_patterns = [
                    # Pattern with extra spaces
                    r'(\d{2}/\d{2}/\d{4})\s+(.+?)\s+(\d{1,2}:\d{2}:\d{2}:\d{2})\s+(.+?)\s+(\d+)\s*$',
                    # Pattern where time might have single digit hour
                    r'(\d{2}/\d{2}/\d{4})\s+(.+?)\s+(\d{1,2}:\d{2}:\d{2}:\d{2})\s+(.+?)\s+(\d+)$',
                ]
                
                for alt_pattern in alt_patterns:
                    alt_match = re.search(alt_pattern, line)
                    if alt_match:
                        try:
                            date = alt_match.group(1)
                            programme = alt_match.group(2).strip()
                            raw_advt_time = alt_match.group(3)
                            advt_theme = alt_match.group(4).strip()
                            duration = alt_match.group(5)
                            
                            # Process the time format
                            advt_time = self.preprocess_time_format(raw_advt_time)
                            
                            if (re.match(r'\d{2}/\d{2}/\d{4}', date) and 
                                programme and advt_theme and 
                                duration.isdigit()):
                                
                                programme = re.sub(r'\s+', ' ', programme).strip()
                                advt_theme = re.sub(r'\s+', ' ', advt_theme).strip()
                                
                                # CORRECTED: Use exact column names
                                data.append({
                                    'Date': date,
                                    'Program': programme,  # Changed from 'Programme' to 'Program'
                                    'Advt_time': advt_time,  # Changed from 'Advt_Time' to 'Advt_time'
                                    'Advt_theme': advt_theme,
                                    'Dur': int(duration)
                                })
                                
                                self.log_status(f"✓ Alt: {date} | {programme[:15]}... | {raw_advt_time}→{advt_time} | {advt_theme[:20]}... | {duration}")
                                break
                                
                        except Exception as e:
                            continue
        
        self.log_status(f"Successfully extracted {len(data)} records")
        return data

    def extract_from_pdf(self, pdf_path):
        """Extract data from PDF"""
        all_data = []
        
        try:
            with pdfplumber.open(pdf_path) as pdf:
                self.log_status(f"Processing {len(pdf.pages)} pages...")
                
                all_text = ""
                
                for page_num, page in enumerate(pdf.pages, 1):
                    self.log_status(f"Processing page {page_num}...")
                    
                    # Extract text from the page
                    page_text = page.extract_text()
                    if page_text:
                        all_text += page_text + "\n"
                
                # Extract data from all text
                if all_text:
                    extracted_data = self.extract_telecast_data(all_text)
                    all_data.extend(extracted_data)
                
                return all_data
                
        except Exception as e:
            self.log_status(f"Error extracting from PDF: {str(e)}")
            return []

    def extract_only(self):
        """Extract data from PDF without immediately exporting"""
        if not self.pdf_path.get():
            messagebox.showerror("Error", "Please select a PDF file")
            return
        
        try:
            self.progress.start()
            self.status_text.delete(1.0, tk.END)  # Clear previous logs
            self.log_status("Starting PDF extraction...")
            
            # Extract data from PDF
            data = self.extract_from_pdf(self.pdf_path.get())
            
            if not data:
                self.log_status("No data extracted from PDF.")
                self.extracted_data = []
                messagebox.showwarning("No Data", "No data could be extracted from the PDF file.\nPlease check the PDF format.")
            else:
                # Remove duplicates
                df = pd.DataFrame(data)
                df = df.drop_duplicates().reset_index(drop=True)
                
                # Sort by date and time - CORRECTED column name
                try:
                    df['datetime_sort'] = pd.to_datetime(df['Date'] + ' ' + df['Advt_time'], 
                                                        format='%m/%d/%Y %H:%M:%S', errors='coerce')
                    df = df.sort_values('datetime_sort').drop('datetime_sort', axis=1)
                    df = df.reset_index(drop=True)
                except:
                    pass
                
                self.extracted_data = df.to_dict('records')
                self.log_status(f"Extraction completed! Found {len(self.extracted_data)} unique records.")
                
                messagebox.showinfo("Extraction Complete", 
                                f"✅ Successfully extracted {len(self.extracted_data)} records!\n\n"
                                f"📋 Data is now displayed in the preview below.\n"
                                f"💾 Click 'Export to Excel' to save the data.")
            
            # Update preview regardless of success/failure
            self.update_preview()
            
        except Exception as e:
            error_msg = f"Extraction failed: {str(e)}"
            self.log_status(error_msg)
            messagebox.showerror("Error", error_msg)
        
        finally:
            self.progress.stop()

    def export_to_excel(self):
        """Export the extracted data to Excel"""
        if not self.extracted_data:
            messagebox.showerror("Error", "No data to export. Please extract data from PDF first.")
            return
        
        if not self.excel_path.get():
            messagebox.showerror("Error", "Please specify output Excel file")
            return
        
        try:
            self.log_status("Exporting to Excel...")
            df = pd.DataFrame(self.extracted_data)
            
            # Save to Excel with formatting
            with pd.ExcelWriter(self.excel_path.get(), engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='Telecast Data', index=False)
                
                # Format the Excel sheet
                worksheet = writer.sheets['Telecast Data']
                
                # Set column widths - CORRECTED comments
                column_widths = {
                    'A': 12,  # Date
                    'B': 30,  # Program
                    'C': 12,  # Advt_time
                    'D': 45,  # Advt_theme
                    'E': 8    # Dur
                }
                
                for col, width in column_widths.items():
                    worksheet.column_dimensions[col].width = width
                
                # Format header row
                try:
                    from openpyxl.styles import Font, PatternFill
                    header_font = Font(bold=True)
                    header_fill = PatternFill(start_color="CCE5FF", end_color="CCE5FF", fill_type="solid")
                    
                    for cell in worksheet[1]:
                        cell.font = header_font
                        cell.fill = header_fill
                except:
                    pass
            
            self.log_status(f"✅ Excel file saved: {self.excel_path.get()}")
            
            # CORRECTED success message
            messagebox.showinfo("Export Success", 
                            f"📊 Excel file created successfully!\n\n"
                            f"📄 Records exported: {len(df)}\n"
                            f"📁 File location: {self.excel_path.get()}\n"
                            f"📋 Columns: Date | Program | Advt_time | Advt_theme | Dur")
            
        except Exception as e:
            error_msg = f"Export failed: {str(e)}"
            self.log_status(error_msg)
            messagebox.showerror("Error", error_msg)

def main():
    try:
        import pdfplumber
        import pandas
        import openpyxl
    except ImportError as e:
        missing_package = str(e).split("'")[1]
        print(f"Required package '{missing_package}' is not installed.")
        print("Please install required packages using:")
        print("pip install pdfplumber pandas openpyxl")
        return

    root = tk.Tk()
    app = PDFToExcelConverter(root)
    root.mainloop()

if __name__ == "__main__":
    main()