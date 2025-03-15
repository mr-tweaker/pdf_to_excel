import os
import sys
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from tkinter.scrolledtext import ScrolledText
import threading
import queue
import time
from datetime import datetime

# Import the converter module
# Make sure this file is in the same directory as the pdf_to_excel_converter.py file
from pdf_to_excel_converter import FinancialStatementConverter

class RedirectText:
    """Class to redirect stdout to a tkinter widget"""
    def __init__(self, text_widget):
        self.text_widget = text_widget
        self.queue = queue.Queue()
        self.updating = True
        threading.Thread(target=self.update_loop, daemon=True).start()
    
    def write(self, string):
        self.queue.put(string)
    
    def flush(self):
        pass
    
    def update_loop(self):
        while self.updating:
            try:
                while True:
                    # Get messages without waiting if queue is not empty
                    string = self.queue.get_nowait()
                    self.text_widget.configure(state='normal')
                    self.text_widget.insert('end', string)
                    self.text_widget.see('end')  # Scroll to the end
                    self.text_widget.configure(state='disabled')
                    self.queue.task_done()
            except queue.Empty:
                time.sleep(0.1)
    
    def stop(self):
        self.updating = False

class FinancialStatementConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Financial Statement PDF to Excel Converter")
        self.root.geometry("800x600")
        self.root.minsize(600, 400)
        
        # Set style
        style = ttk.Style()
        style.theme_use('clam')  # Can use 'clam', 'alt', 'default', 'classic'
        style.configure('TButton', font=('Arial', 10))
        style.configure('TLabel', font=('Arial', 10))
        style.configure('Header.TLabel', font=('Arial', 12, 'bold'))
        
        # Main frame
        main_frame = ttk.Frame(root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Header
        header_label = ttk.Label(
            main_frame, 
            text="Financial Statement PDF to Excel Converter", 
            style='Header.TLabel'
        )
        header_label.pack(pady=10)
        
        # Input frame
        input_frame = ttk.LabelFrame(main_frame, text="Select Files", padding="10")
        input_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # PDF file selection
        pdf_frame = ttk.Frame(input_frame)
        pdf_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(pdf_frame, text="PDF File:").pack(side=tk.LEFT, padx=5)
        
        self.pdf_path_var = tk.StringVar()
        pdf_entry = ttk.Entry(pdf_frame, textvariable=self.pdf_path_var, width=50)
        pdf_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        
        browse_pdf_btn = ttk.Button(pdf_frame, text="Browse...", command=self.browse_pdf)
        browse_pdf_btn.pack(side=tk.LEFT, padx=5)
        
        # Excel file selection
        excel_frame = ttk.Frame(input_frame)
        excel_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(excel_frame, text="Output Excel:").pack(side=tk.LEFT, padx=5)
        
        self.excel_path_var = tk.StringVar()
        excel_entry = ttk.Entry(excel_frame, textvariable=self.excel_path_var, width=50)
        excel_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        
        browse_excel_btn = ttk.Button(excel_frame, text="Browse...", command=self.browse_excel)
        browse_excel_btn.pack(side=tk.LEFT, padx=5)
        
        # Options frame
        options_frame = ttk.LabelFrame(main_frame, text="Options", padding="10")
        options_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # DPI option
        dpi_frame = ttk.Frame(options_frame)
        dpi_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(dpi_frame, text="OCR DPI (higher = better quality but slower):").pack(side=tk.LEFT, padx=5)
        
        self.dpi_var = tk.IntVar(value=300)
        dpi_spinner = ttk.Spinbox(dpi_frame, from_=100, to=600, increment=100, textvariable=self.dpi_var, width=5)
        dpi_spinner.pack(side=tk.LEFT, padx=5)
        
        # Statements to include
        statements_frame = ttk.Frame(options_frame)
        statements_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(statements_frame, text="Include:").pack(side=tk.LEFT, padx=5)
        
        self.include_bs = tk.BooleanVar(value=True)
        bs_check = ttk.Checkbutton(statements_frame, text="Balance Sheet", variable=self.include_bs)
        bs_check.pack(side=tk.LEFT, padx=5)
        
        self.include_pl = tk.BooleanVar(value=True)
        pl_check = ttk.Checkbutton(statements_frame, text="Profit & Loss", variable=self.include_pl)
        pl_check.pack(side=tk.LEFT, padx=5)
        
        self.include_cf = tk.BooleanVar(value=True)
        cf_check = ttk.Checkbutton(statements_frame, text="Cash Flow", variable=self.include_cf)
        cf_check.pack(side=tk.LEFT, padx=5)
        
        self.include_notes = tk.BooleanVar(value=True)
        notes_check = ttk.Checkbutton(statements_frame, text="Notes", variable=self.include_notes)
        notes_check.pack(side=tk.LEFT, padx=5)
        
        # Convert button
        convert_btn = ttk.Button(main_frame, text="Convert PDF to Excel", command=self.start_conversion)
        convert_btn.pack(pady=10)
        
        # Progress frame
        progress_frame = ttk.LabelFrame(main_frame, text="Progress", padding="10")
        progress_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Progress bar
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(
            progress_frame, 
            orient=tk.HORIZONTAL, 
            length=100, 
            mode='indeterminate',
            variable=self.progress_var
        )
        self.progress_bar.pack(fill=tk.X, padx=5, pady=5)
        
        # Log area
        self.log_text = ScrolledText(progress_frame, height=10, wrap=tk.WORD, state='disabled')
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Redirect stdout to log text
        self.stdout_redirect = RedirectText(self.log_text)
        sys.stdout = self.stdout_redirect
        
        # Status bar
        self.status_var = tk.StringVar(value="Ready")
        status_bar = ttk.Label(root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        
        # Set default output path
        self.set_default_paths()
        
        # Bind the close event
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
    
    def set_default_paths(self):
        """Set default paths for input and output files"""
        desktop = os.path.join(os.path.expanduser('~'), 'Desktop')
        
        # Get current date for filename
        date_str = datetime.now().strftime("%Y%m%d")
        
        self.pdf_path_var.set("")  # Empty by default
        self.excel_path_var.set(os.path.join(desktop, f"financial_statements_{date_str}.xlsx"))
    
    def browse_pdf(self):
        """Open file dialog to select PDF file"""
        pdf_file = filedialog.askopenfilename(
            title="Select Financial Statement PDF",
            filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")]
        )
        
        if pdf_file:
            self.pdf_path_var.set(pdf_file)
            
            # Auto-set the Excel output path based on the PDF name
            pdf_basename = os.path.splitext(os.path.basename(pdf_file))[0]
            excel_dir = os.path.dirname(self.excel_path_var.get())
            self.excel_path_var.set(os.path.join(excel_dir, f"{pdf_basename}.xlsx"))
    
    def browse_excel(self):
        """Open file dialog to select Excel output file"""
        excel_file = filedialog.asksaveasfilename(
            title="Save Excel File As",
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")],
            initialfile=os.path.basename(self.excel_path_var.get()),
            initialdir=os.path.dirname(self.excel_path_var.get())
        )
        
        if excel_file:
            self.excel_path_var.set(excel_file)
    
    def start_conversion(self):
        """Start the conversion process in a separate thread"""
        pdf_path = self.pdf_path_var.get()
        excel_path = self.excel_path_var.get()
        
        # Validate inputs
        if not pdf_path:
            messagebox.showerror("Error", "Please select a PDF file.")
            return
        
        if not excel_path:
            messagebox.showerror("Error", "Please specify an output Excel file.")
            return
        
        if not os.path.exists(pdf_path):
            messagebox.showerror("Error", f"PDF file does not exist: {pdf_path}")
            return
        
        # Check if at least one statement is selected
        if not (self.include_bs.get() or self.include_pl.get() or 
                self.include_cf.get() or self.include_notes.get()):
            messagebox.showerror("Error", "Please select at least one statement to include.")
            return
        
        # Check if output directory exists
        output_dir = os.path.dirname(excel_path)
        if output_dir and not os.path.exists(output_dir):
            try:
                os.makedirs(output_dir)
            except Exception as e:
                messagebox.showerror("Error", f"Cannot create output directory: {str(e)}")
                return
        
        # Start progress bar
        self.progress_bar.start()
        self.status_var.set("Converting...")
        
        # Clear log
        self.log_text.configure(state='normal')
        self.log_text.delete(1.0, tk.END)
        self.log_text.configure(state='disabled')
        
        # Log starting info
        print(f"Starting conversion at {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"Input PDF: {pdf_path}")
        print(f"Output Excel: {excel_path}")
        print(f"OCR DPI: {self.dpi_var.get()}")
        print(f"Statements to include: " + 
              ("Balance Sheet " if self.include_bs.get() else "") +
              ("Profit & Loss " if self.include_pl.get() else "") +
              ("Cash Flow " if self.include_cf.get() else "") +
              ("Notes" if self.include_notes.get() else ""))
        print("-" * 50)
        
        # Start conversion in a separate thread
        conversion_thread = threading.Thread(
            target=self.run_conversion,
            args=(pdf_path, excel_path),
            daemon=True
        )
        conversion_thread.start()
    
    def run_conversion(self, pdf_path, excel_path):
        """Run the conversion process (called in a separate thread)"""
        try:
            # Create converter
            converter = FinancialStatementConverter(pdf_path)
            
            # Create workbook structure based on selections
            if self.include_bs.get():
                print("Creating Balance Sheet worksheet...")
                converter.create_balance_sheet()
            
            if self.include_pl.get():
                print("Creating Profit & Loss worksheet...")
                converter.create_profit_loss()
            
            if self.include_cf.get():
                print("Creating Cash Flow worksheet...")
                converter.create_cash_flow()
            
            if self.include_notes.get():
                print("Creating Notes worksheet...")
                converter.create_notes()
            
            # Populate data
            print("Populating data...")
            converter.populate_manual_data()
            
            # Save to Excel
            print("Saving Excel file...")
            success = converter.save_excel(excel_path)
            
            # Show completion message
            if success:
                self.root.after(0, lambda: self.show_completion(True, excel_path))
            else:
                self.root.after(0, lambda: self.show_completion(False, "Failed to save Excel file."))
        
        except Exception as e:
            error_message = f"Error during conversion: {str(e)}"
            print(error_message)
            import traceback
            print(traceback.format_exc())
            self.root.after(0, lambda: self.show_completion(False, error_message))
    
    def show_completion(self, success, message):
        """Show completion message and reset UI state"""
        # Stop progress bar
        self.progress_bar.stop()
        
        if success:
            messagebox.showinfo("Conversion Complete", 
                                f"PDF has been successfully converted to Excel!\n\nFile saved to: {message}")
            self.status_var.set("Conversion completed successfully")
            
            # Ask if user wants to open the file
            if messagebox.askyesno("Open File", "Would you like to open the Excel file now?"):
                try:
                    import subprocess
                    if sys.platform == 'win32':
                        os.startfile(message)
                    elif sys.platform == 'darwin':  # macOS
                        subprocess.call(['open', message])
                    else:  # Linux
                        subprocess.call(['xdg-open', message])
                except Exception as e:
                    messagebox.showerror("Error", f"Could not open file: {str(e)}")
        else:
            messagebox.showerror("Conversion Failed", message)
            self.status_var.set("Conversion failed")
    
    def on_close(self):
        """Handle window close event"""
        # Restore stdout
        sys.stdout = sys.__stdout__
        
        # Stop the text redirection thread
        if hasattr(self, 'stdout_redirect'):
            self.stdout_redirect.stop()
        
        # Destroy the window
        self.root.destroy()

if __name__ == "__main__":
    # Create main window
    root = tk.Tk()
    app = FinancialStatementConverterApp(root)
    
    # Set window icon (if available)
    try:
        # If you have an icon file:
        # root.iconbitmap("icon.ico")  # On Windows
        pass
    except:
        pass
    
    # Start the main loop
    root.mainloop()