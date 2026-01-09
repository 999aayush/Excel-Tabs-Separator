import sys
import os
import threading
import tkinter as tk
from tkinter import filedialog, scrolledtext, messagebox

# --- 1. SETUP PATHS ---

if getattr(sys, 'frozen', False):
    application_path = os.path.dirname(sys.executable)
else:
    application_path = os.path.dirname(os.path.abspath(__file__))

os.chdir(application_path)

# --- 2. IMPORT EXCEL ENGINE ---
try:
    import win32com.client as win32
except ImportError:

    root = tk.Tk()
    root.withdraw()
    messagebox.showerror("Setup Error", 
        "Required libraries not found.\n\n"
        "Please use 'run.bat' to launch this script,\n"
        "or ensure you have installed 'pywin32'.")
    sys.exit()

# --- 3. PROCESSING LOGIC ---
def run_separation_process(input_path, log_func, finish_callback):
    success = False
    excel = None
    wb = None
    
    try:
        if not os.path.exists(input_path):
            log_func(f"ERROR: File not found: {input_path}")
            return

        input_path = os.path.abspath(input_path)
        base_directory = os.path.dirname(input_path)
        file_name_only = os.path.splitext(os.path.basename(input_path))[0]
        output_dir = os.path.join(base_directory, file_name_only)

        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
            log_func(f"Created folder: {output_dir}")
        else:
            log_func(f"Using existing folder: {output_dir}")

        log_func(f"Starting Excel Engine...")
        
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False 
        
        log_func(f"Opening: {os.path.basename(input_path)}...")
        wb = excel.Workbooks.Open(input_path)
        
        total_sheets = wb.Sheets.Count
        log_func(f"Found {total_sheets} tabs. Starting separation...")

        for i in range(1, total_sheets + 1):
            sheet = wb.Sheets(i)
            sheet_name = sheet.Name
            
            log_func(f"[{i}/{total_sheets}] Processing: {sheet_name}...")
            
            safe_name = "".join([c for c in sheet_name if c.isalnum() or c in (' ', '-', '_', '.')]).strip()
            final_output_path = os.path.join(output_dir, f"{safe_name}.xlsx")
            
            sheet.Copy() 
            new_wb = excel.ActiveWorkbook
            
            try:
                new_wb.SaveAs(final_output_path, FileFormat=51)
            except Exception as e:
                log_func(f"Error saving {sheet_name}: {e}")
            finally:
                new_wb.Close(SaveChanges=False)

        log_func("-" * 30)
        log_func("SUCCESS! All tabs separated.")
        log_func(f"Output: {output_dir}")
        success = True

    except Exception as e:
        log_func(f"CRITICAL ERROR: {e}")
        log_func("Ensure Excel is installed and not busy.")
        success = False
    finally:
        if wb:
            try:
                wb.Close(SaveChanges=False)
            except:
                pass
        if excel:
            try:
                excel.Quit()
            except:
                pass
        
        finish_callback(success)

# --- 4. GUI CLASS ---
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Universal Excel Tab Separator")
        self.geometry("600x450")
        self.configure(bg="#2b2b2b")
        
        self.style_bg = "#2b2b2b"
        self.style_fg = "#ffffff"
        self.style_input_bg = "#3d3d3d"
        self.style_btn_bg = "#007acc"

        frame_top = tk.Frame(self, bg=self.style_bg)
        frame_top.pack(fill="x", padx=10, pady=10)

        tk.Label(frame_top, text="Excel File (.xlsx, .xls, .xlsb):", bg=self.style_bg, fg=self.style_fg).pack(anchor="w")
        
        self.entry_path = tk.Entry(frame_top, bg=self.style_input_bg, fg=self.style_fg, insertbackground="white")
        self.entry_path.pack(side="left", fill="x", expand=True, padx=(0, 10))
        
        btn_browse = tk.Button(frame_top, text="Browse...", command=self.browse_file, 
                               bg="#444", fg="white", activebackground="#666", activeforeground="white")
        btn_browse.pack(side="right")

        self.btn_run = tk.Button(self, text="Run Separation", command=self.start_thread,
                                 bg=self.style_btn_bg, fg="white", font=("Segoe UI", 10, "bold"),
                                 activebackground="#005f9e", activeforeground="white", height=2)
        self.btn_run.pack(fill="x", padx=10, pady=5)

        tk.Label(self, text="Terminal Output:", bg=self.style_bg, fg="#aaaaaa").pack(anchor="w", padx=10)
        self.log_window = scrolledtext.ScrolledText(self, height=10, bg="#1e1e1e", fg="#00ff00", 
                                                    font=("Consolas", 9), state="disabled")
        self.log_window.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        if len(sys.argv) > 1:
            self.entry_path.insert(0, sys.argv[1])

    def log(self, message):
        self.log_window.configure(state="normal")
        self.log_window.insert(tk.END, message + "\n")
        self.log_window.see(tk.END)
        self.log_window.configure(state="disabled")

    def browse_file(self):
        filename = filedialog.askopenfilename(
            initialdir=application_path,
            filetypes=[("Excel Files", "*.xlsx *.xls *.xlsb *.xlsm")]
        )
        if filename:
            self.entry_path.delete(0, tk.END)
            self.entry_path.insert(0, filename)

    def start_thread(self):
        path = self.entry_path.get().strip().strip('"')
        if not path:
            self.log("ERROR: Please select a file first.")
            return
        
        self.btn_run.config(state="disabled", text="Processing...")
        self.log_window.configure(state="normal")
        self.log_window.delete(1.0, tk.END)
        self.log_window.configure(state="disabled")
        
        t = threading.Thread(target=run_separation_process, args=(path, self.log_queue_wrapper, self.on_finish))
        t.daemon = True
        t.start()

    def log_queue_wrapper(self, msg):
        self.after(0, lambda: self.log(msg))

    def on_finish(self, success):
        self.after(0, lambda: self.handle_completion(success))

    def handle_completion(self, success):
        if success:
            self.withdraw()
            messagebox.showinfo("Complete", "Task Finished Successfully!")
            self.destroy()
            sys.exit()
        else:
            self.btn_run.config(state="normal", text="Run Separation")
            messagebox.showerror("Failed", "An error occurred.\nPlease check the terminal output below.")

if __name__ == "__main__":
    app = App()
    app.mainloop()