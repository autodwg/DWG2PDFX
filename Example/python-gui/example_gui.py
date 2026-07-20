# ============================================================
# AutoDWG DWG2PDF SDK - Python GUI Application (tkinter)
#
# Prerequisites:
#   1. Register the COM DLL (run as Administrator):
#      64-bit: regsvr32 "path\to\dwgtopdfx64.dll"
#      32-bit: regsvr32 "path\to\dwgtopdfx.dll"
#   2. Install pywin32:
#      pip install pywin32
#
# Run:
#   python example_gui.py
# ============================================================

import os
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime

try:
    import win32com.client
except ImportError:
    print("ERROR: pywin32 is required. Install with: pip install pywin32")
    sys.exit(1)


class DWG2PDFApp:
    def __init__(self, root):
        self.root = root
        self.root.title("AutoDWG DWG to PDF Converter")
        self.root.geometry("600x620")
        self.root.resizable(False, False)

        self._build_ui()

    def _build_ui(self):
        pad = {"padx": 10, "pady": 3}

        # --- Input File ---
        row = ttk.Frame(self.root)
        row.pack(fill="x", **pad)
        ttk.Label(row, text="Input DWG:").pack(side="left")
        self.txt_input = ttk.Entry(row, width=42)
        self.txt_input.pack(side="left", padx=(5, 5))
        ttk.Button(row, text="Browse...", command=self._browse_input).pack(side="left")

        # --- Output File ---
        row = ttk.Frame(self.root)
        row.pack(fill="x", **pad)
        ttk.Label(row, text="Output PDF:").pack(side="left")
        self.txt_output = ttk.Entry(row, width=42)
        self.txt_output.pack(side="left", padx=(5, 5))
        ttk.Button(row, text="Browse...", command=self._browse_output).pack(side="left")

        # --- Separator ---
        ttk.Separator(self.root, orient="horizontal").pack(fill="x", padx=10, pady=8)
        ttk.Label(self.root, text="Conversion Parameters",
                  font=("Segoe UI", 9, "bold")).pack(anchor="w", padx=10)

        # --- Parameters Frame ---
        params = ttk.Frame(self.root)
        params.pack(fill="x", padx=10, pady=5)

        # Row 1: Color Mode + DPI
        r1 = ttk.Frame(params)
        r1.pack(fill="x", pady=2)
        ttk.Label(r1, text="Color Mode:").pack(side="left")
        self.cmb_color = ttk.Combobox(r1, values=["256 Colors", "Black & White"],
                                       state="readonly", width=12)
        self.cmb_color.current(0)
        self.cmb_color.pack(side="left", padx=(5, 20))
        ttk.Label(r1, text="DPI:").pack(side="left")
        self.spn_dpi = ttk.Spinbox(r1, from_=72, to=1200, width=6)
        self.spn_dpi.set(300)
        self.spn_dpi.pack(side="left", padx=5)

        # Row 2: Width + Height
        r2 = ttk.Frame(params)
        r2.pack(fill="x", pady=2)
        ttk.Label(r2, text="Width (mm):").pack(side="left")
        self.spn_width = ttk.Spinbox(r2, from_=50, to=5000, width=6)
        self.spn_width.set(297)
        self.spn_width.pack(side="left", padx=(5, 20))
        ttk.Label(r2, text="Height (mm):").pack(side="left")
        self.spn_height = ttk.Spinbox(r2, from_=50, to=5000, width=6)
        self.spn_height.set(210)
        self.spn_height.pack(side="left", padx=5)

        # Row 3: Checkboxes
        r3 = ttk.Frame(params)
        r3.pack(fill="x", pady=2)
        self.chk_ttf = tk.BooleanVar(value=True)
        ttk.Checkbutton(r3, text="Output TrueType Fonts", variable=self.chk_ttf).pack(side="left")
        self.chk_embed = tk.BooleanVar(value=False)
        ttk.Checkbutton(r3, text="Embed as Image", variable=self.chk_embed).pack(side="left", padx=20)

        # Row 4: Zoom + Layout
        r4 = ttk.Frame(params)
        r4.pack(fill="x", pady=2)
        ttk.Label(r4, text="Zoom:").pack(side="left")
        self.cmb_zoom = ttk.Combobox(r4, values=["Zoom Extents", "Last View"],
                                      state="readonly", width=12)
        self.cmb_zoom.current(0)
        self.cmb_zoom.pack(side="left", padx=(5, 20))
        self.chk_layout = tk.BooleanVar(value=False)
        ttk.Checkbutton(r4, text="Output by Layout Size", variable=self.chk_layout).pack(side="left")

        # Row 5: Font Path
        r5 = ttk.Frame(params)
        r5.pack(fill="x", pady=2)
        ttk.Label(r5, text="Font Path:").pack(side="left")
        self.txt_font = ttk.Entry(r5, width=42)
        self.txt_font.insert(0, r"C:\windows\system32\shxfont")
        self.txt_font.pack(side="left", padx=5)

        # Row 6: PS Path
        r6 = ttk.Frame(params)
        r6.pack(fill="x", pady=2)
        ttk.Label(r6, text="PS Path:").pack(side="left")
        self.txt_ps = ttk.Entry(r6, width=42)
        self.txt_ps.insert(0, r"c:\windows\system32\ps\lib;c:\windows\system32\ps\fonts")
        self.txt_ps.pack(side="left", padx=5)

        # --- Separator ---
        ttk.Separator(self.root, orient="horizontal").pack(fill="x", padx=10, pady=8)

        # --- Convert Button ---
        btn_frame = ttk.Frame(self.root)
        btn_frame.pack(fill="x", padx=10)
        self.btn_convert = ttk.Button(btn_frame, text="Convert to PDF", command=self._convert)
        self.btn_convert.pack(side="left")
        self.lbl_status = ttk.Label(btn_frame, text="Ready.", foreground="green")
        self.lbl_status.pack(side="left", padx=20)

        # --- Log ---
        log_frame = ttk.Frame(self.root)
        log_frame.pack(fill="both", expand=True, padx=10, pady=5)
        self.txt_log = tk.Text(log_frame, height=6, state="disabled",
                                bg="#f5f5f5", font=("Consolas", 9))
        scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.txt_log.yview)
        self.txt_log.configure(yscrollcommand=scrollbar.set)
        self.txt_log.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

    def _browse_input(self):
        path = filedialog.askopenfilename(
            title="Select DWG File",
            filetypes=[("DWG Files", "*.dwg"), ("All Files", "*.*")]
        )
        if path:
            self.txt_input.delete(0, tk.END)
            self.txt_input.insert(0, path)
            # Auto-set output path
            if not self.txt_output.get():
                self.txt_output.insert(0, os.path.splitext(path)[0] + ".pdf")

    def _browse_output(self):
        initial_dir = ""
        initial_file = ""
        inp = self.txt_input.get()
        if inp:
            initial_dir = os.path.dirname(inp)
            initial_file = os.path.splitext(os.path.basename(inp))[0] + ".pdf"

        path = filedialog.asksaveasfilename(
            title="Save PDF As",
            filetypes=[("PDF Files", "*.pdf")],
            initialdir=initial_dir if initial_dir else "",
            initialfile=initial_file
        )
        if path:
            self.txt_output.delete(0, tk.END)
            self.txt_output.insert(0, path)

    def _log(self, message):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.txt_log.configure(state="normal")
        self.txt_log.insert(tk.END, f"[{timestamp}] {message}\n")
        self.txt_log.see(tk.END)
        self.txt_log.configure(state="disabled")

    def _convert(self):
        input_file = self.txt_input.get().strip()
        output_file = self.txt_output.get().strip()

        if not input_file:
            messagebox.showwarning("Warning", "Please select an input DWG file.")
            return
        if not os.path.exists(input_file):
            messagebox.showerror("Error", f"Input file does not exist:\n{input_file}")
            return
        if not output_file:
            messagebox.showwarning("Warning", "Please specify an output PDF path.")
            return

        # Create output directory
        out_dir = os.path.dirname(output_file)
        if out_dir and not os.path.exists(out_dir):
            os.makedirs(out_dir)

        self.btn_convert.configure(state="disabled")
        self.lbl_status.configure(text="Converting...", foreground="orange")
        self.txt_log.configure(state="normal")
        self.txt_log.delete("1.0", tk.END)
        self.txt_log.configure(state="disabled")

        try:
            self._log("Creating COM object...")
            obj = win32com.client.Dispatch("DWGTOPDFX.ConvertPDF")

            # Set parameters
            obj.ColorMode = self.cmb_color.current()
            obj.DPI = int(self.spn_dpi.get())
            obj.OutWithTTF = 1 if self.chk_ttf.get() else 0
            obj.EmbeddedSearchText = 1 if self.chk_embed.get() else 0
            obj.zoomtype = self.cmb_zoom.current()
            obj.OutputByLayoutSize = 1 if self.chk_layout.get() else 0
            obj.nBackgroundColor = 7   # White background
            obj.Width = int(self.spn_width.get())
            obj.Height = int(self.spn_height.get())

            font_path = self.txt_font.get().strip()
            if font_path:
                obj.AddFontPath(font_path)
            ps_path = self.txt_ps.get().strip()
            if ps_path:
                obj.PSPath = ps_path

            obj.InputFile = input_file

            self._log(f"Input:  {input_file}")
            self._log(f"Output: {output_file}")
            self._log("Converting...")

            obj.Convert(output_file)

            self._log("Conversion complete!")
            self.lbl_status.configure(text="Done!", foreground="green")

            # Show result with output file location
            result = messagebox.askyesno(
                "Success",
                f"Conversion complete!\n\nOutput file:\n{output_file}\n\nOpen the output folder?"
            )
            if result:
                os.system(f'explorer.exe /select,"{output_file}"')

            del obj

        except Exception as e:
            self._log(f"ERROR: {e}")
            self.lbl_status.configure(text="Failed!", foreground="red")
            messagebox.showerror("Error", f"Conversion failed. Please check the input file and settings, then try again.\\n\\nError: {e}")

        finally:
            self.btn_convert.configure(state="normal")


if __name__ == "__main__":
    root = tk.Tk()
    app = DWG2PDFApp(root)
    root.mainloop()