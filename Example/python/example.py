# ============================================================
# AutoDWG DWG2PDF SDK - Python Examples
#
# Prerequisites:
#   1. Register the COM DLL (run as Administrator):
#      64-bit: regsvr32 "path\to\dwgtopdfx64.dll"
#      32-bit: regsvr32 "path\to\dwgtopdfx.dll"
#   2. Install pywin32:
#      pip install pywin32
#
# Run:
#   python example.py
# ============================================================

import os
import sys
import glob

try:
    import win32com.client
except ImportError:
    print("ERROR: pywin32 is required. Install with: pip install pywin32")
    sys.exit(1)


def create_converter():
    """Create and return a DWGTOPDFX COM object."""
    try:
        return win32com.client.Dispatch("DWGTOPDFX.ConvertPDF")
    except Exception as e:
        print("ERROR: Cannot create COM object 'DWGTOPDFX.ConvertPDF'.")
        print("Please register the DLL first: regsvr32 dwgtopdfx64.dll")
        print(f"Details: {e}")
        sys.exit(1)


def configure_basic(obj):
    """Apply basic conversion configuration."""
    obj.ConvertNotPlotLayer = 0       # 0=Skip non-printing layers, 1=Convert them
    obj.Width = 240                   # Page width in mm
    obj.Height = 180                  # Page height in mm
    obj.ColorMode = 0                 # 0=256 colors, 1=Black & White
    obj.DPI = 300                     # Output resolution (uniform)
    # obj.DPIWidth = 300              # (Optional) Separate horizontal DPI
    # obj.DPIHeight = 300             # (Optional) Separate vertical DPI
    obj.OutWithTTF = 1                # 1=Preserve TrueType fonts
    obj.EmbeddedSearchText = 0        # 0=Keep text searchable
    obj.zoomtype = 0                  # 0=Zoom extents, 1=Last view
    obj.OutputByLayoutSize = 0        # 0=Use Width/Height, 1=Use layout dims
    obj.nBackgroundColor = 7   # Background color (7=white, 0=black)
    obj.AddFontPath(r"C:\windows\system32\shxfont")
    obj.PSPath = r"c:\windows\system32\ps\lib;c:\windows\system32\ps\fonts"


# ============================================================
# Example 1: Basic single file conversion
# ============================================================
def example1_basic():
    """Convert all layouts of a DWG file into a single PDF."""
    input_file = r"D:\input\dwg\Office.dwg"
    output_file = r"D:\output\pdf\output.pdf"

    os.makedirs(os.path.dirname(output_file), exist_ok=True)

    obj = create_converter()
    try:
        configure_basic(obj)
        obj.PenWidth(1)  # Access pen width property (index 1)
        obj.InputFile = input_file

        print(f"Converting: {input_file}")
        obj.Convert(output_file)
        print(f"Done: {output_file}")
    except Exception as e:
        print(f"Conversion failed. (Error: {e})")
    finally:
        del obj


# ============================================================
# Example 2: One PDF per layout/view
# ============================================================
def example2_one_pdf_per_view():
    """Convert each layout/view to a separate PDF.
    Demonstrates SetOutColor for color index mapping."""
    input_file = r"D:\input\dwg\Office.dwg"
    output_dir = r"D:\output\pdf"
    os.makedirs(output_dir, exist_ok=True)

    obj = create_converter()
    try:
        obj.Addexcudelayer("Defpoints")
        obj.ConvertNotPlotLayer = 0
        obj.Width = 300
        obj.Height = 200
        obj.ColorMode = 1
        obj.PSPath = r"c:\windows\system32\ps\lib;c:\windows\system32\ps\fonts"
        obj.InputFile = input_file

        # Set pen widths for color indices 1-7
        for i in range(1, 8):
            obj.PenWidth(i)

        # Color mapping: map input color index to output color
        # SetOutColor(inputColorIndex, outputColorIndex)
        obj.SetOutColor(1, 5)        # Map red to blue
        # obj.SetOutColor(2, 3)      # Map yellow to green

        view_count = obj.ViewCount
        print(f"Found {view_count} view(s).")

        for i in range(1, view_count + 1):
            view_name = obj.ViewName(i)
            page_setup = obj.PageSetupName(i)
            unit = obj.PageSizeUnit(i)
            w = obj.PageWidth(i)
            h = obj.PageHeight(i)

            print(f"  [{i}] View: {view_name}, Setup: {page_setup}, "
                  f"Unit: {unit}, Size: {w}x{h}mm")

            output_file = os.path.join(output_dir, f"{view_name}.pdf")
            obj.ConvertView(output_file, view_name)

        print("All views converted.")
    except Exception as e:
        print(f"Conversion failed. (Error: {e})")
    finally:
        del obj


# ============================================================
# Example 3: Batch conversion of a folder
# ============================================================
def example3_batch_conversion(input_dir, output_dir):
    """Convert all DWG files in a folder.
    Approach A: Iterate folder, convert each DWG with all views.
    Approach B: Use AddFile + ConvertFiles (commented out)."""
    os.makedirs(output_dir, exist_ok=True)

    dwg_files = glob.glob(os.path.join(input_dir, "*.dwg"))
    if not dwg_files:
        print(f"No DWG files found in: {input_dir}")
        return

    print(f"Found {len(dwg_files)} DWG file(s).")

    obj = create_converter()
    try:
        # --- Approach A: Convert each file individually ---
        for dwg_file in dwg_files:
            obj.Addexcudelayer("Defpoints")
            obj.Width = 1024
            obj.Height = 768
            obj.ColorMode = 1
            obj.InputFile = dwg_file
            obj.PenWidth(7)  # Pen width for index 7
            obj.PSPath = r"c:\windows\system32\ps\lib;c:\windows\system32\ps\fonts"

            view_count = obj.ViewCount
            base_name = os.path.splitext(os.path.basename(dwg_file))[0]

            for i in range(1, view_count + 1):
                view_name = obj.ViewName(i)
                output_file = os.path.join(output_dir,
                                           f"{base_name}_{view_name}.pdf")
                obj.ConvertView(output_file, view_name)

            print(f"  Converted: {os.path.basename(dwg_file)} ({view_count} views)")

        # --- Approach B: Use AddFile + ConvertFiles ---
        # for dwg_file in dwg_files:
        #     obj.AddFile(dwg_file)
        # obj.Width = 1024
        # obj.Height = 768
        # obj.ColorMode = 1
        # obj.PSPath = r"c:\windows\system32\ps\lib;c:\windows\system32\ps\fonts"
        # obj.ConvertFiles(os.path.join(output_dir, "batch_output.pdf"))

        print("Batch conversion complete.")
    except Exception as e:
        print(f"Conversion failed. (Error: {e})")
    finally:
        del obj


# ============================================================
# Example 4: Advanced features (CTB, scale, ConvertPapers)
# ============================================================
def example4_advanced():
    """Demonstrates CTB, scale, ConvertPapers, background color."""
    input_file = r"D:\input\dwg\Office.dwg"
    output_file = r"D:\output\pdf\output_papers.pdf"
    os.makedirs(os.path.dirname(output_file), exist_ok=True)

    obj = create_converter()
    try:
        obj.ConvertNotPlotLayer = 0
        obj.Width = 914
        obj.Height = 610
        obj.ColorMode = 1
        obj.DPI = 100
        obj.OutWithTTF = 1
        obj.EmbeddedSearchText = 0
        obj.ZoomType = 0
        obj.nBackgroundColor = 7   # White background

        obj.SetCTBFile("acad.ctb")       # CTB color book
        obj.SetOutputScale(1.0)          # 1:1 scale (float parameter)

        # Pen widths for different color indices
        pen_widths = {1: 0.02, 2: 0.03, 3: 0.05, 4: 0.08,
                      5: 0.10, 6: 0.15, 7: 0.20}
        for idx, width in pen_widths.items():
            obj.PenWidth(idx)

        obj.PSPath = r"c:\windows\system32\ps\lib;c:\windows\system32\ps\fonts"
        obj.AddFontPath(r"C:\windows\system32\shxfont")
        obj.InputFile = input_file

        print("Converting (papers only, no Model Space)...")
        obj.ConvertPapers(output_file)
        print(f"Done: {output_file}")
    except Exception as e:
        print(f"Conversion failed. (Error: {e})")
    finally:
        del obj


# ============================================================
# Example 5: DWF to DWG/DXF conversion
# ============================================================
def example5_dwf_to_dwg():
    """Convert DWF file to DWG or DXF using ConvertDWFtoDWGDXF."""
    input_file = r"D:\input\dwf\drawing.dwf"
    output_file = r"D:\output\dwg\output.dwg"
    out_version = 24  # Output DWG version code

    os.makedirs(os.path.dirname(output_file), exist_ok=True)

    obj = create_converter()
    try:
        print(f"Converting DWF to DWG...")
        print(f"  Input:  {input_file}")
        print(f"  Output: {output_file}")

        obj.ConvertDWFtoDWGDXF(input_file, output_file, out_version)
        print(f"Done: {output_file}")

        # To convert to DXF instead:
        # dxf_output = os.path.splitext(output_file)[0] + ".dxf"
        # obj.ConvertDWFtoDWGDXF(input_file, dxf_output, out_version)
    except Exception as e:
        print(f"Conversion failed. (Error: {e})")
    finally:
        del obj


if __name__ == "__main__":
    print("=== AutoDWG DWG2PDF SDK - Python Examples ===\n")

    # Uncomment the example you want to run:
    example1_basic()
    # example2_one_pdf_per_view()
    # example3_batch_conversion(r"D:\input\dwg", r"D:\output\pdf")
    # example4_advanced()
    # example5_dwf_to_dwg()

    print("\nDone.")