// ============================================================
// AutoDWG DWG2PDF SDK - C# WinForms GUI Application
//
// Prerequisites:
//   1. Register the COM DLL (run as Administrator):
//      64-bit: regsvr32 "path\to\dwgtopdfx64.dll"
//      32-bit: regsvr32 "path\to\dwgtopdfx.dll"
//   2. .NET 8.0 SDK
//
// Build:
//   dotnet build
//   dotnet run
// ============================================================

using System;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace AutoDWG.DWG2PDF
{
    public class MainForm : Form
    {
        // --- UI Controls ---
        private TextBox txtInputFile;
        private TextBox txtOutputFile;
        private ComboBox cmbColorMode;
        private NumericUpDown nudDPI;
        private NumericUpDown nudWidth;
        private NumericUpDown nudHeight;
        private CheckBox chkOutWithTTF;
        private CheckBox chkEmbeddedSearchText;
        private ComboBox cmbZoomType;
        private CheckBox chkOutputByLayoutSize;
        private TextBox txtFontPath;
        private TextBox txtPSPath;
        private Button btnConvert;
        private TextBox txtLog;
        private Label lblStatus;

        public MainForm()
        {
            InitializeUI();
        }

        private void InitializeUI()
        {
            this.Text = "AutoDWG DWG to PDF Converter";
            this.Size = new Size(620, 680);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.Font = new Font("Segoe UI", 9);

            int y = 15;
            int labelX = 20;
            int controlX = 140;
            int controlW = 380;
            int rowH = 32;

            // --- Input File ---
            AddLabel("Input DWG:", labelX, y);
            txtInputFile = AddTextBox(controlX, y, controlW - 80);
            AddButton("Browse...", controlX + controlW - 75, y, 75, OnBrowseInput);
            y += rowH;

            // --- Output File ---
            AddLabel("Output PDF:", labelX, y);
            txtOutputFile = AddTextBox(controlX, y, controlW - 80);
            AddButton("Browse...", controlX + controlW - 75, y, 75, OnBrowseOutput);
            y += rowH + 5;

            // --- Separator ---
            var sep1 = new Label { Text = "--- Conversion Parameters ---", AutoSize = true,
                Location = new Point(labelX, y), ForeColor = Color.Gray };
            this.Controls.Add(sep1);
            y += 25;

            // --- Color Mode ---
            AddLabel("Color Mode:", labelX, y);
            cmbColorMode = new ComboBox { Location = new Point(controlX, y),
                Width = 150, DropDownStyle = ComboBoxStyle.DropDownList };
            cmbColorMode.Items.AddRange(new object[] { "256 Colors", "Black & White" });
            cmbColorMode.SelectedIndex = 0;
            this.Controls.Add(cmbColorMode);
            y += rowH;

            // --- DPI ---
            AddLabel("DPI:", labelX, y);
            nudDPI = new NumericUpDown { Location = new Point(controlX, y),
                Width = 100, Minimum = 72, Maximum = 1200, Value = 300 };
            this.Controls.Add(nudDPI);
            y += rowH;

            // --- Width / Height ---
            AddLabel("Width (mm):", labelX, y);
            nudWidth = new NumericUpDown { Location = new Point(controlX, y),
                Width = 100, Minimum = 50, Maximum = 5000, Value = 297 };
            this.Controls.Add(nudWidth);
            AddLabel("Height (mm):", 300, y);
            nudHeight = new NumericUpDown { Location = new Point(400, y),
                Width = 100, Minimum = 50, Maximum = 5000, Value = 210 };
            this.Controls.Add(nudHeight);
            y += rowH;

            // --- OutWithTTF ---
            chkOutWithTTF = new CheckBox { Text = "Output TrueType Fonts",
                Location = new Point(controlX, y), AutoSize = true, Checked = true };
            this.Controls.Add(chkOutWithTTF);
            y += rowH;

            // --- EmbeddedSearchText ---
            chkEmbeddedSearchText = new CheckBox { Text = "Embed as Image (no searchable text)",
                Location = new Point(controlX, y), AutoSize = true, Checked = false };
            this.Controls.Add(chkEmbeddedSearchText);
            y += rowH;

            // --- Zoom Type ---
            AddLabel("Zoom Type:", labelX, y);
            cmbZoomType = new ComboBox { Location = new Point(controlX, y),
                Width = 150, DropDownStyle = ComboBoxStyle.DropDownList };
            cmbZoomType.Items.AddRange(new object[] { "Zoom Extents", "Last View" });
            cmbZoomType.SelectedIndex = 0;
            this.Controls.Add(cmbZoomType);
            y += rowH;

            // --- OutputByLayoutSize ---
            chkOutputByLayoutSize = new CheckBox { Text = "Output by Layout Size",
                Location = new Point(controlX, y), AutoSize = true, Checked = false };
            this.Controls.Add(chkOutputByLayoutSize);
            y += rowH + 5;

            // --- Font Path ---
            AddLabel("Font Path:", labelX, y);
            txtFontPath = AddTextBox(controlX, y, controlW);
            txtFontPath.Text = @"C:\windows\system32\shxfont";
            y += rowH;

            // --- PS Path ---
            AddLabel("PS Path:", labelX, y);
            txtPSPath = AddTextBox(controlX, y, controlW);
            txtPSPath.Text = @"c:\windows\system32\ps\lib;c:\windows\system32\ps\fonts";
            y += rowH + 10;

            // --- Convert Button ---
            btnConvert = new Button { Text = "Convert to PDF", Location = new Point(controlX, y),
                Size = new Size(150, 35), BackColor = Color.FromArgb(0, 120, 215),
                ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
            btnConvert.Click += OnConvert;
            this.Controls.Add(btnConvert);
            y += 45;

            // --- Status Label ---
            lblStatus = new Label { Text = "Ready.", Location = new Point(labelX, y),
                AutoSize = true, ForeColor = Color.DarkGreen };
            this.Controls.Add(lblStatus);
            y += 25;

            // --- Log ---
            txtLog = new TextBox { Location = new Point(labelX, y),
                Size = new Size(560, 100), Multiline = true, ReadOnly = true,
                ScrollBars = ScrollBars.Vertical, BackColor = Color.FromArgb(245, 245, 245) };
            this.Controls.Add(txtLog);
        }

        private Label AddLabel(string text, int x, int y)
        {
            var lbl = new Label { Text = text, Location = new Point(x, y + 3), AutoSize = true };
            this.Controls.Add(lbl);
            return lbl;
        }

        private TextBox AddTextBox(int x, int y, int width)
        {
            var txt = new TextBox { Location = new Point(x, y), Width = width };
            this.Controls.Add(txt);
            return txt;
        }

        private Button AddButton(string text, int x, int y, int width, EventHandler onClick)
        {
            var btn = new Button { Text = text, Location = new Point(x, y),
                Size = new Size(width, 25), FlatStyle = FlatStyle.System };
            btn.Click += onClick;
            this.Controls.Add(btn);
            return btn;
        }

        // --- Browse for input DWG file ---
        private void OnBrowseInput(object sender, EventArgs e)
        {
            using (var dlg = new OpenFileDialog())
            {
                dlg.Title = "Select DWG File";
                dlg.Filter = "DWG Files (*.dwg)|*.dwg|All Files (*.*)|*.*";
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    txtInputFile.Text = dlg.FileName;
                    // Auto-set output path
                    if (string.IsNullOrWhiteSpace(txtOutputFile.Text))
                    {
                        txtOutputFile.Text = Path.ChangeExtension(dlg.FileName, ".pdf");
                    }
                }
            }
        }

        // --- Browse for output PDF file ---
        private void OnBrowseOutput(object sender, EventArgs e)
        {
            using (var dlg = new SaveFileDialog())
            {
                dlg.Title = "Save PDF As";
                dlg.Filter = "PDF Files (*.pdf)|*.pdf";
                if (!string.IsNullOrWhiteSpace(txtInputFile.Text))
                {
                    dlg.FileName = Path.ChangeExtension(Path.GetFileName(txtInputFile.Text), ".pdf");
                    dlg.InitialDirectory = Path.GetDirectoryName(txtInputFile.Text);
                }
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    txtOutputFile.Text = dlg.FileName;
                }
            }
        }

        // --- Convert ---
        private void OnConvert(object sender, EventArgs e)
        {
            string inputFile = txtInputFile.Text.Trim();
            string outputFile = txtOutputFile.Text.Trim();

            if (string.IsNullOrWhiteSpace(inputFile))
            {
                MessageBox.Show("Please select an input DWG file.", "Warning",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (!File.Exists(inputFile))
            {
                MessageBox.Show("Input file does not exist:\n" + inputFile, "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (string.IsNullOrWhiteSpace(outputFile))
            {
                MessageBox.Show("Please specify an output PDF path.", "Warning",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Create output directory
            string outDir = Path.GetDirectoryName(outputFile);
            if (!string.IsNullOrEmpty(outDir) && !Directory.Exists(outDir))
                Directory.CreateDirectory(outDir);

            btnConvert.Enabled = false;
            lblStatus.Text = "Converting...";
            lblStatus.ForeColor = Color.DarkOrange;
            txtLog.Clear();

            try
            {
                Log("Creating COM object...");
                Type comType = Type.GetTypeFromProgID("DWGTOPDFX.ConvertPDF");
                if (comType == null)
                    throw new Exception("COM object 'DWGTOPDFX.ConvertPDF' not found.\nPlease register: regsvr32 dwgtopdfx64.dll");

                dynamic obj = Activator.CreateInstance(comType);
                try
                {
                    // Set parameters
                    obj.ColorMode = cmbColorMode.SelectedIndex;
                    obj.DPI = (int)nudDPI.Value;
                    obj.OutWithTTF = chkOutWithTTF.Checked ? 1 : 0;
                    obj.EmbeddedSearchText = chkEmbeddedSearchText.Checked ? 1 : 0;
                    obj.zoomtype = cmbZoomType.SelectedIndex;
                    obj.OutputByLayoutSize = chkOutputByLayoutSize.Checked ? 1 : 0;
                    obj.nBackgroundColor = 7;  // White background
                    obj.Width = (int)nudWidth.Value;
                    obj.Height = (int)nudHeight.Value;

                    if (!string.IsNullOrWhiteSpace(txtFontPath.Text))
                        obj.AddFontPath(txtFontPath.Text);
                    if (!string.IsNullOrWhiteSpace(txtPSPath.Text))
                        obj.PSPath = txtPSPath.Text;

                    obj.InputFile = inputFile;

                    Log("Input:  " + inputFile);
                    Log("Output: " + outputFile);
                    Log("Converting...");

                    obj.Convert(outputFile);

                    Log("Conversion complete!");
                    lblStatus.Text = "Done!";
                    lblStatus.ForeColor = Color.DarkGreen;

                    // Show result with output file location
                    string msg = "Conversion complete!\n\nOutput file:\n" + outputFile;
                    var result = MessageBox.Show(msg + "\n\nOpen the output folder?",
                        "Success", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                    if (result == DialogResult.Yes)
                    {
                        System.Diagnostics.Process.Start("explorer.exe",
                            "/select,\"" + outputFile + "\"");
                    }
                }
                finally
                {
                    Marshal.ReleaseComObject(obj);
                }
            }
            catch (Exception ex)
            {
                Log("ERROR: 0x" + (ex.HResult & 0xFFFFFFFF).ToString("X8"));
                lblStatus.Text = "Failed!";
                lblStatus.ForeColor = Color.Red;
                MessageBox.Show("Conversion failed. Please check the input file and settings, then try again.\n\nError: 0x" + (ex.HResult & 0xFFFFFFFF).ToString("X8"), "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                btnConvert.Enabled = true;
            }
        }

        private void Log(string message)
        {
            txtLog.AppendText("[" + DateTime.Now.ToString("HH:mm:ss") + "] " + message + "\r\n");
        }
    }

    static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }
    }
}