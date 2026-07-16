<?php
// ============================================================
// AutoDWG DWG2PDF SDK - PHP Web Application
//
// Prerequisites:
//   1. Register the COM DLL (run as Administrator):
//      regsvr32 "path\to\dwgtopdfx64.dll"
//   2. PHP with COM extension enabled: extension=com_dotnet
//   3. Run: php -S localhost:9877
//
// Open: http://localhost:9877
// ============================================================

$result = null;
$error = null;

// --- Handle file download ---
if (isset($_GET['download'])) {
    $file = basename($_GET['download']);
    $filePath = __DIR__ . '/output/' . $file;
    if (file_exists($filePath)) {
        header('Content-Type: application/pdf');
        header('Content-Disposition: attachment; filename="' . $file . '"');
        header('Content-Length: ' . filesize($filePath));
        readfile($filePath);
        exit;
    }
}

// --- Handle form submission ---
if ($_SERVER['REQUEST_METHOD'] === 'POST' && isset($_FILES['dwg_file'])) {

    $uploadDir = __DIR__ . '/uploads/';
    $outputDir = __DIR__ . '/output/';

    if (!is_dir($uploadDir)) mkdir($uploadDir, 0755, true);
    if (!is_dir($outputDir)) mkdir($outputDir, 0755, true);

    $uploadedFile = $_FILES['dwg_file'];

    // Validate upload
    if ($uploadedFile['error'] !== UPLOAD_ERR_OK) {
        $error = "File upload error (code: " . $uploadedFile['error'] . ")";
    } elseif (strtolower(pathinfo($uploadedFile['name'], PATHINFO_EXTENSION)) !== 'dwg') {
        $error = "Please upload a .dwg file.";
    } else {
        // Save uploaded file
        $inputPath = $uploadDir . basename($uploadedFile['name']);
        if (move_uploaded_file($uploadedFile['tmp_name'], $inputPath)) {

            $outputFileName = pathinfo($uploadedFile['name'], PATHINFO_FILENAME) . '.pdf';
            $outputPath = $outputDir . $outputFileName;

            try {
                // Create COM object
                $obj = new COM("DWGTOPDFX.ConvertPDF");

                // Set parameters from form
                $obj->ColorMode = intval($_POST['color_mode'] ?? 0);
                $obj->DPI = intval($_POST['dpi'] ?? 300);
                $obj->Width = intval($_POST['width'] ?? 297);
                $obj->Height = intval($_POST['height'] ?? 210);
                $obj->OutWithTTF = isset($_POST['out_with_ttf']) ? 1 : 0;
                $obj->EmbeddedSearchText = isset($_POST['embedded_search_text']) ? 1 : 0;
                $obj->zoomtype = intval($_POST['zoom_type'] ?? 0);
                $obj->OutputByLayoutSize = isset($_POST['output_by_layout_size']) ? 1 : 0;

                $fontPath = trim($_POST['font_path'] ?? '');
                if ($fontPath) $obj->AddFontPath($fontPath);

                $psPath = trim($_POST['ps_path'] ?? '');
                if ($psPath) $obj->PSPath = $psPath;

                $obj->InputFile = $inputPath;

                // Convert
                $obj->Convert($outputPath);

                $result = [
                    'success' => true,
                    'filename' => $outputFileName,
                    'message' => 'Conversion completed successfully!'
                ];

                $obj->Release();
                $obj = null;

            } catch (Exception $e) {
                $result = [
                    'success' => false,
                    'filename' => '',
                    'message' => $e->getMessage()
                ];
            }
        } else {
            $error = "Failed to save uploaded file.";
        }
    }
}
?>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>DWG to PDF Converter - AutoDWG</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: 'Segoe UI', Tahoma, Geneva, sans-serif; background: #f0f2f5; color: #333; }
        .header { background: linear-gradient(135deg, #8E2DE2, #4A00E0); color: white; padding: 20px 0; text-align: center; }
        .header h1 { font-size: 24px; font-weight: 600; }
        .header p { font-size: 14px; opacity: 0.85; margin-top: 5px; }
        .container { max-width: 800px; margin: 30px auto; padding: 0 20px; }
        .card { background: white; border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.1); padding: 30px; margin-bottom: 20px; }
        .card h2 { font-size: 18px; color: #8E2DE2; margin-bottom: 20px; border-bottom: 2px solid #8E2DE2; padding-bottom: 8px; cursor: pointer; user-select: none; }
        .card h2 .toggle-icon { float: right; font-size: 14px; color: #999; }
        .form-group { margin-bottom: 15px; }
        .form-group label { display: block; font-weight: 600; margin-bottom: 5px; font-size: 14px; }
        .form-group input, .form-group select { width: 100%; padding: 8px 12px; border: 1px solid #ddd; border-radius: 4px; font-size: 14px; }
        .form-group input[type="file"] { padding: 6px; }
        .form-row { display: flex; gap: 15px; }
        .form-row .form-group { flex: 1; }
        .checkbox-group { display: flex; gap: 20px; flex-wrap: wrap; margin: 10px 0; }
        .checkbox-group label { font-weight: normal; display: flex; align-items: center; gap: 5px; }
        .btn { display: inline-block; padding: 12px 36px; background: linear-gradient(135deg, #8E2DE2, #4A00E0); color: white; border: none; border-radius: 6px; font-size: 16px; font-weight: 600; cursor: pointer; transition: all 0.2s; box-shadow: 0 3px 10px rgba(142,45,226,0.3); }
        .btn:hover { transform: translateY(-1px); box-shadow: 0 5px 15px rgba(142,45,226,0.4); }
        .btn:disabled { background: #ccc; cursor: not-allowed; box-shadow: none; transform: none; }
        .download-btn { display: inline-block; padding: 12px 32px; background: linear-gradient(135deg, #28a745, #218838); color: white !important; border-radius: 6px; font-size: 16px; font-weight: 700; text-decoration: none; box-shadow: 0 3px 10px rgba(40,167,69,0.3); transition: all 0.2s; }
        .download-btn:hover { transform: translateY(-1px); box-shadow: 0 5px 15px rgba(40,167,69,0.4); }
        .result { padding: 20px; border-radius: 6px; margin-top: 15px; }
        .result.success { background: #d4edda; border: 1px solid #c3e6cb; color: #155724; }
        .result.error { background: #f8d7da; border: 1px solid #f5c6cb; color: #721c24; }
        .footer { text-align: center; padding: 20px; color: #999; font-size: 12px; }
    </style>
</head>
<body>
    <div class="header">
        <h1>AutoDWG DWG to PDF Converter</h1>
        <p>PHP Web Edition - powered by AutoDWG SDK</p>
    </div>

    <div class="container">
        <form method="post" enctype="multipart/form-data" id="convertForm">

            <!-- 1. Upload -->
            <div class="card">
                <h2>1. Select DWG File</h2>
                <div class="form-group">
                    <label for="dwg_file">Upload a DWG file to convert:</label>
                    <input type="file" name="dwg_file" id="dwg_file" accept=".dwg" required />
                </div>
            </div>

            <!-- 2. Settings (collapsible) -->
            <div class="card">
                <h2 onclick="toggleSettings()">
                    2. Conversion Settings
                    <span class="toggle-icon" id="toggleIcon">&#9660;</span>
                </h2>
                <div id="settingsPanel">
                    <div class="form-row">
                        <div class="form-group">
                            <label for="color_mode">Color Mode</label>
                            <select name="color_mode" id="color_mode">
                                <option value="0">256 Colors</option>
                                <option value="1">Black &amp; White</option>
                            </select>
                        </div>
                        <div class="form-group">
                            <label for="dpi">DPI</label>
                            <input type="number" name="dpi" id="dpi" value="300" min="72" max="1200" />
                        </div>
                    </div>
                    <div class="form-row">
                        <div class="form-group">
                            <label for="width">Width (mm)</label>
                            <input type="number" name="width" id="width" value="297" min="50" max="5000" />
                        </div>
                        <div class="form-group">
                            <label for="height">Height (mm)</label>
                            <input type="number" name="height" id="height" value="210" min="50" max="5000" />
                        </div>
                    </div>
                    <div class="form-row">
                        <div class="form-group">
                            <label for="zoom_type">Zoom Type</label>
                            <select name="zoom_type" id="zoom_type">
                                <option value="0">Zoom Extents</option>
                                <option value="1">Last View</option>
                            </select>
                        </div>
                    </div>
                    <div class="checkbox-group">
                        <label><input type="checkbox" name="out_with_ttf" value="1" checked /> Output TrueType Fonts</label>
                        <label><input type="checkbox" name="embedded_search_text" value="1" /> Embed as Image</label>
                        <label><input type="checkbox" name="output_by_layout_size" value="1" /> Output by Layout Size</label>
                    </div>
                    <div class="form-group">
                        <label for="font_path">Font Path</label>
                        <input type="text" name="font_path" id="font_path" value="C:\windows\system32\shxfont" />
                    </div>
                    <div class="form-group">
                        <label for="ps_path">PostScript Path</label>
                        <input type="text" name="ps_path" id="ps_path" value="c:\windows\system32\ps\lib;c:\windows\system32\ps\fonts" />
                    </div>
                </div>
            </div>

            <!-- Convert Button -->
            <div style="text-align: center; margin: 20px 0;">
                <button type="submit" class="btn" id="btnConvert" onclick="this.disabled=true;this.value='Converting, please wait...';">Convert to PDF</button>
            </div>
        </form>

        <!-- 3. Result -->
        <?php if ($result || $error): ?>
        <div class="card" id="resultCard">
            <h2>3. Result</h2>
            <?php if ($error): ?>
                <div class="result error">
                    <p><strong>Error:</strong> <?= htmlspecialchars($error) ?></p>
                </div>
            <?php elseif ($result && $result['success']): ?>
                <div class="result success">
                    <p style="font-size:16px;"><strong><?= htmlspecialchars($result['message']) ?></strong></p>
                    <p style="margin: 10px 0;">Output file: <strong><?= htmlspecialchars($result['filename']) ?></strong></p>
                    <p style="margin-top: 15px;">
                        <a href="?download=<?= urlencode($result['filename']) ?>" class="download-btn">
                            &#11015; Download PDF File
                        </a>
                    </p>
                </div>
            <?php elseif ($result): ?>
                <div class="result error">
                    <p><strong>Conversion failed.</strong></p>
                    <p><?= htmlspecialchars($result['message']) ?></p>
                </div>
            <?php endif; ?>
        </div>
        <?php endif; ?>
    </div>

    <div class="footer">
        &copy; AutoDWG - Professional DWG to PDF Conversion SDK
    </div>

    <script>
        function toggleSettings() {
            var panel = document.getElementById('settingsPanel');
            var icon = document.getElementById('toggleIcon');
            if (panel.style.display === 'none') {
                panel.style.display = 'block';
                icon.innerHTML = '&#9660;';
            } else {
                panel.style.display = 'none';
                icon.innerHTML = '&#9654;';
            }
        }

        <?php if ($result && $result['success']): ?>
        // Auto-collapse settings and scroll to result
        document.getElementById('settingsPanel').style.display = 'none';
        document.getElementById('toggleIcon').innerHTML = '&#9654;';
        document.getElementById('resultCard').scrollIntoView({ behavior: 'smooth', block: 'start' });
        alert("Conversion completed!\n\nOutput file: <?= addslashes($result['filename']) ?>");
        <?php elseif ($result): ?>
        document.getElementById('settingsPanel').style.display = 'none';
        document.getElementById('toggleIcon').innerHTML = '&#9654;';
        document.getElementById('resultCard').scrollIntoView({ behavior: 'smooth', block: 'start' });
        alert("Conversion failed!\n\nError: <?= addslashes($result['message']) ?>");
        <?php endif; ?>
    </script>
</body>
</html>