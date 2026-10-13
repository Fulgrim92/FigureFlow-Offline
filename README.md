# FigureFlow — offline PDF figures for Word (Windows)

A local drag-and-drop converter for unpublished scientific figures. Converts PDFs to **vector SVG** and **600 dpi PNG** (adjustable). Preview the original PDF next to the exported result and review an automated visual difference check. No accounts, cloud APIs, analytics, network calls, or uploads are in the app.

## Start (Windows)

1. Install **Python 3.10+**, with Tkinter included (standard Windows installer).
2. Double-click **`Install_Once.bat`** to install 3 Python packages. This initial dependency download may use the internet; **your PDF data never does**. For air-gapped computers, install from locally provided Python wheels instead.
3. Double-click **`Launch_FigureFlow.bat`**.
4. Drag one or more figure PDFs (or a PDF-containing folder) into FigureFlow.
5. Select **Vector SVG** and **PNG** and click **Convert all figures**.
6. In Microsoft Word (desktop): **Insert → Pictures → This Device → select the SVG**. If Word has rendering issues, insert the 600 dpi PNG instead.

The default output is under `%LOCALAPPDATA%\FigureFlow\Exports`, intentionally away from typical synchronized *Documents* paths. Change it to a local folder if needed. Converted figures appear in `SVG/` and `PNG/` subfolders, with a `quality_report.csv`. Always review complex multipanel figures at final manuscript size.

## What stays high quality?

- PDF **vector paths** (plots, lines, shapes and text outlines) remain vector-based in SVG.
- PDF **raster images** (microscopy) remain raster at their source resolution within SVG: SVG does not add missing detail.
- Font glyphs become **vector paths**, helping avoid font substitution. As a result, labels inside the SVG are not editable as text. SVG physical dimensions are specified in PDF points to help Word respect the figure size.
- 600 dpi PNG is high resolution at the PDF's original physical page dimensions. Large pages can create very large files. The tool automatically refuses oversized renders (>135 megapixels) and records the issue so it does not exhaust RAM.
- Some transparency, masks, blending, or embedded fonts may differ across Word versions. **Visual QC is indicative, not a publication-quality guarantee.** Word may rasterize SVG on PDF export: test the final Word-generated PDF.

## Features

- Drop PDFs or folders, browse for files, process multiple figures.
- Convert page 1 (default for one-page scientific figures), or all pages.
- 300, 450, 600, or 900 dpi PNG.
- Original PDF vs exported SVG/PNG preview.
- Offline preview comparison: average RGB difference and percentage of pixels differing substantially (at 120 dpi).
- CSV report with paths, export status, and preview difference metrics.
- Independent conversion per PDF/page; failure in one file does not stop the batch.
- Cancel after the active page finishes.

## Confidentiality

FigureFlow contains **no networking functions**. Conversion and quality checks use local PyMuPDF/Pillow code; drag-and-drop uses `tkinterdnd2`. The supplied Python source is auditable. **It cannot protect files that you intentionally store in OneDrive, Dropbox, Google Drive, etc.** Check the source files' folder and selected output folder. FigureFlow warns about common cloud-synced output folder names but cannot detect every sync configuration.

## Optional: a standalone Windows executable

Run **`Build_Windows_EXE.bat`** on your own Windows machine. It installs PyInstaller and builds `dist\FigureFlow.exe`. That EXE can be used locally without launching a Python script. This repository does not contain a prebuilt Windows EXE because building a Windows EXE requires a compatible Windows build environment.

If you redistribute the app or compiled binary, review the licenses for PyMuPDF, Pillow, tkinterdnd2, and PyInstaller. PyMuPDF is available under the AGPL or a commercial license.

## Command line (fully local)

```powershell
py -3 figureflow.py --batch "D:\Figures\Figure1.pdf" "D:\Figures\Figure2.pdf" --output "D:\Figures\Converted" --dpi 600
```

Process each page instead of just page 1:

```powershell
py -3 figureflow.py --batch "D:\Figures" --output "D:\Figures\Converted" --all-pages --dpi 300
```

## Common troubleshooting

**Dragging files does nothing:** run `Install_Once.bat`; this installs `tkinterdnd2`. Until installed you can still use *Add PDFs*.

**Black/incorrect fonts or missing chart details:** check PDF vs SVG preview, then try re-exporting the original PDF from Illustrator or use the PNG fallback.

**Quality report says “Low preview difference”:** compare at manuscript size anyway. The app measures differences using its own renderers, not Word's rendering engine.

**Word rasterizes the SVG:** in Word check the PDF export at high zoom; if it fails, use 600 dpi PNG or use the original PDF in a professional layout application.

**Very large PNG:** select 300 or 450 dpi, or crop figure whitespace upstream while keeping the intended physical dimensions.

## Test without unpublished data

A harmless, completely **synthetic** example PDF is included under `examples/Demo_Figure_SYNTHETIC.pdf`. Drag this file into the app to verify that PDF→SVG/PNG conversion and the preview work before using your real figures.

## Check the installation

Double-click **`Run_Self_Test.bat`** to validate the SVG/PNG export using only the bundled synthetic example. The test uses an automatically cleaned local temporary directory; it does not access the internet or upload any files.