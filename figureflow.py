#!/usr/bin/env python3
"""FigureFlow — local PDF figure conversion for Microsoft Word.

No network access, telemetry, subprocess upload, external converters, or cloud APIs.
PDF -> vector SVG (with text as paths) and/or high-DPI PNG using PyMuPDF.
Requires PyMuPDF, Pillow; tkinterdnd2 for native drag-and-drop.
"""

from __future__ import annotations

import argparse
import csv
import hashlib
import os
from pathlib import Path
import queue
import re
import sys
import threading
import traceback
from dataclasses import dataclass
from datetime import datetime
from typing import Callable, Optional

try:
    import pymupdf as fitz
except ImportError:
    try:
        import fitz
    except ImportError as exc:
        raise SystemExit("PyMuPDF is required. Run Install_Once.bat (Windows) or pip install -r requirements.txt") from exc

try:
    from PIL import Image, ImageChops, ImageStat, ImageTk
except ImportError as exc:
    raise SystemExit("Pillow is required. Run Install_Once.bat (Windows) or pip install -r requirements.txt") from exc

try:
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk
except ImportError:
    tk = None

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    DRAG_AVAILABLE = True
except ImportError:
    DND_FILES = None
    TkinterDnD = None
    DRAG_AVAILABLE = False

APP_TITLE = "FigureFlow  |  Offline PDF → Word Figures"
BG = "#101724"
PANEL = "#192334"
PANEL_DARK = "#141d2b"
FG = "#eff5ff"
MUTED = "#9cacc2"
ACCENT = "#5ca8ff"
GREEN = "#65dcaf"
YELLOW = "#f4ce76"
RED = "#fd8e92"
BORDER = "#2a3850"
PREVIEW_SIZE = (525, 355)
MAX_RENDER_PIXELS = 135_000_000


def local_default_output() -> Path:
    # LocalAppData is not normally a cloud-sync location on Windows.
    root = os.environ.get("LOCALAPPDATA")
    if os.name == "nt" and root:
        return Path(root) / "FigureFlow" / "Exports"
    return Path.home() / "FigureFlow_Exports"


def likely_synced_path(path: Path) -> bool:
    name = str(path).lower().replace("/", "\\")
    return any(s in name for s in ("onedrive", "dropbox", "google drive", "icloud", "sharepoint", "syncthing"))


def safe_stem(path: Path) -> str:
    stem = re.sub(r"[^\w. -]+", "_", path.stem, flags=re.UNICODE).strip(" .")
    return stem[:105] or "figure"


def destination_name(path: Path, duplicates: dict[str, int]) -> str:
    """Use stable suffixes only when separate sources have the same basename."""
    stem = safe_stem(path)
    k = stem.casefold()
    duplicates[k] = duplicates.get(k, 0) + 1
    return stem if duplicates[k] == 1 else f"{stem}_{duplicates[k]:02d}"


def open_pdf(path: Path):
    doc = fitz.open(str(path))
    if doc.needs_pass:
        doc.close()
        raise ValueError("Password-protected PDF: unlock a local copy first")
    if not doc.is_pdf:
        doc.close()
        raise ValueError("Not a PDF document")
    return doc


def render_pdf_page(page, dpi: int = 120) -> Image.Image:
    pix = page.get_pixmap(dpi=dpi, colorspace=fitz.csRGB, alpha=False)
    return Image.frombytes("RGB", (pix.width, pix.height), pix.samples)


def render_svg(svg_path: Path, dpi: int = 120) -> Image.Image:
    svg = svg_path.read_bytes()
    with fitz.open(stream=svg, filetype="svg") as document:
        page = document[0]
        return render_pdf_page(page, dpi=dpi)


def preview_image(source: Path, size: tuple[int, int] = PREVIEW_SIZE, pdf_page_index: int = 0) -> Image.Image:
    ext = source.suffix.lower()
    if ext == ".pdf":
        with open_pdf(source) as d:
            img = render_pdf_page(d[pdf_page_index], dpi=110)
    elif ext == ".svg":
        img = render_svg(source, dpi=110)
    else:
        with Image.open(source) as im:
            img = im.convert("RGB")
    img.thumbnail(size, Image.Resampling.LANCZOS)
    return img


def compare_pdf_svg(page, svg_path: Path) -> tuple[float, float, str]:
    """Compare independent renderings at 120 dpi; heuristic, not a guarantee."""
    try:
        before = render_pdf_page(page, dpi=120)
        after = render_svg(svg_path, dpi=120)
        if before.size != after.size:
            return 255.0, 100.0, "Review: different rendering dimensions"
        diff = ImageChops.difference(before, after)
        avg_diff = float(sum(ImageStat.Stat(diff).mean) / 3.0)
        gray = diff.convert("L")
        hist = gray.histogram()
        high_diff_pct = 100.0 * sum(hist[29:]) / (before.width * before.height)
        if avg_diff > 5.0 or high_diff_pct > 2.0:
            label = "Review recommended"
        else:
            label = "Low preview difference"
        return round(avg_diff, 2), round(high_diff_pct, 2), label
    except Exception as exc:
        return -1.0, -1.0, f"Visual check unavailable ({type(exc).__name__})"


def raster_dimensions(page, dpi: int) -> tuple[int, int, int]:
    rect = page.rect
    w = max(1, round(rect.width * dpi / 72))
    h = max(1, round(rect.height * dpi / 72))
    return w, h, w * h


@dataclass
class ExportResult:
    source: str
    page: int
    svg: str = ""
    png: str = ""
    status: str = ""
    mean_diff: str = ""
    changed_pct: str = ""
    detail: str = ""


def export_one(path: Path, output_dir: Path, basename: str, dpi: int, svg_enabled: bool,
               png_enabled: bool, all_pages: bool, cancellation: Optional[threading.Event] = None,
               progress: Optional[Callable[[str], None]] = None) -> list[ExportResult]:
    results = []
    with open_pdf(path) as doc:
        total = doc.page_count if all_pages else min(1, doc.page_count)
        if total == 0:
            raise ValueError("PDF has no pages")
        for index in range(total):
            if cancellation is not None and cancellation.is_set():
                break
            page = doc[index]
            suffix = f"_p{index + 1:02d}" if all_pages or doc.page_count > 1 else ""
            file_stem = basename + suffix
            result = ExportResult(source=str(path), page=index + 1)
            if progress:
                progress(f"{path.name}  •  page {index + 1}/{total}")

            if svg_enabled:
                svg_path = output_dir / "SVG" / (file_stem + ".svg")
                try:
                    svg_path.parent.mkdir(parents=True, exist_ok=True)
                    # Glyphs as vector paths prevent font substitution in Word.
                    svg_markup = page.get_svg_image(text_as_path=True)
                    # SVG defaults bare width/height to CSS px. Word uses these to
                    # decide the inserted image size, so specify PDF units (points)
                    # explicitly while keeping the original viewBox unchanged.
                    svg_markup = re.sub(
                        r'(<svg\b[^>]*\bwidth=")([0-9.]+)("\s+height=")([0-9.]+)(")',
                        lambda m: f'{m.group(1)}{m.group(2)}pt{m.group(3)}{m.group(4)}pt{m.group(5)}',
                        svg_markup, count=1,
                    )
                    svg_path.write_text(svg_markup, encoding="utf-8")
                    result.svg = str(svg_path)
                    avg, changed, label = compare_pdf_svg(page, svg_path)
                    result.mean_diff = "" if avg < 0 else f"{avg:.2f}"
                    result.changed_pct = "" if changed < 0 else f"{changed:.2f}"
                    result.detail = label
                except Exception as exc:
                    result.detail = f"SVG error: {exc}"
                    if svg_path.exists():
                        svg_path.unlink(missing_ok=True)

            if png_enabled:
                png_path = output_dir / "PNG" / (file_stem + ".png")
                try:
                    w, h, pixels = raster_dimensions(page, dpi)
                    if pixels > MAX_RENDER_PIXELS:
                        raise ValueError(f"{w:,}×{h:,} ({pixels / 1e6:.0f} MP) exceeds safety limit; lower DPI")
                    png_path.parent.mkdir(parents=True, exist_ok=True)
                    pix = page.get_pixmap(dpi=dpi, colorspace=fitz.csRGB, alpha=False)
                    pix.save(str(png_path))
                    result.png = str(png_path)
                except Exception as exc:
                    result.detail = (result.detail + " | " if result.detail else "") + f"PNG error: {exc}"
                    if png_path.exists():
                        png_path.unlink(missing_ok=True)

            if result.svg and (not png_enabled or result.png):
                result.status = "Converted"
            elif result.png and not svg_enabled:
                result.status = "Converted"
            elif result.svg or result.png:
                result.status = "Partial"
            else:
                result.status = "Failed"
            results.append(result)
    return results


def write_report(output_dir: Path, results: list[ExportResult]) -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)
    report = output_dir / "quality_report.csv"
    with report.open("w", encoding="utf-8-sig", newline="") as handle:
        wr = csv.writer(handle)
        wr.writerow(("Source PDF", "Page", "Status", "SVG path", "PNG path", "SVG mean RGB diff 0-255",
                     "SVG pixels with luminance diff >28 (%)", "Preview assessment"))
        for r in results:
            wr.writerow((r.source, r.page, r.status, r.svg, r.png, r.mean_diff, r.changed_pct, r.detail))
    return report


def process_files(paths: list[Path], output_dir: Path, dpi: int, svg_enabled: bool,
                  png_enabled: bool, all_pages: bool, cancellation: Optional[threading.Event] = None,
                  emit: Optional[Callable[[str, object], None]] = None):
    """Shared export engine used by GUI and CLI; completely offline."""
    total_results: list[ExportResult] = []
    basename_counts: dict[str, int] = {}
    for path in paths:
        if cancellation is not None and cancellation.is_set():
            break
        base = destination_name(path, basename_counts)
        try:
            results = export_one(
                path, output_dir, base, dpi, svg_enabled, png_enabled, all_pages,
                cancellation=cancellation,
                progress=(lambda s: emit("progress", s)) if emit else None,
            )
        except Exception as exc:
            results = [ExportResult(source=str(path), page=0, status="Failed", detail=str(exc))]
        total_results.extend(results)
        if emit:
            emit("file_done", (str(path), results))
    report = write_report(output_dir, total_results)
    if emit:
        emit("finished", (total_results, report, bool(cancellation and cancellation.is_set())))
    return total_results, report


def open_local_folder(path: Path):
    """Only opens user's local filesystem; does not contact web services."""
    path.mkdir(parents=True, exist_ok=True)
    if os.name == "nt":
        os.startfile(str(path))  # type: ignore[attr-defined]
    elif sys.platform == "darwin":
        import subprocess
        subprocess.Popen(["open", str(path)])
    else:
        import subprocess
        subprocess.Popen(["xdg-open", str(path)])


class FigureFlowApp:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("1240x855")
        self.root.minsize(990, 700)
        self.root.configure(bg=BG)
        self.rows: dict[str, dict] = {}
        self.path_keys: set[str] = set()
        self.output_var = tk.StringVar(value=str(local_default_output()))
        self.dpi_var = tk.StringVar(value="600")
        self.svg_var = tk.BooleanVar(value=True)
        self.png_var = tk.BooleanVar(value=True)
        self.all_pages_var = tk.BooleanVar(value=False)
        self.preview_var = tk.StringVar(value="SVG")
        self.status_var = tk.StringVar(value="Ready · Processing stays on this computer")
        self.preview_note_var = tk.StringVar(value="Select a PDF to preview its first page.")
        self.current_image_a = None
        self.current_image_b = None
        self.active = False
        self.cancel_event = threading.Event()
        self.messages = queue.Queue()
        self._build_style()
        self._build_ui()
        self.root.after(100, self._poll)
        self.root.protocol("WM_DELETE_WINDOW", self._close)

    def _build_style(self):
        sty = ttk.Style(self.root)
        sty.theme_use("clam")
        sty.configure(".", background=BG, foreground=FG, font=("Segoe UI", 10))
        sty.configure("TFrame", background=BG)
        sty.configure("Panel.TFrame", background=PANEL)
        sty.configure("TLabel", background=BG, foreground=FG)
        sty.configure("Muted.TLabel", background=BG, foreground=MUTED)
        sty.configure("Panel.TLabel", background=PANEL, foreground=FG)
        sty.configure("TButton", background=PANEL, foreground=FG, bordercolor=BORDER, lightcolor=PANEL,
                      darkcolor=PANEL, padding=(14, 9), focusthickness=0)
        sty.map("TButton", background=[("active", "#2b4061")], foreground=[("disabled", "#6c7d94")])
        sty.configure("Accent.TButton", background=ACCENT, foreground="#071523", padding=(17, 10),
                      bordercolor=ACCENT, font=("Segoe UI", 10, "bold"))
        sty.map("Accent.TButton", background=[("active", "#9ad0ff"), ("disabled", "#344c67")])
        sty.configure("TCheckbutton", background=BG, foreground=FG)
        sty.map("TCheckbutton", background=[("active", BG)], foreground=[("disabled", MUTED)])
        sty.configure("TCombobox", fieldbackground=PANEL, background=PANEL, foreground=FG,
                      arrowcolor=FG, bordercolor=BORDER)
        sty.map("TCombobox", fieldbackground=[("readonly", PANEL)], foreground=[("readonly", FG)])
        sty.configure("Treeview", background=PANEL_DARK, fieldbackground=PANEL_DARK, foreground=FG,
                      rowheight=31, bordercolor=BORDER, font=("Segoe UI", 10))
        sty.map("Treeview", background=[("selected", "#274a73")], foreground=[("selected", FG)])
        sty.configure("Treeview.Heading", background=PANEL, foreground=FG, relief="flat", padding=(9, 8),
                      font=("Segoe UI", 10, "bold"))
        sty.map("Treeview.Heading", background=[("active", "#243347")])
        sty.configure("Horizontal.TProgressbar", troughcolor=PANEL, background=ACCENT, thickness=8)

    def _label(self, parent, text: str, fg=FG, size=10, weight="normal", **pack):
        label = tk.Label(parent, text=text, bg=parent.cget("bg"), fg=fg, font=("Segoe UI", size, weight),
                         anchor="w")
        label.pack(**pack)
        return label

    def _build_ui(self):
        main = tk.Frame(self.root, bg=BG, padx=24, pady=18)
        main.pack(fill="both", expand=True)
        title_row = tk.Frame(main, bg=BG)
        title_row.pack(fill="x")
        self._label(title_row, "FigureFlow", size=24, weight="bold", side="left")
        tk.Label(title_row, text="OFFLINE  •  NO UPLOADS", bg="#21413d", fg=GREEN,
                 font=("Segoe UI", 9, "bold"), padx=12, pady=6).pack(side="right")
        self._label(main, "Convert scientific figure PDFs to Word-ready SVG and high-resolution PNG without cloud services.",
                    fg=MUTED, size=10, pady=(0, 17))

        drop = tk.Frame(main, bg=PANEL, highlightbackground=BORDER, highlightthickness=1, height=92)
        drop.pack(fill="x")
        drop.pack_propagate(False)
        label = tk.Label(drop, bg=PANEL, fg=FG, text="Drop PDF figures here", font=("Segoe UI", 14, "bold"))
        label.pack(pady=(15, 0))
        self.drop_instruction = tk.Label(drop, bg=PANEL, fg=MUTED,
             text="Multiple PDFs and folders supported  •  Or click Add PDFs below" if DRAG_AVAILABLE else
                  "Install tkinterdnd2 to enable native drag-and-drop  •  Add PDFs works now",
             font=("Segoe UI", 9))
        self.drop_instruction.pack(pady=(1, 10))
        if DRAG_AVAILABLE:
            for widget in (drop, label, self.drop_instruction):
                widget.drop_target_register(DND_FILES)
                widget.dnd_bind("<<Drop>>", self._on_drop)

        controls = tk.Frame(main, bg=BG)
        controls.pack(fill="x", pady=(14, 12))
        ttk.Button(controls, text="+ Add PDFs", command=self._add_pdf_dialog).pack(side="left", padx=(0, 8))
        ttk.Button(controls, text="Add folder", command=self._add_folder_dialog).pack(side="left", padx=(0, 8))
        ttk.Button(controls, text="Clear list", command=self._clear).pack(side="left")
        self._label(controls, "Export:", fg=MUTED, side="left", padx=(25, 8))
        ttk.Checkbutton(controls, text="Vector SVG", variable=self.svg_var).pack(side="left", padx=(0, 12))
        ttk.Checkbutton(controls, text="PNG", variable=self.png_var).pack(side="left", padx=(0, 6))
        self._label(controls, "DPI", fg=MUTED, side="left", padx=(0, 5))
        ttk.Combobox(controls, textvariable=self.dpi_var, values=("300", "450", "600", "900"),
                     state="readonly", width=5).pack(side="left", padx=(0, 12))
        ttk.Checkbutton(controls, text="All PDF pages", variable=self.all_pages_var).pack(side="left")

        dir_line = tk.Frame(main, bg=BG)
        dir_line.pack(fill="x", pady=(0, 12))
        tk.Label(dir_line, text="Save locally to", bg=BG, fg=MUTED, font=("Segoe UI", 10)).pack(side="left", padx=(0, 9))
        self.output_field = tk.Entry(dir_line, textvariable=self.output_var, bg=PANEL, fg=FG,
                           insertbackground=FG, relief="flat", highlightbackground=BORDER,
                           highlightthickness=1, font=("Segoe UI", 10))
        self.output_field.pack(side="left", fill="x", expand=True, ipady=8)
        ttk.Button(dir_line, text="Browse…", command=self._choose_output).pack(side="left", padx=(8, 0))
        ttk.Button(dir_line, text="Open folder", command=self._open_output).pack(side="left", padx=(7, 0))

        queue_top = tk.Frame(main, bg=BG)
        queue_top.pack(fill="x")
        self._label(queue_top, "FIGURE QUEUE", fg=MUTED, size=9, weight="bold", side="left", pady=(0, 7))
        self.count_label = tk.Label(queue_top, text="0 PDFs", bg=BG, fg=MUTED, font=("Segoe UI", 9))
        self.count_label.pack(side="right")

        list_wrap = tk.Frame(main, bg=PANEL_DARK, highlightbackground=BORDER, highlightthickness=1)
        list_wrap.pack(fill="x", ipady=1)
        columns = ("name", "pages", "status")
        self.tree = ttk.Treeview(list_wrap, columns=columns, show="headings", height=5, selectmode="browse")
        self.tree.heading("name", text="FILE")
        self.tree.heading("pages", text="PAGES")
        self.tree.heading("status", text="STATUS / QUALITY")
        self.tree.column("name", width=510, stretch=True)
        self.tree.column("pages", width=90, stretch=False, anchor="center")
        self.tree.column("status", width=440, stretch=True)
        self.tree.pack(side="left", fill="both", expand=True)
        scroll = ttk.Scrollbar(list_wrap, command=self.tree.yview, orient="vertical")
        scroll.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=scroll.set)
        self.tree.bind("<<TreeviewSelect>>", self._on_select)

        action = tk.Frame(main, bg=BG)
        action.pack(fill="x", pady=(12, 11))
        self.convert_btn = ttk.Button(action, text="Convert all figures", style="Accent.TButton", command=self._convert)
        self.convert_btn.pack(side="left")
        self.cancel_btn = ttk.Button(action, text="Cancel", command=self._cancel, state="disabled")
        self.cancel_btn.pack(side="left", padx=(8, 20))
        self.progress = ttk.Progressbar(action, mode="determinate", maximum=100)
        self.progress.pack(side="left", fill="x", expand=True, padx=(0, 10))
        tk.Label(action, textvariable=self.status_var, bg=BG, fg=MUTED,
                 font=("Segoe UI", 9), anchor="e").pack(side="right")

        preview_top = tk.Frame(main, bg=BG)
        preview_top.pack(fill="x")
        self._label(preview_top, "QUALITY PREVIEW", fg=MUTED, size=9, weight="bold", side="left", pady=(0, 7))
        self._label(preview_top, "Compare with:", fg=MUTED, size=9, side="left", padx=(16, 6))
        combo = ttk.Combobox(preview_top, textvariable=self.preview_var, values=("SVG", "PNG"), width=6, state="readonly")
        combo.pack(side="left", pady=(0, 6))
        combo.bind("<<ComboboxSelected>>", self._on_select)
        tk.Label(preview_top, textvariable=self.preview_note_var, bg=BG, fg=MUTED,
                 font=("Segoe UI", 9)).pack(side="right", pady=(0, 6))

        previews = tk.Frame(main, bg=BG)
        previews.pack(fill="both", expand=True)
        previews.grid_columnconfigure(0, weight=1, uniform="p")
        previews.grid_columnconfigure(1, weight=1, uniform="p")
        previews.grid_rowconfigure(0, weight=1)
        self.preview_a = self._preview_pane(previews, 0, "ORIGINAL PDF")
        self.preview_b = self._preview_pane(previews, 1, "CONVERTED OUTPUT")

        foot = tk.Frame(main, bg=BG)
        foot.pack(fill="x", pady=(10, 0))
        self._label(foot, "WORD TIP   Insert → Pictures → This Device → select SVG. Export Word as Standard-quality PDF.",
                    fg=MUTED, size=9, side="left")
        self._label(foot, "Quality report saved as CSV", fg=MUTED, size=9, side="right")

    def _preview_pane(self, parent, column, heading):
        outer = tk.Frame(parent, bg=PANEL_DARK, highlightbackground=BORDER, highlightthickness=1)
        outer.grid(row=0, column=column, sticky="nsew", padx=((0, 7) if column == 0 else (7, 0)))
        tk.Label(outer, text=heading, bg=PANEL_DARK, fg=MUTED,
                 font=("Segoe UI", 9, "bold")).pack(anchor="w", padx=14, pady=(10, 6))
        img_label = tk.Label(outer, text="Preview will appear here", bg=PANEL_DARK, fg="#73849c",
                             font=("Segoe UI", 10))
        img_label.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        return img_label

    def _on_drop(self, event):
        paths = [Path(p) for p in self.root.tk.splitlist(event.data)]
        self._add_paths(paths)

    def _add_pdf_dialog(self):
        files = filedialog.askopenfilenames(title="Choose local PDF figures", filetypes=[("PDF figures", "*.pdf")])
        self._add_paths([Path(p) for p in files])

    def _add_folder_dialog(self):
        folder = filedialog.askdirectory(title="Choose a folder containing PDFs")
        if folder:
            self._add_paths([Path(folder)])

    def _add_paths(self, paths: list[Path]):
        if self.active:
            return
        candidates = []
        for item in paths:
            try:
                if item.is_dir():
                    candidates.extend(sorted(item.rglob("*.pdf"), key=lambda p: str(p).casefold()))
                    candidates.extend(sorted(item.rglob("*.PDF"), key=lambda p: str(p).casefold()))
                elif item.is_file() and item.suffix.lower() == ".pdf":
                    candidates.append(item)
            except OSError:
                pass
        added = 0
        for path in candidates:
            try:
                canonical = str(path.resolve())
                key = os.path.normcase(canonical)
                if key in self.path_keys:
                    continue
                with open_pdf(path) as pdf:
                    num_pages = pdf.page_count
                iid = str(len(self.rows) + 1) + "-" + hashlib.sha1(canonical.encode("utf-8")).hexdigest()[:8]
                self.rows[iid] = {"path": path, "pages": num_pages, "results": []}
                self.path_keys.add(key)
                self.tree.insert("", "end", iid=iid, values=(path.name, num_pages, "Ready"))
                added += 1
            except Exception as exc:
                # Show unusable PDFs too: preserve error visibility in queue.
                canonical = str(path.absolute())
                key = os.path.normcase(canonical)
                if key not in self.path_keys:
                    iid = str(len(self.rows) + 1) + "-bad"
                    self.rows[iid] = {"path": path, "pages": 0, "results": []}
                    self.path_keys.add(key)
                    self.tree.insert("", "end", iid=iid, values=(path.name, "—", f"Cannot read: {str(exc)[:80]}"))
                    added += 1
        self.count_label.configure(text=f"{len(self.rows)} PDFs")
        if added and not self.tree.selection():
            first = self.tree.get_children()[0]
            self.tree.selection_set(first)
            self._on_select()
        if added:
            self.status_var.set(f"Added {added} PDF(s) · Ready")

    def _choose_output(self):
        dest = filedialog.askdirectory(title="Choose a LOCAL folder for converted figures")
        if dest:
            self.output_var.set(dest)
            if likely_synced_path(Path(dest)):
                messagebox.showwarning("Possible synced folder", "This folder may be cloud-synchronized.\n"
                    "Choose a local folder outside OneDrive / Dropbox if the figures are unpublished.")

    def _open_output(self):
        try:
            open_local_folder(Path(self.output_var.get()).expanduser())
        except Exception as exc:
            messagebox.showerror("Cannot open folder", str(exc))

    def _clear(self):
        if self.active:
            return
        self.tree.delete(*self.tree.get_children())
        self.rows.clear()
        self.path_keys.clear()
        self.count_label.configure(text="0 PDFs")
        self.status_var.set("Ready · Processing stays on this computer")
        self.preview_note_var.set("Select a PDF to preview its first page.")
        for label in (self.preview_a, self.preview_b):
            label.configure(image="", text="Preview will appear here")

    def _selected_row(self):
        selection = self.tree.selection()
        return (selection[0], self.rows[selection[0]]) if selection else (None, None)

    def _on_select(self, event=None):
        iid, row = self._selected_row()
        if row is None:
            return
        try:
            if row["pages"] == 0:
                return
            self._show_preview(self.preview_a, preview_image(row["path"]))
        except Exception as exc:
            self.preview_a.configure(image="", text=f"Original preview unavailable\n{str(exc)[:140]}")
        results = row.get("results") or []
        if results:
            result = results[0]
            selected = self.preview_var.get()
            converted = result.svg if selected == "SVG" else result.png
            if converted and Path(converted).exists():
                try:
                    self._show_preview(self.preview_b, preview_image(Path(converted)))
                except Exception as exc:
                    self.preview_b.configure(image="", text=f"Preview unavailable\n{str(exc)[:100]}")
            else:
                self.preview_b.configure(image="", text=f"No {selected} result yet")
            detail = result.detail or result.status
            if result.changed_pct:
                detail += f"  •  >28-tone pixel changes: {result.changed_pct}%"
            self.preview_note_var.set(detail[:108])
        else:
            self.preview_b.configure(image="", text="Convert the figures to compare")
            self.preview_note_var.set("SVG preview checks for visual differences (not a guarantee).")

    def _show_preview(self, label, img: Image.Image):
        photo = ImageTk.PhotoImage(img)
        label.configure(image=photo, text="")
        label.image = photo

    def _convert(self):
        if self.active:
            return
        if not self.rows:
            messagebox.showinfo("No figures", "Drop PDF figures here, or use Add PDFs.")
            return
        if not self.svg_var.get() and not self.png_var.get():
            messagebox.showinfo("Choose a format", "Select SVG, PNG, or both.")
            return
        output = Path(self.output_var.get()).expanduser()
        if not str(output).strip():
            messagebox.showwarning("Choose output folder", "Please choose an output folder.")
            return
        if likely_synced_path(output) and not messagebox.askyesno("Possible cloud-sync folder",
                "The output path appears to be in a cloud-synced folder.\n"
                "Do you still want to export confidential figures there?"):
            return
        try:
            output.mkdir(parents=True, exist_ok=True)
        except Exception as exc:
            messagebox.showerror("Cannot write output", str(exc))
            return
        self.active = True
        self.cancel_event.clear()
        self.progress["value"] = 0
        self.convert_btn.configure(state="disabled")
        self.cancel_btn.configure(state="normal")
        paths = [r["path"] for r in self.rows.values()]
        for iid in self.rows:
            self.rows[iid]["results"] = []
            self.tree.set(iid, "status", "Queued")
        self.status_var.set("Converting locally…")
        args = (paths, output, int(self.dpi_var.get()), bool(self.svg_var.get()),
                bool(self.png_var.get()), bool(self.all_pages_var.get()))
        threading.Thread(target=self._worker, args=args, daemon=True).start()

    def _worker(self, *args):
        try:
            process_files(*args, cancellation=self.cancel_event, emit=lambda kind, payload: self.messages.put((kind, payload)))
        except Exception as exc:
            self.messages.put(("fatal", f"{exc}\n{traceback.format_exc()}"))

    def _poll(self):
        try:
            while True:
                kind, payload = self.messages.get_nowait()
                if kind == "progress":
                    self.status_var.set(str(payload))
                elif kind == "file_done":
                    path, results = payload
                    for iid, row in self.rows.items():
                        if str(row["path"]) == path:
                            row["results"] = results
                            failures = [r for r in results if r.status != "Converted"]
                            checks = [r for r in results if "Review recommended" in r.detail]
                            state = ("Failed" if failures and all(r.status == "Failed" for r in results)
                                     else "Partial" if failures else "Converted")
                            if checks:
                                state += " · visual review"
                            elif state == "Converted" and any(r.mean_diff for r in results):
                                state += " · preview diff low"
                            self.tree.set(iid, "status", state)
                            break
                    done = sum(bool(r["results"]) for r in self.rows.values())
                    self.progress["value"] = 100 * done / max(len(self.rows), 1)
                    self._on_select()
                elif kind == "finished":
                    results, report, was_cancelled = payload
                    self.active = False
                    self.convert_btn.configure(state="normal")
                    self.cancel_btn.configure(state="disabled")
                    successes = sum(r.status == "Converted" for r in results)
                    failed = sum(r.status != "Converted" for r in results)
                    if not was_cancelled:
                        self.progress["value"] = 100
                    self.status_var.set(f"{'Cancelled' if was_cancelled else 'Finished'} · {successes} converted · {failed} partial/failed")
                    self._on_select()
                    if failed:
                        messagebox.showwarning("Conversion completed with warnings",
                            f"Converted: {successes} page(s)\nPartial/failed: {failed} page(s)\n\n"
                            f"See the local report:\n{report}")
                elif kind == "fatal":
                    self.active = False
                    self.convert_btn.configure(state="normal")
                    self.cancel_btn.configure(state="disabled")
                    self.status_var.set("Conversion failed; check details")
                    messagebox.showerror("Unexpected conversion error", str(payload)[:1500])
        except queue.Empty:
            pass
        self.root.after(100, self._poll)

    def _cancel(self):
        self.cancel_event.set()
        self.status_var.set("Stopping after the current page…")

    def _close(self):
        if self.active and not messagebox.askyesno("Conversion running", "Stop conversion and close the app?"):
            return
        self.cancel_event.set()
        self.root.destroy()


def cli_main(arguments):
    path_list: list[Path] = []
    for name in arguments.batch:
        p = Path(name).expanduser()
        if p.is_dir():
            path_list.extend(sorted(x for x in p.rglob("*") if x.is_file() and x.suffix.lower() == ".pdf"))
        elif p.is_file():
            path_list.append(p)
    if not path_list:
        print("No PDFs found", file=sys.stderr)
        return 2
    if not arguments.svg and not arguments.png:
        print("At least one format must be enabled", file=sys.stderr)
        return 2
    out = Path(arguments.output).expanduser()
    results, report = process_files(path_list, out, arguments.dpi, arguments.svg, arguments.png,
                                    arguments.all_pages,
                                    emit=lambda kind, payload: print(payload) if kind == "progress" else None)
    print(f"Saved {sum(bool(r.svg) for r in results)} SVG(s) and {sum(bool(r.png) for r in results)} PNG(s)")
    print(f"Quality report: {report}")
    for r in results:
        print(f"{Path(r.source).name}, page {r.page}: {r.status} | {r.detail}")
    return 1 if any(r.status != "Converted" for r in results) else 0


def main():
    parser = argparse.ArgumentParser(description="Offline scientific PDF figure converter for Word")
    parser.add_argument("--batch", nargs="+", help="Run headless conversion for these PDF files or folders")
    parser.add_argument("--output", default=str(local_default_output()))
    parser.add_argument("--dpi", type=int, default=600, choices=(150, 300, 450, 600, 900))
    parser.add_argument("--all-pages", action="store_true")
    parser.add_argument("--svg", action=argparse.BooleanOptionalAction, default=True)
    parser.add_argument("--png", action=argparse.BooleanOptionalAction, default=True)
    args = parser.parse_args()
    if args.batch:
        return cli_main(args)
    if tk is None:
        print("Tkinter not available: install a Python distribution with Tk support.", file=sys.stderr)
        return 2
    root = TkinterDnD.Tk() if DRAG_AVAILABLE else tk.Tk()
    FigureFlowApp(root)
    root.mainloop()
    return 0


if __name__ == "__main__":
    sys.exit(main())