"""Main application window for Excel Image Extractor Pro.
Provides a modern Windows desktop interface with multi-sheet tabs,
column auto-detection, thumbnail preview, progress tracking, and 1-click JPG export.
"""

import os
import sys
import threading
import subprocess
import queue
from typing import List, Optional
import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk

from core.excel_engine import (
    scan_excel_workbook,
    export_images,
    SheetScanResult
)
from core.ai_assistant import load_config, save_config
from ui.sheet_view import SheetView
from ui.settings_modal import SettingsModal


class AppWindow(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Load preferences
        self.config = load_config()
        ctk.set_appearance_mode(self.config.get("theme", "dark"))
        ctk.set_default_color_theme("blue")

        # Window properties
        self.title("Excel Image Extractor Pro - [Rename to Design Number]")
        self.geometry("1140x820")
        self.minsize(980, 680)

        # State
        self.current_file: Optional[str] = None
        self.output_directory: Optional[str] = None
        self.sheet_results: List[SheetScanResult] = []
        self.is_processing = False

        # Thread-safe UI message queue
        self.ui_queue = queue.Queue()
        self.after(50, self._process_ui_queue)

        self._build_ui()
        self._check_for_sample_file()

    def _build_ui(self):
        # 1. Top Header Bar
        header = ctk.CTkFrame(self, height=64, corner_radius=0, fg_color=("gray85", "gray14"))
        header.pack(fill="x", padx=0, pady=0)

        title_frame = ctk.CTkFrame(header, fg_color="transparent")
        title_frame.pack(side="left", padx=20, pady=10)

        main_title = ctk.CTkLabel(
            title_frame,
            text="💎 Excel Image Extractor Pro",
            font=ctk.CTkFont(size=20, weight="bold")
        )
        main_title.pack(anchor="w")

        sub_title = ctk.CTkLabel(
            title_frame,
            text="Extract & rename embedded Excel images to [DesignNumber].jpg across multiple sheets",
            font=ctk.CTkFont(size=12),
            text_color="gray"
        )
        sub_title.pack(anchor="w")

        # Top right quick action buttons
        nav_btns = ctk.CTkFrame(header, fg_color="transparent")
        nav_btns.pack(side="right", padx=20, pady=10)

        self.sample_btn = ctk.CTkButton(
            nav_btns,
            text="Load Sample Catalog",
            command=self._load_sample_catalog,
            fg_color="#34495e",
            hover_color="#2c3e50",
            width=140,
            height=32,
            font=ctk.CTkFont(size=12)
        )
        self.sample_btn.pack(side="left", padx=6)

        open_out_btn = ctk.CTkButton(
            nav_btns,
            text="📂 Open Output Folder",
            command=self._open_output_folder,
            fg_color="#2c3e50",
            hover_color="#1a252f",
            width=145,
            height=32,
            font=ctk.CTkFont(size=12)
        )
        open_out_btn.pack(side="left", padx=6)

        settings_btn = ctk.CTkButton(
            nav_btns,
            text="⚙ Settings & AI",
            command=self._open_settings,
            fg_color="#4a5568",
            hover_color="#2d3748",
            width=120,
            height=32,
            font=ctk.CTkFont(size=12)
        )
        settings_btn.pack(side="left", padx=6)

        # 2. Main File & Folder Selection Section
        config_card = ctk.CTkFrame(self, corner_radius=10)
        config_card.pack(fill="x", padx=16, pady=(12, 6))

        # Excel File Row
        f_row = ctk.CTkFrame(config_card, fg_color="transparent")
        f_row.pack(fill="x", padx=15, pady=(12, 6))

        ctk.CTkLabel(
            f_row,
            text="Excel File (.xlsx / .xlsm):",
            width=170,
            anchor="w",
            font=ctk.CTkFont(size=12, weight="bold")
        ).pack(side="left")

        self.file_entry = ctk.CTkEntry(
            f_row,
            placeholder_text="Click 'Browse Excel File' or load sample catalog...",
            font=ctk.CTkFont(size=12)
        )
        self.file_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))

        browse_file_btn = ctk.CTkButton(
            f_row,
            text="Browse Excel File",
            command=self._browse_excel_file,
            width=135,
            height=32,
            font=ctk.CTkFont(size=12, weight="bold")
        )
        browse_file_btn.pack(side="left")

        # Output Folder Row
        o_row = ctk.CTkFrame(config_card, fg_color="transparent")
        o_row.pack(fill="x", padx=15, pady=(4, 12))

        ctk.CTkLabel(
            o_row,
            text="Output Directory:",
            width=170,
            anchor="w",
            font=ctk.CTkFont(size=12, weight="bold")
        ).pack(side="left")

        self.out_entry = ctk.CTkEntry(
            o_row,
            placeholder_text="Destination folder for exported .jpg images...",
            font=ctk.CTkFont(size=12)
        )
        self.out_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))

        browse_out_btn = ctk.CTkButton(
            o_row,
            text="Browse Folder",
            command=self._browse_output_folder,
            fg_color="gray30",
            hover_color="gray40",
            width=135,
            height=32,
            font=ctk.CTkFont(size=12)
        )
        browse_out_btn.pack(side="left")

        # 3. Center Workspace: Tabview or Placeholder
        self.workspace_frame = ctk.CTkFrame(self, corner_radius=10)
        self.workspace_frame.pack(fill="both", expand=True, padx=16, pady=6)

        self._show_placeholder_view()

        # 4. Bottom Controls & Progress Bar
        bottom_frame = ctk.CTkFrame(self, corner_radius=10)
        bottom_frame.pack(fill="x", padx=16, pady=(6, 12))

        # Options checkboxes and Extract CTA
        opts_row = ctk.CTkFrame(bottom_frame, fg_color="transparent")
        opts_row.pack(fill="x", padx=15, pady=(10, 6))

        self.subfolder_var = ctk.BooleanVar(value=self.config.get("subfolders_per_sheet", True))
        self.subfolder_cb = ctk.CTkCheckBox(
            opts_row,
            text="Organize into subfolders per sheet",
            variable=self.subfolder_var,
            font=ctk.CTkFont(size=12)
        )
        self.subfolder_cb.pack(side="left", padx=(0, 20))

        self.save_once_var = ctk.BooleanVar(value=True)
        self.save_once_cb = ctk.CTkCheckBox(
            opts_row,
            text="Save same Design # once (No suffix _ saves)",
            variable=self.save_once_var,
            font=ctk.CTkFont(size=12)
        )
        self.save_once_cb.pack(side="left")

        # Large Extract Button
        self.extract_btn = ctk.CTkButton(
            opts_row,
            text="🚀 Extract All Images (JPG)",
            command=self._start_export,
            fg_color="#2b825b",
            hover_color="#216446",
            width=230,
            height=38,
            font=ctk.CTkFont(size=14, weight="bold")
        )
        self.extract_btn.pack(side="right")

        # Progress bar & status row
        prog_row = ctk.CTkFrame(bottom_frame, fg_color="transparent")
        prog_row.pack(fill="x", padx=15, pady=(0, 8))

        self.progress_bar = ctk.CTkProgressBar(prog_row, height=10)
        self.progress_bar.set(0)
        self.progress_bar.pack(fill="x", pady=(2, 4))

        self.status_label = ctk.CTkLabel(
            prog_row,
            text="Ready. Select an Excel file to begin.",
            font=ctk.CTkFont(size=12),
            text_color="gray",
            anchor="w"
        )
        self.status_label.pack(side="left")

        # Toggle log button
        self.log_btn = ctk.CTkButton(
            prog_row,
            text="Show Logs ▼",
            command=self._toggle_log_console,
            fg_color="transparent",
            hover_color="gray25",
            text_color="gray70",
            width=90,
            height=20,
            font=ctk.CTkFont(size=11)
        )
        self.log_btn.pack(side="right")

        # Collapsible log box
        self.log_box = ctk.CTkTextbox(
            bottom_frame,
            height=90,
            font=ctk.CTkFont(family="Consolas", size=11),
            fg_color=("gray90", "gray10")
        )
        self.log_box_visible = False

    def _check_for_sample_file(self):
        sample_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "Jewelry_Catalog_Sample.xlsx")
        if os.path.exists(sample_path):
            self.sample_btn.configure(state="normal")
        else:
            self.sample_btn.configure(state="disabled")

    def _show_placeholder_view(self):
        for w in self.workspace_frame.winfo_children():
            w.destroy()

        card = ctk.CTkFrame(self.workspace_frame, fg_color="transparent")
        card.pack(expand=True, pady=40)

        icon_lbl = ctk.CTkLabel(card, text="📊", font=ctk.CTkFont(size=48))
        icon_lbl.pack(pady=(0, 10))

        ctk.CTkLabel(
            card,
            text="No Excel File Loaded",
            font=ctk.CTkFont(size=20, weight="bold")
        ).pack(pady=(0, 6))

        ctk.CTkLabel(
            card,
            text="Select an Excel file (.xlsx) or click 'Load Sample Catalog' above.\n"
                 "The app will automatically locate your 'Design Number' and 'Image' columns across all sheets.",
            font=ctk.CTkFont(size=13),
            text_color="gray",
            justify="center"
        ).pack(pady=(0, 20))

        load_btn = ctk.CTkButton(
            card,
            text="Browse Excel File",
            command=self._browse_excel_file,
            width=160,
            height=36,
            font=ctk.CTkFont(size=13, weight="bold")
        )
        load_btn.pack()

    def _toggle_log_console(self):
        if self.log_box_visible:
            self.log_box.pack_forget()
            self.log_btn.configure(text="Show Logs ▼")
            self.log_box_visible = False
        else:
            self.log_box.pack(fill="x", padx=15, pady=(0, 8))
            self.log_btn.configure(text="Hide Logs ▲")
            self.log_box_visible = True

    def log(self, message: str):
        self.log_box.insert("end", f"{message}\n")
        self.log_box.see("end")

    def _open_settings(self):
        SettingsModal(self, on_save_callback=self._on_settings_saved)

    def _on_settings_saved(self, new_cfg):
        self.config = new_cfg
        ctk.set_appearance_mode(new_cfg.get("theme", "dark"))
        self.log("[System] Settings updated.")

    def _open_output_folder(self):
        out_dir = self.out_entry.get().strip() or self.output_directory
        if out_dir and os.path.exists(out_dir):
            if sys.platform == "win32":
                os.startfile(out_dir)
            else:
                subprocess.run(["xdg-open", out_dir])
        else:
            messagebox.showinfo("Output Folder", "Output folder does not exist yet. Please run an extraction first.")

    def _browse_excel_file(self):
        path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel Files", "*.xlsx *.xlsm *.xltx"), ("All Files", "*.*")]
        )
        if path:
            self._load_workbook_file(path)

    def _load_sample_catalog(self):
        sample_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "Jewelry_Catalog_Sample.xlsx")
        if os.path.exists(sample_path):
            self._load_workbook_file(sample_path)
        else:
            messagebox.showwarning("Sample Not Found", "Jewelry_Catalog_Sample.xlsx not found.")

    def _browse_output_folder(self):
        folder = filedialog.askdirectory(title="Select Destination Folder for JPG Images")
        if folder:
            self.output_directory = folder
            self.out_entry.delete(0, "end")
            self.out_entry.insert(0, folder)

    def _load_workbook_file(self, file_path: str):
        self.current_file = file_path
        self.file_entry.delete(0, "end")
        self.file_entry.insert(0, file_path)

        # Set default output directory next to file
        file_dir = os.path.dirname(os.path.abspath(file_path))
        base_name = os.path.splitext(os.path.basename(file_path))[0]
        default_out = os.path.join(file_dir, f"Extracted_Images_{base_name}")
        self.output_directory = default_out
        self.out_entry.delete(0, "end")
        self.out_entry.insert(0, default_out)

        # Show scanning status
        self.status_label.configure(text=f"Scanning '{os.path.basename(file_path)}'...")
        self.progress_bar.set(0.1)
        self.log(f"--- Loading Workbook: {file_path} ---")

        # Launch scanning in background thread
        thread = threading.Thread(target=self._scan_thread, args=(file_path,), daemon=True)
        thread.start()

    def _process_ui_queue(self):
        try:
            while not self.ui_queue.empty():
                item = self.ui_queue.get_nowait()
                action = item[0]
                args = item[1:]

                if action == "scan_progress":
                    self._update_scan_progress(args[0], args[1])
                elif action == "scan_completed":
                    self._on_scan_completed(args[0])
                elif action == "scan_failed":
                    self._on_scan_failed(args[0])
                elif action == "export_progress":
                    self._update_export_progress(args[0], args[1], args[2])
                elif action == "export_completed":
                    self._on_export_completed(args[0], args[1])
                elif action == "export_failed":
                    self._on_export_failed(args[0])
        except queue.Empty:
            pass

        self.after(50, self._process_ui_queue)

    def _scan_thread(self, file_path: str):
        try:
            results = scan_excel_workbook(
                file_path,
                progress_callback=lambda msg, p: self.ui_queue.put(("scan_progress", msg, p))
            )
            self.ui_queue.put(("scan_completed", results))
        except Exception as e:
            self.ui_queue.put(("scan_failed", str(e)))

    def _update_scan_progress(self, msg: str, progress: float):
        self.status_label.configure(text=msg)
        self.progress_bar.set(progress)

    def _on_scan_completed(self, results: List[SheetScanResult]):
        self.sheet_results = results
        total_images = sum(len(r.images) for r in results)
        total_sheets = len(results)

        self.status_label.configure(
            text=f"Loaded {total_sheets} sheets. Found {total_images} total images ready for export."
        )
        self.progress_bar.set(1.0)
        self.log(f"Scan complete: {total_sheets} sheets, {total_images} images found.")

        # Build Sheet Tabs UI
        self._build_sheet_tabs()

    def _on_scan_failed(self, error_msg: str):
        self.status_label.configure(text=f"Error reading file: {error_msg}")
        self.progress_bar.set(0)
        self.log(f"[Error] Failed to scan workbook: {error_msg}")
        messagebox.showerror("Error Reading Excel", f"Could not open or parse Excel workbook:\n\n{error_msg}")

    def _build_sheet_tabs(self):
        for w in self.workspace_frame.winfo_children():
            w.destroy()

        if not self.sheet_results:
            self._show_placeholder_view()
            return

        # Create Tabview for sheets
        tabview = ctk.CTkTabview(self.workspace_frame)
        tabview.pack(fill="both", expand=True, padx=4, pady=4)

        for res in self.sheet_results:
            tab_name = f"{res.sheet_name} ({len(res.images)} imgs)"
            tabview.add(tab_name)
            tab_frame = tabview.tab(tab_name)

            # Embed SheetView
            view = SheetView(tab_frame, res, on_log_callback=self.log)
            view.pack(fill="both", expand=True)

    def _start_export(self):
        if not self.sheet_results:
            messagebox.showwarning("No Data", "Please load an Excel file first.")
            return

        out_dir = self.out_entry.get().strip()
        if not out_dir:
            messagebox.showwarning("Missing Output Directory", "Please specify an output folder.")
            return

        # Prepare options
        options = {
            "jpeg_quality": self.config.get("jpeg_quality", 95),
            "duplicate_mode": "once" if self.save_once_var.get() else "overwrite",
            "subfolders_per_sheet": self.subfolder_var.get()
        }

        self.extract_btn.configure(state="disabled", text="⏳ Extracting Images...")
        self.log(f"--- Starting Export to: {out_dir} ---")

        # Run export in background thread
        thread = threading.Thread(
            target=self._export_thread,
            args=(self.sheet_results, out_dir, options),
            daemon=True
        )
        thread.start()

    def _export_thread(self, results, out_dir, options):
        try:
            stats = export_images(
                sheet_results=results,
                output_dir=out_dir,
                options=options,
                progress_callback=lambda msg, p, st: self.ui_queue.put(("export_progress", msg, p, st))
            )
            self.ui_queue.put(("export_completed", stats, out_dir))
        except Exception as e:
            self.ui_queue.put(("export_failed", str(e)))

    def _update_export_progress(self, msg: str, progress: float, stats: dict):
        self.status_label.configure(text=msg)
        self.progress_bar.set(progress)
        self.log(f"[Export] {msg}")

    def _on_export_completed(self, stats: dict, out_dir: str):
        self.extract_btn.configure(state="normal", text="🚀 Extract All Images (JPG)")
        self.progress_bar.set(1.0)
        
        saved = stats.get("saved_count", 0)
        total = stats.get("total_images", 0)
        skipped = stats.get("skipped_count", 0)
        errors = stats.get("error_count", 0)

        summary_msg = f"Extraction Finished!\n\nSuccessfully saved: {saved} of {total} images as JPG.\nSkipped: {skipped}\nErrors: {errors}\n\nSaved location:\n{out_dir}"
        self.status_label.configure(text=f"Extraction complete! Saved {saved} images.")
        self.log(f"=== Export Finished: {saved} saved, {skipped} skipped, {errors} errors ===")

        # Refresh preview in active tabs to reflect "Saved" status
        self._build_sheet_tabs()

        if messagebox.askyesno("Extraction Complete", f"{summary_msg}\n\nWould you like to open the output folder now?"):
            self._open_output_folder()

    def _on_export_failed(self, error_msg: str):
        self.extract_btn.configure(state="normal", text="🚀 Extract All Images (JPG)")
        self.progress_bar.set(0)
        self.status_label.configure(text=f"Export failed: {error_msg}")
        self.log(f"[Error] Export failed: {error_msg}")
        messagebox.showerror("Export Failed", f"An error occurred while saving images:\n\n{error_msg}")

