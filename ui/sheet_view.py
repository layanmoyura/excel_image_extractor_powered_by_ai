"""SheetView component: Displays column mapping controls, AI analysis,
and live preview table with image thumbnails for an individual worksheet.
"""

import customtkinter as ctk
from typing import Callable, Optional, Dict, Any, List
from core.excel_engine import SheetScanResult, SheetRowItem
from core.ai_assistant import analyze_columns_with_ai, load_config
from core.column_detector import detect_sheet_columns


class SheetView(ctk.CTkFrame):
    def __init__(
        self,
        parent,
        sheet_result: SheetScanResult,
        on_log_callback: Optional[Callable[[str], None]] = None
    ):
        super().__init__(parent, fg_color="transparent")
        self.sheet_result = sheet_result
        self.on_log_callback = on_log_callback
        
        # Keep CTkImage references alive to prevent Python garbage collection
        self._thumb_refs: List[ctk.CTkImage] = []

        self._build_ui()
        self.refresh_preview()

    def _log(self, msg: str):
        if self.on_log_callback:
            self.on_log_callback(msg)

    def _build_ui(self):
        # 1. Top Control & Detection Panel
        ctrl_card = ctk.CTkFrame(self, corner_radius=10)
        ctrl_card.pack(fill="x", padx=10, pady=(10, 6))

        # Top row: detection info badge & AI button
        top_row = ctk.CTkFrame(ctrl_card, fg_color="transparent")
        top_row.pack(fill="x", padx=12, pady=(10, 6))

        # Confidence badge
        conf_pct = int(self.sheet_result.detection.confidence * 100)
        conf_color = "#1f6f43" if conf_pct >= 80 else "#996515" if conf_pct >= 50 else "#8b2626"

        self.conf_badge = ctk.CTkLabel(
            top_row,
            text=f" AI Auto-Detection: {conf_pct}% Confidence ",
            font=ctk.CTkFont(size=12, weight="bold"),
            fg_color=conf_color,
            corner_radius=6,
            text_color="white"
        )
        self.conf_badge.pack(side="left")

        img_count = len(self.sheet_result.images)
        self.stats_label = ctk.CTkLabel(
            top_row,
            text=f"  |   {img_count} Images Found   |   {len(self.sheet_result.raw_rows)} Total Rows",
            font=ctk.CTkFont(size=12),
            text_color="gray"
        )
        self.stats_label.pack(side="left", padx=5)

        # AI Assistant button
        ai_btn = ctk.CTkButton(
            top_row,
            text="✨ AI Analyze Sheet",
            command=self._run_ai_analysis,
            fg_color="#3a506b",
            hover_color="#2b3d52",
            height=28,
            width=130,
            font=ctk.CTkFont(size=12, weight="bold")
        )
        ai_btn.pack(side="right")

        # Reasoning banner
        self.reason_label = ctk.CTkLabel(
            ctrl_card,
            text=self.sheet_result.detection.reasoning,
            font=ctk.CTkFont(size=11),
            text_color="#9aa0a6",
            anchor="w",
            justify="left"
        )
        self.reason_label.pack(fill="x", padx=12, pady=(0, 8))

        # Divider
        sep = ctk.CTkFrame(ctrl_card, height=1, fg_color="gray30")
        sep.pack(fill="x", padx=12, pady=(0, 8))

        # Bottom row: Dropdowns for column overrides
        map_row = ctk.CTkFrame(ctrl_card, fg_color="transparent")
        map_row.pack(fill="x", padx=12, pady=(0, 10))

        # Design Number dropdown
        ctk.CTkLabel(
            map_row,
            text="Design Number Column:",
            font=ctk.CTkFont(size=12, weight="bold")
        ).pack(side="left", padx=(0, 6))

        # Column options
        col_options = [c["display"] for c in self.sheet_result.detection.all_columns]
        if not col_options:
            col_options = ["Col A"]

        self.design_var = ctk.StringVar()
        cur_design_idx = self.sheet_result.detection.design_col_idx
        cur_design_disp = col_options[cur_design_idx] if cur_design_idx < len(col_options) else col_options[0]
        self.design_var.set(cur_design_disp)

        self.design_menu = ctk.CTkOptionMenu(
            map_row,
            values=col_options,
            variable=self.design_var,
            command=self._on_column_changed,
            width=210,
            height=28
        )
        self.design_menu.pack(side="left", padx=(0, 20))

        # Image Column dropdown
        ctk.CTkLabel(
            map_row,
            text="Image Column:",
            font=ctk.CTkFont(size=12, weight="bold")
        ).pack(side="left", padx=(0, 6))

        self.image_var = ctk.StringVar()
        cur_img_idx = self.sheet_result.detection.image_col_idx
        cur_img_disp = col_options[cur_img_idx] if cur_img_idx < len(col_options) else col_options[0]
        self.image_var.set(cur_img_disp)

        self.image_menu = ctk.CTkOptionMenu(
            map_row,
            values=col_options,
            variable=self.image_var,
            command=self._on_column_changed,
            width=200,
            height=28
        )
        self.image_menu.pack(side="left", padx=(0, 15))

        # Reset button
        reset_btn = ctk.CTkButton(
            map_row,
            text="Reset to Auto",
            command=self._reset_to_auto,
            fg_color="gray30",
            hover_color="gray40",
            width=100,
            height=28
        )
        reset_btn.pack(side="right")

        # 2. Preview Table Section
        table_container = ctk.CTkFrame(self, corner_radius=10)
        table_container.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        # Table Header
        th = ctk.CTkFrame(table_container, height=32, fg_color="gray20", corner_radius=6)
        th.pack(fill="x", padx=6, pady=(6, 4))

        headers = [
            ("#", 45),
            ("Row", 60),
            ("Preview", 80),
            ("Design Number", 220),
            ("Target Output Filename", 260),
            ("Status", 100)
        ]
        for title, w in headers:
            lbl = ctk.CTkLabel(
                th,
                text=title,
                width=w,
                font=ctk.CTkFont(size=12, weight="bold"),
                anchor="center" if title in ("#", "Row", "Preview", "Status") else "w"
            )
            lbl.pack(side="left", padx=4)

        # Scrollable table body
        self.table_scroll = ctk.CTkScrollableFrame(table_container, fg_color="transparent")
        self.table_scroll.pack(fill="both", expand=True, padx=4, pady=(0, 6))

    def _parse_col_index_from_display(self, display_str: str) -> int:
        for c in self.sheet_result.detection.all_columns:
            if c["display"] == display_str:
                return c["index"]
        return 0

    def _on_column_changed(self, _=None):
        d_idx = self._parse_col_index_from_display(self.design_var.get())
        i_idx = self._parse_col_index_from_display(self.image_var.get())
        self.sheet_result.remap_items(design_col_idx=d_idx, image_col_idx=i_idx)
        self._log(f"[{self.sheet_result.sheet_name}] Manually mapped Design Col -> {self.sheet_result.detection.design_col_letter}, Image Col -> {self.sheet_result.detection.image_col_letter}")
        self.refresh_preview()

    def _reset_to_auto(self):
        anchor_cols = [img.col for img in self.sheet_result.images]
        det = detect_sheet_columns(
            self.sheet_result.sheet_name,
            self.sheet_result.raw_rows,
            anchor_cols
        )
        self.sheet_result.detection = det
        self.sheet_result.remap_items(det.design_col_idx, det.image_col_idx)

        # Update menus
        col_options = [c["display"] for c in det.all_columns]
        self.design_var.set(col_options[det.design_col_idx])
        self.image_var.set(col_options[det.image_col_idx])

        # Update badge
        conf_pct = int(det.confidence * 100)
        conf_color = "#1f6f43" if conf_pct >= 80 else "#996515" if conf_pct >= 50 else "#8b2626"
        self.conf_badge.configure(
            text=f" AI Auto-Detection: {conf_pct}% Confidence ",
            fg_color=conf_color
        )
        self.reason_label.configure(text=det.reasoning)
        self._log(f"[{self.sheet_result.sheet_name}] Re-applied Auto-Detection: Design Col={det.design_col_letter}, Image Col={det.image_col_letter}")
        self.refresh_preview()

    def _run_ai_analysis(self):
        cfg = load_config()
        provider = cfg.get("ai_provider", "local")
        api_key = cfg.get("gemini_api_key" if provider == "gemini" else "openai_api_key", "")
        model_name = cfg.get("gemini_model" if provider == "gemini" else "openai_model", "")

        headers = [c["name"] for c in self.sheet_result.detection.all_columns]
        h_row = self.sheet_result.detection.header_row_idx
        sample_rows = self.sheet_result.raw_rows[h_row:h_row + 8]

        self._log(f"[{self.sheet_result.sheet_name}] Running AI analysis with provider: {provider.upper()}...")

        try:
            res = analyze_columns_with_ai(
                sheet_name=self.sheet_result.sheet_name,
                headers=headers,
                sample_rows=sample_rows,
                provider=provider,
                api_key=api_key,
                model_name=model_name
            )

            if res.get("success"):
                d_idx = res.get("design_col_index", 0)
                i_idx = res.get("image_col_index", 1)
                reason = res.get("reasoning", "")
                conf = float(res.get("confidence", 0.95))

                col_options = [c["display"] for c in self.sheet_result.detection.all_columns]
                if d_idx < len(col_options):
                    self.design_var.set(col_options[d_idx])
                if i_idx < len(col_options):
                    self.image_var.set(col_options[i_idx])

                self.sheet_result.detection.confidence = conf
                self.sheet_result.detection.reasoning = f"[{provider.upper()} AI]: {reason}"
                self.sheet_result.remap_items(d_idx, i_idx)

                conf_pct = int(conf * 100)
                self.conf_badge.configure(
                    text=f" {provider.upper()} AI: {conf_pct}% Confidence ",
                    fg_color="#1f6f43"
                )
                self.reason_label.configure(text=self.sheet_result.detection.reasoning)
                self._log(f"[{self.sheet_result.sheet_name}] AI successfully identified columns: Design={d_idx}, Image={i_idx}")
                self.refresh_preview()
            else:
                self._reset_to_auto()
        except Exception as e:
            self._log(f"[{self.sheet_result.sheet_name}] AI Analysis note: {e}. Falling back to Smart Local AI.")
            self._reset_to_auto()

    def refresh_preview(self):
        """Rebuilds the preview table rows."""
        # Clear existing rows
        for widget in self.table_scroll.winfo_children():
            widget.destroy()
        self._thumb_refs.clear()

        items = self.sheet_result.items
        if not items:
            empty_lbl = ctk.CTkLabel(
                self.table_scroll,
                text="No data or images found in this sheet.",
                font=ctk.CTkFont(size=13),
                text_color="gray"
            )
            empty_lbl.pack(pady=30)
            return

        # Show up to 150 preview rows for instant responsiveness
        for idx, item in enumerate(items[:150], 1):
            row_frame = ctk.CTkFrame(
                self.table_scroll,
                height=52,
                fg_color="gray15" if idx % 2 == 0 else "gray18",
                corner_radius=4
            )
            row_frame.pack(fill="x", pady=2, padx=2)

            # # index
            ctk.CTkLabel(row_frame, text=str(idx), width=45, font=ctk.CTkFont(size=11), text_color="gray").pack(side="left", padx=4)

            # Excel Row
            ctk.CTkLabel(row_frame, text=f"Row {item.excel_row}", width=60, font=ctk.CTkFont(size=11)).pack(side="left", padx=4)

            # Thumbnail preview
            thumb_frame = ctk.CTkFrame(row_frame, width=80, height=48, fg_color="transparent")
            thumb_frame.pack(side="left", padx=4)

            if item.image and item.image.thumbnail:
                try:
                    ctk_img = ctk.CTkImage(light_image=item.image.thumbnail, dark_image=item.image.thumbnail, size=(42, 42))
                    self._thumb_refs.append(ctk_img)
                    img_label = ctk.CTkLabel(thumb_frame, text="", image=ctk_img)
                    img_label.pack(anchor="center")
                except Exception:
                    ctk.CTkLabel(thumb_frame, text="[Img]", font=ctk.CTkFont(size=10), text_color="gray").pack(anchor="center")
            else:
                ctk.CTkLabel(thumb_frame, text="No Image", font=ctk.CTkFont(size=10), text_color="gray50").pack(anchor="center")

            # Design Number
            ctk.CTkLabel(
                row_frame,
                text=item.raw_design_number or "[Missing]",
                width=220,
                anchor="w",
                font=ctk.CTkFont(size=12, weight="bold" if item.raw_design_number else "normal"),
                text_color="white" if item.raw_design_number else "orange"
            ).pack(side="left", padx=4)

            # Proposed Filename
            ctk.CTkLabel(
                row_frame,
                text=item.proposed_filename if item.image else "-",
                width=260,
                anchor="w",
                font=ctk.CTkFont(family="Consolas", size=11),
                text_color="#80c3ff" if item.image else "gray50"
            ).pack(side="left", padx=4)

            # Status pill
            status = item.status
            status_colors = {
                "Ready": ("#1f6f43", "white"),
                "Saved": ("#195b8c", "white"),
                "No Image": ("gray30", "gray70"),
                "Missing Design #": ("#996515", "white"),
                "Error": ("#8b2626", "white"),
                "Skipped": ("gray40", "gray80")
            }
            bg_col, txt_col = status_colors.get(status, ("gray30", "white"))

            st_lbl = ctk.CTkLabel(
                row_frame,
                text=f" {status} ",
                width=90,
                font=ctk.CTkFont(size=10, weight="bold"),
                fg_color=bg_col,
                text_color=txt_col,
                corner_radius=4
            )
            st_lbl.pack(side="left", padx=6)

        if len(items) > 150:
            more_lbl = ctk.CTkLabel(
                self.table_scroll,
                text=f"... and {len(items) - 150} more items (all will be exported when you click Extract).",
                font=ctk.CTkFont(size=11, slant="italic"),
                text_color="gray"
            )
            more_lbl.pack(pady=6)
