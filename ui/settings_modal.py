"""Settings Modal Dialog for configuring AI providers, API keys, and export preferences.
"""

import customtkinter as ctk
from typing import Callable, Optional
from core.ai_assistant import load_config, save_config


class SettingsModal(ctk.CTkToplevel):
    def __init__(self, parent, on_save_callback: Optional[Callable] = None):
        super().__init__(parent)
        self.title("Application Settings & AI Configuration")
        self.geometry("540x580")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()

        self.on_save_callback = on_save_callback
        self.config = load_config()

        self._build_ui()

    def _build_ui(self):
        # Header
        header_frame = ctk.CTkFrame(self, fg_color="transparent")
        header_frame.pack(fill="x", padx=20, pady=(20, 10))

        title = ctk.CTkLabel(
            header_frame,
            text="Settings & Preferences",
            font=ctk.CTkFont(size=20, weight="bold")
        )
        title.pack(anchor="w")

        subtitle = ctk.CTkLabel(
            header_frame,
            text="Configure Smart AI detection, API keys, and JPEG output options",
            font=ctk.CTkFont(size=12),
            text_color="gray"
        )
        subtitle.pack(anchor="w")

        # Scrollable content
        content_frame = ctk.CTkScrollableFrame(self, fg_color="transparent", height=420)
        content_frame.pack(fill="both", expand=True, padx=20, pady=5)

        # Section 1: AI Provider
        ai_box = ctk.CTkFrame(content_frame, corner_radius=10)
        ai_box.pack(fill="x", pady=8, padx=5)

        ctk.CTkLabel(
            ai_box,
            text="AI & Column Detection Engine",
            font=ctk.CTkFont(size=14, weight="bold")
        ).pack(anchor="w", padx=15, pady=(12, 4))

        ctk.CTkLabel(
            ai_box,
            text="Choose between offline Built-in Smart Heuristics or Cloud LLMs:",
            font=ctk.CTkFont(size=11),
            text_color="gray"
        ).pack(anchor="w", padx=15, pady=(0, 8))

        self.provider_var = ctk.StringVar(value=self.config.get("ai_provider", "local"))
        provider_menu = ctk.CTkSegmentedButton(
            ai_box,
            values=["Local Smart AI", "Google Gemini", "OpenAI"],
            command=self._on_provider_change
        )
        # Map values to internal names
        provider_display = {
            "local": "Local Smart AI",
            "gemini": "Google Gemini",
            "openai": "OpenAI"
        }
        provider_menu.set(provider_display.get(self.config.get("ai_provider", "local"), "Local Smart AI"))
        provider_menu.pack(fill="x", padx=15, pady=(0, 12))
        self.provider_menu = provider_menu

        # Gemini settings
        self.gemini_frame = ctk.CTkFrame(ai_box, fg_color="transparent")
        ctk.CTkLabel(self.gemini_frame, text="Gemini API Key:", font=ctk.CTkFont(size=12)).pack(anchor="w", pady=(2, 2))
        self.gemini_key_entry = ctk.CTkEntry(self.gemini_frame, placeholder_text="AIzaSy...", show="*")
        self.gemini_key_entry.insert(0, self.config.get("gemini_api_key", ""))
        self.gemini_key_entry.pack(fill="x", pady=(0, 6))

        ctk.CTkLabel(self.gemini_frame, text="Model:", font=ctk.CTkFont(size=12)).pack(anchor="w", pady=(2, 2))
        self.gemini_model_entry = ctk.CTkEntry(self.gemini_frame, placeholder_text="gemini-1.5-flash")
        self.gemini_model_entry.insert(0, self.config.get("gemini_model", "gemini-1.5-flash"))
        self.gemini_model_entry.pack(fill="x", pady=(0, 12))

        # OpenAI settings
        self.openai_frame = ctk.CTkFrame(ai_box, fg_color="transparent")
        ctk.CTkLabel(self.openai_frame, text="OpenAI API Key:", font=ctk.CTkFont(size=12)).pack(anchor="w", pady=(2, 2))
        self.openai_key_entry = ctk.CTkEntry(self.openai_frame, placeholder_text="sk-...", show="*")
        self.openai_key_entry.insert(0, self.config.get("openai_api_key", ""))
        self.openai_key_entry.pack(fill="x", pady=(0, 6))

        ctk.CTkLabel(self.openai_frame, text="Model:", font=ctk.CTkFont(size=12)).pack(anchor="w", pady=(2, 2))
        self.openai_model_entry = ctk.CTkEntry(self.openai_frame, placeholder_text="gpt-4o-mini")
        self.openai_model_entry.insert(0, self.config.get("openai_model", "gpt-4o-mini"))
        self.openai_model_entry.pack(fill="x", pady=(0, 12))

        # Update initial visibility
        self._update_provider_fields(self.config.get("ai_provider", "local"))

        # Section 2: JPEG Quality & Output
        output_box = ctk.CTkFrame(content_frame, corner_radius=10)
        output_box.pack(fill="x", pady=8, padx=5)

        ctk.CTkLabel(
            output_box,
            text="JPEG Quality & Image Settings",
            font=ctk.CTkFont(size=14, weight="bold")
        ).pack(anchor="w", padx=15, pady=(12, 4))

        quality_val = self.config.get("jpeg_quality", 95)
        self.quality_label = ctk.CTkLabel(output_box, text=f"JPEG Quality: {quality_val}%")
        self.quality_label.pack(anchor="w", padx=15, pady=(4, 0))

        self.quality_slider = ctk.CTkSlider(
            output_box,
            from_=70,
            to=100,
            number_of_steps=30,
            command=self._on_quality_change
        )
        self.quality_slider.set(quality_val)
        self.quality_slider.pack(fill="x", padx=15, pady=(2, 12))

        # Duplicate mode
        ctk.CTkLabel(
            output_box,
            text="When Duplicate Design Numbers Exist:",
            font=ctk.CTkFont(size=12)
        ).pack(anchor="w", padx=15, pady=(4, 2))

        self.dup_mode_var = ctk.StringVar(value=self.config.get("duplicate_mode", "once"))
        dup_options = [
            ("Save same design number once (No suffix _ files)", "once"),
            ("Overwrite existing file", "overwrite")
        ]
        for text, val in dup_options:
            r = ctk.CTkRadioButton(
                output_box,
                text=text,
                variable=self.dup_mode_var,
                value=val
            )
            r.pack(anchor="w", padx=20, pady=3)

        # Section 3: UI Theme
        theme_box = ctk.CTkFrame(content_frame, corner_radius=10)
        theme_box.pack(fill="x", pady=8, padx=5)

        ctk.CTkLabel(
            theme_box,
            text="Appearance Theme",
            font=ctk.CTkFont(size=14, weight="bold")
        ).pack(anchor="w", padx=15, pady=(12, 4))

        self.theme_var = ctk.StringVar(value=self.config.get("theme", "dark").capitalize())
        theme_menu = ctk.CTkOptionMenu(
            theme_box,
            values=["Dark", "Light", "System"],
            variable=self.theme_var,
            command=lambda v: ctk.set_appearance_mode(v.lower())
        )
        theme_menu.pack(anchor="w", padx=15, pady=(4, 12))

        # Bottom buttons
        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.pack(fill="x", padx=20, pady=(5, 15))

        save_btn = ctk.CTkButton(
            btn_frame,
            text="Save Settings",
            command=self._save_and_close,
            fg_color="#2b825b",
            hover_color="#236848",
            width=130,
            height=36
        )
        save_btn.pack(side="right", padx=(10, 0))

        cancel_btn = ctk.CTkButton(
            btn_frame,
            text="Cancel",
            command=self.destroy,
            fg_color="gray30",
            hover_color="gray40",
            width=90,
            height=36
        )
        cancel_btn.pack(side="right")

    def _on_provider_change(self, val):
        mapping = {
            "Local Smart AI": "local",
            "Google Gemini": "gemini",
            "OpenAI": "openai"
        }
        chosen = mapping.get(val, "local")
        self._update_provider_fields(chosen)

    def _update_provider_fields(self, provider: str):
        self.gemini_frame.pack_forget()
        self.openai_frame.pack_forget()

        if provider == "gemini":
            self.gemini_frame.pack(fill="x", padx=15, pady=5)
        elif provider == "openai":
            self.openai_frame.pack(fill="x", padx=15, pady=5)

    def _on_quality_change(self, val):
        self.quality_label.configure(text=f"JPEG Quality: {int(val)}%")

    def _save_and_close(self):
        mapping = {
            "Local Smart AI": "local",
            "Google Gemini": "gemini",
            "OpenAI": "openai"
        }
        prov = mapping.get(self.provider_menu.get(), "local")

        new_config = {
            "ai_provider": prov,
            "gemini_api_key": self.gemini_key_entry.get().strip(),
            "gemini_model": self.gemini_model_entry.get().strip() or "gemini-1.5-flash",
            "openai_api_key": self.openai_key_entry.get().strip(),
            "openai_model": self.openai_model_entry.get().strip() or "gpt-4o-mini",
            "jpeg_quality": int(self.quality_slider.get()),
            "duplicate_mode": self.dup_mode_var.get(),
            "theme": self.theme_var.get().lower()
        }

        save_config(new_config)
        if self.on_save_callback:
            self.on_save_callback(new_config)

        self.destroy()
