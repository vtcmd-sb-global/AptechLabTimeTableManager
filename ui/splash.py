#!/usr/bin/env python3
"""Professional startup splash screen for Aptech Monthly BSR."""

from __future__ import annotations

import tkinter as tk
from typing import Optional


# Match main app header palette
_HEADER = "#1a1f36"
_ACCENT = "#5b6cff"
_MUTED = "#a8b0d0"
_SURFACE = "#ffffff"
_TROUGH = "#e6e6e6"
_GREEN = "#06b025"
_GREEN_TOP = "#3ddc65"
_GREEN_EDGE = "#7aef95"
_GREEN_BOT = "#05941f"


class SplashScreen(tk.Toplevel):
    """Borderless centered splash shown while the main window initializes."""

    def __init__(self, master: tk.Misc, app_name: str = "Aptech Monthly BSR", version: str = "1.0"):
        super().__init__(master)
        self._phase = 0.0
        self._after_id: Optional[str] = None
        self._closed = False

        self.overrideredirect(True)
        self.configure(bg=_HEADER)
        try:
            self.attributes("-topmost", True)
        except Exception:
            pass

        w, h = 480, 300
        self.geometry(f"{w}x{h}")
        self._center(w, h)

        # Outer card
        card = tk.Frame(self, bg=_HEADER, highlightthickness=0)
        card.pack(fill=tk.BOTH, expand=True)

        # Accent top line
        tk.Frame(card, bg=_ACCENT, height=3).pack(fill=tk.X)

        body = tk.Frame(card, bg=_HEADER)
        body.pack(fill=tk.BOTH, expand=True, padx=36, pady=28)

        # Brand mark
        mark = tk.Canvas(body, width=56, height=56, bg=_HEADER, highlightthickness=0)
        mark.pack(pady=(8, 12))
        mark.create_oval(4, 4, 52, 52, fill=_ACCENT, outline="")
        mark.create_text(28, 28, text="A", fill=_SURFACE, font=("Segoe UI", 22, "bold"))

        tk.Label(
            body, text=app_name,
            bg=_HEADER, fg=_SURFACE,
            font=("Segoe UI", 18, "bold"),
        ).pack()

        tk.Label(
            body, text="Batch Status Report  ·  Faculty Workload  ·  Monthly Generation",
            bg=_HEADER, fg=_MUTED,
            font=("Segoe UI", 9),
        ).pack(pady=(6, 0))

        tk.Label(
            body, text=f"Version {version}",
            bg=_HEADER, fg="#6b7288",
            font=("Segoe UI", 8),
        ).pack(pady=(14, 0))

        # Status + Windows-style progress
        bottom = tk.Frame(card, bg=_HEADER)
        bottom.pack(fill=tk.X, side=tk.BOTTOM, padx=36, pady=(0, 24))

        self.status_var = tk.StringVar(value="Starting…")
        tk.Label(
            bottom, textvariable=self.status_var,
            bg=_HEADER, fg=_MUTED,
            font=("Segoe UI", 9), anchor="w",
        ).pack(fill=tk.X, pady=(0, 8))

        self.bar = tk.Canvas(
            bottom, height=14, bg=_TROUGH,
            highlightthickness=1, highlightbackground="#adadad", bd=0,
        )
        self.bar.pack(fill=tk.X)
        self._bar_w = 400
        self.bar.bind("<Configure>", self._on_bar_configure)

        self._draw_bar(0.12)
        self._tick()

        # Ensure splash paints before heavy work
        self.update_idletasks()
        self.update()

    def _center(self, w: int, h: int) -> None:
        try:
            self.update_idletasks()
            sw = self.winfo_screenwidth()
            sh = self.winfo_screenheight()
            x = max(0, (sw - w) // 2)
            y = max(0, (sh - h) // 2)
            self.geometry(f"{w}x{h}+{x}+{y}")
        except Exception:
            pass

    def _on_bar_configure(self, event) -> None:
        self._bar_w = max(40, event.width)
        self._draw_bar(min(0.95, 0.15 + self._phase * 0.8))

    def set_status(self, text: str) -> None:
        if self._closed:
            return
        try:
            self.status_var.set(text)
            self.update_idletasks()
        except Exception:
            pass

    def pulse(self) -> None:
        """Keep animation alive during long synchronous work."""
        if self._closed:
            return
        try:
            self.update()
        except Exception:
            pass

    def _tick(self) -> None:
        if self._closed:
            return
        self._phase = (self._phase + 0.025) % 1.0
        # Soft progress that creeps forward while loading (not a fake 100%)
        frac = 0.12 + 0.78 * self._phase
        # Bounce gently near the end until close()
        if self._phase > 0.92:
            frac = 0.88 + 0.06 * ((self._phase - 0.92) / 0.08)
        self._draw_bar(frac)
        try:
            self._after_id = self.after(22, self._tick)
        except Exception:
            self._after_id = None

    def _draw_bar(self, frac: float) -> None:
        if self._closed:
            return
        try:
            self.bar.delete("all")
            w = max(self._bar_w, 40)
            h = 14
            self.bar.create_rectangle(0, 0, w, h, fill=_TROUGH, outline="")
            fw = max(2, int(w * max(0.0, min(1.0, frac))))
            self.bar.create_rectangle(0, 0, fw, h, fill=_GREEN, outline="")
            self.bar.create_rectangle(0, 0, fw, max(2, h // 2), fill=_GREEN_TOP, outline="")
            self.bar.create_rectangle(0, 0, fw, max(1, h // 5), fill=_GREEN_EDGE, outline="")
            self.bar.create_rectangle(0, max(0, h - max(2, h // 4)), fw, h, fill=_GREEN_BOT, outline="")
            # Moving sheen
            band = max(28, int(fw * 0.4))
            x0 = int(self._phase * (fw + band)) - band
            left = max(0, x0)
            right = min(fw, x0 + band)
            if right > left:
                self.bar.create_rectangle(left, 0, right, h, fill="#ffffff", stipple="gray25", outline="")
                mid = max(4, (right - left) // 4)
                if right - left > mid * 2:
                    self.bar.create_rectangle(
                        left + mid, 0, right - mid, h,
                        fill="#ffffff", stipple="gray50", outline="",
                    )
        except Exception:
            pass

    def close_splash(self) -> None:
        if self._closed:
            return
        self._closed = True
        if self._after_id is not None:
            try:
                self.after_cancel(self._after_id)
            except Exception:
                pass
            self._after_id = None
        try:
            self.destroy()
        except Exception:
            pass
