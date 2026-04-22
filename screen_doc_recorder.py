#!/usr/bin/env python3
"""
IT-Dokumentationsassistent
Zeichnet Bildschirmaktionen auf und erstellt daraus eine HTML-Dokumentation
mit Screenshots und Aktionsbeschreibungen.

Hotkeys:
  F8  - Aufnahme starten / stoppen
  F9  - Manuellen Screenshot erstellen
  F10 - Notiz zum aktuellen Schritt hinzufügen
"""

import os
import sys
import time
import base64
import threading
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
from datetime import datetime
from pathlib import Path
from io import BytesIO
from dataclasses import dataclass, field
from typing import List, Optional
import html as html_module
import queue

# ---------------------------------------------------------------------------
# Dependency check / install
# ---------------------------------------------------------------------------
REQUIRED = {"mss": "mss", "PIL": "Pillow", "pynput": "pynput"}

missing = []
for mod, pkg in REQUIRED.items():
    try:
        __import__(mod)
    except ImportError:
        missing.append(pkg)

if missing:
    print(f"Installiere fehlende Pakete: {', '.join(missing)}")
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install", "--quiet"] + missing)

import mss
import mss.tools
from PIL import Image
from pynput import mouse as pmouse, keyboard as pkeyboard

# ---------------------------------------------------------------------------
# Data model
# ---------------------------------------------------------------------------

@dataclass
class ActionEvent:
    timestamp: float
    action_type: str          # click | scroll | key | screenshot | note | start | stop
    description: str
    screenshot_b64: Optional[str] = None
    note: str = ""
    x: int = 0
    y: int = 0

# ---------------------------------------------------------------------------
# Screen capture
# ---------------------------------------------------------------------------

class ScreenCapture:
    def __init__(self):
        self._lock = threading.Lock()

    def capture(self, quality: int = 85) -> Optional[str]:
        """Capture full screen, return base64-encoded JPEG string."""
        try:
            with self._lock:
                with mss.mss() as sct:
                    # Capture all monitors combined
                    monitor = sct.monitors[0]  # monitor[0] = all screens combined
                    sct_img = sct.grab(monitor)
                img = Image.frombytes("RGB", sct_img.size, sct_img.bgra, "raw", "BGRX")
                buf = BytesIO()
                img.save(buf, format="JPEG", quality=quality, optimize=True)
                return base64.b64encode(buf.getvalue()).decode("utf-8")
        except Exception as e:
            print(f"Screenshot-Fehler: {e}")
            return None

    def capture_thumbnail(self, quality: int = 75, max_width: int = 1280) -> Optional[str]:
        """Capture and downscale for smaller file size."""
        try:
            with self._lock:
                with mss.mss() as sct:
                    monitor = sct.monitors[0]
                    sct_img = sct.grab(monitor)
                img = Image.frombytes("RGB", sct_img.size, sct_img.bgra, "raw", "BGRX")
                if img.width > max_width:
                    ratio = max_width / img.width
                    new_size = (max_width, int(img.height * ratio))
                    img = img.resize(new_size, Image.LANCZOS)
                buf = BytesIO()
                img.save(buf, format="JPEG", quality=quality, optimize=True)
                return base64.b64encode(buf.getvalue()).decode("utf-8")
        except Exception as e:
            print(f"Screenshot-Fehler: {e}")
            return None

# ---------------------------------------------------------------------------
# Event tracker
# ---------------------------------------------------------------------------

class EventTracker:
    def __init__(self, on_event):
        self._on_event = on_event
        self._mouse_listener = None
        self._keyboard_listener = None
        self._active = False
        self._last_click_time = 0.0
        self._min_click_interval = 0.5  # deduplicate rapid clicks

    def start(self):
        self._active = True
        self._mouse_listener = pmouse.Listener(
            on_click=self._on_click,
            on_scroll=self._on_scroll,
        )
        self._keyboard_listener = pkeyboard.Listener(
            on_press=self._on_key_press,
        )
        self._mouse_listener.start()
        self._keyboard_listener.start()

    def stop(self):
        self._active = False
        if self._mouse_listener:
            self._mouse_listener.stop()
        if self._keyboard_listener:
            self._keyboard_listener.stop()

    def _on_click(self, x, y, button, pressed):
        if not self._active or not pressed:
            return
        now = time.time()
        if now - self._last_click_time < self._min_click_interval:
            return
        self._last_click_time = now
        btn_name = "Links" if button == pmouse.Button.left else (
                   "Rechts" if button == pmouse.Button.right else "Mitte")
        self._on_event("click", f"{btn_name}klick bei ({x}, {y})", x, y)

    def _on_scroll(self, x, y, dx, dy):
        if not self._active:
            return
        direction = "unten" if dy < 0 else "oben"
        self._on_event("scroll", f"Gescrollt nach {direction} bei ({x}, {y})", x, y)

    def _on_key_press(self, key):
        if not self._active:
            return
        try:
            key_str = key.char if hasattr(key, "char") and key.char else str(key).replace("Key.", "")
        except AttributeError:
            key_str = str(key).replace("Key.", "")
        # Only log special keys to avoid capturing passwords
        special = {
            "enter", "tab", "backspace", "delete", "escape", "space",
            "f1", "f2", "f3", "f4", "f5", "f6", "f7",
            "f11", "f12", "ctrl_l", "ctrl_r", "alt_l", "alt_r",
            "shift", "cmd", "page_up", "page_down", "home", "end",
        }
        if key_str.lower() in special:
            self._on_event("key", f"Taste gedrückt: {key_str}", 0, 0)

# ---------------------------------------------------------------------------
# HTML report generator
# ---------------------------------------------------------------------------

HTML_STYLE = """
<style>
  :root {
    --primary: #0078d4;
    --bg: #f3f2f1;
    --card: #ffffff;
    --border: #e1dfdd;
    --text: #323130;
    --muted: #605e5c;
    --click: #107c10;
    --key: #8764b8;
    --scroll: #ca5010;
    --note: #0078d4;
    --manual: #004578;
  }
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: 'Segoe UI', Arial, sans-serif; background: var(--bg);
         color: var(--text); font-size: 14px; }
  header { background: var(--primary); color: #fff; padding: 20px 32px; }
  header h1 { font-size: 22px; font-weight: 600; }
  header .meta { margin-top: 6px; font-size: 13px; opacity: .85; }
  main { max-width: 1200px; margin: 24px auto; padding: 0 16px 60px; }
  .summary { background: var(--card); border: 1px solid var(--border);
             border-radius: 4px; padding: 16px 20px; margin-bottom: 24px;
             display: flex; gap: 32px; flex-wrap: wrap; }
  .summary .stat { text-align: center; }
  .summary .stat .val { font-size: 28px; font-weight: 700; color: var(--primary); }
  .summary .stat .lbl { font-size: 12px; color: var(--muted); margin-top: 2px; }
  .timeline { list-style: none; }
  .timeline li { display: flex; gap: 16px; margin-bottom: 16px; }
  .timeline .idx {
    flex-shrink: 0; width: 36px; height: 36px; border-radius: 50%;
    display: flex; align-items: center; justify-content: center;
    font-weight: 700; font-size: 13px; color: #fff;
  }
  .idx.click  { background: var(--click); }
  .idx.key    { background: var(--key); }
  .idx.scroll { background: var(--scroll); }
  .idx.note   { background: var(--note); }
  .idx.screenshot { background: var(--manual); }
  .idx.start, .idx.stop { background: var(--muted); }
  .card { flex: 1; background: var(--card); border: 1px solid var(--border);
          border-radius: 4px; overflow: hidden; }
  .card-head { padding: 10px 14px; border-bottom: 1px solid var(--border);
               display: flex; justify-content: space-between; align-items: center; }
  .card-head .action { font-weight: 600; }
  .card-head .ts { font-size: 12px; color: var(--muted); }
  .card-body { padding: 10px 14px; }
  .note-text { background: #fffbe6; border-left: 3px solid #ffb900;
               padding: 8px 12px; border-radius: 2px; font-style: italic; }
  details summary {
    cursor: pointer; font-size: 12px; color: var(--primary);
    padding: 6px 0; user-select: none;
  }
  details summary:hover { text-decoration: underline; }
  .screenshot-wrap { margin-top: 8px; }
  .screenshot-wrap img {
    max-width: 100%; border: 1px solid var(--border);
    border-radius: 2px; display: block; cursor: zoom-in;
  }
  /* Lightbox */
  #lb { display:none; position:fixed; inset:0; background:rgba(0,0,0,.85);
        z-index:9999; align-items:center; justify-content:center; }
  #lb.show { display:flex; }
  #lb img { max-width:95vw; max-height:95vh; border-radius:4px; }
  #lb-close { position:fixed; top:16px; right:24px; color:#fff; font-size:32px;
              cursor:pointer; line-height:1; }
</style>
"""

HTML_SCRIPT = """
<script>
  const lb = document.getElementById('lb');
  const lbImg = document.getElementById('lb-img');
  document.querySelectorAll('.screenshot-wrap img').forEach(img => {
    img.addEventListener('click', () => {
      lbImg.src = img.src;
      lb.classList.add('show');
    });
  });
  document.getElementById('lb-close').addEventListener('click', () => lb.classList.remove('show'));
  lb.addEventListener('click', e => { if (e.target === lb) lb.classList.remove('show'); });
  document.addEventListener('keydown', e => { if (e.key === 'Escape') lb.classList.remove('show'); });
</script>
"""

class DocumentGenerator:
    @staticmethod
    def generate(events: List[ActionEvent], title: str, output_path: str) -> str:
        start_ts = events[0].timestamp if events else time.time()
        end_ts = events[-1].timestamp if events else time.time()
        duration_s = int(end_ts - start_ts)
        duration_str = f"{duration_s // 60} Min {duration_s % 60} Sek"

        screenshot_count = sum(1 for e in events if e.screenshot_b64)
        action_count = sum(1 for e in events if e.action_type not in ("start", "stop"))

        date_str = datetime.fromtimestamp(start_ts).strftime("%d.%m.%Y %H:%M")
        safe_title = html_module.escape(title)

        items_html = ""
        step = 0
        for ev in events:
            step += 1
            ts_str = datetime.fromtimestamp(ev.timestamp).strftime("%H:%M:%S")
            atype = html_module.escape(ev.action_type)
            desc = html_module.escape(ev.description)

            # Badge label
            badge_map = {
                "click": str(step), "scroll": "↕", "key": "⌨",
                "screenshot": "📷", "note": "✎", "start": "▶", "stop": "■",
            }
            badge = badge_map.get(ev.action_type, str(step))

            # Action label
            label_map = {
                "click": "Mausklick", "scroll": "Scrollen", "key": "Tastatureingabe",
                "screenshot": "Screenshot", "note": "Notiz",
                "start": "Aufnahme gestartet", "stop": "Aufnahme beendet",
            }
            label = label_map.get(ev.action_type, ev.action_type)

            # Note block
            note_html = ""
            if ev.note:
                note_html = f'<div class="note-text">{html_module.escape(ev.note)}</div>'

            # Screenshot block
            ss_html = ""
            if ev.screenshot_b64:
                ss_html = f"""
                <details open>
                  <summary>Screenshot anzeigen</summary>
                  <div class="screenshot-wrap">
                    <img src="data:image/jpeg;base64,{ev.screenshot_b64}" alt="Screenshot">
                  </div>
                </details>"""

            items_html += f"""
      <li>
        <div class="idx {ev.action_type}">{badge}</div>
        <div class="card">
          <div class="card-head">
            <span class="action">{label}: {desc}</span>
            <span class="ts">{ts_str}</span>
          </div>
          {"<div class='card-body'>" + note_html + ss_html + "</div>" if note_html or ss_html else ""}
        </div>
      </li>"""

        html = f"""<!DOCTYPE html>
<html lang="de">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width,initial-scale=1">
  <title>{safe_title}</title>
  {HTML_STYLE}
</head>
<body>
<header>
  <h1>{safe_title}</h1>
  <div class="meta">Erstellt: {date_str} &nbsp;|&nbsp; Dauer: {duration_str}</div>
</header>
<main>
  <div class="summary">
    <div class="stat"><div class="val">{action_count}</div><div class="lbl">Aktionen</div></div>
    <div class="stat"><div class="val">{screenshot_count}</div><div class="lbl">Screenshots</div></div>
    <div class="stat"><div class="val">{duration_str}</div><div class="lbl">Aufnahmedauer</div></div>
  </div>
  <ul class="timeline">
{items_html}
  </ul>
</main>
<div id="lb"><span id="lb-close">&times;</span><img id="lb-img" src="" alt="Vollbild"></div>
{HTML_SCRIPT}
</body>
</html>"""

        with open(output_path, "w", encoding="utf-8") as f:
            f.write(html)
        return output_path

# ---------------------------------------------------------------------------
# Main application
# ---------------------------------------------------------------------------

class RecorderApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("IT-Dokumentationsassistent")
        self.root.resizable(False, False)

        self.events: List[ActionEvent] = []
        self.recording = False
        self.capture = ScreenCapture()
        self.tracker = EventTracker(self._on_tracked_event)
        self.event_queue: queue.Queue = queue.Queue()

        # Config
        self.capture_on_click = tk.BooleanVar(value=True)
        self.capture_on_scroll = tk.BooleanVar(value=False)
        self.capture_on_key = tk.BooleanVar(value=False)
        self.doc_title = tk.StringVar(value=f"IT-Dokumentation {datetime.now().strftime('%Y-%m-%d')}")
        self.screenshot_quality = tk.IntVar(value=75)
        self.max_width = tk.IntVar(value=1280)

        self._build_ui()
        self._setup_hotkeys()

        # Poll queue from main thread
        self.root.after(100, self._process_queue)

    # ------------------------------------------------------------------
    # UI
    # ------------------------------------------------------------------

    def _build_ui(self):
        self.root.configure(bg="#f3f2f1")
        pad = {"padx": 12, "pady": 6}

        # Title bar area
        title_frame = tk.Frame(self.root, bg="#0078d4", pady=10)
        title_frame.pack(fill=tk.X)
        tk.Label(title_frame, text="IT-Dokumentationsassistent",
                 bg="#0078d4", fg="white",
                 font=("Segoe UI", 13, "bold")).pack(padx=14, side=tk.LEFT)

        # Status
        self.status_var = tk.StringVar(value="Bereit")
        self.status_label = tk.Label(self.root, textvariable=self.status_var,
                                     bg="#f3f2f1", font=("Segoe UI", 10),
                                     anchor="w")
        self.status_label.pack(fill=tk.X, **pad)

        # Controls
        btn_frame = tk.Frame(self.root, bg="#f3f2f1")
        btn_frame.pack(fill=tk.X, padx=12, pady=4)

        self.rec_btn = tk.Button(btn_frame, text="▶  Aufnahme starten (F8)",
                                 command=self.toggle_recording,
                                 bg="#107c10", fg="white",
                                 font=("Segoe UI", 10, "bold"),
                                 relief=tk.FLAT, padx=10, pady=6,
                                 cursor="hand2")
        self.rec_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))

        self.manual_btn = tk.Button(btn_frame, text="📷 Screenshot (F9)",
                                    command=self.manual_screenshot,
                                    state=tk.DISABLED,
                                    bg="#004578", fg="white",
                                    font=("Segoe UI", 10),
                                    relief=tk.FLAT, padx=10, pady=6,
                                    cursor="hand2")
        self.manual_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))

        self.note_btn = tk.Button(btn_frame, text="✎ Notiz (F10)",
                                  command=self.add_note,
                                  state=tk.DISABLED,
                                  bg="#8764b8", fg="white",
                                  font=("Segoe UI", 10),
                                  relief=tk.FLAT, padx=10, pady=6,
                                  cursor="hand2")
        self.note_btn.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # Event counter
        self.counter_var = tk.StringVar(value="Ereignisse: 0  |  Screenshots: 0")
        tk.Label(self.root, textvariable=self.counter_var,
                 bg="#f3f2f1", fg="#605e5c",
                 font=("Segoe UI", 9)).pack(**pad)

        # Settings frame
        settings = ttk.LabelFrame(self.root, text="Einstellungen", padding=8)
        settings.pack(fill=tk.X, padx=12, pady=6)

        tk.Label(settings, text="Titel:", font=("Segoe UI", 9)).grid(
            row=0, column=0, sticky="w", padx=(0, 6))
        tk.Entry(settings, textvariable=self.doc_title, width=34,
                 font=("Segoe UI", 9)).grid(row=0, column=1, columnspan=3, sticky="ew")

        tk.Label(settings, text="Screenshot bei:", font=("Segoe UI", 9)).grid(
            row=1, column=0, sticky="w", pady=(6, 0))
        tk.Checkbutton(settings, text="Klick", variable=self.capture_on_click,
                       font=("Segoe UI", 9)).grid(row=1, column=1, sticky="w")
        tk.Checkbutton(settings, text="Scrollen", variable=self.capture_on_scroll,
                       font=("Segoe UI", 9)).grid(row=1, column=2, sticky="w")
        tk.Checkbutton(settings, text="Sondertaste", variable=self.capture_on_key,
                       font=("Segoe UI", 9)).grid(row=1, column=3, sticky="w")

        tk.Label(settings, text="Qualität:", font=("Segoe UI", 9)).grid(
            row=2, column=0, sticky="w", pady=(6, 0))
        quality_scale = ttk.Scale(settings, from_=40, to=95,
                                  variable=self.screenshot_quality,
                                  orient=tk.HORIZONTAL, length=140)
        quality_scale.grid(row=2, column=1, columnspan=2, sticky="ew")
        tk.Label(settings, textvariable=tk.StringVar(), font=("Segoe UI", 9)).grid(
            row=2, column=3, sticky="w")
        self.quality_label = tk.Label(settings, font=("Segoe UI", 9))
        self.quality_label.grid(row=2, column=3, sticky="w")
        self.screenshot_quality.trace_add("write",
            lambda *_: self.quality_label.config(
                text=f"{self.screenshot_quality.get()} %"))
        self.quality_label.config(text=f"{self.screenshot_quality.get()} %")

        # Export button
        self.export_btn = tk.Button(self.root, text="💾  Dokumentation exportieren",
                                    command=self.export_document,
                                    state=tk.DISABLED,
                                    bg="#0078d4", fg="white",
                                    font=("Segoe UI", 10, "bold"),
                                    relief=tk.FLAT, padx=12, pady=8,
                                    cursor="hand2")
        self.export_btn.pack(fill=tk.X, padx=12, pady=(4, 12))

        # Hotkey hint
        tk.Label(self.root,
                 text="F8: Aufnahme  |  F9: Screenshot  |  F10: Notiz",
                 bg="#f3f2f1", fg="#605e5c",
                 font=("Segoe UI", 8)).pack(pady=(0, 8))

        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    # ------------------------------------------------------------------
    # Hotkeys
    # ------------------------------------------------------------------

    def _setup_hotkeys(self):
        self._hk_listener = None
        self._pressed_keys = set()

        def on_press(key):
            try:
                k = key.name if hasattr(key, "name") else str(key)
            except AttributeError:
                k = str(key)
            k = k.lower().replace("key.", "")
            if k == "f8":
                self.root.after(0, self.toggle_recording)
            elif k == "f9":
                self.root.after(0, self.manual_screenshot)
            elif k == "f10":
                self.root.after(0, self.add_note)

        self._hk_listener = pkeyboard.Listener(on_press=on_press, daemon=True)
        self._hk_listener.start()

    # ------------------------------------------------------------------
    # Recording control
    # ------------------------------------------------------------------

    def toggle_recording(self):
        if self.recording:
            self._stop_recording()
        else:
            self._start_recording()

    def _start_recording(self):
        self.recording = True
        self.events.clear()
        ev = ActionEvent(
            timestamp=time.time(),
            action_type="start",
            description="Aufnahme gestartet",
        )
        self.events.append(ev)
        self.tracker.start()

        self.rec_btn.config(text="■  Aufnahme stoppen (F8)", bg="#a4262c")
        self.manual_btn.config(state=tk.NORMAL)
        self.note_btn.config(state=tk.NORMAL)
        self.export_btn.config(state=tk.DISABLED)
        self._set_status("🔴 Aufnahme läuft ...", "#a4262c")
        self._update_counter()

    def _stop_recording(self):
        self.recording = False
        self.tracker.stop()

        ev = ActionEvent(
            timestamp=time.time(),
            action_type="stop",
            description="Aufnahme beendet",
        )
        self.events.append(ev)

        self.rec_btn.config(text="▶  Aufnahme starten (F8)", bg="#107c10")
        self.manual_btn.config(state=tk.DISABLED)
        self.note_btn.config(state=tk.DISABLED)
        if len(self.events) > 2:
            self.export_btn.config(state=tk.NORMAL)
        self._set_status(f"Gestoppt – {len(self.events)-2} Ereignisse erfasst.", "#107c10")
        self._update_counter()

    # ------------------------------------------------------------------
    # Event handling
    # ------------------------------------------------------------------

    def _on_tracked_event(self, action_type: str, description: str, x: int, y: int):
        """Called from pynput threads – put into queue for main thread."""
        self.event_queue.put((action_type, description, x, y))

    def _process_queue(self):
        while not self.event_queue.empty():
            try:
                action_type, description, x, y = self.event_queue.get_nowait()
                self._handle_event(action_type, description, x, y)
            except queue.Empty:
                break
        self.root.after(100, self._process_queue)

    def _handle_event(self, action_type: str, description: str, x: int, y: int):
        if not self.recording:
            return

        # Decide whether to take screenshot
        take_ss = (
            (action_type == "click" and self.capture_on_click.get()) or
            (action_type == "scroll" and self.capture_on_scroll.get()) or
            (action_type == "key" and self.capture_on_key.get())
        )

        ss_b64 = None
        if take_ss:
            ss_b64 = self.capture.capture_thumbnail(
                quality=self.screenshot_quality.get(),
                max_width=self.max_width.get(),
            )

        ev = ActionEvent(
            timestamp=time.time(),
            action_type=action_type,
            description=description,
            screenshot_b64=ss_b64,
            x=x, y=y,
        )
        self.events.append(ev)
        self._update_counter()

    # ------------------------------------------------------------------
    # Manual actions
    # ------------------------------------------------------------------

    def manual_screenshot(self):
        if not self.recording:
            return
        ss_b64 = self.capture.capture_thumbnail(
            quality=self.screenshot_quality.get(),
            max_width=self.max_width.get(),
        )
        ev = ActionEvent(
            timestamp=time.time(),
            action_type="screenshot",
            description="Manueller Screenshot",
            screenshot_b64=ss_b64,
        )
        self.events.append(ev)
        self._set_status("📷 Screenshot gespeichert.", "#004578")
        self._update_counter()

    def add_note(self):
        if not self.recording:
            return
        self.root.lift()
        note = simpledialog.askstring(
            "Notiz hinzufügen",
            "Beschreibung / Notiz zum aktuellen Schritt:",
            parent=self.root,
        )
        if note and note.strip():
            ss_b64 = self.capture.capture_thumbnail(
                quality=self.screenshot_quality.get(),
                max_width=self.max_width.get(),
            )
            ev = ActionEvent(
                timestamp=time.time(),
                action_type="note",
                description=note.strip(),
                screenshot_b64=ss_b64,
                note=note.strip(),
            )
            self.events.append(ev)
            self._set_status(f"✎ Notiz gespeichert: {note[:40]}", "#8764b8")
            self._update_counter()

    # ------------------------------------------------------------------
    # Export
    # ------------------------------------------------------------------

    def export_document(self):
        if not self.events:
            messagebox.showwarning("Keine Daten", "Es wurden noch keine Aktionen aufgezeichnet.")
            return

        default_name = self.doc_title.get().replace(" ", "_").replace("/", "-") + ".html"
        path = filedialog.asksaveasfilename(
            title="Dokumentation speichern",
            defaultextension=".html",
            initialfile=default_name,
            filetypes=[("HTML-Datei", "*.html"), ("Alle Dateien", "*.*")],
        )
        if not path:
            return

        try:
            DocumentGenerator.generate(self.events, self.doc_title.get(), path)
            self._set_status(f"✅ Exportiert: {Path(path).name}", "#107c10")
            if messagebox.askyesno("Export erfolgreich",
                                   f"Dokumentation gespeichert:\n{path}\n\nIm Browser öffnen?"):
                import webbrowser
                webbrowser.open(f"file:///{path}")
        except Exception as e:
            messagebox.showerror("Fehler", f"Export fehlgeschlagen:\n{e}")

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------

    def _set_status(self, text: str, color: str = "#323130"):
        self.status_var.set(text)
        self.status_label.config(fg=color)

    def _update_counter(self):
        total = len([e for e in self.events if e.action_type not in ("start", "stop")])
        screenshots = sum(1 for e in self.events if e.screenshot_b64)
        self.counter_var.set(f"Ereignisse: {total}  |  Screenshots: {screenshots}")

    def _on_close(self):
        if self.recording:
            if not messagebox.askyesno("Beenden",
                                       "Aufnahme läuft noch. Trotzdem beenden?\n"
                                       "(Nicht exportierte Daten gehen verloren.)"):
                return
        if self._hk_listener:
            self._hk_listener.stop()
        if self.recording:
            self.tracker.stop()
        self.root.destroy()

    # ------------------------------------------------------------------
    # Run
    # ------------------------------------------------------------------

    def run(self):
        self.root.mainloop()


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    app = RecorderApp()
    app.run()
