import tkinter as tk
from tkinter import messagebox
import queue
import subprocess
import tempfile
import threading
import time
import re
import os
import sys
from typing import Optional, Tuple

try:
    import pyautogui
    import pytesseract
    from PIL import Image, ImageEnhance
except ImportError as e:
    print(f"Required packages missing: {e}")
    print("  pip install pyautogui pytesseract Pillow opencv-python-headless")
    print("  brew install tesseract tesseract-lang")
    sys.exit(1)

CLICK_DELAY = 3.0
MIN_REGION_SIZE = 10
OCR_CONFIG = "--psm 6"
OVERLAY_ALPHA = 0.5
MATCH_CONFIDENCE = 0.8
IMAGE_RETRY_INTERVAL = 0.5
IMAGE_RETRY_MAX = 10


SCREENSHOT_TMP = os.path.join(tempfile.gettempdir(), "autosword_screen.png")


def take_screenshot() -> Image.Image:
    """Take screenshot using macOS screencapture to avoid Quartz thread deadlock."""
    subprocess.run(
        ["screencapture", "-x", SCREENSHOT_TMP],
        capture_output=True, timeout=5,
    )
    img = Image.open(SCREENSHOT_TMP)
    return img.copy()  # copy so file handle is released


def get_retina_scale() -> float:
    full = take_screenshot()
    logical_w = pyautogui.size()[0]
    return full.size[0] / logical_w


REGION_TMP = os.path.join(tempfile.gettempdir(), "autosword_region.png")


def capture_region(x, y, w, h) -> Image.Image:
    """Capture specific screen region using screencapture -R (handles Retina automatically)."""
    subprocess.run(
        ["screencapture", "-x", "-R", f"{x},{y},{w},{h}", REGION_TMP],
        capture_output=True, timeout=5,
    )
    img = Image.open(REGION_TMP)
    return img.copy()


def ask_text_native(prompt: str) -> Optional[str]:
    script = (
        f'set userInput to text returned of '
        f'(display dialog "{prompt}" default answer "" '
        f'with title "autoSword")'
    )
    try:
        result = subprocess.run(
            ["osascript", "-e", script],
            capture_output=True, text=True, timeout=120,
        )
        if result.returncode == 0:
            return result.stdout.strip()
    except (subprocess.TimeoutExpired, Exception):
        pass
    return None


class RegionSelector:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.region: Optional[Tuple[int, int, int, int]] = None
        self.start_x = 0
        self.start_y = 0
        self.rect_id = None
        self.coord_text_id = None
        self.overlay: Optional[tk.Toplevel] = None
        self.canvas: Optional[tk.Canvas] = None
        self._done = False

    def select(self, restore_window=True) -> Optional[Tuple[int, int, int, int]]:
        self.region = None
        self._done = False
        self.root.withdraw()
        self.root.update()
        time.sleep(0.5)
        self._create_overlay()
        while not self._done:
            try:
                self.root.update()
            except tk.TclError:
                break
            time.sleep(0.01)
        if restore_window:
            self.root.deiconify()
            self.root.lift()
        return self.region

    def _create_overlay(self):
        self.overlay = tk.Toplevel(self.root)
        self.overlay.overrideredirect(True)
        self.overlay.attributes("-topmost", True)

        screen_w = self.root.winfo_screenwidth()
        screen_h = self.root.winfo_screenheight()
        self.overlay.geometry(f"{screen_w}x{screen_h}+0+0")

        self.canvas = tk.Canvas(
            self.overlay, cursor="crosshair",
            bg="black", highlightthickness=0,
        )
        self.canvas.pack(fill=tk.BOTH, expand=True)

        self.overlay.update_idletasks()
        self.overlay.attributes("-alpha", OVERLAY_ALPHA)

        self.canvas.create_text(
            screen_w // 2, 50,
            text="Drag to select  |  ESC: Cancel",
            fill="yellow", font=("Helvetica", 28, "bold"),
        )

        self.canvas.bind("<ButtonPress-1>", self._on_press)
        self.canvas.bind("<B1-Motion>", self._on_drag)
        self.canvas.bind("<ButtonRelease-1>", self._on_release)
        self.overlay.bind("<Escape>", self._on_escape)
        self.overlay.focus_force()
        self.overlay.update()

    def _on_press(self, event):
        self.start_x = event.x
        self.start_y = event.y
        if self.rect_id:
            self.canvas.delete(self.rect_id)
        if self.coord_text_id:
            self.canvas.delete(self.coord_text_id)
        self.rect_id = self.canvas.create_rectangle(
            self.start_x, self.start_y, self.start_x, self.start_y,
            outline="#00FF00", width=3, dash=(6, 4),
        )

    def _on_drag(self, event):
        if self.rect_id:
            self.canvas.coords(
                self.rect_id,
                self.start_x, self.start_y, event.x, event.y,
            )
            w = abs(event.x - self.start_x)
            h = abs(event.y - self.start_y)
            if self.coord_text_id:
                self.canvas.delete(self.coord_text_id)
            self.coord_text_id = self.canvas.create_text(
                (self.start_x + event.x) // 2,
                (self.start_y + event.y) // 2,
                text=f"{w} x {h}",
                fill="yellow", font=("Helvetica", 16, "bold"),
            )

    def _on_release(self, event):
        x1 = min(self.start_x, event.x)
        y1 = min(self.start_y, event.y)
        x2 = max(self.start_x, event.x)
        y2 = max(self.start_y, event.y)
        w = x2 - x1
        h = y2 - y1
        if w >= MIN_REGION_SIZE and h >= MIN_REGION_SIZE:
            self.region = (x1, y1, w, h)
        self._close_overlay()

    def _on_escape(self, event):
        self.region = None
        self._close_overlay()

    def _close_overlay(self):
        if self.overlay:
            self.overlay.destroy()
            self.overlay = None
        self.rect_id = None
        self.coord_text_id = None
        self._done = True


class MacroController:
    def __init__(self, scale: float):
        self.scale = scale
        self.result_region: Optional[Tuple[int, int, int, int]] = None
        self.click_image: Optional[Image.Image] = None
        self.target_text = ""
        self.click_delay = CLICK_DELAY
        self.running = False
        self.attempt_count = 0
        self.last_ocr_text = ""
        self.on_status_update = None
        self.on_match_found = None
        self.on_attempt_update = None
        self.on_ocr_update = None
        self.on_stopped = None

    def _capture_region(self, region: Tuple[int, int, int, int]) -> Image.Image:
        x, y, w, h = region
        return capture_region(x, y, w, h)

    def _find_and_click(self) -> bool:
        rx, ry, rw, rh = self.result_region
        region_img = capture_region(rx, ry, rw, rh)
        needle = self.click_image
        try:
            location = pyautogui.locate(needle, region_img, confidence=MATCH_CONFIDENCE)
        except pyautogui.ImageNotFoundException:
            return False

        if location is None:
            return False

        center = pyautogui.center(location)
        # center is in physical pixels within the region image
        # convert to logical screen coords: region offset + (pixel offset / scale)
        click_x = rx + center.x / self.scale
        click_y = ry + center.y / self.scale
        pyautogui.click(click_x, click_y)
        return True

    def _ocr_image(self, image: Image.Image) -> str:
        gray = image.convert("L")
        enhancer = ImageEnhance.Contrast(gray)
        gray = enhancer.enhance(2.0)

        w, h = gray.size
        if w < 300 or h < 50:
            gray = gray.resize((w * 2, h * 2), Image.LANCZOS)

        gray = gray.point(lambda x: 0 if x < 128 else 255, "1")

        text = pytesseract.image_to_string(
            gray, lang="kor+eng", config=OCR_CONFIG
        )
        return text.strip()

    def _check_match(self, ocr_text: str) -> bool:
        normalized_ocr = re.sub(r"\s+", " ", ocr_text).strip().lower()
        normalized_target = re.sub(r"\s+", " ", self.target_text).strip().lower()
        return normalized_target in normalized_ocr

    def _notify(self, callback, *args):
        if callback:
            callback(*args)

    def _interruptible_sleep(self, seconds: float) -> bool:
        for _ in range(int(seconds * 10)):
            if not self.running:
                return False
            time.sleep(0.1)
        return True

    def run(self):
        self.running = True
        self.attempt_count = 0

        while self.running:
            self.attempt_count += 1
            self._notify(self.on_attempt_update, self.attempt_count)

            # OCR check first (before clicking)
            self._notify(self.on_status_update, f"#{self.attempt_count} OCR...")
            try:
                image = self._capture_region(self.result_region)
            except pyautogui.FailSafeException:
                self._notify(self.on_status_update, "EMERGENCY STOP")
                self.running = False
                self._notify(self.on_stopped)
                return

            ocr_text = self._ocr_image(image)
            self.last_ocr_text = ocr_text
            self._notify(self.on_ocr_update, ocr_text)

            if self._check_match(ocr_text):
                self._notify(self.on_status_update, f"MATCH! (#{self.attempt_count})")
                self.running = False
                self._notify(self.on_match_found)
                return

            # No match → find image and click
            found = False
            for retry in range(IMAGE_RETRY_MAX):
                if not self.running:
                    self._notify(self.on_status_update, "stopped")
                    return
                self._notify(self.on_status_update,
                             f"#{self.attempt_count} searching... ({retry + 1}/{IMAGE_RETRY_MAX})")
                try:
                    found = self._find_and_click()
                except pyautogui.FailSafeException:
                    self._notify(self.on_status_update, "EMERGENCY STOP")
                    self.running = False
                    self._notify(self.on_stopped)
                    return
                if found:
                    break
                if not self._interruptible_sleep(IMAGE_RETRY_INTERVAL):
                    self._notify(self.on_status_update, "stopped")
                    return

            if not found:
                self._notify(self.on_status_update,
                             f"#{self.attempt_count} image not found after {IMAGE_RETRY_MAX} retries. stopping.")
                self.running = False
                self._notify(self.on_stopped)
                return

            self._notify(self.on_status_update, f"#{self.attempt_count} clicked. waiting...")

            if not self._interruptible_sleep(self.click_delay):
                self._notify(self.on_status_update, "stopped")
                return

        self._notify(self.on_status_update, "stopped")

    def stop(self):
        self.running = False


class MacroApp:
    def __init__(self, scale: float):
        self.scale = scale
        self.root = tk.Tk()
        self.root.title("autoSword - OCR Macro")
        self.root.geometry("450x420")
        self.root.resizable(False, False)

        self.controller = MacroController(scale)
        self.selector = RegionSelector(self.root)
        self.macro_thread: Optional[threading.Thread] = None

        # Label references (updated via .config(text=...) instead of StringVar)
        self.lbl_result_region = None
        self.lbl_click_image = None
        self.lbl_target_text = None
        self.lbl_delay = None
        self.lbl_status = None
        self.lbl_attempts = None
        self.lbl_last_ocr = None

        self.msg_queue = queue.Queue()

        self._build_gui()
        self._setup_callbacks()
        self._poll_queue()
        self.root.bind("<Escape>", lambda e: self._on_stop())

    def _set_label(self, widget, text):
        """Update button/label text and force redraw."""
        if widget:
            widget.config(text=text)
            widget.update_idletasks()

    def _make_info_btn(self, parent, text="--"):
        """Create a disabled flat button used as a display label (Tk 8.5 workaround)."""
        btn = tk.Button(
            parent, text=text, relief=tk.FLAT, bd=0,
            state=tk.DISABLED, disabledforeground="blue",
            font=("Helvetica", 12), anchor=tk.W,
        )
        return btn

    def _build_gui(self):
        pad = {"padx": 10, "pady": 4}

        # Step 1 - Result region
        tk.Button(
            self.root, text="[ Step 1 ] Result check region",
            font=("Helvetica", 13, "bold"), anchor=tk.W,
            relief=tk.FLAT, bd=0, state=tk.DISABLED,
            disabledforeground="black",
        ).pack(fill=tk.X, padx=10, pady=(10, 0))

        row1 = tk.Frame(self.root)
        row1.pack(fill=tk.X, **pad)
        self.btn_result_region = tk.Button(
            row1, text="Set Region",
            command=self._on_set_result_region, width=16,
        )
        self.btn_result_region.pack(side=tk.LEFT, padx=(5, 10))
        self.lbl_result_region = self._make_info_btn(row1, "--")
        self.lbl_result_region.pack(side=tk.LEFT)

        # Step 2 - Click target image (capture)
        tk.Button(
            self.root, text="[ Step 2 ] Click target image",
            font=("Helvetica", 13, "bold"), anchor=tk.W,
            relief=tk.FLAT, bd=0, state=tk.DISABLED,
            disabledforeground="black",
        ).pack(fill=tk.X, padx=10, pady=(8, 0))

        row2 = tk.Frame(self.root)
        row2.pack(fill=tk.X, **pad)
        self.btn_click_image = tk.Button(
            row2, text="Capture Image",
            command=self._on_set_click_image, width=16,
        )
        self.btn_click_image.pack(side=tk.LEFT, padx=(5, 5))
        self.btn_preview = tk.Button(
            row2, text="Preview",
            command=self._preview_click_image, width=8,
            state=tk.DISABLED,
        )
        self.btn_preview.pack(side=tk.LEFT, padx=(0, 10))
        self.lbl_click_image = self._make_info_btn(row2, "--")
        self.lbl_click_image.pack(side=tk.LEFT)

        # Step 3 - Match text
        tk.Button(
            self.root, text="[ Step 3 ] Match text",
            font=("Helvetica", 13, "bold"), anchor=tk.W,
            relief=tk.FLAT, bd=0, state=tk.DISABLED,
            disabledforeground="black",
        ).pack(fill=tk.X, padx=10, pady=(8, 0))

        row3 = tk.Frame(self.root)
        row3.pack(fill=tk.X, **pad)
        self.btn_set_text = tk.Button(
            row3, text="Input Text",
            command=self._on_set_target_text, width=16,
        )
        self.btn_set_text.pack(side=tk.LEFT, padx=(5, 10))
        self.lbl_target_text = self._make_info_btn(row3, "--")
        self.lbl_target_text.pack(side=tk.LEFT)

        # Step 4 - Click delay
        tk.Button(
            self.root, text="[ Step 4 ] Click delay",
            font=("Helvetica", 13, "bold"), anchor=tk.W,
            relief=tk.FLAT, bd=0, state=tk.DISABLED,
            disabledforeground="black",
        ).pack(fill=tk.X, padx=10, pady=(8, 0))

        row4 = tk.Frame(self.root)
        row4.pack(fill=tk.X, **pad)
        self.btn_delay_down = tk.Button(
            row4, text="-", command=self._on_delay_down,
            width=3, font=("Helvetica", 12, "bold"),
        )
        self.btn_delay_down.pack(side=tk.LEFT, padx=(5, 2))
        self.lbl_delay = tk.Button(
            row4, text=f"{CLICK_DELAY:.1f}s",
            relief=tk.FLAT, bd=0, state=tk.DISABLED,
            disabledforeground="blue",
            font=("Helvetica", 14, "bold"), width=6,
        )
        self.lbl_delay.pack(side=tk.LEFT, padx=2)
        self.btn_delay_up = tk.Button(
            row4, text="+", command=self._on_delay_up,
            width=3, font=("Helvetica", 12, "bold"),
        )
        self.btn_delay_up.pack(side=tk.LEFT, padx=2)
        tk.Button(
            row4, text="(0.5 ~ 10.0)",
            relief=tk.FLAT, bd=0, state=tk.DISABLED,
            disabledforeground="gray50",
            font=("Helvetica", 11),
        ).pack(side=tk.LEFT, padx=(10, 0))

        # Buttons
        tk.Frame(self.root, height=1, bg="gray70").pack(fill=tk.X, padx=10, pady=8)

        frame_btn = tk.Frame(self.root)
        frame_btn.pack(fill=tk.X, **pad)
        self.btn_start = tk.Button(
            frame_btn, text="START", command=self._on_start,
            width=12, font=("Helvetica", 12, "bold"),
        )
        self.btn_start.pack(side=tk.LEFT, padx=5)
        self.btn_stop = tk.Button(
            frame_btn, text="STOP (ESC)", command=self._on_stop,
            width=12, font=("Helvetica", 12, "bold"),
            state=tk.DISABLED,
        )
        self.btn_stop.pack(side=tk.LEFT, padx=5)

        # Status
        tk.Frame(self.root, height=1, bg="gray70").pack(fill=tk.X, padx=10, pady=8)

        status_frame = tk.Frame(self.root)
        status_frame.pack(fill=tk.X, padx=10)

        for heading, attr_name, init_text in [
            ("Status:", "lbl_status", "Ready"),
            ("Attempts:", "lbl_attempts", "0"),
            ("Last OCR:", "lbl_last_ocr", ""),
        ]:
            row = tk.Frame(status_frame)
            row.pack(fill=tk.X, pady=1)
            tk.Button(
                row, text=heading, width=10, anchor=tk.W,
                relief=tk.FLAT, bd=0, state=tk.DISABLED,
                disabledforeground="black",
            ).pack(side=tk.LEFT)
            btn = tk.Button(
                row, text=init_text, anchor=tk.W,
                relief=tk.FLAT, bd=0, state=tk.DISABLED,
                disabledforeground="darkred" if attr_name == "lbl_status" else "gray30",
                font=("Helvetica", 11),
            )
            if attr_name == "lbl_last_ocr":
                btn.config(wraplength=300, justify=tk.LEFT)
            btn.pack(side=tk.LEFT, padx=5)
            setattr(self, attr_name, btn)

    def _poll_queue(self):
        """Main thread polls the queue every 50ms - no root.after() from worker thread."""
        try:
            while True:
                func, args = self.msg_queue.get_nowait()
                func(*args)
        except queue.Empty:
            pass
        self.root.after(50, self._poll_queue)

    def _enqueue(self, func, *args):
        """Worker thread puts GUI updates in the queue (thread-safe)."""
        self.msg_queue.put((func, args))

    def _setup_callbacks(self):
        self.controller.on_status_update = lambda msg: self._enqueue(
            self._set_label, self.lbl_status, msg
        )
        self.controller.on_attempt_update = lambda n: self._enqueue(
            self._set_label, self.lbl_attempts, str(n)
        )
        self.controller.on_ocr_update = lambda txt: self._enqueue(
            self._set_label, self.lbl_last_ocr, txt[:200] if txt else ""
        )
        self.controller.on_match_found = lambda: self._enqueue(
            self._on_match_found
        )
        self.controller.on_stopped = lambda: self._enqueue(
            self._reset_buttons
        )

    def _on_set_result_region(self):
        region = self.selector.select(restore_window=False)
        if region:
            self.controller.result_region = region
            x, y, w, h = region
            self._set_label(self.lbl_result_region, f"[OK] ({x},{y}) {w}x{h}")
            # preview the captured region
            img = capture_region(x, y, w, h)
            preview_path = os.path.join(tempfile.gettempdir(), "autosword_region_preview.png")
            img.save(preview_path)
            subprocess.Popen(["open", preview_path])
        else:
            self._set_label(self.lbl_result_region, "-- cancelled --")
        self.root.deiconify()
        self.root.lift()

    def _on_set_click_image(self):
        region = self.selector.select(restore_window=False)
        if region:
            x, y, w, h = region
            img = capture_region(x, y, w, h)
            self.controller.click_image = img
            self._set_label(self.lbl_click_image, f"[OK] {w}x{h}")
            self.btn_preview.config(state=tk.NORMAL)
            self._preview_click_image()
        else:
            self._set_label(self.lbl_click_image, "-- cancelled --")
        self.root.deiconify()
        self.root.lift()

    def _preview_click_image(self):
        img = self.controller.click_image
        if img is None:
            return
        preview_path = os.path.join(tempfile.gettempdir(), "autosword_preview.png")
        img.save(preview_path)
        subprocess.Popen(["open", preview_path])

    def _on_set_target_text(self):
        text = ask_text_native("Enter text to match:")
        if text and text.strip():
            self.controller.target_text = text.strip()
            display = text.strip()
            if len(display) > 20:
                display = display[:20] + "..."
            self._set_label(self.lbl_target_text, f'[OK] "{display}"')

    def _on_delay_up(self):
        new_val = min(self.controller.click_delay + 0.5, 10.0)
        self.controller.click_delay = new_val
        self._set_label(self.lbl_delay, f"{new_val:.1f}s")

    def _on_delay_down(self):
        new_val = max(self.controller.click_delay - 0.5, 0.5)
        self.controller.click_delay = new_val
        self._set_label(self.lbl_delay, f"{new_val:.1f}s")

    def _on_start(self):
        if not self.controller.result_region:
            messagebox.showwarning("Warning", "Set the result region first.")
            return
        if self.controller.click_image is None:
            messagebox.showwarning("Warning", "Capture the click target image first.")
            return
        if not self.controller.target_text:
            messagebox.showwarning("Warning", "Input the match text first.")
            return

        self.controller.running = True
        self._set_label(self.lbl_attempts, "0")
        self._set_label(self.lbl_status, "Running...")
        self._set_label(self.lbl_last_ocr, "")

        self.btn_start.config(state=tk.DISABLED)
        self.btn_stop.config(state=tk.NORMAL)
        self.btn_result_region.config(state=tk.DISABLED)
        self.btn_click_image.config(state=tk.DISABLED)
        self.btn_preview.config(state=tk.DISABLED)
        self.btn_set_text.config(state=tk.DISABLED)
        self.btn_delay_up.config(state=tk.DISABLED)
        self.btn_delay_down.config(state=tk.DISABLED)

        self.macro_thread = threading.Thread(
            target=self.controller.run, daemon=True
        )
        self.macro_thread.start()

    def _on_stop(self):
        self.controller.stop()
        self._reset_buttons()

    def _on_match_found(self):
        self._reset_buttons()
        messagebox.showinfo(
            "Match Found",
            f"Found '{self.controller.target_text}'!\n"
            f"Total {self.controller.attempt_count} attempts",
        )

    def _reset_buttons(self):
        self.btn_start.config(state=tk.NORMAL)
        self.btn_stop.config(state=tk.DISABLED)
        self.btn_result_region.config(state=tk.NORMAL)
        self.btn_click_image.config(state=tk.NORMAL)
        if self.controller.click_image is not None:
            self.btn_preview.config(state=tk.NORMAL)
        self.btn_set_text.config(state=tk.NORMAL)
        self.btn_delay_up.config(state=tk.NORMAL)
        self.btn_delay_down.config(state=tk.NORMAL)

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    try:
        pytesseract.get_tesseract_version()
    except pytesseract.TesseractNotFoundError:
        print("tesseract not installed.")
        print("Run: brew install tesseract tesseract-lang")
        sys.exit(1)

    pyautogui.FAILSAFE = True
    pyautogui.PAUSE = 0.1

    print("Detecting Retina scale...")
    scale = get_retina_scale()
    print(f"Scale factor: {scale}x")

    app = MacroApp(scale)
    app.run()
