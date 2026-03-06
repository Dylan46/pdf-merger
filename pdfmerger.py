"""
PDF Merger — scan a folder, preview & reorder your PDFs, then smash them into one.

Requirements (install with: pip install pymupdf pillow):
  - pymupdf  : does the heavy lifting — renders thumbnails AND merges PDFs
  - pillow   : bridges the gap between fitz pixel data and what tkinter can display
  - tkinter  : ships with Python, no install needed
"""

import threading
from dataclasses import dataclass, field
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import fitz  # PyMuPDF
from PIL import Image, ImageTk


# ── Tweak these if you want bigger/smaller thumbnails or a crisper preview ──
THUMBNAIL_SIZE = (100, 140)   # (width, height) of the little thumbnail in the list
THUMBNAIL_ZOOM = 0.5          # zoom level for rendering thumbnails — low is fine, it's tiny
PREVIEW_ZOOM = 1.5            # zoom for the big preview on the right — higher = crisper
LIST_ROW_HEIGHT = 130         # vertical space each PDF row gets in the list
LIST_PANEL_WIDTH = 260        # starting width of the left panel
APP_TITLE = "PDF Merger"


# ── This little guy holds everything we need to know about one PDF ──
@dataclass
class PDFEntry:
    path: Path
    name: str
    page_count: int = 0
    thumbnail: object = field(default=None, repr=False)  # ImageTk.PhotoImage — set async
    broken: bool = False  # True if the file refused to open


def render_thumbnail(path: Path, size: tuple = THUMBNAIL_SIZE) -> ImageTk.PhotoImage | None:
    """
    Opens a PDF, grabs the first page, and turns it into a thumbnail.

    Okay so here's the deal — fitz gives us raw pixel bytes, then we hand those
    to Pillow to resize nicely, and THEN we convert to an ImageTk.PhotoImage for
    tkinter. It's a bit of a relay race but each step earns its keep.
    """
    try:
        doc = fitz.open(path)
        page = doc[0]
        # We calculate a zoom so the output fits roughly within our target size
        zoom = min(size[0] / page.rect.width, size[1] / page.rect.height) * 2
        matrix = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=matrix, alpha=False)
        doc.close()

        # fitz pixmap → PIL Image → resize to thumbnail bounds
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        img.thumbnail(size, Image.LANCZOS)

        return ImageTk.PhotoImage(img)
    except Exception:
        # Corrupt file, password-protected, whatever — we'll show a placeholder
        return None


class PDFMergerApp(tk.Tk):
    """The whole application lives here. One window, one class, no fuss."""

    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1000x680")
        self.minsize(700, 500)

        # The ordered list of PDFs — this IS the merge order
        self.pdf_entries: list[PDFEntry] = []

        # Track which item is selected in the list
        self.selected_index: int = -1

        # For drag-and-drop reordering in the list panel
        self.drag_start_index: int = -1
        self.drag_ghost_id = None

        # For the preview panel — which page we're currently looking at
        self.preview_page_num: int = 0
        self.preview_image = None  # keep a reference so it doesn't get garbage collected

        self._build_ui()

    # ─────────────────────────────────────────────────────────────
    #  UI CONSTRUCTION
    # ─────────────────────────────────────────────────────────────

    def _build_ui(self):
        """Builds the whole window. Toolbar on top, two panels below."""
        self._build_toolbar()

        # PanedWindow lets the user drag the divider between list and preview
        paned = ttk.PanedWindow(self, orient=tk.HORIZONTAL)
        paned.pack(fill=tk.BOTH, expand=True, padx=6, pady=(0, 6))

        left_frame = ttk.Frame(paned, width=LIST_PANEL_WIDTH)
        left_frame.pack_propagate(False)
        paned.add(left_frame, weight=1)

        right_frame = ttk.Frame(paned)
        paned.add(right_frame, weight=3)

        self._build_list_panel(left_frame)
        self._build_preview_panel(right_frame)

    def _build_toolbar(self):
        """Top bar with action buttons. Simple and clean."""
        bar = ttk.Frame(self, padding=(6, 6, 6, 4))
        bar.pack(fill=tk.X)

        ttk.Button(bar, text="Scan Folder", command=self.scan_folder).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(bar, text="Add Files", command=self.add_files).pack(side=tk.LEFT, padx=(0, 4))

        # Separator to visually separate the dangerous/important merge button
        ttk.Separator(bar, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=8)

        self.merge_btn = ttk.Button(bar, text="Merge PDFs →", command=self.merge_pdfs, state=tk.DISABLED)
        self.merge_btn.pack(side=tk.LEFT)

        # File count label on the right — nice little status indicator
        self.status_label = ttk.Label(bar, text="No files loaded", foreground="gray")
        self.status_label.pack(side=tk.RIGHT)

    def _build_list_panel(self, parent):
        """
        Left side: a scrollable canvas that shows all the PDFs with their thumbnails.
        We use a Canvas instead of a Listbox because Listbox can't embed images —
        it's strictly text only, which would be a shame.
        """
        label = ttk.Label(parent, text="PDF Files (drag to reorder)", font=("", 9, "bold"))
        label.pack(anchor=tk.W, padx=4, pady=(4, 2))

        # Canvas + scrollbar wrapped together
        canvas_frame = ttk.Frame(parent)
        canvas_frame.pack(fill=tk.BOTH, expand=True, padx=4)

        self.list_canvas = tk.Canvas(canvas_frame, bg="#f5f5f5", highlightthickness=0)
        scrollbar = ttk.Scrollbar(canvas_frame, orient=tk.VERTICAL, command=self.list_canvas.yview)
        self.list_canvas.configure(yscrollcommand=scrollbar.set)

        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.list_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Mouse bindings for click-to-select and drag-to-reorder
        self.list_canvas.bind("<Button-1>", self._on_list_click)
        self.list_canvas.bind("<B1-Motion>", self._on_drag_motion)
        self.list_canvas.bind("<ButtonRelease-1>", self._on_drag_drop)
        self.list_canvas.bind("<MouseWheel>", lambda e: self.list_canvas.yview_scroll(-1 * (e.delta // 120), "units"))

        # Reorder buttons below the list
        btn_frame = ttk.Frame(parent)
        btn_frame.pack(fill=tk.X, padx=4, pady=(4, 4))

        self.up_btn = ttk.Button(btn_frame, text="▲ Up", command=self.move_up, state=tk.DISABLED, width=8)
        self.up_btn.pack(side=tk.LEFT, padx=(0, 2))

        self.down_btn = ttk.Button(btn_frame, text="▼ Down", command=self.move_down, state=tk.DISABLED, width=8)
        self.down_btn.pack(side=tk.LEFT, padx=(0, 2))

        self.remove_btn = ttk.Button(btn_frame, text="✕ Remove", command=self.remove_selected, state=tk.DISABLED)
        self.remove_btn.pack(side=tk.LEFT)

    def _build_preview_panel(self, parent):
        """
        Right side: shows a rendered page from the selected PDF.
        Scroll through pages with the Prev/Next buttons at the bottom.
        """
        label = ttk.Label(parent, text="Preview", font=("", 9, "bold"))
        label.pack(anchor=tk.W, padx=4, pady=(4, 2))

        # Scrollable canvas for the page image — pages can be tall
        canvas_frame = ttk.Frame(parent, relief=tk.SUNKEN, borderwidth=1)
        canvas_frame.pack(fill=tk.BOTH, expand=True, padx=4)

        self.preview_canvas = tk.Canvas(canvas_frame, bg="#404040", highlightthickness=0)
        v_scroll = ttk.Scrollbar(canvas_frame, orient=tk.VERTICAL, command=self.preview_canvas.yview)
        h_scroll = ttk.Scrollbar(canvas_frame, orient=tk.HORIZONTAL, command=self.preview_canvas.xview)
        self.preview_canvas.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)

        v_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        h_scroll.pack(side=tk.BOTTOM, fill=tk.X)
        self.preview_canvas.pack(fill=tk.BOTH, expand=True)
        self.preview_canvas.bind("<MouseWheel>", lambda e: self.preview_canvas.yview_scroll(-1 * (e.delta // 120), "units"))

        # Initial placeholder text so it doesn't look broken on startup
        self.preview_canvas.create_text(
            300, 200, text="Select a PDF to preview it here",
            fill="#aaaaaa", font=("", 12), tags="placeholder"
        )

        # Page navigation at the bottom
        nav_frame = ttk.Frame(parent)
        nav_frame.pack(fill=tk.X, padx=4, pady=(4, 4))

        self.prev_btn = ttk.Button(nav_frame, text="◀ Prev", command=self.prev_page, state=tk.DISABLED, width=8)
        self.prev_btn.pack(side=tk.LEFT)

        self.page_label = ttk.Label(nav_frame, text="", anchor=tk.CENTER)
        self.page_label.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.next_btn = ttk.Button(nav_frame, text="Next ▶", command=self.next_page, state=tk.DISABLED, width=8)
        self.next_btn.pack(side=tk.RIGHT)

    # ─────────────────────────────────────────────────────────────
    #  FILE LOADING
    # ─────────────────────────────────────────────────────────────

    def scan_folder(self):
        """Ask the user to pick a folder, then load every PDF in it."""
        folder = filedialog.askdirectory(title="Pick a folder with PDFs")
        if not folder:
            return  # user cancelled, no big deal

        folder_path = Path(folder)
        # Grab all PDFs — case-insensitive because Windows will throw curveballs
        pdf_paths = sorted([p for p in folder_path.iterdir() if p.suffix.lower() == ".pdf"])

        if not pdf_paths:
            messagebox.showinfo("No PDFs Found", f"Didn't find any PDF files in:\n{folder}")
            return

        self.title(f"{APP_TITLE} — {folder_path.name}")
        self.load_pdf_entries(pdf_paths)

    def add_files(self):
        """Let the user cherry-pick individual PDFs to add to the list."""
        paths = filedialog.askopenfilenames(
            title="Pick PDF files to add",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if not paths:
            return

        new_paths = [Path(p) for p in paths]
        existing = {e.path for e in self.pdf_entries}
        # Only add files that aren't already in the list
        to_add = [p for p in new_paths if p not in existing]

        if not to_add:
            messagebox.showinfo("Already Added", "All selected files are already in the list.")
            return

        self._append_pdf_entries(to_add)

    def load_pdf_entries(self, paths: list[Path]):
        """Replace the current list with these PDFs. Thumbnails load in the background."""
        self.pdf_entries.clear()
        self.selected_index = -1
        self.preview_page_num = 0
        self._clear_preview()
        self._append_pdf_entries(paths)

    def _append_pdf_entries(self, paths: list[Path]):
        """
        Adds PDFs to the list, then fires off background threads to render thumbnails.
        This way the UI stays responsive even if you've got 50 PDFs to load.
        """
        start_index = len(self.pdf_entries)

        for path in paths:
            try:
                doc = fitz.open(path)
                page_count = doc.page_count
                doc.close()
            except Exception:
                # File is probably corrupted or password-protected
                page_count = 0
                entry = PDFEntry(path=path, name=path.name, page_count=0, broken=True)
                self.pdf_entries.append(entry)
                continue

            entry = PDFEntry(path=path, name=path.name, page_count=page_count)
            self.pdf_entries.append(entry)

        self.refresh_list()
        self._update_status()

        # Now kick off thumbnail threads for the newly added entries
        for i in range(start_index, len(self.pdf_entries)):
            entry = self.pdf_entries[i]
            if not entry.broken:
                t = threading.Thread(target=self._load_thumbnail_async, args=(entry, i), daemon=True)
                t.start()

    def _load_thumbnail_async(self, entry: PDFEntry, _index: int):
        """
        Runs in a background thread. Renders the thumbnail and then hands it back
        to the main thread via self.after() — you CANNOT touch tkinter widgets
        directly from a background thread, that's a one-way ticket to crash city.
        """
        img = render_thumbnail(entry.path)
        if img is not None:
            # Schedule the UI update on the main thread
            self.after(0, lambda e=entry, i=img: self._apply_thumbnail(e, i))

    def _apply_thumbnail(self, entry: PDFEntry, image: ImageTk.PhotoImage):
        """Called on the main thread once a thumbnail is ready. Just update and redraw."""
        entry.thumbnail = image
        self.refresh_list()

    # ─────────────────────────────────────────────────────────────
    #  LIST PANEL
    # ─────────────────────────────────────────────────────────────

    def refresh_list(self):
        """
        Redraws the entire list canvas from scratch. Not the most efficient approach
        for huge lists, but for a PDF merger you're probably not dealing with thousands
        of files, so it's totally fine and way simpler to reason about.
        """
        self.list_canvas.delete("all")

        total_height = max(len(self.pdf_entries) * LIST_ROW_HEIGHT, 1)
        canvas_width = self.list_canvas.winfo_width() or LIST_PANEL_WIDTH

        self.list_canvas.configure(scrollregion=(0, 0, canvas_width, total_height))

        for i, entry in enumerate(self.pdf_entries):
            y_top = i * LIST_ROW_HEIGHT
            y_bot = y_top + LIST_ROW_HEIGHT - 2

            is_selected = (i == self.selected_index)
            bg_color = "#d0e8ff" if is_selected else ("#ffffff" if i % 2 == 0 else "#f0f0f0")

            # Row background rectangle — also serves as the click target
            self.list_canvas.create_rectangle(
                0, y_top, canvas_width, y_bot,
                fill=bg_color, outline="#cccccc", tags=f"row_{i}"
            )

            # Thumbnail on the left side of the row
            thumb_x = 8 + THUMBNAIL_SIZE[0] // 2
            thumb_y = y_top + LIST_ROW_HEIGHT // 2 - 5

            if entry.thumbnail:
                self.list_canvas.create_image(thumb_x, thumb_y, image=entry.thumbnail, tags=f"row_{i}")
            elif entry.broken:
                # Show a little broken indicator instead of a blank space
                self.list_canvas.create_rectangle(
                    8, y_top + 10, 8 + THUMBNAIL_SIZE[0], y_bot - 10,
                    fill="#ffe0e0", outline="#ffaaaa", tags=f"row_{i}"
                )
                self.list_canvas.create_text(
                    thumb_x, thumb_y, text="⚠\nCan't\nread", fill="#cc4444",
                    font=("", 8), justify=tk.CENTER, tags=f"row_{i}"
                )
            else:
                # Thumbnail is still loading — show a placeholder
                self.list_canvas.create_rectangle(
                    8, y_top + 10, 8 + THUMBNAIL_SIZE[0], y_bot - 10,
                    fill="#e8e8e8", outline="#cccccc", tags=f"row_{i}"
                )
                self.list_canvas.create_text(
                    thumb_x, thumb_y, text="Loading…", fill="#999999",
                    font=("", 8), tags=f"row_{i}"
                )

            # File name and page count to the right of the thumbnail
            text_x = 8 + THUMBNAIL_SIZE[0] + 10
            name_display = entry.name if len(entry.name) <= 22 else entry.name[:19] + "…"
            self.list_canvas.create_text(
                text_x, y_top + 20, text=name_display,
                anchor=tk.W, font=("", 9, "bold"), fill="#222222", tags=f"row_{i}"
            )

            page_text = f"{entry.page_count} page{'s' if entry.page_count != 1 else ''}" if not entry.broken else "unreadable"
            self.list_canvas.create_text(
                text_x, y_top + 38, text=page_text,
                anchor=tk.W, font=("", 8), fill="#666666", tags=f"row_{i}"
            )

            # Row number badge — handy for knowing the merge order at a glance
            self.list_canvas.create_oval(
                canvas_width - 26, y_top + 8, canvas_width - 6, y_top + 28,
                fill="#5a7fd4" if is_selected else "#aaaaaa", outline="", tags=f"row_{i}"
            )
            self.list_canvas.create_text(
                canvas_width - 16, y_top + 18, text=str(i + 1),
                fill="white", font=("", 8, "bold"), tags=f"row_{i}"
            )

        self._update_button_states()

    def _update_status(self):
        n = len(self.pdf_entries)
        if n == 0:
            self.status_label.config(text="No files loaded")
        else:
            self.status_label.config(text=f"{n} PDF{'s' if n != 1 else ''} loaded")

    def _update_button_states(self):
        """Enable/disable buttons based on what's currently selected."""
        has_files = len(self.pdf_entries) > 0
        has_selection = 0 <= self.selected_index < len(self.pdf_entries)

        self.merge_btn.config(state=tk.NORMAL if has_files else tk.DISABLED)
        self.up_btn.config(state=tk.NORMAL if has_selection and self.selected_index > 0 else tk.DISABLED)
        self.down_btn.config(state=tk.NORMAL if has_selection and self.selected_index < len(self.pdf_entries) - 1 else tk.DISABLED)
        self.remove_btn.config(state=tk.NORMAL if has_selection else tk.DISABLED)

    def _index_at_y(self, y: int) -> int:
        """Figure out which list row a y-coordinate falls in. Returns -1 if out of bounds."""
        # Account for canvas scroll position
        canvas_y = self.list_canvas.canvasy(y)
        index = int(canvas_y // LIST_ROW_HEIGHT)
        if 0 <= index < len(self.pdf_entries):
            return index
        return -1

    def _on_list_click(self, event):
        """Handle a click on the list — select the item and show its preview."""
        self.drag_start_index = self._index_at_y(event.y)
        index = self._index_at_y(event.y)
        if index == -1:
            return

        self.selected_index = index
        self.preview_page_num = 0
        self.refresh_list()
        self.show_preview(self.pdf_entries[index])

    def move_up(self):
        """Swap the selected item with the one above it."""
        i = self.selected_index
        if i <= 0:
            return
        self.pdf_entries[i], self.pdf_entries[i - 1] = self.pdf_entries[i - 1], self.pdf_entries[i]
        self.selected_index = i - 1
        self.refresh_list()

    def move_down(self):
        """Swap the selected item with the one below it."""
        i = self.selected_index
        if i < 0 or i >= len(self.pdf_entries) - 1:
            return
        self.pdf_entries[i], self.pdf_entries[i + 1] = self.pdf_entries[i + 1], self.pdf_entries[i]
        self.selected_index = i + 1
        self.refresh_list()

    def remove_selected(self):
        """Yeet the selected PDF out of the list."""
        i = self.selected_index
        if i < 0 or i >= len(self.pdf_entries):
            return

        self.pdf_entries.pop(i)
        # Keep the selection on the same position, or clamp if we removed the last item
        self.selected_index = min(i, len(self.pdf_entries) - 1)

        if self.selected_index >= 0:
            self.preview_page_num = 0
            self.show_preview(self.pdf_entries[self.selected_index])
        else:
            self._clear_preview()

        self.refresh_list()
        self._update_status()

    # ─────────────────────────────────────────────────────────────
    #  DRAG AND DROP (list reordering)
    # ─────────────────────────────────────────────────────────────

    def _on_drag_motion(self, event):
        """
        While the user is dragging, draw a little indicator line showing where
        the item will land when they release. Much better than nothing.
        """
        if self.drag_start_index < 0:
            return

        # Clear the previous ghost line
        if self.drag_ghost_id:
            self.list_canvas.delete(self.drag_ghost_id)

        target = self._index_at_y(event.y)
        if target < 0:
            target = len(self.pdf_entries) - 1

        canvas_width = self.list_canvas.winfo_width()
        # Draw the drop-target line between rows
        line_y = self.list_canvas.canvasy(event.y)
        # Snap to the nearest row boundary for a clean look
        snapped_y = round(line_y / LIST_ROW_HEIGHT) * LIST_ROW_HEIGHT

        self.drag_ghost_id = self.list_canvas.create_rectangle(
            2, snapped_y - 2, canvas_width - 2, snapped_y + 2,
            fill="#3a7bd5", outline="", tags="ghost"
        )

    def _on_drag_drop(self, event):
        """When the user lets go, move the item to its new position."""
        if self.drag_ghost_id:
            self.list_canvas.delete(self.drag_ghost_id)
            self.drag_ghost_id = None

        if self.drag_start_index < 0:
            return

        canvas_y = self.list_canvas.canvasy(event.y)
        target_index = max(0, min(int(canvas_y // LIST_ROW_HEIGHT), len(self.pdf_entries) - 1))

        if target_index != self.drag_start_index:
            # Pluck the item out and reinsert it at the target position
            entry = self.pdf_entries.pop(self.drag_start_index)
            self.pdf_entries.insert(target_index, entry)
            self.selected_index = target_index
            self.refresh_list()

        self.drag_start_index = -1

    # ─────────────────────────────────────────────────────────────
    #  PREVIEW PANEL
    # ─────────────────────────────────────────────────────────────

    def show_preview(self, entry: PDFEntry):
        """Render the current page of the selected PDF and display it."""
        if entry.broken:
            self._clear_preview(message="This PDF can't be read — it may be corrupted or password protected.")
            return

        self._render_preview_page(entry, self.preview_page_num)

    def _render_preview_page(self, entry: PDFEntry, page_num: int):
        """
        Does the actual fitz render. We keep self.preview_image alive on purpose —
        tkinter's PhotoImage gets garbage collected if nothing holds a reference to it,
        which gives you a blank canvas and a lot of head-scratching.
        """
        try:
            doc = fitz.open(entry.path)
            # Clamp page_num just in case something got out of sync
            page_num = max(0, min(page_num, doc.page_count - 1))
            self.preview_page_num = page_num

            page = doc[page_num]
            matrix = fitz.Matrix(PREVIEW_ZOOM, PREVIEW_ZOOM)
            pix = page.get_pixmap(matrix=matrix, alpha=False)
            doc.close()

            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            self.preview_image = ImageTk.PhotoImage(img)  # store the reference!

            self.preview_canvas.delete("all")
            self.preview_canvas.create_image(0, 0, image=self.preview_image, anchor=tk.NW)
            self.preview_canvas.configure(scrollregion=(0, 0, pix.width, pix.height))

            self.page_label.config(text=f"Page {page_num + 1} of {entry.page_count}")
            self.prev_btn.config(state=tk.NORMAL if page_num > 0 else tk.DISABLED)
            self.next_btn.config(state=tk.NORMAL if page_num < entry.page_count - 1 else tk.DISABLED)

        except Exception as e:
            self._clear_preview(message=f"Couldn't render this page:\n{e}")

    def _clear_preview(self, message: str = "Select a PDF to preview it here"):
        """Reset the preview panel back to its empty state."""
        self.preview_canvas.delete("all")
        self.preview_image = None
        self.preview_canvas.create_text(
            300, 200, text=message, fill="#aaaaaa",
            font=("", 12), tags="placeholder", justify=tk.CENTER
        )
        self.page_label.config(text="")
        self.prev_btn.config(state=tk.DISABLED)
        self.next_btn.config(state=tk.DISABLED)

    def next_page(self):
        if self.selected_index < 0:
            return
        entry = self.pdf_entries[self.selected_index]
        if self.preview_page_num < entry.page_count - 1:
            self.preview_page_num += 1
            self._render_preview_page(entry, self.preview_page_num)

    def prev_page(self):
        if self.selected_index < 0:
            return
        entry = self.pdf_entries[self.selected_index]
        if self.preview_page_num > 0:
            self.preview_page_num -= 1
            self._render_preview_page(entry, self.preview_page_num)

    # ─────────────────────────────────────────────────────────────
    #  MERGING
    # ─────────────────────────────────────────────────────────────

    def merge_pdfs(self):
        """
        Ask where to save, then spin up a background thread to do the actual merge.
        We use a thread so the UI doesn't freeze — merging a bunch of big PDFs can take
        a second and nobody wants a white window that looks like it crashed.
        """
        if not self.pdf_entries:
            messagebox.showwarning("Nothing to Merge", "Add some PDFs first!")
            return

        output_path = filedialog.asksaveasfilename(
            title="Save merged PDF as",
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")],
            initialfile="merged_output.pdf"
        )
        if not output_path:
            return  # user bailed on the save dialog

        # Only merge the readable ones
        readable = [e for e in self.pdf_entries if not e.broken]
        skipped = [e.name for e in self.pdf_entries if e.broken]

        if not readable:
            messagebox.showerror("Nothing Readable", "All the PDFs in your list are unreadable. Nothing to merge.")
            return

        # Show a progress modal so the user knows something is happening
        progress_win = tk.Toplevel(self)
        progress_win.title("Merging…")
        progress_win.geometry("300x100")
        progress_win.resizable(False, False)
        progress_win.grab_set()  # block interaction with the main window while merging
        progress_win.transient(self)

        ttk.Label(progress_win, text=f"Merging {len(readable)} PDF{'s' if len(readable) != 1 else ''}…").pack(pady=(18, 8))
        bar = ttk.Progressbar(progress_win, mode="indeterminate", length=240)
        bar.pack()
        bar.start(12)

        def on_done(error: str | None):
            """Called back on the main thread when the merge finishes."""
            bar.stop()
            progress_win.destroy()

            if error:
                messagebox.showerror("Merge Failed", f"Something went wrong:\n{error}")
            else:
                msg = f"Merged {len(readable)} PDF{'s' if len(readable) != 1 else ''} successfully!\n\nSaved to:\n{output_path}"
                if skipped:
                    msg += f"\n\nSkipped {len(skipped)} unreadable file{'s' if len(skipped) != 1 else ''}:\n" + "\n".join(skipped)
                messagebox.showinfo("Done!", msg)

        def run_merge():
            error = self._do_merge([e.path for e in readable], output_path)
            self.after(0, lambda: on_done(error))

        threading.Thread(target=run_merge, daemon=True).start()

    def _do_merge(self, paths: list[Path], output_path: str) -> str | None:
        """
        The actual merge logic. Runs in a background thread.
        Returns None on success, or an error message string if something goes wrong.
        """
        merged = fitz.open()
        try:
            for path in paths:
                src = fitz.open(path)
                merged.insert_pdf(src)
                src.close()

            merged.save(output_path, garbage=4, deflate=True)
            return None  # success!

        except Exception as e:
            return str(e)
        finally:
            merged.close()


# ─────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    app = PDFMergerApp()
    app.mainloop()
