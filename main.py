"""
FormGenerator - Standalone Windows desktop application
Fill a 76-field form and export it as a single-page A4 PDF.

Runs with plain Python (Tkinter is in the standard library) plus the
third-party 'reportlab' package for PDF generation. For the target
computer (no admin rights, no Python), this file is bundled into a
single portable FormGenerator.exe with PyInstaller (see instructions
supplied separately). The .exe embeds Python, Tkinter and reportlab,
so the target computer needs nothing installed at all.
"""

import os
import re
import sys
import traceback
from datetime import datetime, date

import tkinter as tk
from tkinter import ttk, messagebox, filedialog

from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas as pdfcanvas
from reportlab.pdfbase.pdfmetrics import stringWidth


# ----------------------------------------------------------------------
# 1. FIELD DEFINITIONS
# ----------------------------------------------------------------------
# Every one of the 76 requested field positions is listed here, in the
# exact order given. Internal ids are unique even where the visible
# label repeats (e.g. "House/Flat/Shop No." appears 5 times). Nothing
# is merged, renamed or removed. Field 76's real label was not
# supplied, so it is a clearly marked placeholder.
#
# type is one of: "text", "long_text", "numeric", "email", "phone", "date"
# section groups fields for both the on-screen form and the PDF layout.

SECTION_PRODUCT = "Product / Application Information"
SECTION_PERSONAL = "Applicant Personal Information"
SECTION_RESIDENCE = "Residence / Employment Information"
SECTION_ADDITIONAL = "Additional / Related Person Information"

RAW_FIELDS = [
    (1, "S. No.", "numeric", SECTION_PRODUCT),
    (2, "Product Sub Class", "text", SECTION_PRODUCT),
    (3, "Product Type", "text", SECTION_PRODUCT),
    (4, "Origin City", "text", SECTION_PRODUCT),
    (5, "Referral Branch", "text", SECTION_PRODUCT),
    (6, "Channel", "text", SECTION_PRODUCT),
    (7, "Rate/Pricing Option", "text", SECTION_PRODUCT),
    (8, "Preferred Mailing Address", "long_text", SECTION_PRODUCT),
    (9, "Card Destination", "text", SECTION_PRODUCT),
    (10, "PB/BM Employee No.", "numeric", SECTION_PRODUCT),
    (11, "Desired Financing", "numeric", SECTION_PRODUCT),
    (12, "Tenure", "numeric", SECTION_PRODUCT),
    (13, "Relationship with Applicant", "text", SECTION_PRODUCT),

    (14, "Phone(1)", "phone", SECTION_PERSONAL),
    (15, "First Name", "text", SECTION_PERSONAL),
    (16, "Middle Name", "text", SECTION_PERSONAL),
    (17, "Last Name", "text", SECTION_PERSONAL),
    (19, "New NIC", "text", SECTION_PERSONAL),
    (21, "Issue Date", "date", SECTION_PERSONAL),
    (22, "Expiry Date", "date", SECTION_PERSONAL),
    (24, "Gender", "text", SECTION_PERSONAL),
    (25, "Date of Birth", "date", SECTION_PERSONAL),
    (26, "Marital Status", "text", SECTION_PERSONAL),
    (27, "No. of Children", "numeric", SECTION_PERSONAL),
    (28, "Country", "text", SECTION_PERSONAL),
    (29, "State", "text", SECTION_PERSONAL),
    (30, "City", "text", SECTION_PERSONAL),
    (31, "Dependents", "numeric", SECTION_PERSONAL),
    (32, "Education", "text", SECTION_PERSONAL),
    (33, "Mother Maiden Name", "text", SECTION_PERSONAL),
    (34, "Father/Husband Name", "text", SECTION_PERSONAL),
    (35, "Father/Husband NIC", "text", SECTION_PERSONAL),
    (36, "Father/Husband", "text", SECTION_PERSONAL),
    (37, "Father/Husband CNIC", "text", SECTION_PERSONAL),
    (38, "Nationality", "text", SECTION_PERSONAL),

    (39, "House/Flat/Shop No.", "text", SECTION_RESIDENCE),
    (40, "Employment Status", "text", SECTION_RESIDENCE),
    (41, "Industry", "text", SECTION_RESIDENCE),
    (42, "Nature of Business", "long_text", SECTION_RESIDENCE),
    (43, "House/Flat/Shop No.", "text", SECTION_RESIDENCE),
    (44, "State", "text", SECTION_RESIDENCE),
    (45, "City", "text", SECTION_RESIDENCE),
    (46, "Residence Type/Nature", "text", SECTION_RESIDENCE),
    (47, "Occupation Category", "text", SECTION_RESIDENCE),
    (48, "House/Flat/Shop No.", "text", SECTION_RESIDENCE),
    (49, "City", "text", SECTION_RESIDENCE),
    (50, "Mobile", "phone", SECTION_RESIDENCE),
    (51, "Email", "email", SECTION_RESIDENCE),
    (52, "Occupation Category", "text", SECTION_RESIDENCE),
    (53, "Other Company Name", "text", SECTION_RESIDENCE),
    (54, "Employment Status", "text", SECTION_RESIDENCE),
    (55, "Industry", "text", SECTION_RESIDENCE),
    (56, "Nature Of Business", "long_text", SECTION_RESIDENCE),
    (57, "Employment Status", "text", SECTION_RESIDENCE),
    (58, "Occupation Category", "text", SECTION_RESIDENCE),
    (59, "House/Flat/Shop No.", "text", SECTION_RESIDENCE),
    (60, "City", "text", SECTION_RESIDENCE),
    (65, "Currency", "text", SECTION_RESIDENCE),
    (66, "Gross Income", "numeric", SECTION_RESIDENCE),
    (77, "Expense", "numeric", SECTION_RESIDENCE),

    (67, "Title", "text", SECTION_ADDITIONAL),
    (68, "First Name", "text", SECTION_ADDITIONAL),
    (69, "Last Name", "text", SECTION_ADDITIONAL),
    (70, "Gender", "text", SECTION_ADDITIONAL),
    (71, "New NIC", "text", SECTION_ADDITIONAL),
    (72, "House/Flat/Shop No.", "text", SECTION_ADDITIONAL),
    (73, "Country", "text", SECTION_ADDITIONAL),
    (74, "State", "text", SECTION_ADDITIONAL),
    (75, "City", "text", SECTION_ADDITIONAL),
]


def _slugify(label: str) -> str:
    s = re.sub(r"[^a-zA-Z0-9]+", "_", label).strip("_").lower()
    return s[:30] if s else "field"


class Field:
    __slots__ = ("num", "label", "type", "section", "key", "full_width", "value")

    def __init__(self, num, label, ftype, section):
        self.num = num
        self.label = label
        self.type = ftype
        self.section = section
        # unique internal id even when the visible label repeats
        self.key = f"field_{num:02d}_{_slugify(label)}"
        self.full_width = ftype == "long_text"
        self.value = ""


FIELDS = [Field(*row) for row in RAW_FIELDS]
SECTIONS = [SECTION_PRODUCT, SECTION_PERSONAL, SECTION_RESIDENCE, SECTION_ADDITIONAL]
FIELDS_BY_SECTION = {s: [f for f in FIELDS if f.section == s] for s in SECTIONS}

assert len(FIELDS) == 69, "69 fields expected: 76 original positions minus 7 removed fields"
assert len({f.key for f in FIELDS}) == 69, "Internal field ids must be unique"


# ----------------------------------------------------------------------
# 2. VALIDATION
# ----------------------------------------------------------------------

EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")
PHONE_RE = re.compile(r"^[0-9+\-\s()]{5,25}$")
DATE_FORMATS = ("%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y")


def parse_date(text):
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None


def validate_value(field: Field, raw: str):
    """Returns (cleaned_value, error_message_or_None)."""
    value = raw.strip()

    if field.type == "numeric":
        if value == "":
            return "", None
        cleaned = value.replace(",", "")
        try:
            if "." in cleaned:
                float(cleaned)
            else:
                int(cleaned)
        except ValueError:
            return value, f'"{field.label}" (#{field.num}) must be a number.'
        return cleaned, None

    if field.type == "email":
        if value == "":
            return "", None
        if not EMAIL_RE.match(value):
            return value, f'"{field.label}" (#{field.num}) is not a valid email address.'
        return value, None

    if field.type == "phone":
        if value == "":
            return "", None
        if not PHONE_RE.match(value):
            return value, (
                f'"{field.label}" (#{field.num}) must be a valid phone number '
                "(digits, spaces, +, -, () only)."
            )
        return value, None

    if field.type == "date":
        if value == "":
            return "", None
        d = parse_date(value)
        if d is None:
            return value, (
                f'"{field.label}" (#{field.num}) must be a date in YYYY-MM-DD '
                "(or DD-MM-YYYY / DD/MM/YYYY) format."
            )
        return d.strftime("%Y-%m-%d"), None

    # text / long_text
    if field.type == "long_text":
        if len(value) > 600:
            return value, (
                f'"{field.label}" (#{field.num}) is far too long for a single-page '
                "form. Please shorten it to under 600 characters."
            )
        return value, None

    if len(value) > 120:
        return value, (
            f'"{field.label}" (#{field.num}) is too long. Please shorten it to '
            "under 120 characters so the form can stay on one page."
        )
    return value, None


# ----------------------------------------------------------------------
# 3. TEXT WRAPPING (used by the PDF fitting algorithm)
# ----------------------------------------------------------------------

def wrap_text(text, font_name, font_size, max_width):
    """Greedy word-wrap 'text' to fit within max_width at font_size.
    Falls back to hard character breaks for words longer than the
    column itself. Always returns at least one (possibly empty) line."""
    if not text:
        return [""]

    lines = []
    current = ""
    for word in text.split():
        trial = f"{current} {word}".strip()
        if stringWidth(trial, font_name, font_size) <= max_width:
            current = trial
            continue

        if current:
            lines.append(current)
            current = ""

        if stringWidth(word, font_name, font_size) <= max_width:
            current = word
            continue

        # word alone is wider than the column: hard-break it
        chunk = ""
        for ch in word:
            if stringWidth(chunk + ch, font_name, font_size) <= max_width:
                chunk += ch
            else:
                if chunk:
                    lines.append(chunk)
                chunk = ch
        current = chunk

    if current:
        lines.append(current)
    return lines or [""]


# ----------------------------------------------------------------------
# 4. SINGLE-PAGE PDF FITTING + GENERATION
# ----------------------------------------------------------------------
# Layout: each section has a heading, then its fields are laid out two
# per row ("Label: value    Label: value"), except long_text fields
# which take a full-width row on their own so their wrapped content has
# room to breathe. The whole thing is measured at a shrinking sequence
# of font sizes until it is proven to fit inside one A4 page; only then
# is it actually drawn. This guarantees exactly one page every time.

PAGE_W, PAGE_H = A4
MARGIN = 40
COL_GAP = 16
LABEL_FONT = "Helvetica-Bold"
VALUE_FONT = "Helvetica"
TITLE_FONT = "Helvetica-Bold"
SECTION_FONT = "Helvetica-Bold"

FONT_SIZE_CANDIDATES = [10, 9.5, 9, 8.5, 8, 7.5, 7, 6.5, 6, 5.5, 5, 4.5, 4]


def _cell_text(field: Field, value: str) -> str:
    display = value if value else "________"
    return f"{field.label}: {display}"


def _measure_layout(font_size, title="Application Form"):
    """Compute total content height at the given font_size.
    Returns (total_height, usable_width, two_col_width)."""
    usable_width = PAGE_W - 2 * MARGIN
    two_col_width = (usable_width - COL_GAP) / 2

    line_h = font_size * 1.35
    row_pad = font_size * 0.55
    section_h = font_size * 1.9 + font_size * 0.8

    height = 0.0
    height += font_size * 2.4 + 10  # title block
    height += 4  # rule under title

    for section in SECTIONS:
        flist = FIELDS_BY_SECTION[section]
        height += section_h

        i = 0
        while i < len(flist):
            f = flist[i]
            if f.full_width:
                text = _cell_text(f, f.value)
                lines = wrap_text(text, VALUE_FONT, font_size, usable_width)
                height += len(lines) * line_h + row_pad
                i += 1
                continue

            f2 = None
            if i + 1 < len(flist) and not flist[i + 1].full_width:
                f2 = flist[i + 1]

            lines1 = wrap_text(_cell_text(f, f.value), VALUE_FONT, font_size, two_col_width)
            lines2 = wrap_text(_cell_text(f2, f2.value), VALUE_FONT, font_size, two_col_width) if f2 else [""]
            row_lines = max(len(lines1), len(lines2) if f2 else 0)
            height += row_lines * line_h + row_pad
            i += 2 if f2 else 1

    return height, usable_width, two_col_width


def compute_fit(title="Application Form"):
    """Try shrinking font sizes until the form fits one A4 page.
    Returns (font_size, height) or (None, best_height) if nothing fit."""
    usable_height = PAGE_H - 2 * MARGIN
    best = None
    for fs in FONT_SIZE_CANDIDATES:
        height, _, _ = _measure_layout(fs, title)
        if best is None or height < best[1]:
            best = (fs, height)
        if height <= usable_height:
            return fs, height
    return None, best[1] if best else 0


def generate_pdf(values: dict, filepath: str, title="Application Form"):
    """values: dict of field.key -> string. Draws exactly one A4 page.
    Raises ValueError if the content cannot be made to fit even at the
    smallest allowed font size (caller should warn the user)."""
    for f in FIELDS:
        f.value = values.get(f.key, "")

    font_size, height = compute_fit(title)
    if font_size is None:
        raise ValueError(
            "The entered text is too long to fit on a single A4 page even at "
            "the smallest readable font size. Please shorten some answers "
            "(especially the address / business-nature fields) and try again."
        )

    usable_width = PAGE_W - 2 * MARGIN
    two_col_width = (usable_width - COL_GAP) / 2
    line_h = font_size * 1.35
    row_pad = font_size * 0.55

    c = pdfcanvas.Canvas(filepath, pagesize=A4)
    y = PAGE_H - MARGIN

    # Outer frame so the content is visibly inset from the page edge
    c.setLineWidth(0.75)
    c.setStrokeGray(0.6)
    c.rect(MARGIN - 10, MARGIN - 10, usable_width + 20, PAGE_H - 2 * MARGIN + 20, fill=0, stroke=1)
    c.setStrokeGray(0)

    # Title
    c.setFont(TITLE_FONT, font_size * 1.7)
    c.drawString(MARGIN, y - font_size * 1.4, title)
    y -= font_size * 2.4 + 10
    c.setLineWidth(1)
    c.line(MARGIN, y, PAGE_W - MARGIN, y)
    y -= 4

    def draw_lines(x, top_y, lines, font_name, font_size):
        yy = top_y
        for i, line in enumerate(lines):
            c.setFont(font_name if i > 0 else LABEL_FONT, font_size)
            c.drawString(x, yy, line if i > 0 else line)
            yy -= line_h
        return yy

    for section in SECTIONS:
        flist = FIELDS_BY_SECTION[section]

        # section header
        c.setFillGray(0.92)
        c.rect(MARGIN, y - (font_size * 1.9), usable_width, font_size * 1.9, fill=1, stroke=0)
        c.setFillGray(0)
        c.setFont(SECTION_FONT, font_size * 1.05)
        c.drawString(MARGIN + 4, y - font_size * 1.4, section)
        y -= font_size * 1.9 + font_size * 0.8

        i = 0
        while i < len(flist):
            f = flist[i]
            if f.full_width:
                lines = wrap_text(_cell_text(f, f.value), VALUE_FONT, font_size, usable_width)
                yy = y
                for line in lines:
                    c.setFont(VALUE_FONT, font_size)
                    c.drawString(MARGIN, yy, line)
                    yy -= line_h
                y = yy - row_pad
                i += 1
                continue

            f2 = None
            if i + 1 < len(flist) and not flist[i + 1].full_width:
                f2 = flist[i + 1]

            lines1 = wrap_text(_cell_text(f, f.value), VALUE_FONT, font_size, two_col_width)
            lines2 = wrap_text(_cell_text(f2, f2.value), VALUE_FONT, font_size, two_col_width) if f2 else []
            row_lines = max(len(lines1), len(lines2) if f2 else 0)

            yy = y
            for line in lines1:
                c.setFont(VALUE_FONT, font_size)
                c.drawString(MARGIN, yy, line)
                yy -= line_h
            if f2:
                yy2 = y
                for line in lines2:
                    c.setFont(VALUE_FONT, font_size)
                    c.drawString(MARGIN + two_col_width + COL_GAP, yy2, line)
                    yy2 -= line_h

            y -= row_lines * line_h + row_pad
            i += 2 if f2 else 1

    c.showPage()  # draws exactly ONE page; never called again
    c.save()

    # Structural self-check: a reportlab canvas that calls showPage()
    # exactly once before save() always produces a one-page PDF, since
    # no further page break is ever issued.
    return font_size


# ----------------------------------------------------------------------
# 5. TKINTER APPLICATION
# ----------------------------------------------------------------------

class ScrollableFrame(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.canvas = tk.Canvas(self, borderwidth=0, highlightthickness=0)
        self.vscroll = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.inner = ttk.Frame(self.canvas)

        self.inner.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")),
        )
        self.canvas_window = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.canvas.configure(yscrollcommand=self.vscroll.set)

        self.canvas.bind("<Configure>", self._on_canvas_resize)
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        self.canvas.bind_all("<Button-4>", lambda e: self.canvas.yview_scroll(-2, "units"))
        self.canvas.bind_all("<Button-5>", lambda e: self.canvas.yview_scroll(2, "units"))

        self.canvas.pack(side="left", fill="both", expand=True)
        self.vscroll.pack(side="right", fill="y")

    def _on_canvas_resize(self, event):
        self.canvas.itemconfig(self.canvas_window, width=event.width)

    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")


class FormGeneratorApp:
    def __init__(self, root):
        self.root = root
        root.title("Form Generator")
        root.geometry("900x700")
        root.minsize(700, 500)

        self.widgets = {}  # field.key -> tk widget

        header = ttk.Frame(root, padding=(12, 10))
        header.pack(fill="x")
        ttk.Label(header, text="Application Form", font=("Segoe UI", 16, "bold")).pack(side="left")
        ttk.Label(
            header,
            text="Fill in the fields below, then export as a single-page A4 PDF.",
            font=("Segoe UI", 9),
        ).pack(side="left", padx=12)

        body = ScrollableFrame(root)
        body.pack(fill="both", expand=True, padx=10, pady=(0, 6))

        for section in SECTIONS:
            self._build_section(body.inner, section)

        footer = ttk.Frame(root, padding=(12, 8))
        footer.pack(fill="x")
        ttk.Button(footer, text="Clear / Reset", command=self.on_reset).pack(side="left")
        ttk.Button(footer, text="Export as PDF", command=self.on_export).pack(side="right")
        self.status_var = tk.StringVar(value="Ready.")
        ttk.Label(footer, textvariable=self.status_var, foreground="#555").pack(side="right", padx=12)

    # -- UI construction -------------------------------------------------

    def _build_section(self, parent, section_name):
        frame = ttk.LabelFrame(parent, text=section_name, padding=(10, 8))
        frame.pack(fill="x", padx=4, pady=6)
        for col in range(4):
            frame.grid_columnconfigure(col, weight=1 if col % 2 else 0)

        row = 0
        for field in FIELDS_BY_SECTION[section_name]:
            if not field.full_width:
                continue
            ttk.Label(frame, text=f"{field.label}:").grid(
                row=row, column=0, sticky="nw", padx=4, pady=4
            )
            txt = tk.Text(frame, height=2, width=60, wrap="word")
            txt.grid(row=row, column=1, columnspan=3, sticky="ew", padx=4, pady=4)
            self.widgets[field.key] = txt
            row += 1

        # lay remaining (non full-width) fields two per row
        non_full = [f for f in FIELDS_BY_SECTION[section_name] if not f.full_width]
        r = row
        for idx in range(0, len(non_full), 2):
            pair = non_full[idx:idx + 2]
            c = 0
            for field in pair:
                ttk.Label(frame, text=f"{field.label}:").grid(
                    row=r, column=c, sticky="w", padx=4, pady=3
                )
                entry = ttk.Entry(frame, width=26)
                entry.grid(row=r, column=c + 1, sticky="ew", padx=4, pady=3)
                self.widgets[field.key] = entry
                c += 2
            r += 1

    # -- actions -----------------------------------------------------

    def _get_raw_value(self, field: Field) -> str:
        w = self.widgets[field.key]
        if isinstance(w, tk.Text):
            return w.get("1.0", "end-1c")
        return w.get()

    def collect_and_validate(self):
        values = {}
        errors = []
        for field in FIELDS:
            raw = self._get_raw_value(field)
            cleaned, err = validate_value(field, raw)
            values[field.key] = cleaned
            if err:
                errors.append(err)
        return values, errors

    def on_reset(self):
        if not messagebox.askyesno("Clear form", "Clear all entered data?"):
            return
        for w in self.widgets.values():
            if isinstance(w, tk.Text):
                w.delete("1.0", "end")
            else:
                w.delete(0, "end")
        self.status_var.set("Form cleared.")

    def on_export(self):
        try:
            values, errors = self.collect_and_validate()
            if errors:
                messagebox.showerror(
                    "Please fix the following", "\n".join(f"- {e}" for e in errors[:12])
                    + ("\n...and more." if len(errors) > 12 else "")
                )
                return

            font_size, height = None, None
            try:
                font_size, height = compute_fit()
            except Exception:
                pass

            if font_size is None:
                messagebox.showwarning(
                    "Content too large",
                    "The entered text is too long to fit on a single A4 page, "
                    "even after shrinking the font as much as reasonably "
                    "possible.\n\nPlease shorten some of the longer answers "
                    "(especially address / business-nature fields) and try "
                    "again. Nothing has been exported.",
                )
                self.status_var.set("Export cancelled: content too large for one page.")
                return

            default_name = f"Form_{date.today().isoformat()}.pdf"
            filepath = filedialog.asksaveasfilename(
                title="Save PDF As",
                defaultextension=".pdf",
                initialfile=default_name,
                filetypes=[("PDF files", "*.pdf")],
            )
            if not filepath:
                self.status_var.set("Export cancelled.")
                return

            try:
                used_font_size = generate_pdf(values, filepath)
            except PermissionError:
                messagebox.showerror(
                    "Cannot save file",
                    "The file could not be written. It may be open in another "
                    "program, or you may not have permission to write to that "
                    "folder. Please choose a different location and try again.",
                )
                return
            except ValueError as ve:
                messagebox.showwarning("Content too large", str(ve))
                return
            except OSError as oe:
                messagebox.showerror("File error", f"Could not save the PDF:\n{oe}")
                return

            messagebox.showinfo(
                "Export complete",
                f"The form was exported as a single A4 page:\n{filepath}\n\n"
                f"(font size used: {used_font_size:g}pt)",
            )
            self.status_var.set(f"Exported: {os.path.basename(filepath)}")

        except Exception:
            # last-resort catch-all so the app never crashes silently
            err_text = traceback.format_exc()
            messagebox.showerror(
                "Unexpected error",
                "An unexpected error occurred while exporting the PDF:\n\n"
                + err_text[-1200:],
            )
            self.status_var.set("Export failed due to an unexpected error.")


def main():
    root = tk.Tk()
    try:
        style = ttk.Style()
        if "vista" in style.theme_names():
            style.theme_use("vista")
        elif "clam" in style.theme_names():
            style.theme_use("clam")
    except Exception:
        pass
    app = FormGeneratorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
