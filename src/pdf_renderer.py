"""Pure-Python PDF export for the single-notice workbooks (no LibreOffice / Excel needed).

Why: LibreOffice cannot be installed on some machines and is not available on Streamlit Community Cloud.
The notice workbooks produced by notice_generator are self-contained (no formulas, one sheet, a print area,
one manual page break), so they can be drawn directly with ReportLab:

    column widths + row heights + merged cells + borders + fonts + alignment + number formats + page breaks

It reads the SAME .xlsx that LibreOffice would convert, so the numbers and text in the PDF are the ones in the
workbook. Only the layout engine differs; validator.py checks the result exactly as it checks the other engine.
"""
import logging
import re
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.utils.cell import range_boundaries
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.pdfgen import canvas

log = logging.getLogger("drawdown")

PAD_X = 3.5          # horizontal cell padding (sheet points)
PAD_Y = 1.0
LINE = 1.18          # line height / font size
BORDER_W = {"hair": 0.25, "thin": 0.5, "dotted": 0.5, "dashed": 0.5, "medium": 1.0, "double": 1.0, "thick": 1.5}
DEFAULT_COL_W, DEFAULT_ROW_H = 8.43, 15.0


# ------------------------------------------------------------------ small helpers
def _font(name, bold=False, italic=False):
    base = "Times" if "times" in (name or "").lower() else "Helvetica"
    if base == "Times":
        return {(0, 0): "Times-Roman", (1, 0): "Times-Bold", (0, 1): "Times-Italic", (1, 1): "Times-BoldItalic"}[
            (int(bool(bold)), int(bool(italic)))]
    return {(0, 0): "Helvetica", (1, 0): "Helvetica-Bold", (0, 1): "Helvetica-Oblique",
            (1, 1): "Helvetica-BoldOblique"}[(int(bool(bold)), int(bool(italic)))]


def _safe(text):
    """Standard PDF fonts are WinAnsi: keep what they can draw, replace the rest."""
    text = text.replace("\xa0", " ").replace("\r", "")
    out = []
    for ch in text:
        try:
            ch.encode("cp1252")
            out.append(ch)
        except UnicodeEncodeError:
            out.append({"₹": "Rs.", "→": "->", "–": "-", "—": "-"}.get(ch, "?"))
    return "".join(out)


def _decimals(nf):
    m = re.search(r"\.(0+)", nf)
    return len(m.group(1)) if m else 0


def format_value(v, nf):
    """Excel-like display text for a plain value (the formats used in the notices: accounting, %, General)."""
    if v is None:
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, (int, float)):
        nf = nf or "General"
        if "%" in nf:
            return f"{v * 100:.{_decimals(nf)}f}%"
        if nf != "General" and ("#" in nf or "0" in nf):
            if v == 0 and '"-"' in nf:
                return "-"
            return f"{v:,.{_decimals(nf)}f}"
        return format(v, ".15g")
    if hasattr(v, "strftime"):
        return v.strftime("%d-%b-%Y")
    return str(v)


def wrap(text, font, size, max_w):
    """Greedy word wrap that keeps explicit newlines and runs of spaces."""
    lines = []
    for para in text.split("\n"):
        cur = ""
        for i, w in enumerate(para.split(" ")):
            trial = w if i == 0 else cur + " " + w
            if cur == "" or stringWidth(trial, font, size) <= max_w:
                cur = trial
            else:
                lines.append(cur)
                cur = w
        lines.append(cur)
    return lines


def _runs(value, cell):
    """Cell value -> list of lines; each line is a list of (text, font, size, underline)."""
    f = cell.font
    base_size = float(f.sz or 11)
    base = (f.name, bool(f.b), bool(f.i), base_size, bool(f.u))
    if isinstance(value, CellRichText):
        segs = []
        for part in value:
            if isinstance(part, TextBlock):
                pf = part.font
                segs.append((_safe(part.text), pf.rFont or f.name, bool(pf.b) if pf.b is not None else bool(f.b),
                             bool(pf.i), float(pf.sz or base_size), bool(pf.u)))
            else:
                segs.append((_safe(str(part)),) + base)
        lines, cur = [], []
        for text, name, b, i, sz, u in segs:
            pieces = text.split("\n")
            for k, piece in enumerate(pieces):
                if k > 0:
                    lines.append(cur)
                    cur = []
                if piece:
                    cur.append((piece, _font(name, b, i), sz, u))
        lines.append(cur)
        return lines, True
    text = _safe(format_value(value, cell.number_format))
    return [[(t, _font(f.name, f.b, f.i), base_size, bool(f.u))] for t in text.split("\n")], False


def _line_width(line):
    return sum(stringWidth(t, fn, sz) for t, fn, sz, _ in line)


# ------------------------------------------------------------------ layout
class _Sheet:
    def __init__(self, ws):
        self.ws = ws
        area = (ws.print_area or ws.dimensions).split("!")[-1].replace("$", "")
        self.c1, self.r1, self.c2, self.r2 = range_boundaries(area.split(",")[0])
        self.col_w = self._col_widths()
        self.merge = {}          # anchor (row, col) -> (r1, c1, r2, c2)
        self.covered = set()
        for m in ws.merged_cells.ranges:
            a = (m.min_row, m.min_col)
            self.merge[a] = (m.min_row, m.min_col, m.max_row, m.max_col)
            for r in range(m.min_row, m.max_row + 1):
                for c in range(m.min_col, m.max_col + 1):
                    if (r, c) != a:
                        self.covered.add((r, c))
        self.row_h = self._row_heights()

    def _col_widths(self):
        ws, widths = self.ws, {}
        dflt = ws.sheet_format.defaultColWidth or DEFAULT_COL_W
        for c in range(self.c1, self.c2 + 1):
            widths[c] = dflt
        for dim in ws.column_dimensions.values():
            if dim.min is None or dim.width is None:
                continue
            for c in range(dim.min, min(dim.max or dim.min, self.c2) + 1):
                if c in widths:
                    widths[c] = 0 if dim.hidden else dim.width
        return {c: (w * 7 + 5) * 0.75 for c, w in widths.items()}          # Excel chars -> points

    def _row_heights(self):
        ws = self.ws
        dflt = ws.sheet_format.defaultRowHeight or DEFAULT_ROW_H
        heights, explicit = {}, set()
        for r in range(self.r1, self.r2 + 1):
            d = ws.row_dimensions[r] if r in ws.row_dimensions else None
            if d is not None and d.hidden:
                heights[r] = 0
                explicit.add(r)
            elif d is not None and d.height:
                heights[r] = float(d.height)
                explicit.add(r)
            else:
                heights[r] = float(dflt)
        # grow rows so wrapped text is never cut off (rows with an explicit height only grow for merged blocks)
        for r in range(self.r1, self.r2 + 1):
            for c in range(self.c1, self.c2 + 1):
                if (r, c) in self.covered:
                    continue
                cell = ws.cell(r, c)
                if cell.value in (None, ""):
                    continue
                r1, c1, r2, c2 = self.merge.get((r, c), (r, c, r, c))
                need = self.text_height(cell, self.box_width(c1, c2))
                have = sum(heights[x] for x in range(r1, r2 + 1))
                if need > have + 0.5:
                    target = r2                                              # extra height goes to the last row
                    if r1 == r2 and r1 in explicit and not cell.alignment.wrap_text:
                        continue
                    heights[target] += need - have
        return heights

    def box_width(self, c1, c2):
        return sum(self.col_w[c] for c in range(c1, c2 + 1))

    def text_height(self, cell, width):
        lines, rich = _runs(cell.value, cell)
        if rich:
            return sum(max([sz for _, _, sz, _ in ln] or [cell.font.sz or 11]) * LINE for ln in lines) + 2 * PAD_Y
        if not cell.alignment.wrap_text:
            return len(lines) * float(cell.font.sz or 11) * LINE + 2 * PAD_Y
        font, size = lines[0][0][1], lines[0][0][2]
        text = "\n".join(t for ln in lines for t, *_ in ln)
        return len(wrap(text, font, size, max(width - 2 * PAD_X, 1))) * size * LINE + 2 * PAD_Y

    def page_breaks(self):
        """Rows after which a manual page break occurs, within the print area."""
        return sorted(b.id for b in self.ws.row_breaks.brk if self.r1 <= b.id < self.r2)


# ------------------------------------------------------------------ drawing
def _draw_text(c, sh, cell, x0, y0, x1, y1, ox, oy, s, page_h):
    """Draw one cell's text inside box (x0,y0)-(x1,y1) in sheet points (y grows downward)."""
    lines, rich = _runs(cell.value, cell)
    if not any(ln for ln in lines):
        return
    al = cell.alignment
    wrap_on = bool(al.wrap_text) and not rich
    inner_w = max(x1 - x0 - 2 * PAD_X, 1)
    laid = []                                    # list of (runs, line_height, justify_ok)
    if rich:
        for ln in lines:
            sz = max([r[2] for r in ln] or [cell.font.sz or 11])
            w = _line_width(ln)
            if w > inner_w:                      # shrink an over-long letterhead line to fit
                k = inner_w / w
                ln = [(t, fn, sz2 * k, u) for t, fn, sz2, u in ln]
                sz *= k
            laid.append((ln, sz * LINE, False))
    else:
        font, size = lines[0][0][1], lines[0][0][2]
        text = "\n".join(t for ln in lines for t, *_ in ln)
        under = lines[0][0][3]
        parts = wrap(text, font, size, inner_w) if wrap_on else text.split("\n")
        for i, p in enumerate(parts):
            last = i == len(parts) - 1 or (wrap_on and False)
            laid.append(([(p, font, size, under)], size * LINE, wrap_on and not last))
    total = sum(h for _, h, _ in laid)
    v = al.vertical or "bottom"
    top = y0 + PAD_Y if v == "top" else (y0 + y1) / 2 - total / 2 if v in ("center", "distributed", "justify") \
        else y1 - PAD_Y - total
    h_al = al.horizontal if al.horizontal not in (None, "general") else (
        "right" if isinstance(cell.value, (int, float)) and not isinstance(cell.value, bool) else "left")
    color = cell.font.color
    rgb = color.rgb if color is not None and color.type == "rgb" and isinstance(color.rgb, str) and not rich else None
    ytop = top
    for runs, lh, justify in laid:
        base_y = ytop + lh * 0.80                                  # baseline inside the line box
        width = _line_width(runs)
        spaces = sum(t.count(" ") for t, *_ in runs) - (len(runs[0][0]) - len(runs[0][0].lstrip(" ")) if runs else 0)
        extra = 0.0
        if h_al == "right":
            x = x1 - PAD_X - width
        elif h_al in ("center", "centerContinuous"):
            x = (x0 + x1) / 2 - width / 2
        else:
            x = x0 + PAD_X
            if h_al == "justify" and justify and spaces > 0 and width < inner_w:
                extra = (inner_w - width) / spaces
        for t, fn, sz, u in runs:
            if not t:
                continue
            c.setFont(fn, sz * s)
            if rgb and rgb[2:] != "000000":
                c.setFillColorRGB(*(int(rgb[i:i + 2], 16) / 255 for i in (2, 4, 6)))
            else:
                c.setFillColorRGB(0, 0, 0)
            px, py = ox + x * s, page_h - (oy + base_y * s)
            if extra:
                tobj = c.beginText(px, py)
                tobj.setFont(fn, sz * s)
                tobj.setWordSpace(extra * s)
                tobj.textOut(t)
                c.drawText(tobj)
                w_t = stringWidth(t, fn, sz) + extra * t.count(" ")
            else:
                c.drawString(px, py, t)
                w_t = stringWidth(t, fn, sz)
            if u:
                c.setLineWidth(0.5 * s)
                c.line(px, py - 1.5 * s, px + w_t * s, py - 1.5 * s)
            x += w_t
        ytop += lh


def _draw_borders(c, sh, r, col, x0, y0, x1, y1, ox, oy, s, page_h):
    b = sh.ws.cell(r, col).border
    for side, pts in (("top", (x0, y0, x1, y0)), ("bottom", (x0, y1, x1, y1)),
                      ("left", (x0, y0, x0, y1)), ("right", (x1, y0, x1, y1))):
        st = getattr(b, side)
        if st is not None and st.style:
            c.setLineWidth(BORDER_W.get(st.style, 0.5) * s)
            c.setStrokeColorRGB(0, 0, 0)
            ax, ay, bx, by = pts
            c.line(ox + ax * s, page_h - (oy + ay * s), ox + bx * s, page_h - (oy + by * s))


def render_notice_pdf(xlsx_path, pdf_path):
    """Draw one notice workbook (single sheet) into a PDF. Returns the number of pages."""
    wb = load_workbook(xlsx_path, rich_text=True)
    ws = wb.active
    sh = _Sheet(ws)
    page_w, page_h = A4
    m = ws.page_margins
    left, right, top, bottom = (m.left or 0.7) * 72, (m.right or 0.7) * 72, (m.top or 0.75) * 72, (m.bottom or 0.75) * 72
    total_w = sh.box_width(sh.c1, sh.c2)
    s = min(1.0, (page_w - left - right) / total_w)                 # "fit to 1 page wide": shrink only
    avail_h = (page_h - top - bottom) / s
    ox = left + ((page_w - left - right) - total_w * s) / 2 if ws.print_options.horizontalCentered else left

    # split rows into pages: manual breaks first, then automatically if a block is too tall
    cuts = sh.page_breaks()
    segments, start = [], sh.r1
    for end in cuts + [sh.r2]:
        segments.append((start, end))
        start = end + 1
    pages = []
    for a, b in segments:
        cur_a, h = a, 0.0
        for r in range(a, b + 1):
            if h + sh.row_h[r] > avail_h and r > cur_a:
                pages.append((cur_a, r - 1))
                cur_a, h = r, 0.0
            h += sh.row_h[r]
        pages.append((cur_a, b))

    c = canvas.Canvas(str(pdf_path), pagesize=A4, pageCompression=1, invariant=1)
    c.setTitle(Path(xlsx_path).stem)
    xs = {}
    x = 0.0
    for col in range(sh.c1, sh.c2 + 1):
        xs[col] = x
        x += sh.col_w[col]
    xs[sh.c2 + 1] = x
    for a, b in pages:
        ys, y = {}, 0.0
        for r in range(a, b + 1):
            ys[r] = y
            y += sh.row_h[r]
        ys[b + 1] = y
        oy = top
        # borders first, then text on top
        for r in range(a, b + 1):
            for col in range(sh.c1, sh.c2 + 1):
                _draw_borders(c, sh, r, col, xs[col], ys[r], xs[col + 1], ys[r + 1], ox, oy, s, page_h)
        for r in range(a, b + 1):
            for col in range(sh.c1, sh.c2 + 1):
                if (r, col) in sh.covered:
                    continue
                cell = ws.cell(r, col)
                if cell.value in (None, ""):
                    continue
                r1, c1, r2, c2 = sh.merge.get((r, col), (r, col, r, col))
                r2 = min(r2, b)
                _draw_text(c, sh, cell, xs[c1], ys[r1], xs[c2 + 1], ys[r2 + 1], ox, oy, s, page_h)
        c.showPage()
    c.save()
    wb.close()
    return len(pages)


def export_pdf(xlsx_files, pdf_dir):
    """Same contract as pdf_generator.export_pdf: {xlsx_stem: pdf_path or None}."""
    pdf_dir = Path(pdf_dir)
    pdf_dir.mkdir(parents=True, exist_ok=True)
    results = {}
    for f in map(Path, xlsx_files):
        pdf = pdf_dir / (f.stem + ".pdf")
        try:
            pdf.unlink(missing_ok=True)
            n = render_notice_pdf(f, pdf)
            results[f.stem] = pdf
            log.info("PDF created (built-in renderer): %s (%d pages)", pdf.name, n)
        except Exception as exc:                                    # noqa: BLE001 - one bad file must not stop the batch
            log.error("PDF NOT created for %s: %s", f.name, exc)
            results[f.stem] = None
    return results
