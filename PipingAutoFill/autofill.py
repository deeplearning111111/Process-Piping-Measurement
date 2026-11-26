# -*- coding: utf-8 -*-
"""
Process Piping Measurement - Auto Fill  (version: see VERSION below)

Reads isometric drawing PDFs and appends rows to the Measurement Table (计量表格 / Measurement Sheet).

Drawing types
  SCAN  scanned drawings and AutoCAD PDFs in the older 'STANDART / ДОКУМЕНТ' layout -> OCR (Tesseract),
        plus the exact text layer where a page has one (cut lengths, supports, valve tags, line no.)
  CAD   AutoCAD PDFs with an ITEM CODE column (Priority Unit TKM1, ISO-00159 ...) -> text layer only

Filled : A B C D E F G H I J K M N V (X Y Z when rules.xlsx > Settings allows D5 codes)
         O / P Q  fitting length / reducing-tee run + branch   learned from reference answers
         R  valve tag, T support no. (+ support rows)            read from the drawing
         S  valve length                                         learned, always yellow (check flanges)
         L U Painting Tracing Insulation                         AutoCAD (ITEM CODE) drawings only
Blank  : W

Usage
  python autofill.py [FOLDER]                 fill: PDFs + workbook in FOLDER (default: .\\input)
  python autofill.py FOLDER --selftest        compare with drawings already filled by hand, write nothing
  python autofill.py --learn REFERENCE.xlsx   learn D5 codes, fitting/valve lengths from filled rows
Options
  --engine auto|rapidocr|tesseract   (auto: Tesseract if installed, else RapidOCR)
  --workbook PATH  --force  --rules PATH  --out PATH  --jobs N
"""
import argparse
import copy
import datetime as dt
import difflib
import glob
import math
import os
import re
import shutil
import sys
from collections import Counter, OrderedDict

import numpy as np
import openpyxl
from openpyxl.comments import Comment
from openpyxl.styles import Font, PatternFill
from PIL import Image, ImageFilter

VERSION = "0.9.11"
if getattr(sys, "frozen", False):                  # running as PipingAutoFill.exe
    HERE = os.path.dirname(sys.executable)
else:
    HERE = os.path.dirname(os.path.abspath(__file__))
# HERE = the PROGRAM folder (code, bundled Tesseract, default rules) - may be read-only when installed.
# DATA = the WORKSPACE (rules.xlsx, input, output, backup) - your files, kept when the program is
#        updated or uninstalled. See data_dir() below.


def _documents_dir():
    """The user's Documents folder (also when Windows moved it, e.g. to OneDrive)."""
    if sys.platform.startswith("win"):
        try:
            import ctypes
            from ctypes import wintypes
            import uuid
            guid = uuid.UUID("{FDD39AD0-238F-46AF-ADB4-6C85480369C7}")       # FOLDERID_Documents
            buf = ctypes.c_wchar_p()
            fn = ctypes.windll.shell32.SHGetKnownFolderPath
            fn.argtypes = [ctypes.c_char_p, wintypes.DWORD, wintypes.HANDLE, ctypes.POINTER(ctypes.c_wchar_p)]
            if fn(guid.bytes_le, 0, None, ctypes.byref(buf)) == 0 and buf.value:
                path = buf.value
                ctypes.windll.ole32.CoTaskMemFree(buf)
                return path
        except Exception:
            pass
    return os.path.join(os.path.expanduser("~"), "Documents")


def data_dir():
    """Where rules.xlsx, input, output and backup live.
    - PIPINGAUTOFILL_HOME set          -> that folder
    - source code (python autofill.py) -> next to the code, as before
    - exe with 'portable.txt' beside it -> next to the exe (portable zip edition)
    - installed exe                    -> Documents\\PipingAutoFill"""
    env = os.environ.get("PIPINGAUTOFILL_HOME")
    if env:
        return os.path.abspath(env)
    if not getattr(sys, "frozen", False) or os.path.isfile(os.path.join(HERE, "portable.txt")):
        return HERE
    return os.path.join(_documents_dir(), "PipingAutoFill")


DATA = data_dir()
WORKSPACE_README = r"""Piping AutoFill - your workspace
================================

This folder holds YOUR files. Updating or uninstalling the program never changes them.

  rules.xlsx   your rules (categories, lengths, D5 codes ...). "Learn" updates this file.
               First copied from the rules that came with the program.
  input\       put drawings (PDF) + the allocation workbook here - or drag them onto the window
  output\      the filled workbooks ("..._autofill_<date>.xlsx") and output\ocr_debug
  backup\      a copy of rules.xlsx is saved here before every "Learn"

To start again with the rules that came with the program: rename rules.xlsx (e.g. rules_old.xlsx)
and start Piping AutoFill - it copies the program's rules here again.
"""


def default_rules_path():
    """The rules.xlsx shipped with the program (installed: <program>\\defaults\\rules.xlsx)."""
    for p in (os.path.join(HERE, "defaults", "rules.xlsx"), os.path.join(HERE, "rules.xlsx")):
        if os.path.isfile(p):
            return p
    return None


def ensure_workspace():
    """Create the workspace on first use. Returns a list of what was created (empty if nothing)."""
    made = []
    try:
        if not os.path.isdir(DATA):
            os.makedirs(DATA, exist_ok=True)
            made.append(DATA)
        for sub in ("input", "output"):
            d = os.path.join(DATA, sub)
            if not os.path.isdir(d):
                os.makedirs(d, exist_ok=True)
                made.append(d)
        rules = os.path.join(DATA, "rules.xlsx")
        src = default_rules_path()
        if not os.path.isfile(rules) and src and os.path.abspath(src) != os.path.abspath(rules):
            shutil.copy2(src, rules)
            made.append(rules)
        readme = os.path.join(DATA, "README - workspace.txt")
        if DATA != HERE and not os.path.isfile(readme):
            with open(readme, "w", encoding="utf-8") as f:
                f.write(WORKSPACE_README)
    except Exception as e:                       # e.g. Documents not writable - reported by the caller
        made.append("ERROR: %s" % e)
    return made
SHEET_NAMES = ("Measurement Table", "Measurement Sheet", "计量表格")
ROSTER_NAMES = ("Personnel Roster", "人员名单")
FIRST_DATA_ROW = 4
DPI = 300

# Page zones as fractions of the rendered (landscape) page: (x0, y0, x1, y1)
ZONE_PANEL = (0.566, 0.040, 0.965, 0.590)      # material list + cut pipe length + surface area
ZONE_DRAWING = (0.070, 0.050, 0.550, 0.660)    # isometric drawing area (valve tags, supports)
ZONE_SHEETNO = (0.887, 0.705, 0.916, 0.725)    # "Sheet No" cell
ZONE_TOTAL = (0.921, 0.705, 0.962, 0.725)      # "Total Sheets" cell

YELLOW = PatternFill("solid", fgColor="FFFF00")   # please check
ORANGE = PatternFill("solid", fgColor="FFD8A8")   # from a rule not marked verified yet
VOID_FILL = PatternFill("solid", fgColor="D9D9D9")  # v0.9.2: row of a cancelled (VOID) drawing

VALID_DN = {6, 8, 10, 15, 20, 25, 32, 40, 50, 65, 80, 90, 100, 125, 150, 200, 250, 300, 350,
            400, 450, 500, 550, 600, 650, 700, 750, 800, 850, 900, 1000, 1050, 1100, 1200,
            1300, 1400, 1500, 1600, 1800, 2000}
PIPE_OD = {10.3, 13.7, 17.1, 21.3, 26.7, 33.4, 42.2, 48.3, 60.3, 73.0, 88.9, 101.6, 114.3,
           141.3, 168.3, 219.1, 273.0, 323.8, 323.9, 355.6, 406.4, 457.0, 508.0, 559.0,
           610.0, 660.0, 711.0, 762.0, 813.0, 864.0, 914.0, 1016.0, 1067.0, 1219.0}
SECTION_WORDS = [("PIPE", "Pipe"), ("FITTING", "Fittings"), ("FLANGE", "Flanges"),
                 ("VALVE", "Valve"), ("IN-LINE", "Valve"), ("GASKET", "Gasket"),
                 ("BOLT", "Bolt"), ("SUPPORT", "Support"),
                 ("ITEMS IN", None)]       # v0.9.4, LGN 'ITEMS IN COMPLECT': no section, keywords decide
HEADER_WORDS = ("FABRICATION", "POS ", "COMPONENT", "CUT PIPE", "PIECE", "LENGTH", "STANDART",
                "SURFACE", "STRESS", "WELD DIAMETER", "CONT.", "ERECTION", "QTY")
# valve tag code -> word expected in the valve description (None = any valve)
TAG_TYPES = {"GLV": "GLOBE", "GTV": "GATE", "GAV": "GATE", "BLV": "BALL", "BAV": "BALL", "BV": "BALL",
             "BFV": "BUTTERFLY", "CV": "CHECK", "GV": "GATE",
             "CHV": "CHECK", "CKV": "CHECK", "BFV": "BUTTERFLY", "NDV": "NEEDLE", "PLV": "PLUG",
             "PSV": "SAFETY", "SV": "SAFETY", "PCV": "CONTROL", "FCV": "CONTROL", "LCV": "CONTROL",
             "TCV": "CONTROL"}                                         # v0.9: safety / control valves
COLS = [openpyxl.utils.get_column_letter(c) for c in range(1, 27)]


# ============================================================================ OCR engines
class Box(object):
    __slots__ = ("x0", "y0", "x1", "y1", "text", "conf")

    def __init__(self, x0, y0, x1, y1, text, conf):
        self.x0, self.y0, self.x1, self.y1, self.text, self.conf = x0, y0, x1, y1, text, conf

    @property
    def yc(self):
        return (self.y0 + self.y1) / 2.0

    @property
    def h(self):
        return self.y1 - self.y0


def find_tesseract():
    """Bundled copy next to the program first, then TESSERACT_PATH, PATH and the usual folders
    on every drive (e.g. D:\\Tesseract-OCR)."""
    cands = [os.path.join(HERE, "Tesseract-OCR", "tesseract.exe"),
             os.environ.get("TESSERACT_PATH", "")]
    w = shutil.which("tesseract")
    if w:
        cands.append(w)
    for drive in "CDEFGH":
        for sub in (r"Tesseract-OCR", r"Program Files\Tesseract-OCR",
                    r"Program Files (x86)\Tesseract-OCR", r"Tools\Tesseract-OCR"):
            cands.append(r"%s:\%s\tesseract.exe" % (drive, sub))
    cands += [os.path.expandvars(r"%LOCALAPPDATA%\Programs\Tesseract-OCR\tesseract.exe"),
              os.path.expandvars(r"%LOCALAPPDATA%\Tesseract-OCR\tesseract.exe")]
    for p in cands:
        if p and os.path.isfile(p):
            return p
    return None


class TesseractEngine(object):
    name = "tesseract"

    def __init__(self):
        import pytesseract
        self.pt = pytesseract
        exe = find_tesseract()
        if not exe:
            raise RuntimeError("Tesseract is not installed (or not on PATH).")
        pytesseract.pytesseract.tesseract_cmd = exe
        tessdata = os.path.join(os.path.dirname(exe), "tessdata")
        if os.path.isdir(tessdata) and os.path.abspath(exe).startswith(os.path.abspath(HERE)):
            os.environ["TESSDATA_PREFIX"] = tessdata      # bundled copy: use its own language files
        try:
            self.name = "tesseract %s (%s)" % (pytesseract.get_tesseract_version(), exe)
        except Exception:
            pass

    def boxes(self, img):
        d = self.pt.image_to_data(img, config="--psm 6", output_type=self.pt.Output.DICT)
        out = []
        for i, t in enumerate(d["text"]):
            t = (t or "").strip()
            if t:
                x, y, w, h = d["left"][i], d["top"][i], d["width"][i], d["height"][i]
                out.append(Box(x, y, x + w, y + h, t, float(d["conf"][i]) / 100.0))
        return out

    def sparse_text(self, img):
        return self.pt.image_to_string(img, config="--psm 11")

    def token(self, img, chars="0123456789A"):
        return self.pt.image_to_string(
            img, config="--psm 7 -c tessedit_char_whitelist=%s" % chars).strip()

    def char(self, img, chars="0123456789"):
        return self.pt.image_to_string(
            img, config="--psm 10 -c tessedit_char_whitelist=%s" % chars).strip()


class RapidEngine(object):
    name = "rapidocr"

    def __init__(self):
        try:
            from rapidocr_onnxruntime import RapidOCR   # rapidocr_onnxruntime 1.x
            self.new_api = False
        except ImportError:
            from rapidocr import RapidOCR               # rapidocr 2.x / 3.x
            self.new_api = True
        self.eng = RapidOCR()

    def _items(self, res):
        if self.new_api:
            if res is not None and getattr(res, "boxes", None) is not None:
                return list(zip(res.boxes, res.txts, res.scores))
            return []
        res = res[0] if isinstance(res, tuple) else res
        return [(r[0], r[1], r[2]) for r in (res or [])]

    def boxes(self, img):
        out = []
        for pts, txt, score in self._items(self.eng(np.array(img.convert("RGB")))):
            pts = np.array(pts, dtype=float)
            out.append(Box(pts[:, 0].min(), pts[:, 1].min(), pts[:, 0].max(), pts[:, 1].max(),
                           str(txt).strip(), float(score)))
        return out

    def sparse_text(self, img):
        return "\n".join(line_text(ln) for ln in group_lines(self.boxes(img)))

    def token(self, img, chars="0123456789A"):
        """Recognition only (no text detection) - reliable for a single small cell like '1'."""
        arr = np.array(img.convert("RGB"))
        txt = ""
        try:
            res = self.eng(arr, use_det=False, use_cls=False, use_rec=True)
            txt = "".join(_strings(res))
        except Exception:
            pass
        if not txt:
            txt = "".join(b.text for b in sorted(self.boxes(img), key=lambda b: b.x0))
        return "".join(ch for ch in txt.upper() if ch in chars + "O")


def _strings(obj):
    """Collect recognised strings from any RapidOCR result shape."""
    if obj is None:
        return []
    if isinstance(obj, str):
        return [obj]
    if hasattr(obj, "txts") and obj.txts is not None:
        return [str(t) for t in obj.txts]
    if isinstance(obj, (list, tuple)):
        out = []
        for x in obj:
            if isinstance(x, str):
                out.append(x)
            elif isinstance(x, (list, tuple)):
                out += [y for y in x if isinstance(y, str)]
        return out
    return []


def make_engine(choice):
    errors = []
    order = {"auto": ["tesseract", "rapidocr"], "rapidocr": ["rapidocr"],
             "tesseract": ["tesseract"]}[choice]
    for name in order:
        try:
            return RapidEngine() if name == "rapidocr" else TesseractEngine()
        except Exception as e:  # noqa
            errors.append("%s: %s" % (name, e))
    raise UserError("No OCR engine available.\n  " + "\n  ".join(errors) +
                    "\nRun setup.bat first, or install Tesseract-OCR.")


# ============================================================================ image helpers
A3_LONG_PT = 1190.0   # long side of an A3 sheet in PDF points - OCR thresholds are tuned for this


def _prepare_page(page, strip_images=True):
    """v0.9.11: some subcontractors export the drawing on a reduced sheet (790 x 558 pt instead
    of 1190 x 842) with the text drawn as vector outlines in an unusual font.  At the normal
    300 DPI the glyphs come out about 1.5x smaller than on a standard sheet and Tesseract
    misreads them.  Return the render scale that brings any sheet up to A3 pixel size.

    When the sheet is vector (not a scan), the AFC stamp and the signatures are embedded
    images lying on top of the materials table; the blue stamp pixels hide the black text
    underneath.  Those small images are removed before rendering so the table is read clean.
    A full-page scan (one big image) is never touched."""
    w, h = page.get_size()
    up = max(1.0, A3_LONG_PT / max(w, h, 1.0))
    up = 1.0 + (up - 1.0) * float(os.environ.get("AF_SMALL_UP", "1"))
    scale = DPI / 72.0 * up
    if strip_images:
        try:
            import ctypes
            import pypdfium2.raw as pdfium_c
            objs = list(page.get_objects(max_depth=1))
            paths = sum(1 for o in objs if o.type == pdfium_c.FPDF_PAGEOBJ_PATH)
            if paths >= 200:                       # vector drawing, not a scan
                drop = []
                for o in objs:
                    if o.type not in (pdfium_c.FPDF_PAGEOBJ_IMAGE, pdfium_c.FPDF_PAGEOBJ_FORM):
                        continue
                    l, b, r, t = o.get_bounds()
                    if (r - l) * (t - b) >= 0.10 * w * h:
                        continue                   # big image / block: real content, keep it
                    if o.type == pdfium_c.FPDF_PAGEOBJ_FORM:
                        # the stamp block: a frame, the stamp picture, signature, name, date
                        n = pdfium_c.FPDFFormObj_CountObjects(o.raw)
                        kinds = [pdfium_c.FPDFPageObj_GetType(pdfium_c.FPDFFormObj_GetObject(o.raw, k))
                                 for k in range(max(n, 0))]
                        if pdfium_c.FPDF_PAGEOBJ_IMAGE not in kinds or \
                                kinds.count(pdfium_c.FPDF_PAGEOBJ_PATH) > 40:
                            continue
                    drop.append(o)
                for o in drop:
                    page.remove_obj(o)
                if drop:
                    page.gen_content()
                for o in drop:
                    try:
                        o.close()
                    except Exception:  # noqa
                        pass
        except Exception:  # noqa - never let this stop a drawing
            pass
    return scale


def render_pages(pdf_path):
    import pypdfium2 as pdfium
    pdf = pdfium.PdfDocument(pdf_path)
    try:
        for i in range(len(pdf)):
            page = pdf[i]
            img = page.render(scale=_prepare_page(page)).to_pil()
            if img.height > img.width:
                img = img.rotate(90, expand=True)
            yield i + 1, img
            page.close()
    finally:
        pdf.close()


def clean(img):
    """Keep only neutral (black/grey) drawing ink: removes the blue AFC stamp, the signature
    and any other coloured mark-up, then binarises."""
    a = np.asarray(img.convert("RGB")).astype(np.int16)
    spread = a.max(axis=2) - a.min(axis=2)
    g = a.mean(axis=2)
    colour = spread > 28
    fringe = np.zeros_like(colour)
    fringe[1:, :] |= colour[:-1, :]
    fringe[:-1, :] |= colour[1:, :]
    fringe[:, 1:] |= colour[:, :-1]
    fringe[:, :-1] |= colour[:, 1:]
    ink = (g < 150) & ~colour & ~(fringe & (g > 60))
    return Image.fromarray(np.where(ink, 0, 255).astype(np.uint8)).convert("RGB")


def clean_variant(img, spread_max, grey_max):
    """Softer cleaning: keeps black digits that sit on top of the blue stamp text."""
    a = np.asarray(img.convert("RGB")).astype(np.int16)
    ink = ((a.max(axis=2) - a.min(axis=2)) <= spread_max) & (a.mean(axis=2) < grey_max)
    return Image.fromarray(np.where(ink, 0, 255).astype(np.uint8)).convert("RGB")


def remove_lines(img, minlen=45):
    """Erase long horizontal / vertical strokes (tag boxes, underlines, pipes) so the
    text inside boxes can be read."""
    a = np.asarray(img.convert("L")) < 128
    out = a.copy()
    for b, o in ((a, out), (a.T, out.T)):
        for i in np.where(b.sum(axis=1) >= minlen)[0]:
            d = np.diff(np.concatenate(([0], b[i].astype(np.int8), [0])))
            for s, e in zip(np.where(d == 1)[0], np.where(d == -1)[0]):
                if e - s >= minlen:
                    o[i, s:e] = False
    return Image.fromarray(np.where(out, 0, 255).astype(np.uint8)).convert("RGB")


def crop(img, zone):
    W, H = img.size
    return img.crop((int(W * zone[0]), int(H * zone[1]), int(W * zone[2]), int(H * zone[3])))


def pad(img, p=60):
    out = Image.new("RGB", (img.width + 2 * p, img.height + 2 * p), "white")
    out.paste(img, (p, p))
    return out


def group_lines(boxes):
    boxes = sorted(boxes, key=lambda b: b.yc)
    lines = []
    for b in boxes:
        if lines:
            ln = lines[-1]
            ref = sum(x.yc for x in ln) / len(ln)
            if abs(b.yc - ref) <= 0.55 * max(max(x.h for x in ln), b.h):
                ln.append(b)
                continue
        lines.append([b])
    return [sorted(ln, key=lambda b: b.x0) for ln in lines]


def line_text(ln):
    return " ".join(b.text for b in ln)


# ============================================================================ rules.xlsx
def _yes(v):
    return str(v or "").strip().upper() in ("Y", "YES", "TRUE", "1", "是", "OK")


def _num(v, default):
    try:
        return float(v)
    except (TypeError, ValueError):
        return default


LOOKALIKE = str.maketrans("АВЕКМНОРСТХаеорсхукМ", "ABEKMHOPCTXaeopcxykM")


def latin_part(text):
    """'Elbow P90 273х7 09G2S / Отвод П90 273х7 09Г2С' -> 'Elbow P90 273x7 09G2S'
    (keeps the segments between '/' that are mostly Latin; Cyrillic look-alike letters -> Latin)."""
    t = str(text or "")
    if not re.search(r"[А-Яа-яЁё]", t):
        return t
    keep = []
    # IOT (v0.9.5) writes 'Russian ⁄ English' with a fraction slash - their sizes contain '/'
    # themselves ('TEE TSH 89x6/4-57x4/3'), so split on the fraction slash only
    for seg in (t.split("\u2044") if "\u2044" in t else t.split("/")):
        cyr = len(re.findall(r"[А-Яа-яЁё]", seg))
        lat = len(re.findall(r"[A-Za-z]", seg))
        if lat >= cyr:
            keep.append(seg)
    return " ".join(keep).translate(LOOKALIKE)


_STD_WORDS = re.compile(r"\b(?:G[O0]ST\s+R|G[O0]ST|[O0]ST|ATK|TU|STO|ASME|API|MSS|DIN|EN|ISO)\s+[A-Z]?\d[\w.]*(?:-\d[\w.]*)*",
                        re.I)


def _strip_std(desc):
    return _STD_WORDS.sub(" ", latin_part(str(desc or "")))


# material grades written into IOT / GOST item names ('-09G2S', '-08KH18N10T-1')
_MAT_GRADE = re.compile(r"-?\d{0,2}(?:G2S|G2C|KH\d+N\d+T?|KH17N14M3|H18N10T|X18H10T)\S*", re.I)


def core_key(desc):
    """comp_key without the standard and the material grade (v0.9.5)."""
    return comp_key(_MAT_GRADE.sub(" ", _strip_std(desc)))


def comp_key(desc):
    """Component identity used to look up lengths: the part before ';' without punctuation,
    e.g. 'ASME B16.11 ELBOW S90E SW DN25 CL3000;A350 ...' -> 'ASMEB1611ELBOWS90ESWDN25CL3000'.
    For bilingual (English / Russian) descriptions only the English part is used."""
    s = latin_part(desc).upper()
    # diameter written after Φ: letters that are really digits ('ΦBB.9' -> '88.9')
    s = re.sub(r"[Φ©](\S+?)(?=X)", lambda m: m.group(1).translate(str.maketrans("BOSIL", "80511")), s)
    s = s.replace("Φ", "")
    s = s.split(";")[0]
    s = re.sub(r"[^A-Z0-9]", "", s)
    s = re.sub(r"(?<=[A-Z])0(?=\d{2,3}\d)", "", s)      # 'BE0219' (Φ read as 0) -> 'BE219'
    return s


class Rules(object):
    SHEETS = {
        "Categories": ["Keyword (in description)", "Category (col E)", "中文描述 (col F)",
                       "Verified (Y)", "Note", "Drawing type (blank = all, SCAN, CAD)",
                       "Overrides section (Y)"],
        "D5 Codes": ["Category (col E)", "DN from", "DN to", "D5 Code (col X)",
                     "Description (col Y)", "Unit (col Z)", "Verified (Y)", "Note",
                     "Drawing type (blank = all, SCAN, CAD)"],
        "Fitting Length": ["Component (description up to ';')", "Length (col O)", "Times seen",
                           "Verified (Y)", "Note", "Tee run length (col P)", "Tee branch length (col Q)",
                           "Drawing No. pattern (blank = all)", "DN (from the length key)"],
        "Valve Length": ["Component (description)", "Typical length (col S)", "Times seen",
                         "Other values seen", "Verified (Y)", "Note"],
        "OCR Fixes": ["Find (regular expression)", "Replace with", "Note"],
        "Support Weight": ["Support (description)", "Weight (col U)", "Times seen", "Verified (Y)", "Note"],
        "Settings": ["Drawing No. pattern (* = any text)", "Split bolt kits (Y/N)",
                     "Split assemblies 'consisting of' (Y/N)", "One row per support (Y/N)",
                     "Support QTY = 1 (Y/N)", "Fill D5 codes (Y/N)", "Note",
                     "Pipe line no. short (Y/N)", "J/K also for equal sizes (Y/N)",
                     "DN in J on every row (Y/N)", "Item code from description (Y/N)",
                     "Surface area once per line (Y/N)", "Support weight (Y/N)",
                     "Painting / tracing / insulation once per line (Y/N)",
                     "Fitting lengths from drawing (Y/N)"],
        "D5 List": ["D5 Code", "Unit", "Description", "DN from", "DN to", "Section"],
        "Family Names": ["Drawing No. pattern (* = any text)", "Keyword (in description)", "Category (col E)",
                         "中文描述 (col F)", "Verified (Y)", "Note"],
        "D5 Sections": ["Drawing No. pattern (* = any text)", "Category (col E)", "Code section (e.g. 1300-01-01)",
                        "Verified (Y)", "Note"],
    }

    def __init__(self, path):
        self.path = path
        wb = openpyxl.load_workbook(path, data_only=False)
        # (keyword, E, F, verified, drawing type: "" = all, "SCAN" = scanned/OCR, "CAD" = AutoCAD text)
        self.categories = [(str(r[0]).upper(), r[1], r[2], _yes(r[3]), str(r[5] or "").strip().upper(),
                            _yes(r[6])) for r in self._rows(wb, "Categories") if r[0]]
        self.sup_wt = [(comp_key(r[0]), r[1], _yes(r[3]), str(r[0]))
                       for r in self._rows(wb, "Support Weight") if r[0] and r[1] is not None]
        self.d5 = [(str(r[0]).strip().lower(), _num(r[1], 0), _num(r[2], 99999), r[3], r[4], r[5],
                    _yes(r[6]), str(r[8] or "").strip().upper()) for r in self._rows(wb, "D5 Codes") if r[0] and r[3]]
        # (key, O, verified, note, P, Q) - P/Q = run / branch length of a reducing tee
        # v0.9.5: rows with a drawing-number pattern (learned from a fitting-length key) belong to
        # that subcontractor only - each one dimensions its fittings its own way
        self.fit_len, self.fit_len_fam, self.fit_len_dn = [], [], {}
        for r in self._rows(wb, "Fitting Length"):
            if r[0] and (r[1] is not None or r[5] is not None or r[6] is not None):
                row = (comp_key(r[0]), r[1], _yes(r[3]), r[4], r[5], r[6], str(r[0]))
                if r[7]:
                    self.fit_len_fam.append((str(r[7]).strip().upper(), row))
                    self.fit_len_dn[id(row)] = parse_dn(str(r[8] or ""))
                else:
                    self.fit_len.append(row)
        self.valve_len = [(comp_key(r[0]), r[1], r[3], _yes(r[4]), r[0])
                          for r in self._rows(wb, "Valve Length") if r[0] and r[1] is not None]
        # v0.9: official D5 code list (code -> unit, description, DN band) and code section per family
        self.d5_list = OrderedDict((str(r[0]).strip(), (r[2], r[1], _num(r[3], None), _num(r[4], None), r[5]))
                                   for r in self._rows(wb, "D5 List") if r[0])
        self.family_names = [(str(r[0]).strip().upper(), str(r[1]).upper(), r[2], r[3], _yes(r[4]))
                             for r in self._rows(wb, "Family Names") if r[0] and r[1]]
        self.d5_sections = [(str(r[0]).strip().upper(), str(r[1] or "").strip().lower(), str(r[2]).strip(), _yes(r[3]))
                            for r in self._rows(wb, "D5 Sections") if r[0] and r[2]]
        self.fixes = [(re.compile(str(r[0])), "" if r[1] is None else str(r[1]))
                      for r in self._rows(wb, "OCR Fixes") if r[0]]
        # per drawing family (first matching drawing-number pattern wins); columns found by header
        self.options = []
        if "Settings" in wb.sheetnames:
            ws = wb["Settings"]
            head = [str(c.value or "").upper() for c in ws[1]]

            def find(word, default):
                return next((i for i, h in enumerate(head) if word in h), default)
            ix = {"kits": find("KIT", 1), "assemblies": find("ASSEMBL", 2), "supports": find("ROW PER SUPPORT", 3),
                  "sup_qty": find("SUPPORT QTY", None), "d5": find("D5", None),
                  "line_short": find("LINE NO. SHORT", None), "jk_equal": find("EQUAL SIZES", None),
                  "j_all": find("DN IN J", None), "code_desc": find("ITEM CODE FROM", None),
                  "area_once": find("SURFACE AREA ONCE", None), "sup_weight": find("SUPPORT WEIGHT", None),
                  "pdata_once": find("INSULATION ONCE", None), "fit_len": find("FITTING LENGTHS", None)}
            for r in self._rows(wb, "Settings"):
                if not r[0]:
                    continue
                opt = {k: _yes(r[i]) for k, i in ix.items() if i is not None and k not in ("d5", "jk_equal")}
                opt.setdefault("sup_qty", False)
                opt.setdefault("line_short", False)
                for k in ("j_all", "code_desc", "area_once", "pdata_once", "fit_len"):
                    opt.setdefault(k, False)
                opt["sup_weight"] = True if ix["sup_weight"] is None or r[ix["sup_weight"]] in (None, "") \
                    else _yes(r[ix["sup_weight"]])
                opt["jk_equal"] = True if ix["jk_equal"] is None or r[ix["jk_equal"]] in (None, "") \
                    else _yes(r[ix["jk_equal"]])
                opt["d5"] = True if ix["d5"] is None or r[ix["d5"]] in (None, "") else _yes(r[ix["d5"]])
                self.options.append((str(r[0]).strip().upper(), opt))

    def options_for(self, iso):
        import fnmatch
        for pat, opt in self.options:
            if fnmatch.fnmatch(str(iso).upper(), pat):
                return dict(opt)
        return {"kits": False, "assemblies": False, "supports": False, "sup_qty": False, "d5": True,
                "line_short": False, "jk_equal": True, "j_all": False, "code_desc": False, "area_once": False,
                "sup_weight": True, "pdata_once": False, "fit_len": False}

    @staticmethod
    def _rows(wb, name):
        if name not in wb.sheetnames:
            return []
        return [list(r) + [None] * 8 for r in wb[name].iter_rows(min_row=2, values_only=True)
                if r and any(v is not None for v in r)]

    def classify(self, desc, section, dtype="SCAN", iso=None):
        u = latin_part(desc).upper() if dtype == "CAD" else desc.upper()
        if iso:
            # v0.9.2: 'Family Names' rows decide first for their own drawing family (e.g. *-LGN-*),
            # so a family can name its items its own way without changing the rules for the others
            import fnmatch
            for pat, kw, e, f, ver in self.family_names:
                if e and fnmatch.fnmatch(str(iso).upper(), pat) and \
                        re.search(r"(?<![A-Z])%s(?![A-Z])" % re.escape(kw), u):
                    return e, f, ver
        cats = [c for c in self.categories if not c[4] or c[4] == dtype]
        cats.sort(key=lambda c: 0 if c[4] == dtype else 1)        # rules for this drawing type first

        def hit(kw):
            return re.search(r"(?<![A-Z])%s(?![A-Z])" % re.escape(kw), u) is not None
        for kw, e, f, ver, _, over in cats:          # item-level keywords (bolt, nut, gasket ...)
            if over and e and hit(kw):
                return e, f, ver
        for kw, e, f, ver, _, over in cats:
            if hit(kw) and (not section or not e or e == section):
                return (section or e), f, ver
        for kw, e, f, ver, _, over in cats:
            if hit(kw):
                return (section or e), f, ver
        return section, None, False

    def support_weight(self, desc):
        return self._lookup(self.sup_wt, desc)

    def _lookup(self, table, desc):
        """Exact component first; then (v0.9.2) the same component with its standard left out on both
        sides - hand-filled tables often drop 'GOST 17375-2001' or write 'G0ST'."""
        hit = self._best(table, comp_key(desc))
        if hit or not table:
            return hit
        key = comp_key(_strip_std(desc))
        if not key:
            return None
        cache = self.__dict__.setdefault("_nostd", {})
        alt = cache.get(id(table))
        if alt is None:
            alt = [(comp_key(_strip_std(r[-1])),) + tuple(r[1:]) for r in table]
            cache[id(table)] = alt
        return self._best(alt, key)

    def d5_for(self, cat, dn, dtype="SCAN"):
        for e, lo, hi, code, desc, unit, ver, scope in self.d5:
            if scope and scope != dtype:
                continue
            if e == (cat or "").lower() and dn is not None and lo <= dn <= hi:
                return code, desc, unit, ver
        return None

    def d5_band(self, iso, cat, dn):
        """v0.9: D5 code = section of this family + category, band chosen by the DN from the code list."""
        import fnmatch
        if not dn:
            return None
        dn = dn[0] if isinstance(dn, (tuple, list)) else dn
        canon = lambda c: {"fitting": "fittings", "valves": "valve", "flange": "flanges"}.get(c, c)
        sec = next(((code, ver) for pat, c, code, ver in self.d5_sections
                    if fnmatch.fnmatch(str(iso).upper(), pat) and canon(c) == canon(str(cat).lower())), None)
        if not sec:
            return None
        band = [(k, v) for k, v in self.d5_list.items() if k.startswith(sec[0] + "-") and v[2] is not None]
        if not band:
            if sec[0] in self.d5_list:                       # a single code for any DN (supports)
                v = self.d5_list[sec[0]]
                return sec[0], v[0], v[1], sec[1]
            return None
        inside = [(k, v) for k, v in band if v[2] <= dn <= v[3]]
        above = sorted([(k, v) for k, v in band if v[2] > dn], key=lambda kv: kv[1][2])
        k, v = (inside or above or [max(band, key=lambda kv: kv[1][3])])[0]
        return k, v[0], v[1], sec[1]

    @staticmethod
    def _best(table, key, cutoff=0.88):
        best, score = None, 0.0
        digits = re.findall(r"\d+", key)
        for row in table:
            if row[0] != key and re.findall(r"\d+", row[0]) != digits:
                continue                     # sizes / classes must be identical (CL150 != CL300)
            s = 1.0 if row[0] == key else difflib.SequenceMatcher(None, row[0], key).ratio()
            if s > score:
                best, score = row, s
        return best if score >= cutoff else None

    def fitting_length(self, desc):
        return self._lookup(self.fit_len, desc)

    def family_lengths(self, iso):
        """v0.9.5: Fitting Length rows learned for the drawing family of this ISO No."""
        import fnmatch
        return [row for pat, row in self.fit_len_fam if fnmatch.fnmatch(str(iso).upper(), pat)]

    def prior_rows(self, iso, cond, need=1):
        """v0.9.9: rows to estimate a length from - this family's learned rows that pass cond; a
        family with none (a new subcontractor) borrows the rows of the families that have them."""
        fam = [row for row in self.family_lengths(iso) if cond(row)]
        if len(fam) >= need:
            return fam
        return [row for _, row in self.fit_len_fam if cond(row)]

    def family_length(self, iso, desc, how="exact"):
        """exact: the same item; nostd: the same item without its standard;
        core: the same type and sizes in another material / standard
        ('ELBOW P 90-108x5-09G2S GOST 17375' ~ 'ELBOW P 90-108x5-08KH18N10T GOST 17375')."""
        table = self.family_lengths(iso)
        if not table:
            return None
        if how == "exact":                   # (no fuzzy match: 'TEE TSH 377x10/6-325x8/6' has the
            k = comp_key(desc)               # same digits as 'REDUCER PKSH 377x10/6-325x8/6')
            return next((row for row in table if row[0] == k), None)
        f = (lambda t: comp_key(_strip_std(t))) if how == "nostd" else core_key
        k = f(desc)
        if not k:
            return None
        return next((row for row in table if f(row[-1]) == k), None)

    def valve_length(self, desc):
        return self._lookup(self.valve_len, desc)


# ============================================================================ text helpers
def tidy_description(s, rules):
    s = re.sub(r"\s+", " ", s).strip()
    for stop in ("CUT PIPE LENGTH", "CONT. ON", "CONT ON", "PIPE SURFACE", "FABRICATION MATERIALS"):
        k = s.upper().find(stop)
        if k > 0:
            s = s[:k]
    s = " ".join(t for t in s.split(" ") if re.search(r"[A-Za-z0-9]", t))

    def od(m):
        pre, d, x, wt = m.group(1), m.group(2), m.group(3), m.group(4)
        cands = [d]
        if not pre and len(d) >= 3 and d[0] in "02689":
            cands.insert(0 if d[0] == "0" else 1, d[1:])     # 'Φ' read as 0 / 6 / 8 / 9
        for c in cands:
            try:
                if float(c) in PIPE_OD and ("." in c or x == "X") and not c.startswith("0"):
                    return "Φ%sX%s" % (c, wt)
            except ValueError:
                pass
        return m.group(0)
    s = re.sub(r"(?<![\w.])([©®@Φφ]?)(\d{2,5}(?:\.\d)?)([Xx×])(\d+(?:\.\d+)?)", od, s)
    s = re.sub(r"(?<![\w.])([©®@Φφ]?)(\d{2,5}\.\d)()(\d{1,2}\.\d{1,2})(?=;|\s|$)", od, s)
    s = re.sub(r"(-\d{2,3})\s+2\s+(?=KCV)", r"\1°C ", s)
    s = re.sub(r"(-\d{2,3})\s+(?=KCV)", r"\1°C ", s)          # ℃ lost (vector PDFs draw it apart)
    s = re.sub(r'(-\d{1,3})\s*["°º*]?\s*C\b', r"\1°C", s)
    for rx, rep in rules.fixes:
        s = rx.sub(rep, s)
    return re.sub(r"[\s_~™|—–\-:,'\"`]+$", "", s).strip()


NPS_DN = {"1/2": 15, "3/4": 20, "1": 25, "1 1/4": 32, "1 1/2": 40, "1.5": 40, "2": 50, "2 1/2": 65,
          "3": 80, "4": 100, "5": 125, "6": 150, "8": 200, "10": 250, "12": 300, "14": 350, "16": 400,
          "18": 450, "20": 500, "24": 600}


def parse_dn(text):
    nums = [int(n) for n in re.findall(r"\d+", text or "") if int(n) in VALID_DN]
    if not nums:
        return None
    return nums[:2] if len(nums) >= 2 else nums[:1]


def parse_qty(text):
    t = (text or "").replace(" ", "").replace(",", ".")
    m = re.search(r"(\d+(?:\.\d+)?)", t)
    if not m:
        return None
    v = float(m.group(1))
    return int(v) if v.is_integer() else v


def eval_len(v):
    """'=203+60*2' -> 323.0 ; 280 -> 280.0"""
    if v is None or v == "":
        return None
    s = str(v).strip().lstrip("=")
    if not re.fullmatch(r"[\d.+\-*/() ]+", s):
        return None
    try:
        return float(eval(s, {"__builtins__": {}}))
    except Exception:
        return None


def parse_cut_line(txt):
    """'<1> 500 200 <2> 100 25'  ->  [(1, 500, 200), (2, 100, 25)]"""
    out = []
    marks = list(re.finditer(r"[<«(\[]\s*\S{0,2}?\s*[>»)\]]|[<«]\S{0,2}", txt))
    if marks:
        for i, m in enumerate(marks):
            seg = txt[m.end(): marks[i + 1].start() if i + 1 < len(marks) else len(txt)]
            nums = [int(n) for n in re.findall(r"\d+", seg)]
            piece = re.sub(r"\D", "", m.group(0))
            if len(nums) >= 2 and 10 <= nums[0] <= 99999 and nums[1] in VALID_DN:
                out.append((int(piece) if piece else 0, nums[0], nums[1]))
        if out:
            return out
    nums = [int(n) for n in re.findall(r"\d+", txt)]
    i = 0
    while i + 2 < len(nums):
        p, L, d = nums[i], nums[i + 1], nums[i + 2]
        if 0 < p < 100 and 10 <= L <= 99999 and d in VALID_DN:
            out.append((p, L, d))
            i += 3
        else:
            i += 1
    return out


def qty_ok(lengths, qty):
    """Material-list metres vs cut lengths: rounded UP to 0.1 m (scanned drawings) or rounded to
    0.01 m (AutoCAD drawings)."""
    if not lengths or qty is None:
        return False
    tot, q = sum(lengths), float(qty)
    return (abs(math.ceil(tot / 100.0 - 1e-9) / 10.0 - q) < 1e-6
            or abs(round(tot / 1000.0 + 1e-9, 2) - q) < 1e-6
            or abs(round(tot / 1000.0 + 1e-9, 1) - q) < 1e-6
            or abs(math.floor(tot / 10.0 + 1e-9) / 100.0 - q) < 1e-6)       # 0.415 -> 0.41


def repair_cuts(lengths, qty):
    """One cut length misread (a digit hidden by the signature / an extra digit)?  Try single
    edits - most likely kind first - and accept only an unambiguous fix that makes the total
    match the material-list quantity."""
    kinds = [
        lambda s: [d + s for d in "123456789"],                               # lost leading digit
        lambda s: [s[:k] + s[k + 1:] for k in range(len(s))],                 # extra digit
        lambda s: [s[:k] + d + s[k + 1:] for k in range(len(s)) for d in "0123456789"],  # wrong digit
    ]
    for kind in kinds:
        sols = {}
        for i, L in enumerate(lengths):
            for c in kind(str(L)):
                if c and not c.startswith("0") and 10 <= int(c) <= 99999 and int(c) != L:
                    new = lengths[:i] + [int(c)] + lengths[i + 1:]
                    if qty_ok(new, qty):
                        sols[(i, int(c))] = (new, L, int(c))
        if len(sols) == 1:
            return list(sols.values())[0]
        if sols:
            return None          # several equally likely fixes - leave it to the user
    return None


def restore_decimal(text, qty, cutsets):
    """v0.9.11: the small decimal point of a pipe quantity ('9.1M', '59.7M') is often lost by OCR
    ('901M', '597M').  When the quantity as read does not match any reading of the cut lengths but
    exactly one way of putting the point back does, return that value; otherwise None."""
    t = re.sub(r"[^0-9.]", "", (text or "").replace(",", "."))
    if not t or not any(c for c in cutsets):
        return None
    if qty is not None and any(c and qty_ok(c, qty) for c in cutsets):
        return None
    if "." in t:
        t = ""                                   # the point was read: only the odd-character repair
    cands = set()
    # one character the OCR could not make a digit of ('>.1M' = '9.1M' in some fonts)
    raw = re.sub(r"[\sA-Za-zА-Яа-я]+$", "", (text or "").replace(",", ".")).replace(" ", "")
    odd = [i for i, ch in enumerate(raw) if not (ch.isdigit() or ch == ".")]
    if len(odd) == 1 and len(raw) <= 6:
        for dgt in "0123456789":
            cands.add(raw[:odd[0]] + dgt + raw[odd[0] + 1:])
    for i in range(1, len(t)):
        cands.add(t[:i] + "." + t[i:])                      # point dropped
        if t[i] == "0" and i < len(t) - 1:
            cands.add(t[:i] + "." + t[i + 1:])              # point read as 0
    good = set()
    for c in cands:
        try:
            v = float(c)
        except ValueError:
            continue
        if any(cs and qty_ok(cs, v) for cs in cutsets):
            good.add(v)
    return good.pop() if len(good) == 1 else None


def choose_cuts(cands, qty):
    """cands: cut lengths of one DN from the 3 readings. Returns (lengths, message, mismatch)."""
    ok = [c for c in cands if c and qty_ok(c, qty)]
    if ok:
        votes = Counter(tuple(sorted(c)) for c in ok)
        best = max(ok, key=lambda c: (votes[tuple(sorted(c))], -ok.index(c)))
        sums = sorted({sum(c) for c in ok})
        if len(sums) > 1:
            return best, ("Readings of the cut lengths differ (%s mm); used %s - please check"
                          % (" / ".join(map(str, sums)), "+".join(map(str, best)))), False
        return best, None, False
    for c in cands:
        if c and qty is not None:
            fix = repair_cuts(c, qty)
            if fix:
                return fix[0], ("Auto-corrected a cut length %d -> %d so the total matches QTY %s m"
                                " (digit hidden by stamp?) - please confirm" % (fix[1], fix[2], qty)), False
    first = next((c for c in cands if c), [])
    if first:
        return first, "Cut lengths sum %d mm but material list says %s m" % (sum(first), qty), True
    return [], None, False


# ============================================================================ page reading
ZONE_TITLE = (0.80, 0.64, 0.995, 0.78)
ZONE_TITLE_LOW = (0.80, 0.86, 0.995, 0.96)     # v0.9.2: LGN title block (STAGE / SHEET / TOTAL SHEETS)


def _cells_right_of(reg, box, engine, n=2):
    """Read the next n table cells to the right of an OCR box (digits only)."""
    hh = box.y1 - box.y0
    y0, y1 = int(box.y0 - 0.3 * hh), int(box.y1 + 0.3 * hh)
    band = np.asarray(reg.convert("L").crop((0, y0, reg.width, y1))) < 128
    lines = [x for x in np.where(band.mean(axis=0) > 0.8)[0] if x > box.x1]
    edges = []
    for x in lines:                          # one x per vertical line (lines are a few px wide)
        if not edges or x - edges[-1][-1] > 3:
            edges.append([x])
        else:
            edges[-1].append(x)
    xs = [g[0] for g in edges]
    out = []
    for a, b in zip(xs, xs[1:] + [reg.width]):
        if len(out) == n:
            break
        if b - a < 2 * hh:
            continue
        cell = reg.crop((int(a + 4), y0, int(b - 4), y1))
        t = engine.token(pad(cell)).upper().replace("O", "0")
        t = re.sub(r"[^0-9A]", "", t)
        out.append(t if re.fullmatch(r"\d{1,3}A?", t) else None)
    return (out + [None, None])[:2]


def read_sheet_no(img, engine, pre_cropped=False, title_zone=None):
    """Find the 'Sheet No' / 'Total Sheets' headers (their position differs slightly between
    sheets), then read the cells under them. Fixed zones are the fallback."""
    out = [None, None]
    try:
        reg = img if pre_cropped else crop(img, ZONE_TITLE)
        boxes = engine.boxes(reg)
        heads = {"sheet": None, "total": None}
        for i, b in enumerate(boxes):
            t = b.text.upper()
            if t.startswith("SHEET") and heads["sheet"] is None:
                heads["sheet"] = b
            if t.startswith("TOTAL") and heads["total"] is None:
                heads["total"] = b
        for k, idx in (("sheet", 0), ("total", 1)):
            h = heads[k]
            if h is None:
                continue
            hh = h.y1 - h.y0
            y0, y1 = int(h.y1 + 0.5 * hh), int(h.y1 + 3.3 * hh)
            # cell = between the vertical table lines left and right of the header word
            band = np.asarray(reg.convert("L").crop((0, y0, reg.width, y1))) < 128
            full = np.where(band.mean(axis=0) > 0.8)[0]
            xc = (h.x0 + h.x1) / 2.0
            left = [x for x in full if x < xc]
            right = [x for x in full if x > xc]
            x0 = (max(left) + 4) if left else int(h.x0 - hh)
            x1 = (min(right) - 4) if right else int(h.x1 + 3 * hh)
            cell = reg.crop((int(x0), y0, int(x1), y1))
            # trim horizontal border lines at top/bottom
            arr = np.asarray(cell.convert("L")) < 128
            rows_ = [i for i in range(arr.shape[0]) if arr[i].mean() > 0.8]
            if rows_:
                top = max([i for i in rows_ if i < arr.shape[0] / 2] or [-1]) + 3
                bot = min([i for i in rows_ if i > arr.shape[0] / 2] or [arr.shape[0] + 2]) - 3
                cell = cell.crop((0, top, cell.width, bot))
            t = engine.token(pad(cell)).upper().replace("O", "0")
            t = re.sub(r"[^0-9A]", "", t)
            if re.fullmatch(r"\d{1,3}A?", t):
                out[idx] = t
        stage = next((b for b in boxes if "DD" in b.text.upper() or b.text.upper().startswith("P/")), None)
        if out[0] is None and heads["sheet"] is None and stage is not None:
            # v0.9.2, LGN title block: the small SHEET header is often not read, but the stage
            # 'P/DDD' is; SHEET and TOTAL SHEETS are the next two cells to the right of it
            out[0], out[1] = _cells_right_of(reg, stage, engine)
            if out[0] is None and out[1] == "1":
                out[0] = "1"                 # single-sheet drawing: the SHEET cell is left empty
    except Exception:
        pass
    for idx, zone in enumerate((ZONE_SHEETNO, ZONE_TOTAL)):
        if out[idx] is None:
            if title_zone is not None and title_zone != ZONE_TITLE:
                break                    # the fixed cells are only known for the FGG title block
            if pre_cropped:              # zone relative to the title-block crop
                tx0, ty0, tx1, ty1 = ZONE_TITLE
                zone = tuple((v - o) / (e - o) for v, o, e in
                             zip(zone, (tx0, ty0, tx0, ty0), (tx1, ty1, tx1, ty1)))
            t = engine.token(pad(crop(img, zone))).upper().replace("O", "0").replace(" ", "")
            t = re.sub(r"[^0-9A]", "", t)
            out[idx] = t if re.fullmatch(r"\d{1,3}A?", t) else None
    return out[0], out[1]


def _box_ok(left, right, min_w=4):
    """A crop box is only usable when right is clearly to the right of left."""
    return right - left >= min_w


def parse_panel(img, engine):
    W = img.size[0]
    lines = group_lines(engine.boxes(img))
    col = {"pos": 0.027 * W, "desc": 0.078 * W, "dn": 0.517 * W, "std": 0.614 * W, "qty": 0.83 * W}
    for ln in lines:
        txt = line_text(ln).upper()
        if txt.startswith("POS") and "DN" in txt:
            for b in ln:
                t = b.text.upper()
                if t.startswith("POS"):
                    col["pos"] = b.x0
                elif t.startswith("COMPONENT"):
                    col["desc"] = b.x0
                elif t.startswith("DN"):
                    col["dn"] = b.x0
                elif t.startswith("STANDART") or t.startswith("STANDARD"):
                    col["std"] = b.x0
                elif t.startswith("QTY"):
                    col["qty"] = b.x0
            break
    margin = 0.02 * W

    items, cuts, area, raw = [], [], None, []
    section, cur, last_y, line_h = None, None, None, None
    cut_zone = False
    band = [None, None]
    for ln in lines:
        txt = line_text(ln)
        raw.append(txt)
        up = txt.upper()
        lh = max(b.h for b in ln)
        y = sum(b.yc for b in ln) / len(ln)
        if "SURFACE" in up:
            m = re.search(r"(\d{5,})", up.replace(" ", ""))
            if m:
                area = int(m.group(1))
            if cut_zone and band[1] is None:
                band[1] = min(b.y0 for b in ln)
            cut_zone, cur = False, None
            continue
        if "CUT PIPE" in up or up.startswith("PIECE"):
            if band[0] is None:
                band[0] = min(b.y0 for b in ln)
            cut_zone, cur = True, None
            continue
        if cut_zone:
            cuts += parse_cut_line(txt)
            continue
        first = ln[0]
        if first.x0 > col["pos"] + margin and first.x0 < col["dn"] and not re.match(r"^\d", first.text):
            head = up.split("/")[0].strip()
            if len(head) < 30 and "ASME" not in head and "MSS" not in head:
                for w, cat in SECTION_WORDS:
                    if head.startswith(w) or (w in head and len(head) < 20):
                        section, cur = cat, None
                        break
                else:
                    # v0.9.11: one misread letter ('SUPPGRTS/') must not keep the previous section
                    w1 = re.sub(r"[^A-Z]", "", head.split()[0]) if head.split() else ""
                    for w, cat in SECTION_WORDS:
                        if cat and len(w1) >= 4 and difflib.SequenceMatcher(
                                None, w1[:len(w) + 1], w).ratio() >= 0.8:
                            section, cur = cat, None
                            break
        if any(up.startswith(h) for h in HEADER_WORDS):
            cur = None
            continue

        pos_txt, desc, dn_txt, qty_txt, confs = "", [], [], [], []
        wconf, qbox = [], None
        for b in ln:
            t = b.text
            if b.x0 < col["desc"] - margin / 2:
                m = re.match(r"^(\d{1,2})\b\s*(.*)$", t)
                if m:
                    pos_txt = m.group(1)
                    if m.group(2):
                        desc.append(m.group(2))
                elif re.search(r"[A-Za-z]{3}", t):
                    desc.append(t)
                continue
            if b.x0 < col["dn"] - margin:
                desc.append(t)
                confs.append(b.conf)
                wconf.append((t, b.conf))
            elif b.x0 < col["std"] - margin / 2:
                dn_txt.append(t)
            elif b.x0 >= col["qty"] - 2 * margin:
                qty_txt.append(t)
                qbox = (min(qbox[0], b.x0), max(qbox[1], b.x1)) if qbox else (b.x0, b.x1)
        if desc and not dn_txt:
            m = re.search(r"[^\w.]+(\d{2,4}(?:\s*[xX]\s*\d{2,4})?)\W*$", " ".join(desc))
            if m and all(int(n) in VALID_DN for n in re.findall(r"\d+", m.group(1))):
                dn_txt = [m.group(1)]
                desc = [" ".join(desc)[:m.start()]]
        # a material line always has a DN; a description's 2nd line never has one
        is_item = desc and (pos_txt or (dn_txt and parse_dn(" ".join(dn_txt))))
        if is_item:
            cur = {"pos": int(pos_txt) if pos_txt else None, "section": section,
                   "desc": " ".join(desc), "dn": " ".join(dn_txt), "qty": " ".join(qty_txt),
                   "confs": list(confs), "wconf": list(wconf), "qbox": qbox,
                   "y0": min(b.y0 for b in ln), "y1": max(b.y1 for b in ln)}
            items.append(cur)
            last_y, line_h = y, lh
        elif cur is not None and desc and not dn_txt and not qty_txt \
                and last_y is not None and (y - last_y) < 2.2 * max(line_h, lh):
            cur["desc"] += " " + " ".join(desc)
            cur["confs"] += confs
            cur.setdefault("wconf", []).extend(wconf)
            last_y = y
        else:
            cur = None

    # v0.9.11: a reducer's DN is written over two lines ('80 x' / '50'); the second line is not a
    # new item but the rest of the one above
    for i in range(len(items) - 1, 0, -1):
        prev, it = items[i - 1], items[i]
        cont = re.search(r"\d\s*[xXхХ]$", prev["dn"].strip()) or (
            "REDUC" in prev["desc"].upper() and re.fullmatch(r"\s*\d{1,4}\s*", prev["dn"] or ""))
        if it["pos"] is None and cont and \
                re.fullmatch(r"\s*\d{1,4}\s*", it["dn"] or "") and not parse_qty(it["qty"]):
            sep = " " if re.search(r"[xXхХ]$", prev["dn"].strip()) else " x "
            prev["dn"] = prev["dn"].strip() + sep + it["dn"].strip()
            prev["desc"] += " " + it["desc"]
            prev["confs"] += it["confs"]
            prev.setdefault("wconf", []).extend(it.get("wconf") or [])
            del items[i]

    # v0.9.11: a position number the OCR lost ('$ Kamlok ...') between two that were read
    for i in range(1, len(items) - 1):
        a, b, c = items[i - 1]["pos"], items[i]["pos"], items[i + 1]["pos"]
        if b is None and a is not None and c is not None and c == a + 2:
            items[i]["pos"] = a + 1

    # second look at quantities / DNs the first pass missed (e.g. a lone "1")
    for it in items:
        q = re.sub(r"[\s.,'`~]", "", it["qty"])
        if re.fullmatch(r"[iIl|!\[\]/}{)(]", q):          # a '1' partly hidden by the stamp
            it["qty"], it["qty_reread"] = "1", True
        if parse_qty(it["qty"]) is None and _box_ok(col["qty"] - 2 * margin, min(W, col["qty"] + 0.12 * W)):
            box = (col["qty"] - 2 * margin, it["y0"] - 10, min(W, col["qty"] + 0.12 * W), it["y1"] + 10)
            cell0 = remove_lines(img.crop(tuple(int(v) for v in box)), minlen=30)
            # 2x first (small scanned digits); v0.9.11: then 1x - on an enlarged sheet the
            # digits are already big and Tesseract gives up on them when doubled again
            for f in (2, 1):
                cell = pad(cell0.resize((cell0.width * f, cell0.height * f)))
                it["qty"] = engine.token(cell, "0123456789.M")
                if parse_qty(it["qty"]) is None and hasattr(engine, "char"):
                    it["qty"] = engine.char(cell, "0123456789.M")
                if parse_qty(it["qty"]) is not None:
                    break
            it["qty_reread"] = True
        # the column x positions come from the header words; on a page that is not a normal
        # material list they can come out in the wrong order - skip the re-read then
        if not parse_dn(it["dn"]) and _box_ok(col["dn"] - margin, col["std"] - margin / 2):
            box = (col["dn"] - margin, it["y0"] - 8, col["std"] - margin / 2, it["y1"] + 8)
            it["dn"] = engine.token(pad(img.crop(tuple(int(v) for v in box))), "0123456789xX ")
    return items, cuts, area, raw, band


VALVE_TAG = re.compile(r"\b(\d{3}-([A-Z]{1,3}V)-\d{3,5}[A-Z]?)\b")      # v0.9: suffix A/B
SUP_NO = re.compile(r"(\d{5})-(\d\d)-SUP-(\d{4,6}(?:\.\d{1,2}(?!\d))?)")   # v0.9: .1 / .2 sub-numbers
# PGZ drawings also number special supports '13240-11-SPS-00013' (vector PDFs only)
SUP_NO_CAD = re.compile(r"(\d{5})-(\d\d)-(?:SUP|SPS)-(\d{4,6}(?:\.\d{1,2}(?!\d))?)")
ORI = re.compile(r"ORI\s*[=:-]?\s*([NSEWUD](?:\s*\d+(?:\.\d+)?\s*[NSEWUD])?)\b")


SUP_OCR = re.compile(r"(\d{4,6})-(\d\d)-SUP-(\d{3,6}(?:\.\d{1,2}(?!\d))?)")       # OCR may lose / add a digit of the prefix
SUP_SPLIT = re.compile(r"(SUP-\d{2,4}) (\d{1,2})\b")             # 'SUP-0041 1' -> 'SUP-00411'
SUP_DIGITS = 5                                                    # support numbers have 5 digits (00411)


def expected_sup_prefix(path_or_iso):
    """'GCC-FGG-DDD-12401-10-1100-TKM-ISO-00003' -> '12401-10' (support numbers use unit + area)."""
    m = re.search(r"-(\d{5})-(\d\d)-\d{4}-", os.path.basename(str(path_or_iso or "")))
    return "%s-%s" % (m.group(1), m.group(2)) if m else None


def _near_prefix(a, b):
    """'2401-10' / '42801-10' are misreads of '12401-10': same area part, unit part within 2 edits.
    A different area ('12401-25-SUP-...') is a real other prefix and is never changed."""
    ua, _, aa = a.partition("-")
    ub, _, ab = b.partition("-")
    if aa != ab or abs(len(ua) - len(ub)) > 1:
        return False
    prev = list(range(len(ub) + 1))                # edit (Levenshtein) distance
    for i, ca in enumerate(ua, 1):
        cur = [i]
        for j, cb in enumerate(ub, 1):
            cur.append(min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (ca != cb)))
        prev = cur
    return prev[-1] <= 2


def _ori_by_position(boxes):
    """Support number -> ORI from the label layout: 'ORI  E 89.7723 D' is printed directly UNDER the
    support number, left-aligned. Reading by position avoids mixing up labels that sit close together."""
    out = {}
    sups = [b for b in boxes if SUP_OCR.search(b.text.upper())]
    oris = [b for b in boxes if re.match(r"^[O0][RP][I1L|]", b.text.upper())]
    for s in sups:
        m = SUP_OCR.search(s.text.upper())
        h = max(s.h, 1)
        cand = [o for o in oris if 0 < o.y0 - s.y0 < 3.0 * h and abs(o.x0 - s.x0) < 2.5 * h]
        if not cand:
            continue
        o = min(cand, key=lambda o: (o.y0 - s.y0) + abs(o.x0 - s.x0))
        right = sorted([b for b in boxes if b is not o and abs(b.yc - o.yc) < 0.6 * h
                        and o.x1 - 2 <= b.x0 < o.x1 + 14 * h], key=lambda b: b.x0)
        vals, last = [], o.x1                 # only words that continue the label ('=', 'E', '89.77', 'D')
        for b in right:
            if b.x0 - last > 2.5 * h:
                break
            vals.append(b)
            last = b.x1
        text = re.sub(r"^[O0][RP][I1L|]", "ORI", o.text.upper()) + " " + " ".join(b.text for b in vals).upper()
        om = ORI.search(text.replace("=", " = "))
        if om:
            out[m.group(3)] = _ori_text(om.group(1))
    return out


def _ori_text(v):
    """'E 89.7722D' / 'E  89.7722 D' -> 'E 89.7722 D'."""
    v = re.sub(r"(\d)([NSEWUD])$", r"\1 \2", re.sub(r"\s+", " ", v).strip())
    return re.sub(r"^([NSEWUD])(\d)", r"\1 \2", v)


def read_drawing(img, engine, pre_cropped=False, sup_prefix=None):
    """Valve tags and pipe supports from the isometric drawing area.
    Three OCR readings: two as running text (tags, support numbers) and one with word positions,
    which ties each support number to the ORI printed under it."""
    d = img if pre_cropped else crop(img, ZONE_DRAWING)
    nl = remove_lines(d)
    texts = [engine.sparse_text(nl.resize((int(nl.width * 1.5), int(nl.height * 1.5)))),
             engine.sparse_text(d)]
    try:
        pos_boxes = engine.boxes(d)
    except Exception:
        pos_boxes = []
    by_pos = _ori_by_position(pos_boxes)
    tags = []
    for t in texts:
        for m in VALVE_TAG.finditer(t.upper()):
            if m.group(1) not in [x[0] for x in tags]:
                tags.append((m.group(1), m.group(2)))
    found = OrderedDict()                # '00283' -> {"prefix": Counter, "ori": Counter}
    for t in texts:                      # (the position reading only adds ORIs, never new numbers)
        t = SUP_SPLIT.sub(lambda m: m.group(1) + m.group(2)
                          if len(m.group(1)) - 4 + len(m.group(2)) == SUP_DIGITS else m.group(0), t.upper())
        ms = list(SUP_OCR.finditer(t))
        for i, m in enumerate(ms):
            end = ms[i + 1].start() if i + 1 < len(ms) else len(t)
            om = ORI.search(t, m.end(), min(end, m.end() + 120))
            key = m.group(3)
            e = found.setdefault(key, {"prefix": Counter(), "ori": Counter()})
            e["prefix"]["%s-%s" % (m.group(1), m.group(2))] += 1
            if om:
                e["ori"][_ori_text(om.group(1))] += 1
    # a number read with a digit missing ('0041') next to the full one ('00411') is the same support
    # v0.9: '00025' read once without its sub-number next to '00025.1' is the same label
    for key in [k for k in found if "." not in k]:
        subs = [k for k in found if k.startswith(key + ".")]
        if len(subs) == 1:
            found[subs[0]]["prefix"].update(found[key]["prefix"])
            found[subs[0]]["ori"].update(found[key]["ori"])
            del found[key]
        elif len(subs) > 1:
            del found[key]
    for key in [k for k in found if len(k.split(".")[0]) < SUP_DIGITS]:
        full = [k for k in found if len(k.split(".")[0]) == SUP_DIGITS and (k.startswith(key) or k.endswith(key))]
        if len(full) == 1:
            found[full[0]]["prefix"].update(found[key]["prefix"])
            found[full[0]]["ori"].update(found[key]["ori"])
            del found[key]
    allpref = Counter()
    for e in found.values():
        allpref.update(e["prefix"])
    main_pref = allpref.most_common(1)[0][0] if allpref else None
    sups = []
    for key, e in found.items():
        pref = e["prefix"].most_common(1)[0][0]
        if sup_prefix and pref != sup_prefix and _near_prefix(pref, sup_prefix):
            pref = sup_prefix                      # '2401-10' / '42401-10' misread -> '12401-10' (drawing no.)
        elif main_pref and pref != main_pref and allpref[main_pref] > allpref[pref]:
            pref = main_pref
        no = "%s-SUP-%s" % (pref, key)
        unsure_no = len(key.split(".")[0]) != SUP_DIGITS
        if key in by_pos:
            # the label UNDER the number decides; the running text may have taken a neighbour's ORI.
            # Only if the position reading is incomplete ('E') and the text has the full value, use that.
            ori = by_pos[key]
            longer = [o for o in e["ori"] if o.startswith(ori) and len(o) > len(ori)]
            if longer:
                ori = max(longer, key=lambda o: e["ori"][o])
            short = len(ori) == 1 and any(len(o) > 1 for o in e["ori"])
            sups.append((no, ori, unsure_no or short))
        elif e["ori"]:
            ori, n = e["ori"].most_common(1)[0]
            sups.append((no, ori, unsure_no or len(e["ori"]) > 1))
        else:
            sups.append((no, None, True))
    return tags, sups, texts


# ============================================================================ CAD (AutoCAD text) PDFs
class TObj(object):
    """One text entity of a vector PDF (AutoCAD exports each text as one object)."""
    __slots__ = ("x0", "y0", "x1", "y1", "text", "rgb")

    def __init__(self, x0, y0, x1, y1, text, rgb=None):
        self.x0, self.y0, self.x1, self.y1, self.text, self.rgb = x0, y0, x1, y1, text, rgb

    @property
    def red(self):
        """Revision clouds write the new value in red (the old one is struck through)."""
        return bool(self.rgb) and self.rgb[0] > 150 and self.rgb[1] < 110 and self.rgb[2] < 110

    @property
    def t(self):
        return self.text.strip()


def _fill_rgb(pdfium_c, o):
    try:
        import ctypes
        r, g, b, a = (ctypes.c_uint(), ctypes.c_uint(), ctypes.c_uint(), ctypes.c_uint())
        if pdfium_c.FPDFPageObj_GetFillColor(o.raw, r, g, b, a):
            return (r.value, g.value, b.value)
    except Exception:
        pass
    return None


def _merge_chars(objs):
    """Some AutoCAD exports (LGN 12510-27 ISO-00019) write every character as its own text object.
    The characters of one string still come one after the other in the page's content stream, with
    the space kept on the character before it ('S ', 'H '). Join consecutive characters that sit on
    the same line, close together, back into one text object."""
    out, cur = [], None
    for o in objs:
        if cur is not None:
            h = max(cur.y1 - cur.y0, 1)
            gap = o.x0 - cur.x1
            if (-0.5 <= gap < 1.5 * h and o.y0 <= cur.y1 and o.y1 >= cur.y0 and cur.rgb == o.rgb):
                cur = TObj(cur.x0, min(cur.y0, o.y0), max(cur.x1, o.x1), max(cur.y1, o.y1),
                           cur.text + o.text, cur.rgb)
                continue
            out.append(cur)
        cur = TObj(o.x0, o.y0, o.x1, o.y1, o.text, o.rgb)
    if cur is not None:
        out.append(cur)
    for o in out:
        o.text = re.sub(r"\s+", " ", o.text).strip()
    return [o for o in out if o.text]


def cad_objects(path, index=0):
    """Text objects of a page (top-left origin, points) - or [] when the page has no real text."""
    import pypdfium2 as pdfium
    import pypdfium2.raw as pdfium_c
    pdf = pdfium.PdfDocument(path)
    try:
        page = pdf[index]
        W, H = page.get_size()
        rot = page.get_rotation() % 360
        bx0, by0, bx1, by1 = page.get_cropbox()
        pw, ph = bx1 - bx0, by1 - by0
        tp = page.get_textpage()
        out, blanks = [], set()
        for o in page.get_objects(filter=[pdfium_c.FPDF_PAGEOBJ_TEXT], max_depth=8):
            try:
                if pdfium_c.FPDFTextObj_GetTextRenderMode(o.raw) == 3:
                    continue                 # invisible text = OCR layer of a scanned drawing
            except Exception:
                pass
            o.textpage = tp
            try:
                txt = o.extract()
            except Exception:
                continue
            # signatures drawn with a symbol font come out as control characters - Excel refuses them
            txt = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]", "", txt or "")
            if not txt.strip():
                blank = True                 # v0.9.5: a space of a one-character-per-object sheet
                txt = " "
            else:
                blank = False
            l, b, r, t = o.get_bounds()
            l, r, b, t = l - bx0, r - bx0, b - by0, t - by0
            # page coordinates -> the page as it is displayed (top-left origin). Drawings from
            # some contractors (PGZ / PEGAZ) are stored portrait with /Rotate 270.
            if rot == 90:
                box = (b, l, t, r)
            elif rot == 180:
                box = (pw - r, b, pw - l, t)
            elif rot == 270:
                box = (ph - t, pw - r, ph - b, pw - l)
            else:
                box = (l, ph - t, r, ph - b)
            out.append(TObj(*box, txt.replace("\r", " ").replace("\n", " "), rgb=_fill_rgb(pdfium_c, o)))
            if blank:
                blanks.add(id(out[-1]))
        real = [o for o in out if id(o) not in blanks]
        if len(real) > 150 and sum(len(o.text.strip()) == 1 for o in real) > 0.6 * len(real):
            # IOT sheets of this kind store the spaces as objects of their own - keep them so the
            # words stay apart ('PIPE 57x3,5 GOST', not 'PIPE57x3,5GOST')
            out = _merge_chars(out)
        else:
            out = real
        if sum(len(_MOJI.findall(o.text)) for o in out) >= 10:
            for o in out:
                o.text = _fix_cyrillic(o.text)
        return out, W, H
    finally:
        pdf.close()


_MOJI = re.compile("[\u023a-\u0279]")


def _fix_cyrillic(t):
    """v0.9.5: IOT PDFs map their Cyrillic font 0x1D6 code points too low ('Ɉɬɜɨɞ' = 'Отвод',
    'ʋ' = '№', '»' = the '/' between the Russian and the English text)."""
    t = _MOJI.sub(lambda m: chr(ord(m.group(0)) + 0x1D6), t)
    return t.replace("\u028b", "№").replace("\u00bb", "\u2044")


def is_cad_pdf(path):
    """True for AutoCAD PDFs whose material list has the ITEM CODE layout the text reader knows
    (Priority Unit TKM1, ISO-00159 ...). AutoCAD PDFs in the older 'STANDART / ДОКУМЕНТ' layout
    (e.g. 12401-50 ISO-00051 rev 03) are read like the scanned drawings of that layout."""
    return cad_layout(path) == "ITEM CODE"


def cad_layout(path):
    """Layout of the first page that has a material list. Drawing sets from PGZ (PEGAZ) start with
    a 'Symbols / Условные обозначения' cover sheet, so up to three pages are looked at."""
    try:
        n = page_count(path)
    except Exception:
        return None
    for index in range(min(n, 3)):
        lay = _cad_layout_page(path, index)
        if lay:
            return lay
    return None


def _cad_layout_page(path, index):
    try:
        objs, _, _ = cad_objects(path, index)
    except Exception:
        return None
    texts = " ".join(o.t.upper() for o in objs)
    # v0.9.5: IOT (16100) drawings have only an 'ERECTION MATERIALS' list - same ITEM CODE columns
    if len(objs) <= 40 or not ("FABRICATION" in texts or "ERECTION MATERIALS" in texts) or "POS" not in texts:
        return None
    if "ITEM CODE" in texts:
        return "ITEM CODE"
    return "STANDART" if "STANDART" in texts or "ДОКУМЕНТ" in texts else "OTHER"


def _lines(objs, tol=3.5):
    """Group objects into text lines by their top y."""
    lines = []
    for o in sorted(objs, key=lambda o: (o.y0, o.x0)):
        if lines and abs(o.y0 - lines[-1][0].y0) <= tol:
            lines[-1].append(o)
        else:
            lines.append([o])
    return [sorted(ln, key=lambda o: o.x0) for ln in lines]


def _join(objs):
    return re.sub(r"\s+", " ", " ".join(o.t for o in sorted(objs, key=lambda o: (round(o.y0 / 3.5), o.x0)))).strip()


PCS = re.compile(r"\s*-\s*(\d+)\s*(?:pcs|шт)\.?\s*,?\s*$", re.I)


def _stamp_blue(o):
    return bool(o.rgb) and o.rgb[2] > 120 and o.rgb[2] > o.rgb[0] + 60


# '... 09G2S/09Г2С 80 GOST 33259-2015 /' : description | DN | standard
STD_TAIL = re.compile(r"^(.*\S)\s+(\d{1,4}(?:\s*[xXхХ×]\s*\d{1,4})?)\s+((?:GOST|ГОСТ|OST|ОСТ|ASME|ASTM|MSS|API|"
                      r"TU|ТУ|STO|СТО|EN|DIN|ISO|ATK|АТК|RD|РД|Z-)\b.*)$")


def _std_qty(o, col):
    """'ATK 24.200.02-90 1' running on into the QTY column -> standard | quantity (| unit)."""
    m = re.match(r"^(.*\S)\s+(\d+(?:[.,]\d+)?)\s*([MmМм])?$", o.text.strip())
    if not m or o.x1 < col["qty"]:
        return [o]
    out = [TObj(o.x0, o.y0, col["code"] - 4, o.y1, m.group(1)),
           TObj(col["qty"], o.y0, col["qty"] + 10, o.y1, m.group(2))]
    if m.group(3):
        out.append(TObj(col["qty"] + 20, o.y0, col["qty"] + 26, o.y1, m.group(3)))
    return out


def _split_cad_words(objs, col, std=None):
    """PGZ / LGN drawings sometimes put two table cells into one text object. Split them back:
    '1 Pipe 57x4 ...'  -> position | description
    '... Material 09G2S / 15' running into the DN column -> description | DN
    '200 GOST 8732-78 /' (LGN, DN + STANDART column) -> DN | standard
    '0.9M' in the QTY column -> quantity | unit"""
    out = []
    for o in objs:
        t = o.t
        m = re.match(r"^([1-9]\d{0,2})\s+(\S.*)$", t)
        if m and std and o.x0 < col["desc"] - 3 and o.x1 > std:
            # LGN: 'pos description DN standard' all in one text object
            m2 = STD_TAIL.match(m.group(2))
            if m2:
                out.append(TObj(o.x0, o.y0, o.x0 + 4 * len(m.group(1)), o.y1, m.group(1)))
                out.append(TObj(col["desc"], o.y0, col["dn"] - 4, o.y1, m2.group(1)))
                out.append(TObj(col["dn"], o.y0, std - 4, o.y1, m2.group(2)))
                out += _std_qty(TObj(std, o.y0, o.x1, o.y1, m2.group(3)), col)
                continue
        m2 = STD_TAIL.match(t)
        if m2 and std and o.x0 < col["dn"] - 3 and o.x1 > std:
            out.append(TObj(o.x0, o.y0, col["dn"] - 4, o.y1, m2.group(1)))
            out.append(TObj(col["dn"], o.y0, std - 4, o.y1, m2.group(2)))
            out += _std_qty(TObj(std, o.y0, o.x1, o.y1, m2.group(3)), col)
            continue
        m = re.match(r"^(\d{1,4}(?:\s*[xXхХ×]\s*\d{1,4})?)\s+(\S*[A-Za-zА-Яа-я].*)$", t)
        if m and std and col["dn"] - 3 <= o.x0 < std - 3 and o.x1 > std:
            out.append(TObj(o.x0, o.y0, std - 4, o.y1, m.group(1)))
            out += _std_qty(TObj(std, o.y0, o.x1, o.y1, m.group(2)), col)
            continue
        if std and std - 3 <= o.x0 < col["code"] and o.x1 >= col["qty"]:
            # 'ASME B16.9 1': standard running on into the QTY column
            parts = _std_qty(o, col)
            if len(parts) > 1:
                out += parts
                continue
        m = re.match(r"^([1-9]\d{0,2})\s+(\S.*)$", t)
        if m and o.x0 < col["desc"] - 3:
            out.append(TObj(o.x0, o.y0, o.x0 + 4 * len(m.group(1)), o.y1, m.group(1)))
            o = TObj(col["desc"], o.y0, o.x1, o.y1, m.group(2))
            t = o.t
        m = re.match(r"^(.*\S)\s+(\d{1,4}(?:\s*[xXхХ×]\s*\d{1,4})?)$", t)
        if m and o.x0 < col["dn"] - 3 and o.x1 > col["dn"] + 3 and o.x1 < col["code"] - 3:
            out.append(TObj(o.x0, o.y0, col["dn"] - 4, o.y1, m.group(1)))
            out.append(TObj(col["dn"], o.y0, o.x1, o.y1, m.group(2)))
            continue
        m = re.match(r"^([\d.,]+)\s*([MmМм]|pcs|шт\.?)$", t)
        if m and col["qty"] - 8 <= o.x0 < col["qty"] + 18:
            out.append(TObj(o.x0, o.y0, o.x0 + 10, o.y1, m.group(1)))
            out.append(TObj(col["qty"] + 20, o.y0, col["qty"] + 26, o.y1, m.group(2)))
            continue
        out.append(o)
    return out


def _split_assembly(desc_objs, desc_x):
    """'Embedded structure ..., consisting of:' + '- part' lines (English list, then Russian list).
    Returns [(english, russian, pieces)]."""
    dashes = [o for o in desc_objs if o.t == "-" and o.x0 < desc_x + 15]
    texts = [o for o in desc_objs if o not in dashes and re.search(r"\w", o.t)]
    lines = _lines(texts, tol=4.5)
    starts = set()
    for d in dashes:                       # a dash sits ~5 pt below the top of its text line
        cand = [i for i, ln in enumerate(lines) if d.y0 - 9 <= ln[0].y0 <= d.y0 + 1]
        if cand:
            starts.add(max(cand, key=lambda i: lines[i][0].y0))
    eng, rus, cur, section = [], [], None, "head"
    for i, ln in enumerate(lines):
        txt = _join(ln)
        if section != "head" and ("в составе" in txt or re.search(r"[А-Яа-я].*:\s*$", txt)) and i not in starts:
            section, cur = "rus", None
            continue
        if i in starts:
            if section == "head":
                section = "eng"
            cur = [txt]
            (eng if section == "eng" else rus).append(cur)
        elif cur is not None:
            cur.append(txt)
    parts = []
    for k, e in enumerate(eng):
        et = " ".join(e)
        rt = " ".join(rus[k]) if k < len(rus) else ""
        m = PCS.search(et) or PCS.search(rt)
        n = int(m.group(1)) if m else 1
        et, rt = PCS.sub("", et).strip(" ,"), PCS.sub("", rt).strip(" ,")
        parts.append((et, rt, n))
    return parts


# group captions inside a PGZ material list: 'PIPE/ ТРУБА', 'VALVES/IN-LINE ITEMS/ АРМАТУРА' ...
MAT_GROUP = re.compile(r"^[A-Z][A-Z0-9 ,.()&/-]*/\s*[А-ЯЁ][А-ЯЁ0-9 ,.()-]*$")


def _words(o):
    """Split a text object into words with an estimated x position each (monospaced CAD font)."""
    out, n = [], max(len(o.text), 1)
    for m in re.finditer(r"\S+", o.text):
        out.append((o.x0 + (o.x1 - o.x0) * m.start() / n, m.group(0), o.red))
    return out


def _pgz_cuts(objs, x_left, y_stop=None):
    """PGZ 'CUT PIPE LENGTH / ДЛИНА УЧАСТКА ТРУБЫ' table: two entries side by side per row,
    each '<mark>  length  DN  [SETON TEE]'. The words come as separate or merged text objects,
    the mark is sometimes missing, and a revised length is written in red next to the old one.
    Returns [(mark, length, dn)]."""
    cap = next((o for o in objs if "CUT PIPE LENGTH" in o.t.upper() and o.x0 >= x_left - 15), None)
    if cap is None:
        return []
    hdr = [o for o in objs if "CL LENGTH" in o.t.upper() and o.x0 >= x_left - 15 and 0 < o.y0 - cap.y0 < 30]
    if not hdr:
        return []
    hy = min(o.y0 for o in hdr)
    starts = sorted(o.x0 for o in objs if abs(o.y0 - hy) < 3 and o.t.startswith("№") and o.x0 >= x_left - 15)
    if len(starts) < 2:
        cl = sorted(o.x0 for o in hdr)
        starts = [x_left] + ([cl[1] - 25] if len(cl) > 1 else [])
    split = starts[1] - 5 if len(starts) > 1 else 1e9
    stop = min([o.y0 for o in objs if o.y0 > hy + 5 and o.x0 >= x_left - 15
                and ("SURFACE" in o.t.upper() or "CUT PIPE" in o.t.upper())] + [y_stop or 1e9])
    body = [o for o in objs if o.x0 >= x_left - 10 and hy + 5 < o.y0 < stop - 0.5]
    out = []
    for ln in _lines(body, tol=3):
        for side in (0, 1):
            ws = sorted(w for o in ln for w in _words(o) if (w[0] >= split) == bool(side))
            mark = next((re.sub(r"\D", "", t) for _, t, _ in ws if re.fullmatch(r"<\d{1,3}>", t)), None)
            nums = [(int(t), red) for _, t, red in ws if re.fullmatch(r"\d{1,6}", t)]
            if len(nums) < 2:
                continue
            dn, lengths = nums[-1][0], nums[:-1]
            if dn not in VALID_DN:
                continue
            if len(lengths) > 1:             # '569' struck through, '587' written in red
                lengths = [n for n in lengths if n[1]] or lengths[-1:]
            out.append((mark or "?", lengths[0][0], dn))
    return out


KIT_PART = re.compile(r"([A-Za-z][A-Za-z ]*/[^,]*?)\s+-\s+(\d+)\s*pcs/шт", re.I)


class NoMaterialList(RuntimeError):
    pass


def read_cad_page(path, index=0, prev_cols=None):
    """Material list, cut lengths, tags, supports and title data from an AutoCAD-made isometric.
    prev_cols: the column layout of the previous sheet, for a PGZ sheet whose material list is
    continued from there ('CONT. FROM SHEET 1', no header row of its own)."""
    objs, W, H = cad_objects(path, index)
    up = lambda o: o.t.upper()
    head = next((o for o in objs if up(o) == "POS" or up(o).startswith("POSCOMPONENT")), None)
    if head is None:
        # a sheet that continues the list ('CONT. FROM SHEET 1') or only carries the cut table
        # of the previous sheet's list (LGN: 'CUT PIPE LENGTH' at the top, no header row)
        cont = next((o for o in objs if (re.match(r"CONT\.?\s*FROM", up(o)) or up(o).startswith("CUT PIPE LENGTH"))
                     and prev_cols and o.x0 > prev_cols["pos"] - 20), None)
        if cont is None:
            raise NoMaterialList("Material list header (POS) not found")
        col, pgz = dict(prev_cols), True
        iot = bool(col.pop("iot", False))
        row_y = cont.y0 - 6
        head = TObj(col["pos"], row_y, col["pos"] + 10, row_y + 5, "POS")
        hdr = []
    else:
        row_y = head.y0
        # the POS word can sit a few points off its header row (LGN 12510-24 TS-ISO-00005)
        comp = next((o for o in objs if up(o).startswith("COMPONENT") and 0 < abs(o.y0 - head.y0) < 8
                     and o.x0 > head.x0), None)
        if comp is not None and not any(abs(o.y0 - row_y) < 3 and up(o).startswith("COMPONENT") for o in objs):
            row_y = comp.y0
        hdr = [o for o in objs if abs(o.y0 - row_y) < 3]
        col = {"pos": head.x0, "desc": head.x0 + 14, "dn": None, "code": None, "qty": None}
        pgz = iot = False                    # PGZ / PEGAZ variant of the ITEM CODE layout
        # v0.9.5: IOT lists ('ERECTION MATERIALS' only, group captions, '6.2M') are read the PGZ way
        if not any(up(o).startswith("FABRICATION") for o in objs) and \
                any(up(o).startswith("ERECTION MATERIALS") for o in objs):
            pgz = iot = True
    for o in hdr:
        u = up(o)
        if u.startswith("COMPONENT"):
            col["desc"] = o.x0
        elif u == "DN":
            col["dn"] = o.x0
        elif re.fullmatch(r"DN\s*\(\s*MM\s*\)", u):                    # PGZ writes 'DN(MM)'
            col["dn"], pgz = o.x0, True
        elif u.replace(" ", "").startswith("ITEMCODE"):     # (char-split IOT sheets: 'ITEMCODE/КОДПОЗ.')
            col["code"] = o.x0
        elif u.startswith("QTY"):
            col["qty"] = o.x0
        elif u.startswith("STANDART") or u.startswith("STANDARD"):
            # LGN drawings: a STANDART / ДОКУМЕНТ column between DN and ITEM CODE. Their list is laid
            # out like the PGZ one (group captions, '12.9M' quantities), so it gets the same handling
            col["std"], pgz = o.x0, True
    col_std = col.pop("std", None)
    if None in col.values():
        missing = ", ".join(k for k, v in col.items() if v is None)
        raise RuntimeError("Material list columns not recognised (missing: %s; header: %s)"
                           % (missing, " | ".join(o.t for o in sorted(hdr, key=lambda o: o.x0))))
    x_left = col["pos"] - 4
    # end of the material list = first cut-table row / table caption / surface area below it
    cut_re = re.compile(r"^([A-Z])\s+(\d{1,6})\s+(\d{1,4})$")
    below = [o for o in objs if o.x0 >= x_left - 10 and o.y0 > row_y + 12 and
             (cut_re.match(o.t) or "CL LENGTH" in up(o) or "CUT PIPE" in up(o) or "SURFACE" in up(o))]
    # (a PGZ list runs down to the PIPE SURFACE AREA line at about 2/3 of the sheet)
    y_end = min([o.y0 for o in below] + [H * (0.68 if pgz else 0.6)])
    area_objs = [o for o in objs if o.x0 >= x_left - 15 and o.y0 > row_y + 12 and o.y0 < y_end - 0.5]
    # PGZ lists: a second table ('ERECTION MATERIALS', with its own POS header row) and group
    # captions ('PIPE/ ТРУБА', 'FITTINGS/ ФИТИНГИ' ...) sit between the items - they are not text
    # of the item above them
    if pgz:
        more_hdr = [o for o in area_objs if (up(o) == "POS" and abs(o.x0 - head.x0) < 6)
                    or (up(o).startswith("COMPONENT DESCRIPTION") and abs(o.x0 - col["desc"]) < 8)]
        hdr_rows = [h.y0 for h in more_hdr]
        area_objs = [o for o in area_objs
                     if not any(-14 <= o.y0 - y <= 12 for y in hdr_rows)
                     and not (MAT_GROUP.match(o.t) and o.x0 < col["dn"] - 3 and not re.search(r"\d", o.t))
                     and "MATERIALS" not in up(o)
                     and not re.fullmatch(r"\*+", o.t) and not re.match(r"CONT\.?\s*(ON|FROM)", up(o))]
    sub_hdr = [o for o in area_objs if o.y0 < row_y + 12]
    area_objs = [o for o in area_objs if o not in sub_hdr and abs(o.y0 - (row_y + 9)) > 3]
    if pgz:
        # the blue 'В ПРОИЗВОДСТВО РАБОТ / AFC' stamp lies over the list; red revision-triangle
        # numbers ('03', '09') are not table text either
        area_objs = [o for o in area_objs if not _stamp_blue(o)
                     and not (o.red and re.fullmatch(r"0\d", o.t))]
        area_objs = _split_cad_words(area_objs, col, col_std)
        # '09' next to an item is a revision-cloud mark, not a position number
        pos_re, tol, q_lo, d_lo = r"[1-9]\d{0,2}", 3.5, 20, 6
    else:
        pos_re, tol, q_lo, d_lo = r"\d{1,3}", 1.5, 8, 3
    pos_objs = sorted([o for o in area_objs if o.x0 < col["desc"] - 3 and re.fullmatch(pos_re, o.t)
                       and (not pgz or o.x0 > col["pos"] - 5)],    # not a note number left of the table
                      key=lambda o: o.y0)
    # the same number twice on one row (a separate word and a merged '6 Stud ...') is one position
    pos_objs = [p for k, p in enumerate(pos_objs)
                if not any(q.t == p.t and abs(q.y0 - p.y0) < 3 for q in pos_objs[:k])]
    items = []
    for k, p in enumerate(pos_objs):
        # PGZ: the number can sit a few points below the first line of its description
        y_next = pos_objs[k + 1].y0 - (tol if pgz else 1) if k + 1 < len(pos_objs) else y_end
        mine = [o for o in area_objs if p.y0 - (tol if pgz else 2) <= o.y0 < y_next and o is not p]
        first = [o for o in mine if o.y0 <= p.y0 + 12]
        dn = [o for o in (mine if pgz else first) if col["dn"] - (6 if pgz else 3) <= o.x0 < (col_std or col["code"]) - 3
              and re.fullmatch(r"[\d.,xх×\s-]+", o.t) and (not pgz or o.y0 <= p.y0 + 22)]
        dn1 = [o for o in dn if o.y0 <= p.y0 + 3] or dn[:1]
        if pgz and dn1:
            # reducer / tee: '100' on the first line, 'x 80' under it (or '80 x' + '80')
            for o in sorted([o for o in dn if o not in dn1 and o.y0 <= p.y0 + 22], key=lambda o: o.y0):
                # '100' / 'x 80', '80 x' / '80', or '150' / 'x' / '100' on three lines
                if re.match(r"^\s*[xхХ×]", o.t) or re.search(r"[xхХ×]\s*$", _join(dn1)):
                    dn1 = dn1 + [o]
                else:
                    break
        dn = dn1
        code = [o for o in mine if col["code"] - 3 <= o.x0 < col["qty"] - 5 and o.y0 <= p.y0 + 14
                and (o.y0 <= p.y0 + 3 or re.search(r"\d{3}", o.t))]
        qty = [o for o in first if col["qty"] - q_lo <= o.x0 < col["qty"] + 18 and re.fullmatch(r"[\d.,]+", o.t)]
        unit = [o for o in first if o.x0 >= col["qty"] + 18]
        used = set(map(id, dn + code + qty + unit))
        wide = [o for o in mine if id(o) not in used and col["desc"] - d_lo <= o.x0 < col["qty"] - 5]
        text = _join(wide)
        if re.search(r"consisting of|в составе", text, re.I):
            desc = wide                      # assemblies: Russian fragments spill over the DN column
        else:
            desc = [o for o in wide if o.x0 < col["dn"] - 3]
            # a later line that starts in the DN column and is not a number continues the item code
            spill = [o for o in wide if o.x0 >= col["dn"] - 3 and (not re.fullmatch(r"[\d.,\s]+", o.t)
                                                                   or (col_std and o.x0 >= col_std - 3))]
            code = code + spill
            text = _join(desc)
        code_txt = _join(code)
        if pgz:
            code_txt = re.sub(r"(?<=\w)- (?=\w)", "-", code_txt)     # 'GCC-LGN-DDD-00000- 00-0000-…'
        if pgz and not text.strip() and not dn and not qty:
            continue                         # a tag number of the isometric reaching into the list
        it = {"pos": int(p.t), "desc": text, "dn": _join(dn), "code": code_txt,
              "qty": _join(qty), "unit": _join(unit), "parts": None, "kit": None}
        if re.search(r"consisting of|в составе", text, re.I):
            parts = _split_assembly(desc, col["desc"])
            if parts:
                it["parts"] = parts
                it["head"] = text.split(":")[0].strip()
        elif re.search(r"kit|компл", it["unit"], re.I) and KIT_PART.search(text):
            it["kit"] = [(m.group(1).strip(), int(m.group(2))) for m in KIT_PART.finditer(text)]
        items.append(it)

    cuts = []
    for o in objs:
        m = cut_re.match(o.t)
        if m and o.x0 >= x_left - 10:
            cuts.append((m.group(1), int(m.group(2)), int(m.group(3))))
    if not cuts:
        cuts = _pgz_cuts(objs, x_left)
    alltext = "\n".join(o.t for o in objs)
    info = {"area": None, "line": None, "paint": None, "trace": None, "insul": None, "rev": None}
    for o in objs:
        u = up(o)
        if "SURFACE AREA" in u:
            m = re.search(r"\)\s*([\d.,]+)\s*$", o.t) or re.search(r"([\d]+[.,]\d+)\s*$", o.t)
            if m:
                info["area"] = float(m.group(1).replace(",", "."))
            else:
                nxt = [x for x in objs if abs(x.y0 - o.y0) < 3 and x.x0 > o.x1 and re.fullmatch(r"[\d.,]+", x.t)]
                if not nxt:
                    # IOT: '58.1 (TOTAL FOR / ОБЩАЯ ДЛЯ 030-RW-0028)'
                    nxt = [TObj(x.x0, x.y0, x.x1, x.y1, re.match(r"\d+(?:[.,]\d+)?", x.t).group(0))
                           for x in objs if abs(x.y0 - o.y0) < 3 and x.x0 > o.x1
                           and re.match(r"\d+(?:[.,]\d+)?\s*\(TOTAL", x.t.upper())]
                if nxt:
                    info["area"] = float(nxt[0].t.replace(",", "."))
    lab = next((o for o in objs if up(o).startswith("LINE NUMBER")), None)
    if lab:
        cand = [o for o in objs if abs(o.y0 - lab.y0) < 4 and o.x0 > lab.x1 and re.search(r"\d{3}-[A-Z][A-Z0-9]*-\d", o.t)]
        if cand:
            info["line"] = min(cand, key=lambda o: o.x0).t
    for key, word in (("paint", "PAINTING"), ("trace", "TRACING"), ("insul", "INSULATION")):
        lab = next((o for o in objs if up(o).startswith(word)), None)
        if lab:
            # IOT title blocks have the next label column ('PIPE CATEGORY GOST') right of the value
            stop = min([o.x0 for o in objs if o.x0 > lab.x1 and abs(o.y0 - lab.y0) < 4
                        and re.match(r"PIPE (CATEGORY|DIAMETER|CLASS)", up(o))] + [10 ** 6])
            vals = [o for o in objs if lab.x1 + 20 < o.x0 < lab.x1 + 150 and -4 <= o.y0 - lab.y0 <= 8
                    and o.x0 < lab.x0 + 140 and o.x0 < stop - 2
                    and not re.fullmatch(r"0\d", o.t)]     # '01' = revision mark
            if vals:
                seen_ = set()           # (IOT draws some title texts twice on top of each other)
                vals = [o for o in vals if not ((o.t, round(o.x0), round(o.y0)) in seen_
                                                or seen_.add((o.t, round(o.x0), round(o.y0))))]
                info[key] = re.sub(r"\s+", "", _join(vals))
    tags = []
    for m in VALVE_TAG.finditer(alltext.upper()):
        if m.group(1) not in [t[0] for t in tags]:
            tags.append((m.group(1), m.group(2)))
    sups = sorted({m.group(0) for m in SUP_NO_CAD.finditer(alltext.upper())})
    return {"items": items, "cuts": cuts, "info": info, "tags": tags, "sups": sups,
            "cols": dict(col, std=col_std, iot=iot) if pgz else None, "iot": iot,
            "title_sheet": _title_sheet(objs, W, H) if pgz else None,
            "total_sheets": _title_total(objs, W, H) if pgz else None,
            "dbg": "\n".join("%5.0f %5.0f  %s" % (o.y0, o.x0, o.t) for o in sorted(objs, key=lambda o: (o.y0, o.x0)))}


LINE_FULL = re.compile(r"(\d{3}-[A-Z][A-Z0-9]*-\d{3,5}(?:-[A-Z0-9]+)*)")


def _line_list(objs):
    """Cover sheet of a PGZ drawing set: '№ Line / № Линии' + 'Count sheet / Кол. листов' table.
    Returns [(line, sheets)] in the order of the PDF."""
    head = next((o for o in objs if o.t.upper().replace(" ", "").startswith("№LINE")), None)
    if head is None:
        return []
    out = []
    for o in sorted(objs, key=lambda o: o.y0):
        if o.y0 <= head.y0 + 5 or abs(o.x0 - head.x0) > 120:
            continue
        m = LINE_FULL.search(o.t.split("/")[-1].upper())
        if not m:
            continue
        cnt = [x for x in objs if abs(x.y0 - o.y0) < 3 and x.x0 > o.x1 and x.x0 < o.x1 + 200
               and re.fullmatch(r"\d{1,3}", x.t)]
        if cnt:
            out.append((m.group(1), int(cnt[0].t)))
    return out


def _full_line_numbers(pages, line_list, log, iso):
    """The title block of each sheet has the short line number (005-PA-0223); the cover list has the
    full one (005-PA-0223-GCB2B01SN-02) and how many sheets follow for it. Use the full one where the
    short one matches."""
    start, lst = line_list
    owner, p = {}, start
    for line, n in lst:
        for k in range(n):
            owner[p + k] = line
        p += n
    used = 0
    for sheet, pg in pages:
        full, short = owner.get(sheet), pg["info"].get("line")
        if full and short and full != short and full.startswith(short):
            pg["info"]["line"] = full
            used += 1
    if used:
        log(iso, "", "INFO", "Full pipe line numbers taken from the list on the cover sheet (%d sheet(s))" % used)


# outside diameter (mm) in a description -> DN (GOST and ASME sizes)
OD_DN = {14: 10, 17: 10, 18: 15, 21: 15, 22: 15, 25: 20, 26: 20, 27: 20, 32: 25, 33: 25, 34: 25, 38: 32,
         42: 32, 45: 40, 48: 40, 57: 50, 60: 50, 73: 65, 76: 65, 88: 80, 89: 80, 108: 100, 114: 100,
         133: 125, 141: 125, 159: 150, 168: 150, 219: 200, 273: 250, 323: 300, 325: 300, 355: 350,
         377: 350, 406: 400, 426: 400, 457: 450, 508: 500, 530: 500, 610: 600, 630: 600, 711: 700,
         720: 700, 813: 800, 820: 800, 1016: 1000, 1020: 1000, 1219: 1200, 1220: 1200, 1321: 1300}


def validate_sheet(rows, pg, iso, sheet, log):
    """v0.9.4 cross-checks of one PGZ / LGN sheet (rules found by reviewing ~170 LGN drawings).
    They do not change any value; they mark cells yellow when the drawing contradicts itself."""
    def q(r):
        try:
            return float(r.get("M") or 0)
        except (TypeError, ValueError):
            return 0.0
    # 1) two nuts per stud (LGN: 84 of 84 sheets)
    studs = [r for r in rows if re.match(r"\s*stud", latin_part(str(r.get("H") or "")), re.I)]
    nuts = [r for r in rows if re.match(r"\s*nut\b", latin_part(str(r.get("H") or "")), re.I)]
    if studs and nuts:
        ns, nn = sum(q(r) for r in studs), sum(q(r) for r in nuts)
        if abs(nn - 2 * ns) > 1e-6:
            for r in nuts:
                r["_flags"]["M"] = "Nuts should be 2 per stud: studs %g, nuts %g on this sheet" % (ns, nn)
    # 2) DN against the pipe size in the description ('Pipe 108x5' is DN100)
    for r in rows:
        u = latin_part(str(r.get("H") or ""))
        m = re.search(r"(?<![\dA-Za-z.,-])(\d{2,4})(?:[.,]\d+)?\s*[xхХ×]\s*\d", u)
        dns = parse_dn(str(r.get("I") or ""))
        if not m or not dns or r["_flags"].get("I"):
            continue
        exp = OD_DN.get(int(m.group(1)))
        if exp and exp not in dns and not (r.get("J") == exp or r.get("K") == exp):
            r["_flags"]["I"] = "DN %s does not fit the size %s in the description (DN%d expected)" % (
                r.get("I"), m.group(0).strip(), exp)


def validate_drawing(pages, npages, iso, log, bundle=False):
    """v0.9.4: one pipe line per drawing and all its sheets in the PDF (LGN: 150 of 150 drawings)."""
    lines = {pg["info"].get("line") for _, pg in pages if pg["info"].get("line")}
    if len(lines) > 1 and not bundle:
        log(iso, "", "CHECK", "Different pipe line numbers on the sheets of one drawing: %s" % ", ".join(sorted(lines)))
    # cut marks <1>, <2> ... run through the whole drawing without gaps (LGN: 146 of 146 drawings;
    # PGZ starts again at <1> on every sheet); a gap means a cut-table row was not read (e.g. under the
    # AFC stamp)
    per = [[str(c[0]) for c in pg.get("cuts_own", pg["cuts"])] for _, pg in pages]
    per = [m for m in per if m]
    groups = per if len(per) > 1 and any(m and min(int(x) for x in m if x.isdigit() or x == "1") == 1
                                         for m in per[1:] if any(x.isdigit() for x in m)) else [sum(per, [])]
    for marks in groups:
        nums = [int(m) for m in marks if m.isdigit()]
        if not nums or "?" in marks:
            continue
        gaps = sorted(set(range(1, max(nums) + 1)) - set(nums))
        dups = sorted(n for n, k in Counter(nums).items() if k > 1)
        if gaps or dups:
            what = (["missing " + ", ".join("<%d>" % n for n in gaps[:10])] if gaps else []) + \
                   (["twice " + ", ".join("<%d>" % n for n in dups[:10])] if dups else [])
            log(iso, "", "CHECK", "Cut length marks %s - check the cut tables" % "; ".join(what))
    totals = [pg.get("total_sheets") for _, pg in pages if pg.get("total_sheets")]
    if totals and max(totals) > npages:
        log(iso, "", "CHECK", "Title block says %d sheets, the PDF has %d - sheets missing?" % (max(totals), npages))


def _title_sheet(objs, W, H):
    """PGZ / LGN title block: the number under 'SHEET' (next to STAGE / TOTAL SHEETS), if it is text."""
    hd = next((o for o in objs if o.t.upper() == "SHEET" and o.x0 > 0.8 * W and o.y0 > 0.8 * H), None)
    if hd is None:
        return None
    cx = (hd.x0 + hd.x1) / 2
    c = [o for o in objs if re.fullmatch(r"\d{1,3}", o.t) and 8 < o.y0 - hd.y0 < 40
         and abs((o.x0 + o.x1) / 2 - cx) < 20]
    return int(min(c, key=lambda o: o.y0).t) if c else None


def _title_total(objs, W, H):
    hd = next((o for o in objs if o.t.upper().startswith("TOTAL SHEETS") and o.x0 > 0.8 * W and o.y0 > 0.8 * H), None)
    if hd is None:
        return None
    cx = (hd.x0 + hd.x1) / 2
    c = [o for o in objs if re.fullmatch(r"\d{1,3}", o.t) and 8 < o.y0 - hd.y0 < 40 and abs((o.x0 + o.x1) / 2 - cx) < 25]
    return int(min(c, key=lambda o: o.y0).t) if c else None


def _sheet_numbers(pages):
    """v0.9.2, PGZ / LGN: column G = the sheet number printed in the title block. Drawing sets
    (12751-00 ISO-00001, 13240-11 ISO-00001: one PDF, many pipe lines) often do not have it as text;
    there the sheets of each pipe line are counted from 1, as in the hand-filled table."""
    lines = {pg["info"].get("line") for _, pg in pages if pg["info"].get("line")}
    out, prev, run = [], None, 0
    for sheet, pg in pages:
        line = pg["info"].get("line")
        run = run + 1 if line and line == prev else 1
        prev = line
        if pg.get("title_sheet"):
            out.append(pg["title_sheet"])
        elif pg.get("cols") and len(lines) > 1:
            out.append(run)
        else:
            out.append(sheet)
    return out


def _pipe_dns(pg):
    return {d for it in pg["items"] if re.match(r"\s*pipe\b", it["desc"], re.I)
            for d in (parse_dn(it["dn"]) or [])[:1]}


def _lend_cuts(pages, log, iso):
    """PGZ: when the material list runs over two sheets, the CUT PIPE LENGTH table is only on the
    last one. Give the pipes of the earlier sheet the cut lengths of the sheet that continues it
    (same line number), for the sizes that sheet has no pipe of its own."""
    for k, (sheet, pg) in enumerate(pages):
        need = {d for d in _pipe_dns(pg) if d not in {c[2] for c in pg["cuts"]}}
        if not need:
            continue
        for sheet2, pg2 in pages[k + 1:k + 4]:
            if pg2["info"].get("line") != pg["info"].get("line"):
                break
            give = [c for c in pg2["cuts"] if c[2] in need and c[2] not in _pipe_dns(pg2)]
            if give:
                pg.setdefault("cuts_own", list(pg["cuts"]))
                pg["cuts"] = pg["cuts"] + give
                log(iso, sheet, "INFO", "Cut lengths for DN %s taken from sheet %d (material list continues there)"
                    % ("/".join(str(d) for d in sorted({c[2] for c in give})), sheet2))
                break
            if pg2["cuts"]:
                break


def process_cad_pdf(path, rules, log, debug_dir=None):
    iso, rev = iso_from_filename(path)
    if not iso:
        log("?", "", "ERROR", "Cannot read ISO No. from file name: %s" % os.path.basename(path))
        return None, []
    rows, dbg, pages = [], [], []
    npages = page_count(path)
    cols, line_list = None, None
    for index in range(npages):
        emit("pages", done=index, total=npages)
        try:
            pg = read_cad_page(path, index, cols)
            pg["_path"], pg["_index"] = path, index
            cols = pg.get("cols")
        except NoMaterialList:
            # cover sheet of a drawing set ('Symbols / Условные обозначения', list of lines)
            log(iso, index + 1, "INFO", "Page %d has no material list (cover / legend sheet) - skipped"
                % (index + 1))
            try:
                lst = _line_list(cad_objects(path, index)[0])
            except Exception:
                lst = []
            if lst:
                line_list = (index + 2, lst)        # sheets of the listed lines start on the next page
            continue
        except Exception as e:
            log(iso, index + 1, "ERROR", "Page %d could not be read: %s" % (index + 1, e))
            continue
        dbg.append("=== page %d\n%s" % (index + 1, pg["dbg"]))
        pages.append((index + 1, pg))
    pages_idx = {id(pg): s - 1 for s, pg in pages}
    if line_list:
        _full_line_numbers(pages, line_list, log, iso)
    _lend_cuts(pages, log, iso)
    for (_, pg), sheet in zip(pages, _sheet_numbers(pages)):
        if pg.get("cols") and crossed_out(path, pages_idx[id(pg)]):
            log(iso, sheet, "WARN", "Sheet %s is crossed out (cancelled) - not filled" % sheet)
            continue
        new = cad_rows(iso, rev, sheet, pg, rules, log)
        if pg.get("cols"):
            validate_sheet(new, pg, iso, sheet, log)
        rows += new
    if pages and pages[0][1].get("cols"):
        validate_drawing(pages, npages, iso, log, bundle=bool(line_list))
    if debug_dir:
        os.makedirs(debug_dir, exist_ok=True)
        with open(os.path.join(debug_dir, "%s_text.txt" % iso), "w", encoding="utf-8") as f:
            f.write("CAD text layer (no OCR needed)\n" + "\n".join(dbg))
    return iso, rows


def standard_length(desc, dn, rules=None, iso=None):
    """v0.9.10: (values, source) of a fitting whose length follows from its standard alone:
    GOST elbows, ASME B16.9 elbows / equal tees / reducers / caps, IOT-type nipolets."""
    f = elbow_length(desc, dn)
    if f:
        return {"O": f}, "elbow 2 x R x tan(angle/2)"
    a = asme_b169(desc, dn)
    if a:
        return {a[0]: a[1]}, "ASME B16.9 table"
    if rules is not None and iso:
        f = nipolet_length(rules, iso, desc)
        if f:
            return {"O": f}, "nipolet: header OD / 2 + constant"
    return None, None


def apply_fitting_length(r, desc, rules, flags, required=True, dn=None, pg=None):
    """Column O (or P/Q for a reducing tee) from rules.xlsx > Fitting Length."""
    hit = rules.fitting_length(desc)
    if not hit:
        # v0.9.10: not learned yet - work it out from the standard where it is a standard size
        vals, src = standard_length(desc, dn, rules, r.get("A"))
        if vals:
            r.update(vals)
            dims = _sheet_dims(pg) if pg and pg.get("_path") else None
            if not (dims and all(length_on_drawing(v, dims) for v in vals.values())):
                flags[next(iter(vals))] = "Calculated from the standard (%s), not in the learned table - check" % src
            return
        if required:
            flags["O"] = "Fitting length unknown - add it to rules.xlsx > Fitting Length (or run learn)"
        return
    if hit[1] not in (None, ""):
        r["O"] = hit[1]
    if hit[4] not in (None, ""):
        r["P"] = hit[4]
    if hit[5] not in (None, ""):
        r["Q"] = hit[5]
    if hit[3] and "also seen" in str(hit[3]):
        key = "O" if "O" in r else "P"
        flags[key] = "Reference answers used more than one value: %s" % (
            str(hit[3]).split("also seen:")[-1].strip())
    elif not hit[2]:
        flags["O_unverified"] = True


# ---------------------------------------------------------------- v0.9.5 fitting lengths
def _page_xform(page):
    """v0.9.10: page coordinates -> the page as displayed (top-left origin), the same way
    cad_objects turns them, for /Rotate 0/90/180/270 (PGZ sheets are stored rotated).
    Returns (point(x, y), angle(dx, dy)) functions; angle in degrees, 0 = left to right."""
    rot = page.get_rotation() % 360
    bx0, by0, bx1, by1 = page.get_cropbox()
    pw, ph = bx1 - bx0, by1 - by0
    if rot == 90:
        pt, d = (lambda x, y: (y - by0, x - bx0)), (lambda dx, dy: (dy, dx))
    elif rot == 180:
        pt, d = (lambda x, y: (pw - (x - bx0), y - by0)), (lambda dx, dy: (-dx, dy))
    elif rot == 270:
        pt, d = (lambda x, y: (ph - (y - by0), pw - (x - bx0))), (lambda dx, dy: (-dy, -dx))
    else:
        pt, d = (lambda x, y: (x - bx0, ph - (y - by0))), (lambda dx, dy: (dx, -dy))

    def ang(dx, dy):
        DX, DY = d(dx, dy)
        return math.degrees(math.atan2(-DY, DX))
    return pt, ang


def cad_numbers(path, index, xmax):
    """Numbers written on the isometric part of a vector sheet (left of the material list), as a
    Counter {value: times}. Dimension texts along a sloped pipe are stored one character per
    object, so neighbouring digits with the same text angle are joined in reading order."""
    import pypdfium2 as pdfium
    import pypdfium2.raw as pdfium_c
    pdf = pdfium.PdfDocument(path)
    try:
        page = pdf[index]
        pt, ang = _page_xform(page)
        tp = page.get_textpage()
        singles, out = [], Counter()
        for o in page.get_objects(filter=[pdfium_c.FPDF_PAGEOBJ_TEXT], max_depth=8):
            o.textpage = tp
            try:
                t = (o.extract() or "").strip()
            except Exception:
                continue
            if not t:
                continue
            l, b, r, tt = o.get_bounds()
            x, y = pt((l + r) / 2, (b + tt) / 2)
            if x > xmax:
                continue
            if len(t) == 1 and t.isdigit():
                m = o.get_matrix()
                singles.append((t, x, y, ang(m.a, m.b)))
            elif re.fullmatch(r"\d{1,5}(?:[.,]\d)?", t):
                out[float(t.replace(",", "."))] += 1
    finally:
        pdf.close()
    grid = {}
    for j, c in enumerate(singles):
        grid.setdefault((int(c[1] // 7), int(c[2] // 7)), []).append(j)
    used = [False] * len(singles)
    for i in range(len(singles)):
        if used[i]:
            continue
        grp, used[i], k = [i], True, 0
        while k < len(grp):
            a = singles[grp[k]]
            for dx in (-1, 0, 1):
                for dy in (-1, 0, 1):
                    for j in grid.get((int(a[1] // 7) + dx, int(a[2] // 7) + dy), ()):
                        d = singles[j]
                        if not used[j] and abs(d[3] - a[3]) < 3 and math.hypot(d[1] - a[1], d[2] - a[2]) < 6.5:
                            used[j] = True
                            grp.append(j)
            k += 1
        ang = math.radians(singles[i][3])
        ux, uy = math.cos(ang), -math.sin(ang)
        txt = "".join(c[0] for c in sorted((singles[j] for j in grp), key=lambda c: c[1] * ux + c[2] * uy))
        if len(txt) <= 5:
            out[float(txt)] += 1
    return out


def _sheet_dims(pg):
    if pg.get("_dims") is None:
        try:
            xmax = (pg.get("cols") or {}).get("pos", 10 ** 6) - 8
            pg["_dims"] = cad_numbers(pg["_path"], pg["_index"], xmax)
        except Exception:
            pg["_dims"] = Counter()
    return pg["_dims"]


# ---------------------------------------------------------------- v0.9.10 ASME B16.9 sizes
ASME_NPS = {15: .5, 20: .75, 25: 1, 32: 1.25, 40: 1.5, 50: 2, 65: 2.5, 80: 3, 90: 3.5, 100: 4, 125: 5, 150: 6, 200: 8,
       250: 10, 300: 12, 350: 14, 400: 16, 450: 18, 500: 20, 550: 22, 600: 24, 650: 26, 700: 28, 750: 30, 800: 32,
       850: 34, 900: 36, 1050: 42, 1200: 48}
B169 = {  # NPS: (90 LR A, 90 SR A, 45 LR B, tee C, reducer H, cap E)
    .5: (38, None, 16, 25, None, 25), .75: (38, None, 19, 29, 38, 25), 1: (38, 25, 22, 38, 51, 38),
    1.25: (48, 32, 25, 48, 51, 38), 1.5: (57, 38, 29, 57, 64, 38), 2: (76, 51, 35, 64, 76, 38),
    2.5: (95, 64, 44, 76, 89, 38), 3: (114, 76, 51, 86, 89, 51), 3.5: (133, 89, 57, 95, 102, 64),
    4: (152, 102, 64, 105, 102, 64), 5: (190, 127, 79, 124, 127, 76), 6: (229, 152, 95, 143, 140, 89),
    8: (305, 203, 127, 178, 152, 102), 10: (381, 254, 159, 216, 178, 127), 12: (457, 305, 190, 254, 203, 152),
    14: (533, 356, 222, 279, 330, 165), 16: (610, 406, 254, 305, 356, 178), 18: (686, 457, 286, 343, 381, 203),
    20: (762, 508, 318, 381, 508, 229), 22: (838, 559, 343, 419, 508, 254), 24: (914, 610, 381, 432, 508, 267),
    26: (991, 660, 406, 495, 610, 267), 28: (1067, 711, 438, 521, 610, 267), 30: (1143, 762, 470, 559, 610, 267),
    32: (1219, 813, 502, 597, 610, 267), 34: (1295, 864, 533, 635, 610, 267), 36: (1372, 914, 565, 673, 610, 267),
    42: (1600, 1067, 660, 762, 610, 305), 48: (1829, 1219, 749, 889, 610, 343)}
OD_ASME_NPS = {21.3: .5, 26.7: .75, 33.4: 1, 42.2: 1.25, 48.3: 1.5, 60.3: 2, 73: 2.5, 88.9: 3, 101.6: 3.5, 114.3: 4,
          141.3: 5, 168.3: 6, 219.1: 8, 273: 10, 323.8: 12, 355.6: 14, 406.4: 16, 457: 18, 508: 20, 559: 22,
          610: 24, 660: 26, 711: 28, 762: 30, 813: 32, 864: 34, 914: 36, 1067: 42, 1219: 48}

def _asme_nps(e, dn):
    m = re.search(r'(\d{2,4}(?:\.\d)?)\s*X\s*\d', e.replace(',', '.'))
    if m:
        od = float(m.group(1))
        best = min(OD_ASME_NPS, key=lambda o: abs(o - od))
        if abs(best - od) <= 1.5: return OD_ASME_NPS[best]
    return ASME_NPS.get(dn)


def asme_b169(desc, dn):
    """v0.9.10: (column, length) of an ASME B16.9 butt-weld fitting from the standard's tables, as the
    FGG answer keys write them: elbow '=A*2' (90° long / short radius, 45°), equal tee '=C*3',
    concentric / eccentric reducer H, cap E. 266 of 286 FGG key rows agree. None otherwise."""
    e = latin_part(desc).upper()
    inch = re.search(r'(\d{1,2})"', e)
    if '16.9' not in e and not (inch and re.search(r'\b(SR|LR|SCH|STD|XS|XXS)\b', e)): return None
    if not dn and inch:
        dn = next((d for d, n in ASME_NPS.items() if n == float(inch.group(1))), None)
    if 'B16.11' in e or re.search(r'\bSW\b|SOCKET|THREAD', e): return None
    n = _asme_nps(e, dn)
    if n not in B169: return None
    A90L, A90S, B45, C, H, E = B169[n]
    if re.search(r'\bELBOW', e):
        if re.search(r'45\s*E|45\s*DEG|45°|\b(LR|SR)\s+45\b', e): return ('O', '=%d*2' % B45) if B45 else None
        if re.search(r'90\s*ES|\bSR\b|SHORT', e): return ('O', '=%d*2' % A90S) if A90S else None
        return ('O', '=%d*2' % A90L)
    if re.search(r'\bTEE', e) and re.search(r'\bTEE\s+T?S\b|EQUAL|STRAIGHT', e): return ('O', '=%d*3' % C)
    if re.search(r'\bREDUCER', e) and H: return ('O', H)
    if re.search(r'\bCAP\b', e): return ('O', E)
    return None


_ELBOW_WORD = re.compile(r"\b(?:ELBOW|BEND)\b")


def elbow_length(desc, dn):
    """Butt-weld elbow length as the answer keys write it: both centre-to-end dimensions,
    2 x R x tan(angle / 2). R = the 'kDN' written in the name (OKSH 90°-1,5DN-...), 1 x DN for
    GOST 30753, otherwise 1.5 x DN (GOST 17375). PVC / union / solvent elbows: None."""
    e = latin_part(desc).upper()
    if not dn or not _ELBOW_WORD.search(e) or re.search(r"PVC|SOLVENT|SOLWENT|UNION|THREAD|\bPE\b|\bPP\b", e):
        return None
    if re.search(r"ASME|MSS|B16\.|\d\s*\"", e):
        return None                                  # ASME / inch elbows: asme_b169 (v0.9.10)
    m = re.search(r"(?:ELBOW|BEND)\D{0,12}?(\d{2,3})\s*°?\s*-", e) or re.search(r"(\d{2,3})\s*°", e)
    ang = float(m.group(1)) if m and 10 <= float(m.group(1)) <= 180 else 90.0
    k = re.search(r"-(\d+(?:[.,]\d)?)\s*DN", e)
    k = float(k.group(1).replace(",", ".")) if k else (1.0 if "30753" in e else 1.5)
    R = k * dn
    if k == 1.5 and dn in (15, 20):
        R = {15: 28, 20: 38}[dn]          # GOST 17375 small sizes (PGZ key: 21.3x4 = 56, 25x4 = 76)
    return int(round(2 * R * math.tan(math.radians(ang / 2))))


def _len_parts(v):
    """'=144+40' -> [144, 40]; '=130*2+110' -> [130, 110] (the '*2' is a count, not a length)."""
    s = str(v)
    if not s.startswith("="):
        return []
    return [float(x) for x in re.findall(r"(?<![*\d.])(\d+(?:\.\d+)?)", s)]


def length_on_drawing(v, dims):
    """True when the length (or its halves / thirds - elbows and equal tees are dimensioned
    per leg - or every term of a '=144+40' formula) is written on the sheet."""
    x = eval_len(v)
    if x is None or not dims:
        return False
    parts = _len_parts(v)
    if parts and all(p in dims for p in parts):
        return True
    # 1 mm slack: hand-filled keys write 749 for an elbow dimensioned 375 + 375
    return any(abs(x - k * d) <= 1.0 for d in dims for k in (1, 2, 3))


# ---------------------------------------------------------------- v0.9.5 step 2: lengths read on the drawing
def cad_geometry(path, index):
    """Texts (one-character objects joined along their text angle), line segments and small closed
    boxes of a vector sheet, in the coordinates cad_numbers uses (top-left origin)."""
    import ctypes
    import pypdfium2 as pdfium
    import pypdfium2.raw as pdfium_c
    pdf = pdfium.PdfDocument(path)
    try:
        page = pdf[index]
        pt, ang = _page_xform(page)
        tp = page.get_textpage()
        chars, texts, segs, boxes = [], [], [], []
        for o in page.get_objects(filter=[pdfium_c.FPDF_PAGEOBJ_TEXT, pdfium_c.FPDF_PAGEOBJ_PATH], max_depth=8):
            if o.type == pdfium_c.FPDF_PAGEOBJ_TEXT:
                o.textpage = tp
                try:
                    t = (o.extract() or "").strip()
                except Exception:
                    continue
                if not t:
                    continue
                l, b, rr, tt = o.get_bounds()
                m = o.get_matrix()
                item = (t,) + pt((l + rr) / 2, (b + tt) / 2) + (ang(m.a, m.b),)
                (chars if len(t) == 1 else texts).append(item)
                continue
            n = pdfium_c.FPDFPath_CountSegments(o.raw)
            if n <= 0 or n > 400:
                continue
            m = o.get_matrix()
            pts = []
            for i in range(n):
                seg = pdfium_c.FPDFPath_GetPathSegment(o.raw, i)
                x, y = ctypes.c_float(), ctypes.c_float()
                pdfium_c.FPDFPathSegment_GetPoint(seg, ctypes.byref(x), ctypes.byref(y))
                pts.append(pt(m.a * x.value + m.c * y.value + m.e, m.b * x.value + m.d * y.value + m.f)
                           + (pdfium_c.FPDFPathSegment_GetType(seg),))
            cur = None
            for x, y, typ in pts:
                if typ == pdfium_c.FPDF_SEGMENT_MOVETO:
                    cur = (x, y)
                    continue
                if cur is not None:
                    segs.append((cur[0], cur[1], x, y))
                cur = (x, y)
            l, b, rr, tt = o.get_bounds()
            (x0, y0), (x1, y1) = pt(l, b), pt(rr, tt)
            x0, x1, y0, y1 = min(x0, x1), max(x0, x1), min(y0, y1), max(y0, y1)
            if 4 <= x1 - x0 <= 30 and 6 <= y1 - y0 <= 14 and 4 <= n <= 6:
                boxes.append((x0, y0, x1, y1))
    finally:
        pdf.close()
    grid = {}
    for j, c in enumerate(chars):
        grid.setdefault((int(c[1] // 7), int(c[2] // 7)), []).append(j)
    used = [False] * len(chars)
    for i in range(len(chars)):
        if used[i]:
            continue
        grp, used[i], k = [i], True, 0
        while k < len(grp):
            a = chars[grp[k]]
            for dx in (-1, 0, 1):
                for dy in (-1, 0, 1):
                    for j in grid.get((int(a[1] // 7) + dx, int(a[2] // 7) + dy), ()):
                        d = chars[j]
                        if not used[j] and abs(d[3] - a[3]) < 3 and math.hypot(d[1] - a[1], d[2] - a[2]) < 6.5:
                            used[j] = True
                            grp.append(j)
            k += 1
        ang = math.radians(chars[i][3])
        ux, uy = math.cos(ang), -math.sin(ang)
        g = sorted((chars[j] for j in grp), key=lambda c: c[1] * ux + c[2] * uy)
        texts.append(("".join(c[0] for c in g), sum(c[1] for c in g) / len(g), sum(c[2] for c in g) / len(g),
                      chars[i][3]))
    return {"texts": texts, "segs": segs, "boxes": boxes}


def _sheet_geom(pg):
    if pg.get("_geom") is None:
        try:
            pg["_geom"] = cad_geometry(pg["_path"], pg["_index"])
        except Exception:
            pg["_geom"] = {"texts": [], "segs": [], "boxes": []}
    return pg["_geom"]


def _tag_box(x, y, segs, boxes):
    """the box drawn around a POS number: a closed path, or lines on three / four sides."""
    for b in boxes:
        if b[0] <= x <= b[2] and b[1] <= y <= b[3]:
            return b
    top = bot = left = right = None
    for x0, y0, x1, y1 in segs:
        if abs(y0 - y1) < 0.3 and min(x0, x1) - 1 <= x <= max(x0, x1) + 1:
            dy = y0 - y
            if -10 <= dy <= -2 and (top is None or dy > top[0]):
                top = (dy, min(x0, x1), max(x0, x1), y0)
            if 2 <= dy <= 10 and (bot is None or dy < bot[0]):
                bot = (dy, y0)
        if abs(x0 - x1) < 0.3 and min(y0, y1) - 1 <= y <= max(y0, y1) + 1:
            dx = x0 - x
            if -10 <= dx <= -1.5 and (left is None or dx > left[0]):
                left = (dx, x0)
            if 1.5 <= dx <= 40 and (right is None or dx < right[0]):
                right = (dx, x0)
    if top and left and (bot or right):
        return (left[1], top[3], right[1] if right else top[2], bot[1] if bot else y + 5)
    return None


def _near_box(px, py, bb, tol):
    return bb[0] - tol <= px <= bb[2] + tol and bb[1] - tol <= py <= bb[3] + tol


def _leader_end(b, texts, segs):
    """Where the leader of a POS tag ends (= on the fitting). The leader starts at the box or at
    the label written under it ('DN 100x25', 'Elev. …'); a kinked leader is followed."""
    x0, y0, x1, y1 = b
    for t, x, y, a in texts:          # label block under the tag
        if abs(a) < 1 and b[0] - 3 <= x <= b[0] + 120 and b[3] - 1 <= y <= b[3] + 22 and len(t) >= 2:
            x0, x1 = min(x0, x - 1.3 * len(t) - 2), max(x1, x + 1.3 * len(t) + 2)
            y1 = max(y1, y + 4)
    blk, best = (x0, y0, x1, y1), None
    for s in segs:
        L = math.hypot(s[2] - s[0], s[3] - s[1])
        if L < 8 or L > 500:
            continue
        for ax, ay, bx, by in ((s[0], s[1], s[2], s[3]), (s[2], s[3], s[0], s[1])):
            if _near_box(ax, ay, blk, 3.5) and not _near_box(bx, by, blk, 1) and (best is None or L > best[2]):
                best = (bx, by, L, ax, ay)
    if not best:
        return None
    ex, ey = best[0], best[1]
    for _ in range(2):
        nxt = None
        for s in segs:
            L = math.hypot(s[2] - s[0], s[3] - s[1])
            if L < 8 or L > 500:
                continue
            for ax, ay, bx, by in ((s[0], s[1], s[2], s[3]), (s[2], s[3], s[0], s[1])):
                if math.hypot(ax - ex, ay - ey) < 0.8 and \
                        math.hypot(bx - best[3], by - best[4]) > math.hypot(ex - best[3], ey - best[4]) + 5 and \
                        (nxt is None or L > nxt[2]):
                    nxt = (bx, by, L)
        if not nxt:
            break
        ex, ey = nxt[0], nxt[1]
    return ex, ey


def _number_anchor(t, x, y, a, segs):
    """A horizontal dimension number placed away from its dimension line has its own leader -
    the number belongs where that leader ends."""
    if abs(a) > 1:
        return x, y
    hw, hh, best = 2.7 * len(t) + 2.5, 6.1, None
    for s in segs:
        L = math.hypot(s[2] - s[0], s[3] - s[1])
        if L < 6 or L > 250:
            continue
        for ax, ay, bx, by in ((s[0], s[1], s[2], s[3]), (s[2], s[3], s[0], s[1])):
            if abs(ax - x) <= hw and abs(ay - y) <= hh and not (abs(bx - x) <= hw and abs(by - y) <= hh) and \
                    (best is None or L > best[2]):
                best = (bx, by, L)
    return (best[0], best[1]) if best else (x, y)


def numbers_at_pos(pg, pos, radius=90, counts=None):
    """{number: distance} of the dimension numbers around the fitting a POS tag points to.
    counts (a dict), if given, receives how many times each number is written there."""
    geo = _sheet_geom(pg)
    texts, segs, boxes = geo["texts"], geo["segs"], geo["boxes"]
    xmax = (pg.get("cols") or {}).get("pos", 10 ** 6) - 8
    tags = []
    for t, x, y, a in texts:
        if t == str(pos) and x < xmax and abs(a) < 1:
            b = _tag_box(x, y, segs, boxes)
            if b and b not in tags:
                tags.append(b)
    out = {}
    for b in tags:
        F = _leader_end(b, texts, segs)
        if not F:
            continue
        for t, x, y, a in texts:
            if not re.fullmatch(r"[1-9]\d{0,4}", t) or x >= xmax or float(t) < 3:
                continue
            if any(bb[0] - 0.5 <= x <= bb[2] + 0.5 and bb[1] - 0.5 <= y <= bb[3] + 0.5 for bb in tags + boxes):
                continue                     # a POS / tag number in its box, not a dimension
            ax, ay = _number_anchor(t, x, y, a, segs)
            d = min(math.hypot(ax - F[0], ay - F[1]), 1.5 * math.hypot(x - F[0], y - F[1]))
            if d < radius:
                out[float(t)] = min(out.get(float(t), 10 ** 9), d)
                if counts is not None:
                    counts.setdefault(float(t), set()).add((round(x), round(y)))
    if counts is not None:
        for k in list(counts):
            counts[k] = len(counts[k])
    return out


DRAWING_READ_TYPES = ("REDUCER", "NIPOLET", "FLANGE")
_TYPE_WORDS = ("NIPOLET", "SPECTACLE", "BLIND", "REDUCER", "TEE", "ELBOW", "FLANGE", "CAP", "COUPL", "NIPPLE",
               "BUSH", "GASKET", "FLOW", "HEAD", "JOINT")


def _fit_type(desc):
    e = latin_part(desc).upper()
    # at the start of a word ('TEE', not the TEE in 'STEEL')
    return next((w for w in _TYPE_WORDS if re.search(r"\b" + w, e)), None)


def _fit_shape(desc):
    e = re.sub(r"\b(GOST|TU|OST|ATK|ASME)\b", " ", latin_part(desc).upper())
    return re.sub(r"[^A-Z]+", " ", e).strip()


def length_prior(rules, iso, desc, dn):
    """What a fitting of this type and size probably measures in this family: the most alike learned
    item (same type word, similar name, nearest DN), scaled to this DN (flanges / blinds / gaskets
    not scaled). Returns (value, column) or None."""
    typ = _fit_type(desc)
    if not typ or not dn:
        return None
    sh, best = _fit_shape(desc), None
    for row in rules.prior_rows(iso, lambda rw: _fit_type(rw[-1]) == typ and rules.fit_len_dn.get(id(rw))):
        dns = rules.fit_len_dn.get(id(row))
        col = "O" if row[1] not in (None, "") else "P" if row[4] not in (None, "") else None
        v = eval_len(row[1] if col == "O" else row[4]) if col else None
        if not v:
            continue
        score = difflib.SequenceMatcher(None, sh, _fit_shape(row[-1])).ratio() * 2 - abs(math.log(dns[0] / float(dn)))
        if best is None or score > best[0]:
            scale = 1.0 if typ in ("SPECTACLE", "BLIND", "FLANGE", "GASKET") else float(dn) / dns[0]
            best = (score, v * scale, col)
    return (best[1], best[2]) if best else None


def read_length_on_drawing(pg, pos, desc, expect):
    """Pick the fitting length among the numbers around POS on the sheet: near the fitting and
    close to what such a fitting usually measures (half / third: elbow legs, equal tees; nipolets
    are sometimes written as two numbers added up). Returns (value as written, numbers seen)."""
    c = numbers_at_pos(pg, pos)
    if not c or not expect:
        return None, c
    typ = _fit_type(desc)
    opts = []
    for v, d in c.items():
        for k in ((1, 2, 3) if typ in ("ELBOW", "TEE") else (1,)):      # elbows: dimensioned per leg
            opts.append((d / 180.0 + abs(math.log(v * k / expect)) + 0.15 * (k - 1), v * k,
                         _num_or_text(v) if k == 1 else "=%s*%d" % (_num_or_text(v), k)))
    near = sorted(c, key=c.get)[:5]
    if typ == "NIPOLET":
        for x in near[:2]:
            for y in near:
                if x != y:
                    opts.append(((c[x] + c[y]) / 360.0 + 0.2 + abs(math.log((x + y) / expect)), x + y,
                                 "=%s+%s" % (_num_or_text(x), _num_or_text(y))))
    s, val, txt = min(opts)
    return txt, c


# ---------------------------------------------------------------- v0.9.7 size models (reducers, nipolets, tees)
DN_OD = {10: 14, 15: 18, 20: 25, 25: 32, 32: 38, 40: 45, 50: 57, 65: 76, 80: 89, 100: 108, 125: 133, 150: 159,
         200: 219, 250: 273, 300: 325, 350: 377, 400: 426, 450: 457, 500: 530, 600: 630, 700: 720, 800: 820,
         1000: 1020, 1200: 1220}


def _eng(desc):
    return re.sub(r"\s+", " ", latin_part(desc).upper()).strip()


def fitting_ods(desc):
    """Outside diameters in the item name, in order: 'REDUCER P C-108x6-57x4' -> [108, 57],
    'REDUCER C 1020X820X10' -> [1020, 820], 'TEE TSH 377x10/6-273x7/6' -> [377, 273]."""
    e = _eng(desc).split("(")[0].replace(",", ".")
    e = re.sub(r"\b(GOST|TU|OST|ATK)\s*[\d.\- ]+", " ", e)
    out = []
    for m in re.finditer(r"(\d{2,4})(?:X(\d{2,4}))?X\d+(?:\.\d+)?(?:/\d+(?:\.\d+)?)?", e):
        out.append(float(m.group(1)))
        if m.group(2) and 8 < float(m.group(2)) < float(m.group(1)):
            out.append(float(m.group(2)))
    if not out:                                   # PVC parts: 'OD 63-25', 'OD 110'
        m = re.search(r"\bOD\s*(\d{2,4})(?:\s*[-X]\s*(\d{2,4}))?", e)
        if m:
            out = [float(m.group(1))] + ([float(m.group(2))] if m.group(2) else [])
    return out


def _nipolet_header_od(desc):
    m = re.search(r"NIPOLET\s+SHUT\s+(\d{2,4})-", _eng(desc))
    return float(DN_OD.get(int(m.group(1)), int(m.group(1)))) if m else None


def nipolet_length(rules, iso, desc):
    """IOT nipolets (ШУТ, TU 3600-010-88626180) measure half the header OD + a constant (90 mm in the
    key: SHUT 1000-… = 1020/2 + 90 = 600, SHUT 100-… = 108/2 + 90 = 144). The constant is learned per
    family from its key rows (median); at least three rows are needed."""
    h = _nipolet_header_od(desc)
    if not h:
        return None
    cs = []
    ok = lambda rw: _nipolet_header_od(rw[-1]) and rw[1] not in (None, "") and not str(rw[1]).startswith("=")
    for row in rules.prior_rows(iso, ok, need=3):
        cs.append(float(row[1]) - _nipolet_header_od(row[-1]) / 2)
    if len(cs) < 3:
        return None
    cs.sort()
    c = cs[len(cs) // 2]
    return int(math.floor(h / 2 + c + 0.5))


def _reducer_group(desc):
    e = _eng(desc)
    for pat, g in ((r"PKT", "PKT"), (r"REDUCER\s+P[KE]SH", "PKSH"), (r"REDUCER\s+PKS\b", "PKS"),
                   (r"REDUCER\s+P\s+[CE]\b", "P"), (r"OST 36-22", "OST")):
        if re.search(pat, e):
            return g
    return "other"


def _tee_group(desc):
    e = _eng(desc)
    for pat, g in ((r"PVC", "PVC"), (r"UNION TEE", "TT"), (r"TSHS", "TSHS"), (r"TEE\s+TSH\b", "TSH"),
                   (r"TEE\s+TS\b", "TS"), (r"TEE\s+P\b", "P"), (r"OST 36-24|COUNTERBORE", "OST")):
        if re.search(pat, e):
            return g
    return "other"


REDUCER_TRUST = 0.5


def reducer_prior(rules, iso, desc):
    """Length of a reducer of this kind and size: the learned reducer of the same kind (GOST 17378
    'P C/E', 'PKSH/PESH', 'PKS', 'PKT-2', OST 36-22) with the nearest big / small OD; scaled a little
    when the big OD differs. Other kinds of reducer are used only when this kind was never learned."""
    o = fitting_ods(desc)
    if len(o) < 2:
        return None
    D, d, g = o[0], o[1], _reducer_group(desc)
    pts = []
    ok = lambda rw: "REDUCER" in _eng(rw[-1]) and rw[1] not in (None, "") and not str(rw[1]).startswith("=")
    for row in rules.prior_rows(iso, ok):
        oo = fitting_ods(row[-1])
        if len(oo) >= 2:
            pts.append((_reducer_group(row[-1]) == g, oo[0], oo[1], float(row[1])))
    same = [p for p in pts if p[0]] or pts
    if not same:
        return None
    best = min(same, key=lambda p: abs(math.log(p[1] / D)) * 3 + abs(math.log(p[2] / d)))
    v = best[3] * ((D / best[1]) ** 0.7 if abs(best[1] - D) > 1 else 1.0)
    # how much to trust it: same kind and both sizes -> 1, otherwise less
    trust = 1.0 if best[0] and abs(best[1] - D) <= 1 and abs(best[2] - d) <= 1 else REDUCER_TRUST
    return v, trust


def _tee_parts(row):
    """(run OD, branch OD, run half a, branch b) of a learned tee row: P / Q = 2a / b, '=a*2+b',
    '=a*3' or one total of an equal tee (= 3a)."""
    o = fitting_ods(row[-1])
    if not o:
        return None
    D = o[0]
    d = o[1] if len(o) > 1 and o[1] < o[0] else D
    if row[4] not in (None, "") and row[5] not in (None, ""):
        p, q = eval_len(row[4]), eval_len(row[5])
        return (D, d, p / 2, q) if p and q else None
    s = str(row[1])
    m = re.fullmatch(r"=(\d+(?:\.\d+)?)\*2\+(\d+(?:\.\d+)?)", s)
    if m:
        return D, d, float(m.group(1)), float(m.group(2))
    m = re.fullmatch(r"=(\d+(?:\.\d+)?)\*3", s)
    if m:
        return D, d, float(m.group(1)), float(m.group(1))
    v = eval_len(row[1])
    return (D, d, v / 3, v / 3) if v and d == D else None


def tee_prior(rules, iso, desc):
    """(run half a, branch b, reducing?) of a tee of this kind and size from the learned tees."""
    o = fitting_ods(desc)
    if not o:
        return None
    D = o[0]
    d = o[1] if len(o) > 1 and o[1] < o[0] else D
    g = _tee_group(desc)
    pts = []
    ok = lambda rw: re.search(r"\bTEE\b", _eng(rw[-1])) and _tee_parts(rw)
    for row in rules.prior_rows(iso, ok):
        pts.append((_tee_group(row[-1]) == g,) + _tee_parts(row))
    same = [p for p in pts if p[0]] or pts
    if not same:
        return None
    pa = min(same, key=lambda p: abs(math.log(p[1] / D)))
    ra = pa[3] * ((D / pa[1]) ** 0.8 if abs(pa[1] - D) > 1 else 1.0)
    # the branch length of these tees follows the run size much more than the branch size
    # (TSH 426x159 and 426x325: both 250)
    pb = min(same, key=lambda p: abs(math.log(p[1] / D)) * 3 + abs(math.log(p[2] / d)))
    rb = pb[4] * ((D / pb[1]) ** 0.8 if abs(pb[1] - D) > 1 else 1.0)
    if d == D:
        rb = min(rb, ra)
    return ra, rb, d != D


TEE_TWICE = 0.3


def read_tee_on_drawing(pg, pos, prior):
    """Run half a and branch b of a tee among the numbers around POS: both near the tee, close to
    what this tee usually measures, the run normally not shorter than the branch."""
    n = {}
    c = numbers_at_pos(pg, pos, counts=n)
    if not c:
        return None, c
    ra, rb, reducing = prior
    if not reducing:
        # equal tee (CC, 2026-11-17): only the run length is needed - O = run x 3
        x = min(c, key=lambda v: c[v] / 180.0 + abs(math.log(v / ra)) - (TEE_TWICE if n.get(v, 0) >= 2 else 0))
        return (x, x), c
    near = sorted(c, key=c.get)[:8]
    best = None
    for x in near:
        for y in near:
            cost = (c[x] + c[y]) / 180.0 + abs(math.log(x / ra)) + abs(math.log(y / rb))
            if y > x * 1.02:
                cost += 0.3
            if x == y and abs(ra - rb) > 0.08 * ra:
                cost += 0.2
            if n.get(x, 0) >= 2:                  # the run is dimensioned on both sides of the tee
                cost -= TEE_TWICE
            if best is None or cost < best[0]:
                best = (cost, x, y)
    return (best[1], best[2]), c


def apply_fitting_length_v095(r, desc, dn, rules, flags, pg):
    """Columns O / P / Q for families with 'Fitting lengths from drawing' = Y (IOT, v0.9.5).
    1. rules.xlsx > Fitting Length (same item; same item without its standard),
    2. elbow formula, 3. the same type and sizes in another material / standard.
    Each value is looked for among the dimensions written on this sheet; values that are not
    found there are marked for checking."""
    iso = r.get("A")
    if re.search(r"\b(STUD|NUT|BOLT|WASHER|SCREW)", latin_part(desc).upper()):
        return                                       # fasteners have no material length (v0.9.9)
    take = lambda h: {c: h[i] for c, i in (("O", 1), ("P", 4), ("Q", 5)) if h[i] not in (None, "")}
    hit, src = rules.family_length(iso, desc), "rules.xlsx > Fitting Length"
    vals = take(hit) if hit else {}
    if not vals:
        f = elbow_length(desc, dn)
        if f:
            vals, src = {"O": f}, "elbow 2 x R x tan(angle/2)"
    if not vals:
        a = asme_b169(desc, dn)                       # v0.9.10
        if a:
            vals, src = {a[0]: a[1]}, "ASME B16.9 table"
    if not vals:
        f = nipolet_length(rules, iso, desc)          # v0.9.7
        if f:
            vals, src = {"O": f}, "nipolet: header OD / 2 + the family's constant"
    for how in ("nostd", "core"):
        if not vals:
            hit = rules.family_length(iso, desc, how)
            if hit:
                vals = take(hit)
                src = "same size in rules.xlsx: %s" % latin_part(re.sub(r"\s+", " ", str(hit[-1]))).strip()[:70]
    if not vals:
        hit = rules.fitting_length(desc)              # v0.9.9: rows learned without a family
        if hit:
            vals, src = take(hit), "rules.xlsx > Fitting Length (learned from another table)"
    if not vals:
        # step 2: read it on the drawing
        pos = r.get("_pos")
        typ = _fit_type(desc)
        ok_pg = bool(pos and pg.get("_path"))
        listed = lambda seen: ", ".join(str(_num_or_text(v)) for v in sorted(seen, key=seen.get)[:6])
        if ok_pg and typ == "TEE":                    # v0.9.7: run and branch
            pri = tee_prior(rules, iso, desc)
            if pri:
                got, seen = read_tee_on_drawing(pg, pos, pri)
                if got:
                    a, b = (_num_or_text(v) for v in got)
                    if pri[2]:
                        r["P"], r["Q"] = _num_or_text(2 * got[0]), b
                        flags["P"] = ("Read on the drawing at POS %s: run %s x 2, branch %s (numbers there: %s)"
                                      " - please check" % (pos, a, b, listed(seen)))
                    else:
                        r["O"] = "=%s*3" % a
                        flags["O"] = ("Read on the drawing at POS %s: equal tee, run %s x 3 (numbers there: %s)"
                                      " - please check" % (pos, a, listed(seen)))
                    return
        if ok_pg and typ == "REDUCER":               # v0.9.7: prior from the reducers of the same kind
            pri = reducer_prior(rules, iso, desc)
            if pri:
                pri, trust = pri
                seen = numbers_at_pos(pg, pos)
                if seen:
                    on = [v for v in seen if abs(v - pri) <= 1]
                    if on:                               # the expected length is written at the fitting
                        r["O"] = _num_or_text(on[0])
                        return
                    v = min(seen, key=lambda v: seen[v] / 180.0 + trust * abs(math.log(v / pri)))
                    r["O"] = _num_or_text(v)
                    flags["O"] = "Read on the drawing at POS %s (numbers there: %s) - please check" % (pos, listed(seen))
                    return
        pri = length_prior(rules, iso, desc, dn) if ok_pg else None
        if pri and typ in DRAWING_READ_TYPES:
            txt, seen = read_length_on_drawing(pg, pos, desc, pri[0])
            if txt is not None:
                r["O"] = txt
                flags["O"] = "Read on the drawing at POS %s (numbers there: %s) - please check" % (pos, listed(seen))
                return
        seen = numbers_at_pos(pg, pos) if ok_pg else {}
        hint = length_prior(rules, iso, desc, dn) if typ else None
        flags["O"] = "Fitting length unknown - read it on the drawing at POS %s%s%s" % (
            pos or "?", (" (numbers there: %s)" % listed(seen)) if seen else "",
            ("; similar learned items measure about %s" % _num_or_text(round(hint[0]))) if hint else "")
        return
    if _fit_type(desc) == "TEE" and "O" in vals:
        o = fitting_ods(desc)
        m = re.fullmatch(r"=(\d+(?:\.\d+)?)\*2\+\d+(?:\.\d+)?", str(vals["O"]))
        if m and (len(o) < 2 or o[1] >= o[0]):
            vals["O"] = "=%s*3" % m.group(1)        # equal tee: run x 3 (CC, 2026-11-17)
    r.update(vals)
    dims = _sheet_dims(pg) if pg.get("_path") else Counter()
    missing = [c for c, v in vals.items() if not length_on_drawing(v, dims)]
    if missing:
        flags[missing[0]] = "Length %s (%s) is not among the dimensions on this sheet - check the drawing at POS %s" % (
            " / ".join(str(vals[c]) for c in ("O", "P", "Q") if c in vals), src, r.get("_pos") or "?")


def cad_rows(iso, rev, sheet, pg, rules, log):
    info = pg["info"]
    base = {"A": iso, "B": _int_or(iso.split("-")[3] if iso.count("-") >= 3 else None), "C": rev,
            "D": info["line"], "G": sheet, "V": info["area"], "PAINT": info["paint"],
            "TRACE": info["trace"], "INSUL": info["insul"]}
    if not info["line"]:
        log(iso, sheet, "CHECK", "Line number not found in the title block")
    by_dn = OrderedDict()
    for _, L, dn in pg["cuts"]:
        by_dn.setdefault(dn, []).append(L)
    out = []

    def add(desc, dn_text, qty, code=None, unit=None, piece_len=None, as_text=None, pos=None):
        flags = {}
        cat, zh, ver = rules.classify(as_text or desc, None, "CAD", iso)
        dec = dn_text and re.fullmatch(r"\s*\d+[.,]\d+\s*", dn_text)
        dns = parse_dn(dn_text) if dn_text and not dec else None
        r = dict(base)
        r.update({"E": cat, "F": zh, "H": desc, "L": code or None, "M": qty, "_pos": pos})
        if not cat:
            flags["E"] = "No keyword in rules.xlsx > Categories matches this description"
        elif not ver:
            flags["F_unverified"] = True
        main_dn = None
        two = re.search(r"\(DN\s*(\d+)\s*-\s*DN\s*(\d+)\)", latin_part(desc))
        if two and dns and len(dns) == 1 and two.group(1) != two.group(2):
            r["J"], r["K"] = int(two.group(1)), int(two.group(2))   # reducer / reducing tee sizes
        if dns:
            if len(dns) == 2 and dns[0] != dns[1]:
                r["I"] = "%dX%d" % (dns[0], dns[1])
                r["J"], r["K"] = dns[0], dns[1]
            else:
                r["I"] = dns[0]
            main_dn = dns[0]
        elif dec:
            r["I"] = float(dn_text.replace(",", "."))   # bolt diameter in mm, e.g. 15.875 (5/8")
        elif dn_text and re.fullmatch(r"\d{1,3}", dn_text.strip()):
            r["I"] = int(dn_text)               # bolt size (M12) is written in the DN column
        if qty is None:
            flags["M"] = "QTY not found"
        if cat == "Pipe":
            if piece_len:
                r["N"] = piece_len
            elif main_dn:
                L = by_dn.get(main_dn, [])
                if L:
                    r["N"] = L[0] if len(L) == 1 else "=" + "+".join(map(str, L))
                    if qty is not None and not qty_ok(L, qty):
                        flags["N"] = "Cut lengths sum %d mm but material list says %s m" % (sum(L), qty)
                else:
                    flags["N"] = "No cut pipe length found for DN%s" % main_dn
        elif cat == "Valve":
            hit = rules.valve_length(desc)
            if hit:
                r["S"] = hit[1]
                flags["S"] = "Typical value from reference answers - check the drawing." + (
                    " Also seen: %s" % hit[2] if hit[2] else "")
        elif cat in ("Fittings", "Flanges", "Gasket"):
            if opt.get("fit_len"):
                apply_fitting_length_v095(r, desc, main_dn, rules, flags, pg)
            else:
                apply_fitting_length(r, desc, rules, flags, required=False, dn=main_dn, pg=pg)
        hit = rules.d5_for(cat, main_dn, "CAD") if opt.get("d5", True) else None
        if hit:
            r["X"], r["Y"], r["Z"] = hit[0], hit[1], hit[2]
            if not hit[3]:
                flags["X_unverified"] = True
        r["_flags"], r["_desc"] = flags, desc
        out.append(r)
        return r

    opt = rules.options_for(iso)
    sups = list(pg["sups"])
    for it in pg["items"]:
        qty = parse_qty(it["qty"])
        if it["parts"] and not opt["assemblies"]:
            it = dict(it, parts=None)
        if it["kit"] and not opt["kits"]:
            it = dict(it, kit=None)
        if it["parts"]:
            for eng, rus, n in it["parts"]:
                desc = eng + (" / " + rus if rus else "")
                m = re.search(r"DN\s*(\d+)", eng)
                nps = re.search(r"NPS\s*(\d+(?:\s*\d/\d)?(?:\.\d)?)", eng)
                dn = m.group(1) if m else (str(NPS_DN[nps.group(1)]) if nps and nps.group(1) in NPS_DN
                                           else it["dn"])
                L = re.search(r"L\s*=\s*(\d+)\s*mm", eng, re.I)
                cnt = n * (qty or 1)
                if L and re.match(r"\s*pipe", eng, re.I):
                    add(desc, dn, round(int(L.group(1)) * cnt / 1000.0, 3), piece_len=int(L.group(1)) * cnt)
                else:
                    add(desc, dn, cnt)
            continue
        if it["kit"]:
            for part, n in it["kit"]:
                add(part, it["dn"], n * (qty or 1))
            continue
        cat, _, _ = rules.classify(it["desc"], None, "CAD", iso)
        if cat == "Support" and qty and qty >= 1 and not opt["supports"]:
            take = sups[:int(qty)]
            sups = sups[int(qty):]
            r = add(it["desc"], it["dn"], qty, it["code"], it["unit"])
            if take:
                r["T"] = ", ".join(take)
            if len(take) < int(qty):
                r["_flags"]["T"] = "Not all support numbers found on the drawing"
            w = rules.support_weight(it["desc"])
            if w:
                r["U"] = w[1] if qty == 1 else "=%s*%s" % (w[1], qty)
            continue
        if cat == "Support" and qty and qty >= 1:
            take = sups[:int(qty)]
            sups = sups[int(qty):]
            for k in range(int(qty)):
                r = add(it["desc"], it["dn"], 1, it["code"], it["unit"])
                if k < len(take):
                    r["T"] = take[k]
                else:
                    r["_flags"]["T"] = "Support number not found on the drawing"
                w = rules.support_weight(it["desc"])
                if w:
                    r["U"] = w[1]
            continue
        kit = re.search(r"kit|компл", it["unit"] or "", re.I)
        if cat == "Pipe" and not kit:
            # v0.9.5 IOT: every pipe piece has its own POS, the same number as its cut mark <4> -
            # only when every cut mark of that DN on the sheet is such a POS (PGZ: POS 1 = pipe,
            # cut marks <1> <2> of that pipe)
            piece = [L for m, L, d in pg["cuts"] if str(m) == str(it["pos"])]
            dns_ = parse_dn(it["dn"])
            pipe_pos = {str(x["pos"]) for x in pg["items"]
                        if re.match(r"\s*(PIPE|ТРУБА)\b", latin_part(x["desc"]).strip().upper())}
            marks = [str(m) for m, L, d in pg["cuts"] if dns_ and d == dns_[0]]
            if pg.get("iot") and len(piece) == 1 and qty is not None and qty_ok(piece, qty) and dns_ \
                    and str(it["pos"]) in marks and len(marks) > 1 and all(m in pipe_pos for m in marks):
                add(it["desc"], it["dn"], qty, it["code"], it["unit"], piece_len=piece[0], pos=it["pos"])
                continue
        r = add(it["desc"], it["dn"], qty, it["code"], it["unit"], as_text="FASTENING KIT" if kit else None,
                pos=it["pos"])
    if sups:
        log(iso, sheet, "CHECK", "Support number(s) on the drawing without a support in the material list: %s"
            % ", ".join(sups))
    # valve tags
    tags = list(pg["tags"])
    for r in out:
        if r["E"] == "Valve":
            want = next((t for t in tags if TAG_TYPES.get(t[1]) and
                         TAG_TYPES[t[1]] in latin_part(r["_desc"]).upper()), None)
            if want is None:
                want = next((t for t in tags if not TAG_TYPES.get(t[1])), tags[0] if tags else None)
            if want:
                r["R"] = want[0]
                tags.remove(want)
            else:
                r["_flags"]["R"] = "Valve tag not found on the drawing"
    for r in out:
        for k, v in r["_flags"].items():
            if not k.endswith("_unverified"):
                log(iso, sheet, "CHECK", "%s: %s" % (k, v))
    return out


# ============================================================================ drawing -> rows
def iso_from_filename(path):
    base = os.path.splitext(os.path.basename(path))[0]
    m = re.search(r"([A-Z0-9]+(?:-[A-Z0-9]+)*-ISO-\d+)(?:_(\d{1,2}))?", base, re.I)
    if not m:
        return None, None
    return m.group(1).upper(), m.group(2)


def find_sheet(names_in_book, wanted):
    for n in wanted:
        if n in names_in_book:
            return n
    return None


def roster_lookup(wb):
    out = {}
    name = find_sheet(wb.sheetnames, ROSTER_NAMES)
    if not name:
        return out
    for r in wb[name].iter_rows(min_row=2, values_only=True):
        if r and r[0]:
            title = str(r[7] or "")
            m = re.search(r"(?:/|^|\s)(\d{3}-[A-Z]{1,5}-\d{3,5})", title)
            out[str(r[0]).strip().upper()] = {"rev": str(r[3]).strip() if r[3] else None,
                                              "line": m.group(1) if m else None}
    return out


def _int_or(v):
    try:
        return int(v)
    except (TypeError, ValueError):
        return v


_ENGINE = None
_POOLS = []

# ---- hooks for the GUI (gui.py): progress events and a cancel switch
HOOK = None                 # callable(event, **info) or None
CANCEL = None               # threading.Event or None


class UserError(Exception):
    """A problem the user can fix (no PDFs, no workbook ...) - shown without a traceback."""


class Cancelled(Exception):
    pass


def emit(event, **info):
    if CANCEL is not None and CANCEL.is_set():
        raise Cancelled()
    if HOOK is not None:
        try:
            HOOK(event, **info)
        except Exception:
            pass


def _init_worker(choice):
    global _ENGINE
    os.environ.setdefault("OMP_THREAD_LIMIT", "1")
    for name in ("stdout", "stderr"):               # windowed exe: no console in the workers
        if getattr(sys, name) is None:
            setattr(sys, name, open(os.devnull, "w"))
    _ENGINE = make_engine(choice)


def page_count(path):
    import pypdfium2 as pdfium
    pdf = pdfium.PdfDocument(path)
    try:
        return len(pdf)
    finally:
        pdf.close()


def render_page(path, index, upscale=True):
    import pypdfium2 as pdfium
    pdf = pdfium.PdfDocument(path)
    try:
        page = pdf[index]
        scale = _prepare_page(page)
        up = scale / (DPI / 72.0)
        if not upscale:
            scale = DPI / 72.0
        img = page.render(scale=scale).to_pil()
        img.info["af_up"] = up if upscale else 1.0
        page.close()
    finally:
        pdf.close()
    if img.height > img.width:
        img = img.rotate(90, expand=True)
    return img


ZONE_LINENO = (0.69, 0.925, 0.88, 0.99)          # pipe line number under "Isometric drawing"
LINE_NO = re.compile(r"\b(\d{3})\s*-\s*([A-Z][A-Z0-9]{0,4})\s*-\s*(\d{3,5})\s*-\s*([A-Z0-9]{9,13})\b")
SPEC_SHAPE = "LLLDLDDLLDD"                        # e.g. ACB2B02SO32, ASB0B02SY34, ACB3B05SW35
TO_LETTER = str.maketrans("0815264", "OBISZGA")
TO_DIGIT = str.maketrans("OQDBISZGA", "000815264")


def fix_spec(code):
    """Piping class suffix of a line number: letters and digits have fixed positions."""
    def shaped(c):
        return "".join(ch.translate(TO_LETTER) if k == "L" else ch.translate(TO_DIGIT)
                       for ch, k in zip(c, SPEC_SHAPE))
    ok = lambda c: re.fullmatch(r"[A-Z]{3}\d[A-Z]\d{2}[A-Z]{2}\d{2}", c) is not None
    if len(code) == len(SPEC_SHAPE) and ok(shaped(code)):
        return shaped(code), shaped(code) == code
    if len(code) == len(SPEC_SHAPE) + 1:          # one extra character read: fewest changes wins
        best = None
        for i in range(len(code)):
            raw = code[:i] + code[i + 1:]
            c = shaped(raw)
            if ok(c):
                n = sum(1 for a, b in zip(raw, c) if a != b)
                if best is None or n < best[0]:
                    best = (n, c)
        if best:
            return best[1], False
    return code, False


def read_line_no(img, engine, path=None, index=0):
    """Pipe line number from the title block: text layer if the page has one, else OCR.
    Returns (line, sure)."""
    if path is not None:
        try:
            objs, W, H = cad_objects(path, index)
            for o in objs:
                m = LINE_NO.search(o.t.upper())
                if m and o.y0 > 0.9 * H:
                    return "-".join(m.groups()), True
        except Exception:
            pass
    try:
        txt = engine.sparse_text(clean(crop(img, ZONE_LINENO))).upper()
    except Exception:
        return None, False
    m = LINE_NO.search(txt.replace(" ", ""))
    if not m:
        return None, False
    spec, sure = fix_spec(m.group(4))
    return "%s-%s-%s-%s" % (m.group(1), m.group(2), m.group(3), spec), False


TL_CUT = re.compile(r"<\s*(\d{1,2})\s*>\s+(\d{2,6})\s+(\d{1,4})(?![\d.])")


def text_layer_extras(path, index):
    """Exact cut lengths, valve tags and supports from the text layer of a vector page (or a
    scanned page with some vector text). Returns None when the page has no usable text."""
    try:
        objs, W, H = cad_objects(path, index)
    except Exception:
        return None
    if len(objs) < 5:
        return None
    cuts = []
    head = next((o for o in objs if "CUT PIPE LENGTH" in o.t.upper()), None)
    if head is not None and head.x0 >= 0.5 * W:          # normal orientation: table in the right panel
        zone = [o for o in objs if o.x0 >= 0.55 * W and head.y0 < o.y0 < 0.65 * H]
        for ln in _lines(zone):
            for m in TL_CUT.finditer(" ".join(o.t for o in ln)):
                p, L, dn = int(m.group(1)), int(m.group(2)), int(m.group(3))
                if dn in VALID_DN and (p, L, dn) not in cuts:
                    cuts.append((p, L, dn))
    tags = []
    for o in objs:
        for m in VALVE_TAG.finditer(o.t.upper()):
            if m.group(1) not in [t[0] for t in tags]:
                tags.append((m.group(1), m.group(2)))
    sups = []
    oris = [o for o in objs if ORI.search(o.t.upper())]
    for o in objs:
        m = SUP_NO.search(o.t.upper())
        if not m:
            continue
        no = m.group(0)
        if no in [x[0] for x in sups]:
            continue
        near = sorted(oris, key=lambda q: abs(q.x0 - o.x0) + abs(q.y0 - o.y0))
        ori = None
        om = ORI.search(o.t.upper())
        if om:
            ori = om.group(1)
        elif near and abs(near[0].x0 - o.x0) + abs(near[0].y0 - o.y0) < 40:
            ori = ORI.search(near[0].t.upper()).group(1)
        sups.append((no, re.sub(r"\s+", " ", ori).strip() if ori else None, ori is None))
    return {"cuts": cuts, "tags": tags, "sups": sups}


PDATA_WORDS = (("PAINT", "PAINTING"), ("TRACE", "TRACING"), ("INSUL", "INSULATION"))
ZONE_PDATA = (0.29, 0.90, 0.43, 0.985)       # x0, y0, x1, y1 of the PAINTING / TRACING / INSULATION rows


def _pdata_value(v):
    """'A1→A4' / 'A1->A4' / 'A1—A4' -> 'A1---->A4' (the way the Measurement Table writes it); 'N' stays 'N'."""
    v = re.sub(r"\s+", "", str(v or "")).upper()
    v = re.sub(r"(?<=[A-Z])[IL|](?=\W|$)", "1", v)
    m = re.search(r"([A-Z]\d)\D{0,4}?([A-Z]\d)", v)          # 'FB1>B22' / 'KB1B2>' -> B1, B2
    if m:
        return "%s---->%s" % (m.group(1), m.group(2))
    return v or None


def read_pipeline_data(path, index, engine, img):
    """v0.9: painting system, tracing and insulation from the PIPELINE DATA block (title block, page 1).
    Text layer first; otherwise OCR of the block. Returns {"PAINT": .., "TRACE": .., "INSUL": ..}."""
    out = {}
    try:
        objs, W, H = cad_objects(path, index)
    except Exception:
        objs = []
    if objs:
        for key, word in PDATA_WORDS:
            lab = next((o for o in objs if o.t.upper().startswith(word) and o.y0 > 0.8 * H), None)
            if not lab:
                continue
            h = max(lab.y1 - lab.y0, 3)
            vals = [o for o in objs if lab.x1 < o.x0 < lab.x1 + 0.12 * W and lab.y0 - h <= o.y0 <= lab.y0 + 2.2 * h
                    and not re.search(r"[A-Za-zА-Яа-я]{4}", o.t)]
            if vals:
                out[key] = _pdata_value(min(vals, key=lambda o: o.x0).t)
        if out:
            return out
    try:
        W, H = img.size
        x0, y0, x1, y1 = ZONE_PDATA
        zone = clean(img.crop((int(x0 * W), int(y0 * H), int(x1 * W), int(y1 * H)))).convert("L")
        import numpy as np
        a = np.asarray(zone) < 128
    except Exception:
        return out
    zh, zw = a.shape
    # grid: horizontal lines (rows that are mostly ink) and the column line between label and value
    rows_ = np.where(a[:, int(0.15 * zw):int(0.95 * zw)].mean(axis=1) > 0.6)[0]
    lines = []
    for y in rows_:
        if lines and y - lines[-1][-1] <= 2:
            lines[-1].append(int(y))
        else:
            lines.append([int(y)])
    ys = [(ln[0], ln[-1]) for ln in lines]
    bands = [(ys[i][1] + 2, ys[i + 1][0] - 2) for i in range(len(ys) - 1) if ys[i + 1][0] - ys[i][1] > 0.06 * zh]
    if len(bands) < 3:
        return out
    hb = max(int(np.mean([b - a_ for a_, b in bands])), 1)
    colp = a[bands[0][0]:bands[-1][1], :].mean(axis=0)
    mid = [x for x in range(int(0.45 * zw), int(0.8 * zw)) if colp[x] > 0.8]
    split = (mid[0] + 3) if mid else int(0.6 * zw)
    labels = []
    for (t, b) in bands:
        lab = zone.crop((0, t, split - 4, b))
        try:
            labels.append(engine.sparse_text(lab.resize((lab.width * 2, lab.height * 2))).upper())
        except Exception:
            labels.append("")
    start = next((i for i, l in enumerate(labels) if re.search(r"PAINT|CACTEM|ПОКРАС|MOKPAC|POKRAS", l)), None)
    if start is None:                                   # template order: ... PAINTING, TRACING, INSULATION
        thick = [i for i in range(len(ys) - 1) if ys[i + 1][1] - ys[i + 1][0] >= 3]
        start = (thick[0] - 2) if thick and thick[0] >= 2 else None
    if start is None or start + 2 >= len(bands):
        return out
    for (key, _), (t, b) in zip(PDATA_WORDS, bands[start:start + 3]):
        cell = zone.crop((split + 2, t, zw - 2, b))
        cell = pad(cell.resize((cell.width * 2, cell.height * 2)))
        try:
            v = engine.token(cell, "ABCDEFGHKLMNPRSTUVWXY0123456789-> ")
        except Exception:
            v = ""
        out[key] = _pdata_value(v)
    return out


def _wide_panel_zone(img, raw):
    """v0.9.11: on some subcontractor sheets (e.g. LGN 12510 'TK-ISO' on a 790 x 558 pt page) the
    materials table sits between x = 0.62 and 0.99 of the sheet, so the usual panel zone cuts the
    QTY column off and the GOST number (8732, 17375, ...) was taken as the quantity, while the
    left edge picked up support tags from the isometric.  When the header row is read but has no
    QTY, find the two long vertical frame lines that bound the table and crop between them."""
    head = next((t.upper() for t in raw if t.upper().startswith("POS") and "DN" in t.upper()), None)
    if head is None or "QT" in head:
        return None
    try:
        zx0, zy0, zx1, zy1 = 0.55, 0.02, 1.0, 0.75
        reg = clean(crop(img, (zx0, zy0, zx1, zy1)))
        a = np.asarray(reg.convert("L")) < 128
        W = img.size[0]
        xs = [zx0 + x / W for x in np.where(a.mean(axis=0) > 0.5)[0]]
        left = [x for x in xs if 0.56 < x < 0.70]
        right = [x for x in xs if x > 0.95]
        if not left or not right:
            return None
        x0, x1 = max(left) + 0.003, min(right) - 0.002
        if x1 <= ZONE_PANEL[2]:
            return None
        return (x0, ZONE_PANEL[1] - 0.02, x1, 0.70)
    except Exception:  # noqa
        return None


def pipe_qty_second_look(panel, items, engine):
    """v0.9.11: pipe metres ('11.7M') are read once more from their own cell.  A different
    reading is kept as qty_alt; process_pdf uses it only when it - and not the first reading -
    agrees with the cut lengths (one OCR pass read '11.7M' as '14.7M')."""
    for it in items:
        q = it.get("qty") or ""
        if not it.get("qbox") or not re.search(r"\d[.,]\d|\dM", q.upper()) or it.get("qty_alt"):
            continue
        x0, x1 = it["qbox"]
        box = (max(0, x0 - 25), max(0, it["y0"] - 12), min(panel.size[0], x1 + 25),
               min(panel.size[1], it["y1"] + 12))
        cell0 = panel.crop(tuple(int(v) for v in box))
        for f in (1, 2):
            t = engine.token(pad(cell0.resize((cell0.width * f, cell0.height * f))), "0123456789.M")
            if parse_qty(t) is not None:
                if parse_qty(t) != parse_qty(q):
                    it["qty_alt"] = t
                break


def qty_column_pass(panel, items, engine):
    """v0.9.11, reduced sheets with the hairline font: Tesseract reads a thin '1' as '4' and a
    '9' as '>' in the material list.  Read the QTY column once more on its own, with the strokes
    thickened, and use that reading wherever it gives a number."""
    W = panel.size[0]
    x0 = int(0.89 * W)
    col = panel.crop((x0, 0, W, panel.size[1])).filter(ImageFilter.MinFilter(5))
    p = 60
    boxes = engine.boxes(pad(col, p))
    for it in items:
        y0, y1 = it.get("y0"), it.get("y1")
        if y0 is None or y1 is None:
            continue
        hits = [b for b in boxes if y0 - 8 <= b.yc - p <= y1 + 8]
        if len(hits) != 1:
            continue
        t = hits[0].text.replace("B", "8").replace("O", "0").replace("D", "0")
        t = re.sub(r"[^0-9.,]", "", t)
        if not re.fullmatch(r"\d+(?:[.,]\d+)?", t):
            continue
        old = it["qty"] or ""
        if parse_qty(t) == parse_qty(old):
            it["qty_reread"] = False        # two independent readings agree
            continue
        metres = "." in t or "." in old or "M" in old.upper()
        if parse_qty(old) is None:
            it["qty"], it["qty_reread"] = t + ("M" if metres else ""), True
        elif metres:
            it["qty_alt"] = t + "M"         # pipe metres: the cut lengths decide (process_pdf)
        else:
            it["qty"] = t                   # a count: the thickened reading is the better one


def read_page(job):
    """Everything the tool reads from one PDF page (runs in a worker process)."""
    path, index = job
    engine = _ENGINE
    img = render_page(path, index)
    pno = index + 1
    # v0.9.11: a reduced sheet is rendered larger (see _prepare_page); the title-block reader is
    # tuned for the normal size, so it gets its own render at the usual 300 dpi
    timg = render_page(path, index, upscale=False) if img.info.get("af_up", 1.0) > 1.01 else img
    sheet, total = read_sheet_no(clean(crop(timg, ZONE_TITLE)), engine, pre_cropped=True)
    if sheet is None:
        # LGN / PGZ title block: SHEET / ЛИСТ cell near the bottom right corner
        s2, t2 = read_sheet_no(clean(crop(timg, ZONE_TITLE_LOW)), engine, pre_cropped=True,
                               title_zone=ZONE_TITLE_LOW)
        if s2 is not None:
            sheet, total = s2, t2 or total
    panel_raw = crop(img, ZONE_PANEL)
    items, cuts, area, raw, band = parse_panel(clean(panel_raw), engine)
    wide = _wide_panel_zone(img, raw)
    if wide is not None:
        # v0.9.11: materials table laid out further right (QTY column cut off by the A3 zone)
        panel_raw = crop(img, wide)
        items, cuts, area, raw, band = parse_panel(clean(panel_raw), engine)
    if img.info.get("af_up", 1.0) > 1.01:
        qty_column_pass(clean(panel_raw), items, engine)
    pipe_qty_second_look(clean(panel_raw), items, engine)
    # the cut-length table is often under the stamp: read it 2 more ways and vote later
    cutsets = [cuts]
    for spread_max, grey_max in ((60, 100), (40, 120)):
        part = panel_raw
        if band[0] is not None:
            y1 = band[1] if band[1] is not None else panel_raw.height
            part = panel_raw.crop((0, max(0, int(band[0]) - 10), panel_raw.width, int(y1) + 5))
        res = parse_panel(clean_variant(part, spread_max, grey_max), engine)
        cutsets.append(res[1])
    # v0.9.11: with the table further right the isometric reaches x = 0.62 and down to the legend
    dzone = ZONE_DRAWING if wide is None else (ZONE_DRAWING[0], ZONE_DRAWING[1], wide[0] - 0.006, 0.79)
    tags, sups, dtexts = read_drawing(clean(crop(img, dzone)), engine, pre_cropped=True,
                                      sup_prefix=expected_sup_prefix(path))
    tl = text_layer_extras(path, index)
    if tl:                                   # exact text beats OCR (stamp-free, no misreads)
        if tl["cuts"]:
            cuts = tl["cuts"]
            cutsets = [list(tl["cuts"]), list(tl["cuts"]), cutsets[0]]
        tags = tl["tags"] + [t for t in tags if t[0] not in [x[0] for x in tl["tags"]]]
        sups = tl["sups"] + [x for x in sups if x[0] not in [y[0] for y in tl["sups"]]]
    line = read_line_no(img, engine, path, index) if index == 0 else (None, False)
    pdata = read_pipeline_data(path, index, engine, img) if index == 0 else {}
    dbg = ("=== PDF page %d  sheet %s / %s\n--- panel\n%s\n--- cuts %s (other readings %s)\n--- tags %s\n"
           "--- supports %s\n--- line no. %s\n--- drawing text\n%s\n"
           % (pno, sheet, total, "\n".join(raw), cuts, cutsets[1:], tags, sups, line, dtexts[0]))
    return {"sheet": sheet, "total": total, "items": items, "cuts": cuts, "area": area, "tags": tags,
            "sups": sups, "page": pno, "cutsets": cutsets, "dbg": dbg, "line": line, "pdata": pdata}


def crossed_out(path, index):
    """v0.9.4: a sheet struck through with two lines from corner to corner (how LGN cancels a
    drawing, also on scanned sheets). Looks at the vector lines of the page only."""
    try:
        import ctypes
        import pypdfium2 as pdfium
        import pypdfium2.raw as pdfium_c
    except Exception:
        return False
    pdf = pdfium.PdfDocument(path)
    try:
        page = pdf[index]
        l0, b0, r0, t0 = page.get_mediabox()
        W, H = abs(r0 - l0), abs(t0 - b0)
        n = 0
        for o in page.get_objects(filter=[pdfium_c.FPDF_PAGEOBJ_PATH], max_depth=6):
            l, b, r, t = o.get_bounds()
            if abs(r - l) < 0.6 * W or abs(t - b) < 0.6 * H:
                continue
            if pdfium_c.FPDFPath_CountSegments(o.raw) != 2:
                continue                       # a straight line = move-to + line-to
            p = []
            for k in range(2):
                seg = pdfium_c.FPDFPath_GetPathSegment(o.raw, k)
                x, y = ctypes.c_float(), ctypes.c_float()
                pdfium_c.FPDFPathSegment_GetPoint(seg, x, y)
                p.append((x.value, y.value))
            if p[0][0] != p[1][0] and p[0][1] != p[1][1]:
                n += 1
        return n >= 2
    except Exception:
        return False
    finally:
        pdf.close()


def is_void(path):
    """A cancelled drawing: revision 'VD' in the file name (…-ISO-00083_VD_1_publication.pdf), the
    'АННУЛИРОВАН / VOID' stamp in its text layer, or (v0.9.4) every sheet crossed out.
    Its material is not built, so it is not filled."""
    if re.search(r"-ISO-\d+_VD(_|\b)", os.path.basename(path), re.I):
        return True
    try:
        objs, _, _ = cad_objects(path, 0)
        txt = " ".join(o.t for o in objs).upper()
        if "АННУЛИРОВАН" in txt and "VOID" in txt:
            return True
        n = page_count(path)
        return n > 0 and all(crossed_out(path, i) for i in range(n))
    except Exception:
        return False


def process_pdf(path, engine, rules, roster, log, debug_dir=None, pool=None):
    """Read one drawing and apply the per-family conventions of the Measurement Table (v0.9)."""
    if is_void(path):
        # v0.9.4 (CC, 2026-11-12): cancelled / crossed-out drawings are not filled at all; the log
        # lists them. (v0.9.2 wrote them as grey VOID rows - write_rows still knows '_void'.)
        iso, _ = iso_from_filename(path)
        print("    VOID / АННУЛИРОВАН (cancelled, crossed out) drawing - not filled", flush=True)
        log(iso or os.path.basename(path), "", "WARN",
            "Cancelled drawing (revision VD / 'АННУЛИРОВАН / VOID' / crossed out) - not filled")
        return iso, []
    iso, rows = _process_pdf(path, engine, rules, roster, log, debug_dir, pool)
    return iso, apply_conventions(iso, rows, rules)


def apply_conventions(iso, rows, rules):
    """v0.9 - things that are a matter of convention, not of reading (Settings sheet per drawing family):
    short pipe line number, J/K for equal sizes, D5 code from the D5 code list by DN band,
    description and unit of every D5 code taken from the code list."""
    if not rows:
        return rows
    import fnmatch
    opt = rules.options_for(iso)
    names = [n for n in rules.family_names if fnmatch.fnmatch(str(iso).upper(), n[0])]
    for r in rows:
        if names and r.get("E") != "Support" and r.get("H"):
            u = latin_part(str(r["H"])).upper()
            hit = next((n for n in names if re.search(r"(?<![A-Z])%s(?![A-Z])" % re.escape(n[1]), u)), None)
            if hit:
                r["E"], r["F"] = hit[2] or r["E"], hit[3] or r["F"]
                r["_flags"].pop("F", None)
                r["_flags"].pop("F_unverified", None)
                if hit[2]:
                    r["_flags"].pop("E", None)          # v0.9.4: the family rule named it
                if not hit[4]:
                    r["_flags"]["F_unverified"] = True
        if opt.get("line_short") and r.get("D"):
            r["D"] = "-".join(str(r["D"]).split("-")[:3])
        if not opt.get("jk_equal", True) and r.get("J") not in (None, "") and str(r.get("J")) == str(r.get("K")):
            r["J"] = r["K"] = None
        if opt.get("d5") and not r.get("X") and r.get("E"):
            hit = rules.d5_band(iso, r["E"], parse_dn(str(r.get("I") or "")))
            if hit:
                r["X"], r["Y"], r["Z"] = hit[0], hit[1], hit[2]
                for k in [k for k in r["_flags"] if k in ("X", "Y", "Z")]:
                    del r["_flags"][k]
                if not hit[3]:
                    r["_flags"]["X_unverified"] = "D5 code from the D5 code list by DN band - section not verified"
        if r.get("X") and rules.d5_list.get(str(r["X"])):
            r["Y"], r["Z"] = rules.d5_list[str(r["X"])][:2]
        # v0.9.2 conventions learned per family (PGZ answer key)
        if opt.get("j_all") and r.get("J") in (None, ""):
            dns = parse_dn(str(r.get("I") or ""))
            if dns:
                r["J"] = dns[0]
            elif re.fullmatch(r"\d{1,4}", str(r.get("I") or "").strip()):
                r["J"] = int(str(r["I"]).strip())       # bolt size M12 -> 12
        if not opt.get("sup_weight", True) and r.get("E") == "Support":
            r["U"] = None
            r["_flags"].pop("U", None)
        if opt.get("code_desc") and r.get("L") in (None, "") and r.get("H"):
            m = STD_IN_DESC.search(latin_part(str(r["H"])))
            if m:
                r["L"] = re.sub(r"\s+", " ", m.group(1)).strip()
    for flag, keys in (("area_once", ("V",)), ("pdata_once", ("PAINT", "TRACE", "INSUL"))):
        if opt.get(flag):
            seen = set()
            for r in rows:                   # only on the first row of each pipe line
                k = r.get("D")
                if k in seen:
                    for c in keys:
                        r[c] = None
                        r["_flags"].pop(c, None)
                seen.add(k)
    return rows


# standard in a description: 'GOST 33259-2015', 'OST 26.260.454-99', 'GOST R 52376-2005', 'ATK 24.200.02-90'
STD_IN_DESC = re.compile(r"\b((?:GOST\s+R|GOST|OST|ATK|TU|STO|ASME|API|MSS|DIN|EN|ISO)\s+[A-Z]?\d[\w.]*(?:-\d[\w.]*)*)")


def _process_pdf(path, engine, rules, roster, log, debug_dir=None, pool=None):
    layout = cad_layout(path)
    if layout == "ITEM CODE":               # AutoCAD PDF with real text: read it exactly, no OCR
        print("    vector PDF with text layer - reading text directly", flush=True)
        return process_cad_pdf(path, rules, log, debug_dir)
    iso, rev = iso_from_filename(path)
    if not iso:
        log("?", "", "ERROR", "Cannot read ISO No. from file name: %s" % os.path.basename(path))
        return None, []
    if layout:
        print("    vector PDF in the 'STANDART / DOKUMENT' layout - read like a scanned drawing", flush=True)
        log(iso, "", "INFO", "AutoCAD PDF in the older material-list layout - read with OCR like the scanned "
                             "drawings of this layout (no stamp, so the reading is clean)")
    info = roster.get(iso, {})
    if not rev:
        rev = info.get("rev")
    elif info.get("rev") and info["rev"] != rev:
        log(iso, "", "WARN", "Rev in file name (%s) differs from Personnel Roster (%s); used %s"
            % (rev, info["rev"], rev))
    line_no = info.get("line")

    jobs = [(path, i) for i in range(page_count(path))]
    emit("pages", done=0, total=len(jobs))
    pages = []
    if pool is not None:
        for k, pg in enumerate(pool.imap(read_page, jobs), 1):
            pages.append(pg)
            emit("pages", done=k, total=len(jobs))
    else:
        global _ENGINE
        _ENGINE = engine
        for k, j in enumerate(jobs, 1):
            print("    page %d ..." % (j[1] + 1), flush=True)
            pages.append(read_page(j))
            emit("pages", done=k, total=len(jobs))
    line_from_title = False
    if not line_no:
        line_no, sure = next((p["line"] for p in pages if p.get("line") and p["line"][0]), (None, False))
        if line_no:
            line_from_title = not sure
            log(iso, "", "INFO", "Pipe line %s read from the title block (not in Personnel Roster)%s"
                % (line_no, "" if sure else " - please check"))
        else:
            log(iso, "", "WARN", "Pipe line not found in Personnel Roster or title block")
    opt = rules.options_for(iso)
    sheets, prev, dbg = [], None, []
    for pg in pages:
        sheet = pg["sheet"]
        if not sheet:
            guess = str(int(re.sub(r"\D", "", prev) or 0) + 1) if prev else "1"
            log(iso, guess, "CHECK", "Sheet No. not readable on PDF page %d, assumed %s" % (pg["page"], guess))
            sheet = pg["sheet"] = guess + "?"
        prev = sheet.rstrip("?")
        sheets.append(pg)
        dbg.append(pg["dbg"])
    if debug_dir:
        os.makedirs(debug_dir, exist_ok=True)
        with open(os.path.join(debug_dir, "%s_ocr.txt" % iso), "w", encoding="utf-8") as f:
            f.write("engine: %s\n" % engine.name + "\n".join(dbg))

    area_all = next((s["area"] for s in sheets if s["area"]), None)
    merged, order = {}, []
    for s in sheets:
        base = re.sub(r"[A-Z?]+$", "", s["sheet"]) or s["sheet"]
        if base not in merged:
            merged[base] = {"items": [], "cuts": [], "tags": [], "sups": [], "area": s["area"],
                            "unsure": s["sheet"].endswith("?"), "cutsets": [[], [], []]}
            order.append(base)
        for k in ("items", "cuts", "tags", "sups"):
            merged[base][k] += s[k]
        for i, cs in enumerate(s["cutsets"]):
            merged[base]["cutsets"][i] += cs
    totals = {s["total"] for s in sheets if s["total"]}
    if len(totals) == 1:
        t = totals.pop()
        if t.isdigit() and int(t) != len(order):
            log(iso, "", "CHECK", "Title block says %s sheets, found %d" % (t, len(order)))

    rows = []
    for base in order:
        d = merged[base]
        # v0.9.4: LGN drawings have no 'ORI =' under their support numbers - only ask for an ORI when
        # this drawing writes it somewhere
        drawing_ori = any(o for dd in merged.values() for _, o, _ in dd["sups"])
        sheet_rows, pipe_dns = [], []
        for it in d["items"]:
            flags = {}
            desc = tidy_description(it["desc"], rules)
            cat, zh, zh_ver = rules.classify(desc, it["section"])
            if cat == "Pipe" and not zh:
                zh, zh_ver = "管道", True
            dns = parse_dn(it["dn"])
            from_desc = False
            if not dns and cat not in ("Bolt", "Nut"):
                m = re.search(r"\b(?:DN|SW|BL|WN|SO|LJ|TH)\s?(\d{1,4})(?![\d.])", desc)
                if m and int(m.group(1)) in VALID_DN:
                    dns, from_desc = [int(m.group(1))], True
            qty = parse_qty(it["qty"])
            r = {"A": iso, "B": _int_or(iso.split("-")[3] if iso.count("-") >= 3 else None),
                 "C": rev, "D": line_no, "E": cat, "F": zh, "G": _int_or(base), "H": desc,
                 "M": qty, "V": d["area"] or area_all}
            if not cat:
                flags["E"] = "Section not recognised"
            if not zh:
                flags["F"] = "No keyword in rules.xlsx > Categories matches this description"
            elif not zh_ver:
                flags["F_unverified"] = True
            if d["unsure"]:
                flags["G"] = "Sheet No. not readable, assumed"
            cf = it["confs"]
            # v0.9.11: on bilingual lines ('Pipe 89x5-V / Труба 89x5-B ...') only the English half
            # is judged - the OCR reads Latin letters only, so the Russian half always scores low
            wc = it.get("wconf") or []
            cut = next((i for i, (t, _) in enumerate(wc) if t == "/" or t.endswith("/")), None)
            if cut is not None and cut >= 1 and "-LGN-" in (iso or "").upper():
                cf = [c for t, c in wc[:cut + 1] if t != "/"] or cf
            if cf and sum(1 for x in cf if x < 0.6) > 0.3 * len(cf):
                flags["H"] = "Low OCR confidence - check text"
            main_dn = None
            inch = bolt_size(desc, it["dn"]) if cat in ("Bolt", "Nut") else None
            if inch:
                r["I"] = inch                   # bolts / nuts: the inch size, as written
            elif not dns:
                flags["I"] = "DN not readable (%r)" % it["dn"]
            elif len(dns) == 2:
                r["I"] = "%dX%d" % (dns[0], dns[1])
                if dns[0] != dns[1]:
                    r["J"], r["K"] = dns[0], dns[1]
                main_dn = dns[0]
            else:
                r["I"] = main_dn = dns[0]
            if from_desc:
                flags["I"] = "DN column not readable (%r) - taken from the description" % it["dn"]
            if qty is None:
                flags["M"] = "QTY not readable (%r)" % it["qty"]
            elif it.get("qty_reread"):
                flags["M"] = "QTY read on second attempt - please confirm"

            if cat == "Pipe" and main_dn:
                pipe_dns.append(main_dn)
                css = [[x[1] for x in cs if x[2] == main_dn] for cs in d["cutsets"]]
                alt = parse_qty(it.get("qty_alt"))
                if alt is not None and not any(c and qty_ok(c, qty) for c in css) and \
                        any(c and qty_ok(c, alt) for c in css):
                    flags["M"] = ("QTY read as %r and %r - used %s m, which matches the cut lengths"
                                  % (it["qty"], it["qty_alt"], alt))
                    qty = r["M"] = alt
                for txt in (it["qty"], it.get("qty_alt")):
                    fixed = restore_decimal(txt, qty, css)
                    if fixed is not None:
                        flags["M"] = ("Decimal point restored from the cut lengths (OCR read %r) - "
                                      "please confirm" % txt)
                        qty = r["M"] = fixed
                        break
                L, msg, bad = choose_cuts(
                    [[x[1] for x in cs if x[2] == main_dn] for cs in d["cutsets"]], qty)
                if msg:
                    flags["N"] = msg
                    if bad:
                        flags["M"] = msg
                if not L:
                    flags["N"] = ("No cut pipe length found for DN%d - the row may be hidden under the "
                                  "AFC stamp, please read it on the drawing" % main_dn)
                if L:
                    r["N"] = L[0] if len(L) == 1 else "=" + "+".join(str(x) for x in L)
            elif cat in ("Fittings", "Flanges", "Gasket"):
                apply_fitting_length(r, desc, rules, flags, required=(cat == "Fittings"), dn=main_dn)
            elif cat == "Valve":
                hit = rules.valve_length(desc)
                if hit and comp_key(desc) != hit[0] and hit[4] and all(
                        op[0] in ("equal", "delete") for op in
                        difflib.SequenceMatcher(None, hit[0], comp_key(desc)).get_opcodes()):
                    desc = r["H"] = tidy_description(hit[4], rules)   # snap to the known spelling
                if hit:
                    r["S"] = hit[1]
                    flags["S"] = ("Typical value from reference answers. Check the flanges shown at this "
                                  "valve on the drawing (each flange +60 here)."
                                  + (" Also seen: %s" % hit[2] if hit[2] else ""))
                else:
                    flags["S"] = "Valve length unknown - add it to rules.xlsx > Valve Length (or run learn)"
            if opt["d5"]:
                hit = rules.d5_for(cat, main_dn)
                if hit:
                    r["X"], r["Y"], r["Z"] = hit[0], hit[1], hit[2]
                    if not hit[3]:
                        flags["X_unverified"] = True
                else:
                    flags["X"] = "No D5 code in rules.xlsx for %s DN%s" % (cat, main_dn)
            r["_flags"], r["_pos"], r["_desc"] = flags, it["pos"], desc
            # one row per valve (each valve has its own tag), like the reference answers
            if cat == "Valve" and isinstance(qty, (int, float)) and float(qty).is_integer() and 2 <= qty <= 20:
                for k in range(int(qty)):
                    v = dict(r, M=1, _flags=dict(flags))
                    sheet_rows.append(v)
                continue
            sheet_rows.append(r)

        unused = sorted({x[2] for cs in d["cutsets"][:1] for x in cs} - set(pipe_dns))
        if unused and pipe_dns is not None:
            log(iso, base, "CHECK", "Cut lengths for DN%s on this sheet, but no pipe of that size in its "
                "material list - check whether they belong to another sheet" % "/".join(map(str, unused)))
        # ---- valve tags -> valve rows (by type, then in order)
        valves = [r for r in sheet_rows if r["E"] == "Valve"]
        tags = list(d["tags"])
        for r in valves:
            want = None
            for t in tags:
                kw = TAG_TYPES.get(t[1])
                if kw and kw in r["_desc"].upper():
                    want = t
                    break
            if want is None:
                for t in tags:
                    kw = TAG_TYPES.get(t[1])
                    if not kw or not any(kw in v["_desc"].upper() for v in valves):
                        want = t
                        break
            if want is None and tags:
                want = tags[0]
            if want:
                r["R"] = want[0]
                tags.remove(want)
            else:
                r["_flags"]["R"] = "Valve tag not found on the drawing"
        if tags:
            log(iso, base, "CHECK", "Valve tag(s) found but no matching valve row: %s"
                % ", ".join(t[0] for t in tags))

        # ---- supports (after the material rows, like the reference answers)
        sup_dn = pipe_dns[0] if pipe_dns else None
        known = Counter(o for _, o, _ in d["sups"] if o)
        sheet_ori = known.most_common(1)[0][0] if len(known) == 1 else None
        # v0.9.11: a support number read incompletely ('…-SUP-001', '…-SUP-001711') is dropped when
        # the complete numbers on the sheet already add up to the supports in the material list
        sup_bom = sum(r["M"] for r in sheet_rows if r.get("E") == "Support" and isinstance(r.get("M"), int))
        whole = [x for x in d["sups"] if len(x[0].split("-SUP-")[-1].split(".")[0]) == SUP_DIGITS]
        sups_here = d["sups"]
        if sup_bom and len(whole) < len(d["sups"]) and len(whole) >= sup_bom:
            sups_here = whole
            log(iso, base, "INFO", "Support number(s) read incompletely, left out (the %d complete ones "
                "match the material list): %s" % (len(whole), ", ".join(x[0] for x in d["sups"]
                                                                       if x not in whole)))
        for no, ori, unsure in sups_here:
            copied = False
            if not ori and sheet_ori:
                ori, copied = sheet_ori, True
            hit = rules.d5_for("Support", sup_dn or 0) if opt["d5"] else None
            cat, zh, zh_ver = rules.classify("SUPPORT", "Support")
            r = {"A": iso, "B": _int_or(iso.split("-")[3] if iso.count("-") >= 3 else None),
                 "C": rev, "D": line_no, "E": "Support", "F": zh or "支架", "G": _int_or(base),
                 "H": ("%sORI=%s" % (no, ori)) if ori else no, "I": sup_dn, "T": no, "V": d["area"] or area_all,
                 "_flags": {}, "_pos": None, "_desc": no}
            if opt.get("sup_qty"):
                r["M"] = 1
            if hit:
                r["X"], r["Y"], r["Z"] = hit[0], hit[1], hit[2]
            if copied:
                r["_flags"]["H"] = "ORI not read for this support - copied from the other supports on this sheet"
            elif unsure and (ori or drawing_ori):
                r["_flags"]["H"] = ("ORI not found next to the support number - please check" if not ori
                                    else "ORI read differently in different attempts - please check")
            if len(set(pipe_dns)) > 1:
                r["_flags"]["I"] = "Sheet has several pipe sizes - check which one the support is on"
            sheet_rows.append(r)

        for r in sheet_rows:
            for k, v in r["_flags"].items():
                if not k.endswith("_unverified"):
                    log(iso, base, "CHECK", "%s: %s" % (k, v))
        rows += sheet_rows
    if rows and line_from_title:
        rows[0]["_flags"]["D"] = "Pipe line read from the title block by OCR (not in Personnel Roster) - please check"
    pdata = next((p.get("pdata") for p in pages if p.get("pdata")), {}) or {}
    for r in rows:                                   # v0.9: painting / tracing / insulation on every row
        for k, v in pdata.items():
            if v and not r.get(k):
                r[k] = v
    return iso, rows


def bolt_size(desc, dn_text=""):
    """'ASME B18.2.1,Hex bolt 1 1/4" UNC...' -> '1 1/4"'."""
    m = re.search(r'(?:BOLT|NUT|STUD)\S*\s+((?:\d\s+)?\d{1,2}\s*/\s*\d{1,2}|\d{1,2})\s*["”\'\u2033]', desc, re.I)
    if not m:
        m = re.fullmatch(r'\s*((?:\d\s+)?\d{1,2}\s*/\s*\d{1,2})\s*["”\'\u2033]?\s*', str(dn_text or ""))
    if not m:
        return None
    size = re.sub(r"\s*/\s*", "/", m.group(1).strip())
    f = re.search(r"(\d+)/(\d+)$", size)
    if f and (int(f.group(1)) >= int(f.group(2)) or int(f.group(2)) not in (2, 4, 8, 16)):
        return None
    return size + '"'


# ============================================================================ template columns
# internal key -> (must contain one of, must not contain); found by reading header rows 1-2.
COLUMN_SPEC = OrderedDict([
    ("A", (["ISO NO", "单线图"], [])), ("B", (["TITLE", "主项"], [])), ("C", (["REV", "版本"], [])),
    ("D", (["PIPE LINE", "管线号"], ["LENGTH"])), ("E", (["CATEGORY", "大类"], [])),
    ("F", (["中文描述"], [])), ("G", (["页码"], [])),
    ("H", (["COMPONENT DESCRIPTION", "材料描述"], [])),
    ("J", (["管件大管径"], ["长度"])), ("K", (["管件小管径"], ["长度"])),
    ("I", (["DN公称直径"], [])),
    ("L", (["ITEM CODE"], [])), ("M", (["AFC QTY", "图纸量"], ["SURFACE", "表面积"])),
    ("N", (["NET LENGTH", "管道总净长"], [])), ("O", (["MATERIAL LENGTH", "管件长度"], [])),
    ("P", (["大管径长度"], [])), ("Q", (["小管径长度"], [])),
    ("R", (["VALVE NO", "阀门位号"], [])), ("S", (["阀门长度"], [])),
    ("T", (["SUPPORT NO", "支架编号"], [])), ("U", (["SUPPORT WEIGHT", "支架重量"], [])),
    ("V", (["PIPE SURFACE AREA", "管道表面积"], [])),
    ("PAINT", (["PAINTING", "油漆"], [])), ("TRACE", (["TRACING", "伴热"], [])),
    ("INSUL", (["INSULATION", "保温"], [])),
    ("W", (["PIPELINE LENGTH", "管道延长米"], [])), ("X", (["D5 CODE", "D5代码"], [])),
    ("Y", (["UNIT RATE", "清单名称"], [])), ("Z", (["UNIT", "单位"], ["UNIT RATE", "清单名称"])),
])


class ColumnMap(object):
    """Which worksheet column holds which field (the old and the 'Priority Unit' templates differ)."""

    def __init__(self, ws):
        heads = {}
        for c in range(1, min(ws.max_column, 60) + 1):
            t = " ".join(str(ws.cell(r, c).value or "") for r in (1, 2))
            heads[c] = re.sub(r"\s+", " ", t).upper()
        self.col = {}
        used = set()
        for key, (inc, exc) in COLUMN_SPEC.items():
            for c, t in heads.items():
                if c in used:
                    continue
                if any(k.upper() in t for k in inc) and not any(k.upper() in t for k in exc):
                    self.col[key] = openpyxl.utils.get_column_letter(c)
                    used.add(c)
                    break
        for key in "ABCDEFGHIJKLMNOPQRSTUV":            # old template: fixed positions as fallback
            if key not in self.col and key not in used:
                self.col[key] = key
        self.ncols = max(openpyxl.utils.column_index_from_string(v) for v in self.col.values())

    def get(self, key):
        return self.col.get(key)

    def read_row(self, ws, r):
        out = {}
        for k, L in self.col.items():
            v = ws["%s%d" % (L, r)].value
            if v is not None and not isinstance(v, (str, int, float)):
                v = None                       # array formulas etc. are template helpers, not data
            out[k] = v
        return out


# ============================================================================ workbook output
def last_used_row(ws, ncols=26):
    for r in range(ws.max_row, FIRST_DATA_ROW - 1, -1):
        if any(ws.cell(r, c).value not in (None, "") for c in range(1, ncols + 1)):
            return r
    return FIRST_DATA_ROW - 1


def write_rows(ws, rows, cmap):
    start = last_used_row(ws, cmap.ncols) + 1
    tmpl = [ws.cell(FIRST_DATA_ROW, c) for c in range(1, cmap.ncols + 1)]
    for i, r in enumerate(rows):
        rr = start + i
        for c in range(1, cmap.ncols + 1):
            cell = ws.cell(rr, c)
            if tmpl[c - 1].has_style:
                cell._style = copy.copy(tmpl[c - 1]._style)
                if str(cell.fill.fgColor.rgb)[-6:] in ("FFFF00", "FFD8A8"):
                    cell.fill = PatternFill(fill_type=None)
        for key, val in r.items():
            L = cmap.get(key) if not key.startswith("_") else None
            if L:
                ws["%s%d" % (L, rr)].value = val
        if r.get("_void"):
            for c in range(1, cmap.ncols + 1):
                ws.cell(rr, c).fill = VOID_FILL
            L = cmap.get("C") or cmap.get("A")
            if L:
                ws["%s%d" % (L, rr)].comment = Comment(
                    "VOID / АННУЛИРОВАН - cancelled drawing (revision VD). Not to be built.", "AutoFill")
        for k, msg in r["_flags"].items():
            keys = ("X", "Y", "Z") if k.startswith("X_") else (k.split("_")[0],)
            for key in keys:
                L = cmap.get(key)
                if not L:
                    continue
                cell = ws["%s%d" % (L, rr)]
                if k.endswith("_unverified"):
                    cell.fill = ORANGE
                else:
                    cell.fill = YELLOW
                    cell.comment = Comment(msg, "AutoFill")
    return start, start + len(rows) - 1


# ============================================================================ self-test
TEST_COLS = ["D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V",
             "PAINT", "TRACE", "INSUL"]


def _norm(v, col):
    if v is None or v == "":
        return ""
    s = str(v).strip()
    if col in ("N", "O", "P", "Q", "S", "M", "U", "V", "I", "G"):
        x = eval_len(s.replace(",", "."))
        if x is not None:
            return "%g" % x
    if col == "E":
        return {"FITTING": "FITTINGS", "VALVES": "VALVE", "FLANGE": "FLANGES", "GASKETS": "GASKET",
                "BOLTS": "BOLT", "NUTS": "NUT"}.get(s.upper(), s.upper())
    if col == "H":
        s = latin_part(s).upper().replace("Φ", "").replace("°", "").replace("℃", "C")
        s = re.sub(r"[^A-Z0-9.]", "", s)
        return re.sub(r"(?<=\D)0(?=\d)", "", s)
    s = s.translate(LOOKALIKE).upper().replace("×", "X")
    return re.sub(r"\s+", "", s)


def selftest(ws, iso, rows, cmap):
    cols = [c for c in TEST_COLS if cmap.get(c)]
    manual = [cmap.read_row(ws, r) for r in range(FIRST_DATA_ROW, ws.max_row + 1)
              if str(ws["%s%d" % (cmap.get("A"), r)].value or "").strip().upper() == iso]

    def sort_support_runs(lst):          # support order on a sheet is not meaningful
        out, i = [], 0
        while i < len(lst):
            j = i
            while j < len(lst) and lst[j].get("E") == "Support" and lst[j].get("G") == lst[i].get("G"):
                j += 1
            if j > i:
                out += sorted(lst[i:j], key=lambda x: str(x.get("T")))
                i = j
            else:
                out.append(lst[i])
                i += 1
        return out
    manual, rows = sort_support_runs(manual), sort_support_runs(rows)
    ref_sheets = {str(_int_or(m.get("G"))) for m in manual if m.get("G") not in (None, "")}
    if ref_sheets:                       # compare only the sheets that were filled in by hand
        rows = [r for r in rows if str(r.get("G")) in ref_sheets]
    key = lambda x: "%s|%s|%s" % (x.get("G"), x.get("E"), comp_key(x.get("H"))[:12])
    sm = difflib.SequenceMatcher(None, [key(m) for m in manual], [key(a) for a in rows], autojunk=False)
    pairs = []
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal" or (tag == "replace" and i2 - i1 == j2 - j1):
            pairs += list(zip(range(i1, i2), range(j1, j2)))
        else:
            pairs += [(i, None) for i in range(i1, i2)] + [(None, j) for j in range(j1, j2)]
    ok = tot = 0
    diffs, percol = [], Counter()
    for i, j in pairs:
        m = manual[i] if i is not None else {}
        a = rows[j] if j is not None else {}
        for c in cols:
            av, mv = _norm(a.get(c), c), _norm(m.get(c), c)
            tot += 1
            if av == mv:
                ok += 1
            else:
                percol[c] += 1
                diffs.append(("ref row %s" % (i + 1) if i is not None else "extra", c, a.get(c), m.get(c)))
    print("  self-test %s: %d auto rows / %d reference rows, %d/%d cells equal (%.1f%%)"
          % (iso, len(rows), len(manual), ok, tot, 100.0 * ok / max(tot, 1)))
    if percol:
        print("    differences per column: " + ", ".join("%s=%d" % kv for kv in sorted(percol.items())))
    for d in diffs[:60]:
        print("    %s col %s: auto=%r  reference=%r" % d)
    return ok, tot


# ============================================================================ learning
def length_key_sheet(wb):
    """v0.9.5: a fitting-length answer key (sheet '管件长度': ISO No. | ... | description | DN |
    'Length, mm 大管径长度' | 'Length, mm 小管径长度'). Returns (sheet, {field: column}) or None."""
    for ws in wb.worksheets:
        head = {}
        for c in next(ws.iter_rows(min_row=1, max_row=1), []):
            h = str(c.value or "").replace("\n", " ").strip()
            u = h.upper()
            if "大管径长度" in h:
                head["M"] = c.column - 1
            elif "小管径长度" in h:
                head["N"] = c.column - 1
            elif u.startswith("ISO"):
                head["iso"] = c.column - 1
            elif "DESCRIPTION" in u:
                head["desc"] = c.column - 1
            elif u.split()[:1] == ["DN"]:          # 'DN\n公称直径'
                head["dn"] = c.column - 1
            elif u.startswith("CATEGORY"):
                head["cat"] = c.column - 1
            elif u.startswith("SUBCATEGORY"):
                head["sub"] = c.column - 1
            elif "中文描述" in h:
                head["zh"] = c.column - 1
        if {"M", "iso", "desc"} <= set(head):
            return ws, head
    return None


def learn_length_key(paths, rules_path):
    """v0.9.5: learn fitting lengths (O, or P / Q for a tee with a branch length), the Chinese names
    of the subcategories and 'Fitting lengths from drawing = Y' for the drawing families of a
    fitting-length answer key. Lengths are filled per drawing family from the same item."""
    import fnmatch
    rows = []
    for p in paths:
        found = length_key_sheet(openpyxl.load_workbook(p))
        if not found:
            raise UserError("No Measurement Table / 计量表格 sheet and no fitting-length key in %s" % p)
        ws, h = found
        for r in ws.iter_rows(min_row=2, values_only=True):
            g = lambda k: r[h[k]] if k in h and h[k] < len(r) else None
            if not g("iso") or not g("desc") or (g("M") in (None, "") and g("N") in (None, "")):
                continue
            rows.append({k: g(k) for k in ("iso", "desc", "dn", "cat", "sub", "zh", "M", "N")})
    if not rows:
        raise UserError("The fitting-length key has no rows with a length")
    names = " + ".join(os.path.basename(p) for p in paths)
    stamp = "learned %s from %s" % (dt.date.today().isoformat(), names)
    rules0 = Rules(rules_path)
    rwb = openpyxl.load_workbook(rules_path)

    def sheet(name):
        if name not in rwb.sheetnames:
            ws_ = rwb.create_sheet(name)
            ws_.append(Rules.SHEETS[name])
        ws_ = rwb[name]
        for i, hd in enumerate(Rules.SHEETS[name], 1):              # add missing (newer) columns
            if ws_.cell(1, i).value in (None, ""):
                ws_.cell(1, i, hd)
            ws_.cell(1, i).font = Font(bold=True, color="FFFFFF")
            ws_.cell(1, i).fill = PatternFill("solid", fgColor="305496")
        return ws_

    # 1. Fitting Length: key -> (O) or (P, Q); several answers -> the most frequent one
    vals, label, fam_of, dn_of = OrderedDict(), {}, {}, {}
    for v in rows:
        k = comp_key(v["desc"])
        if not k:
            continue
        m = re.match(r"[A-Z0-9]+-([A-Z0-9]+)-", str(v["iso"]).upper())
        fam_of[k] = "*-%s-*" % m.group(1) if m else None
        dn_of[k] = str(v["dn"]).strip() if v["dn"] not in (None, "") else None
        tgt = (None, v["M"], v["N"]) if v["N"] not in (None, "") else (v["M"], None, None)
        vals.setdefault(k, Counter())[tuple(_num_or_text(x) if x not in (None, "") else None for x in tgt)] += 1
        label[k] = re.sub(r"\s+", " ", str(v["desc"])).strip()
    ws_f = sheet("Fitting Length")
    # rows of other subcontractors stay as they are (key = component + family)
    old = OrderedDict(((comp_key(r[0]), (str(r[7]).upper() if len(r) > 7 and r[7] else None)), (list(r) + [None] * 9)[:9])
                      for r in ws_f.iter_rows(min_row=2, values_only=True) if r and r[0])
    for k, cnt in vals.items():
        byval = Counter()
        for (o, p, q), n in cnt.items():
            byval[(eval_len(o), eval_len(p), eval_len(q))] += n
        best_v = byval.most_common(1)[0][0]
        best = next(t for t, n in cnt.most_common() if (eval_len(t[0]), eval_len(t[1]), eval_len(t[2])) == best_v)
        others = ", ".join("/".join(str(x) for x in t if x is not None) + " x%d" % n
                           for t, n in cnt.items() if (eval_len(t[0]), eval_len(t[1]), eval_len(t[2])) != best_v)
        note = stamp + (("; also seen: " + others) if others else "")
        old[(k, fam_of[k])] = [label[k], best[0], sum(cnt.values()), "Y", note, best[1], best[2], fam_of[k], dn_of[k]]
    _rewrite(ws_f, list(old.values()))

    # 2. Family Names where the current rules name a subcategory differently from the key
    fams = OrderedDict()
    for v in rows:
        m = re.match(r"[A-Z0-9]+-([A-Z0-9]+)-", str(v["iso"]).upper())
        if m:
            fams.setdefault("*-%s-*" % m.group(1), []).append(v)
    ws_n = sheet("Family Names")
    have = {(str(r[0]).upper(), str(r[1]).upper()) for r in ws_n.iter_rows(min_row=2, values_only=True) if r and r[0]}
    n_names = 0
    for pat, vs in fams.items():
        seen = OrderedDict()
        for v in vs:
            sub, zh = str(v["sub"] or "").strip(), str(v["zh"] or "").strip()
            u = latin_part(v["desc"]).upper()
            if not sub or not zh or sub.upper() not in u:
                continue
            e, f, _ = rules0.classify(v["desc"], None, "CAD", v["iso"])
            if f == zh:
                continue
            cat = "Fittings" if str(v["cat"] or "").strip().upper().startswith("FITTING") else e
            if cat:
                seen.setdefault(sub.upper(), (cat, zh))
        for kw, (cat, zh) in seen.items():
            if (pat, kw) not in have:
                ws_n.append([pat, kw, cat, zh, "Y", stamp])
                have.add((pat, kw))
                n_names += 1

    # 3. Settings: 'Fitting lengths from drawing' = Y for these families
    ws_s = sheet("Settings")
    head = [str(c.value or "") for c in ws_s[1]]
    col = next(i for i, hd in enumerate(head, 1) if hd.upper().startswith("FITTING LENGTHS"))
    for pat in fams:
        row = next((r for r in ws_s.iter_rows(min_row=2) if str(r[0].value or "").strip().upper() == pat), None)
        if row is None:
            # a new row on top (the first matching pattern wins), with the settings the family
            # gets today ('*-TKM1-ISO-*' of the Priority Unit must not decide for IOT TKM1 sheets)
            now = next((r for r in ws_s.iter_rows(min_row=2, values_only=True) if r[0] and str(r[0]).strip() == "*"), None)
            new = list(now or [None] * len(head))[:len(head)] + [None] * max(0, len(head) - len(now or []))
            new[0] = pat
            new[head.index("Note")] = stamp + " (other settings copied from the '*' row)"
            ws_s.insert_rows(2)
            for i, val in enumerate(new, 1):
                ws_s.cell(2, i, val)
            row = list(ws_s.iter_rows(min_row=2, max_row=2))[0]
        row[col - 1].value = "Y"
    rwb.save(rules_path)
    print("Updated %s from the fitting-length key %s: %d rows of %d drawing(s) -> %d fitting lengths, "
          "%d family names, families %s." % (os.path.basename(rules_path), names, len(rows),
                                            len({v['iso'] for v in rows}), len(vals), n_names, ", ".join(fams)))
    return {"kind": "learn", "rules": rules_path, "rows": len(rows), "drawings": len({v["iso"] for v in rows}),
            "d5": 0, "fitting": len(vals), "valve": 0, "support": 0, "output": rules_path}


def learn(ref_paths, rules_path):
    """Read every filled row of one or more reference Measurement Tables and update rules.xlsx.
    Rows of the same ISO No. in several files: the file given last wins.
    v0.9.5: fitting-length answer keys ('管件长度') are learned by learn_length_key."""
    if isinstance(ref_paths, str):
        ref_paths = [ref_paths]
    keys, tables = [], []
    for p in ref_paths:
        wb = openpyxl.load_workbook(p, read_only=True)
        try:
            (tables if find_sheet(wb.sheetnames, SHEET_NAMES) else keys).append(p)
        finally:
            wb.close()
    res = learn_length_key(keys, rules_path) if keys else None
    if not tables:
        return res
    ref_paths = tables
    sheets_ = []
    for ref_path in ref_paths:
        wb = openpyxl.load_workbook(ref_path)
        name = find_sheet(wb.sheetnames, SHEET_NAMES)
        if not name:
            raise UserError("No Measurement Table / 计量表格 sheet in %s" % ref_path)
        sheets_.append(wb[name])
    ref_path = " + ".join(os.path.basename(p) for p in ref_paths)
    canon = {"FITTING": "Fittings", "FITTINGS": "Fittings", "VALVE": "Valve", "VALVES": "Valve",
             "FLANGE": "Flanges", "FLANGES": "Flanges", "PIPE": "Pipe", "PIPES": "Pipe",
             "SUPPORT": "Support", "SUPPORTS": "Support"}
    by_iso, renamed = OrderedDict(), Counter()
    for ws in sheets_:
        this = OrderedDict()
        cm = ColumnMap(ws)
        for r in range(FIRST_DATA_ROW, ws.max_row + 1):
            v = cm.read_row(ws, r)
            for c in COLS:
                v.setdefault(c, None)
            if isinstance(v.get("B"), str) and v["B"].startswith("="):
                v["B"] = None
            if v["A"] and v["E"]:
                e = canon.get(str(v["E"]).strip().upper(), str(v["E"]).strip())
                if e != v["E"]:
                    renamed["%s -> %s" % (v["E"], e)] += 1
                    v["E"] = e
                this.setdefault(str(v["A"]).strip().upper(), []).append(v)
        by_iso.update(this)
    rows = [v for lst in by_iso.values() for v in lst]
    for k, n in renamed.items():
        print("  note: category %s (%d rows) treated as the same category" % (k, n))
    isos = sorted({str(v["A"]) for v in rows})
    print("Learning from %d rows of %d drawing(s): %s" % (len(rows), len(isos), ", ".join(i[-5:] for i in isos)))

    d5 = {}
    for v in rows:
        dn = parse_dn(str(v["I"])) if v["I"] is not None else None
        if v["X"] and dn:
            dtype = "CAD" if re.search(r"[А-Яа-я]", str(v["H"])) else "SCAN"
            d5.setdefault((v["E"], dn[0], dtype), Counter())[(v["X"], v["Y"], v["Z"])] += 1
    fit, valve, last, supw, tee = {}, {}, {}, {}, {}
    rules0 = Rules(rules_path)
    for idx, v in enumerate(rows):
        if v["H"] and v["E"] != "Support":       # same clean-up as when filling (Φ, °C ...)
            v["H"] = tidy_description(str(v["H"]), rules0)
        k = comp_key(v["H"])
        if v["E"] == "Support" and v.get("U") not in (None, ""):
            supw.setdefault(k, [v["H"], Counter()])[1][str(v["U"])] += 1
            last[(k, str(v["U"]))] = idx
        if v["E"] != "Pipe" and v["E"] != "Valve" and v["O"] not in (None, ""):
            fit.setdefault(k, [v["H"], Counter()])[1][str(v["O"])] += 1
            last[(k, str(v["O"]))] = idx
        if v["E"] not in ("Pipe", "Valve", "Support") and (v.get("P") not in (None, "") or v.get("Q") not in (None, "")):
            tee[k] = [v["H"], v.get("P"), v.get("Q")]
        if v["E"] == "Valve" and v["S"] not in (None, ""):
            valve.setdefault(k, [v["H"], Counter()])[1][str(v["S"])] += 1
            last[(k, str(v["S"]))] = idx

    rwb = openpyxl.load_workbook(rules_path)
    stamp = "learned %s from %s" % (dt.date.today().isoformat(), ref_path)

    def sheet(name):
        if name not in rwb.sheetnames:
            ws_ = rwb.create_sheet(name)
            ws_.append(Rules.SHEETS[name])
            for c in ws_[1]:
                c.font = Font(bold=True, color="FFFFFF")
                c.fill = PatternFill("solid", fgColor="305496")
        return rwb[name]

    # D5: learned exact-DN rows go first (first match wins), old rows stay below
    ws_d5 = sheet("D5 Codes")
    old = [list(r) for r in ws_d5.iter_rows(min_row=2, values_only=True) if r and r[0]]
    learned_keys = set()
    new = []
    for (cat, dn, dtype), cnt in sorted(d5.items(), key=lambda kv: (str(kv[0][0]), kv[0][1], kv[0][2])):
        (code, desc, unit), n = cnt.most_common(1)[0]
        note = "%s, %d row(s)" % (stamp, n)
        others = Counter()
        for k, c in cnt.items():
            if k[0] != code:
                others[k[0]] += c
        if others:
            note += "; also used: " + ", ".join("%s x%d" % kv for kv in others.items())
        new.append([cat, dn, dn, code, desc, unit, "Y", note, dtype])
        learned_keys.add((str(cat).lower(), dn, dtype))
    old = [(r + [None] * 9)[:9] for r in old]
    keep = [r for r in old if not (str(r[0]).lower(), _num(r[1], -1), str(r[8] or "").upper()) in learned_keys
            or _num(r[1], -1) != _num(r[2], -2)]
    _rewrite(ws_d5, new + keep)

    def table(name, data, valve_mode=False, extra=None):
        ws_ = sheet(name)
        if name == "Fitting Length" and ws_.max_column < 7:          # older rules.xlsx: add P/Q headers
            for i, h in enumerate(Rules.SHEETS[name][5:], 6):
                c = ws_.cell(1, i, h)
                c.font = Font(bold=True, color="FFFFFF")
                c.fill = PatternFill("solid", fgColor="305496")
        old = OrderedDict()
        fam_rows = []           # v0.9.5: rows learned for one family (col H) are kept as they are
        for r in ws_.iter_rows(min_row=2, values_only=True):
            if r and r[0]:
                if name == "Fitting Length" and len(r) > 7 and r[7]:
                    fam_rows.append(list(r))
                else:
                    old[comp_key(r[0])] = list(r)
        for k, (label, cnt) in data.items():
            # group spellings of the same length ('=27*2' and 54) by value
            # most frequent value wins; among its spellings / ties, the most recent row wins
            byval = OrderedDict()
            for val, n in sorted(cnt.items(), key=lambda kv: (-kv[1], len(kv[0]), -last[(k, kv[0])])):
                byval.setdefault(eval_len(val), []).append((val, n))
            ranked = sorted(byval.values(), key=lambda lst: (-sum(n for _, n in lst),
                                                             -max(last[(k, v)] for v, _ in lst)))
            best = ranked[0][0][0]
            seen = sum(cnt.values())
            others = ", ".join("%s x%d" % (lst[0][0], sum(n for _, n in lst)) for lst in ranked[1:])
            label = re.sub(r"\s+", " ", str(label)).strip()
            if valve_mode:
                old[k] = [label, _num_or_text(best), seen, others, "Y", stamp]
            else:
                note = stamp + (("; also seen: " + others) if others else "")
                old[k] = [label, _num_or_text(best), seen, "Y", note]
        for k, (label, p, q) in (extra or {}).items():
            row = (old.get(k) or [re.sub(r"\s+", " ", str(label)).strip(), None, 1, "Y", stamp]) + [None] * 7
            row[5], row[6] = _num_or_text(p), _num_or_text(q)
            old[k] = row[:7]
        _rewrite(ws_, list(old.values()) + fam_rows)
        return len(data) + len([k for k in (extra or {}) if k not in data])

    nf = table("Fitting Length", fit, extra=tee)
    nv = table("Valve Length", valve, valve_mode=True)
    nw = table("Support Weight", supw)

    # mark Categories rows confirmed by the reference as Verified
    rules = Rules(rules_path)
    ws_c = sheet("Categories")
    confirmed, conflict = set(), {}
    for v in rows:
        if not v["F"] or v["E"] == "Support":
            continue
        dtype = "CAD" if re.search(r"[А-Яа-я]", str(v["H"])) else "SCAN"
        u = latin_part(v["H"]).upper() if dtype == "CAD" else str(v["H"]).upper()
        for kw, e, f, ver, scope, _ in sorted(rules.categories, key=lambda c: 0 if c[4] == dtype else 1):
            if scope and scope != dtype:
                continue
            if re.search(r"(?<![A-Z])%s(?![A-Z])" % re.escape(kw), u) and (not e or e == v["E"]):
                if f == v["F"]:
                    confirmed.add(kw)
                else:
                    conflict.setdefault(kw, Counter())[v["F"]] += 1
                break
    for row in ws_c.iter_rows(min_row=2):
        kw = str(row[0].value or "").upper()
        if kw in confirmed and not _yes(row[3].value):
            row[3].value = "Y"
            row[4].value = ((row[4].value or "") + "; confirmed by " + ref_path).strip("; ")
        if kw in conflict:
            print("  note: keyword %s -> %s in rules.xlsx, but the reference also uses %s"
                  % (kw, row[2].value, ", ".join("%s x%d" % kv for kv in conflict[kw].items())))

    ns, nk, nd = _learn_v09(rows, rwb, rules0, stamp, sheet)
    rwb.save(rules_path)
    print("Updated %s: %d D5 rows, %d fitting lengths, %d valve lengths, %d support weights, "
          "%d family settings, %d new keywords, %d D5 sections." % (
              os.path.basename(rules_path), len(new), nf, nv, nw, ns, nk, nd))
    out = {"kind": "learn", "rules": rules_path, "rows": len(rows), "drawings": len(isos), "d5": len(new),
           "fitting": nf, "valve": nv, "support": nw, "output": rules_path}
    if res:                                  # a fitting-length key learned in the same run
        for k in ("rows", "drawings", "fitting"):
            out[k] += res[k]
    return out


LEARN_STOP = {"ASME", "SEAMLESS", "TEST", "KCV", "WELDED", "CLASS", "WITH", "PACKAGE", "VALVE", "ONE",
              "TYPE", "MSS", "EN", "GOST", "API", "BW", "SW", "RF", "FF", "NPT", "SCH", "STD", "WPL", "LCB",
              "GR", "LF2", "A333", "A350", "A420", "CL1", "PE", "BE", "TS", "DN", "AND", "FOR", "THE"}


def _family_pattern(iso):
    m = re.search(r"-(\d{5}-\d\d)-\d{4}-", str(iso))
    return "*-%s-*" % m.group(1) if m else None


def _learn_v09(rows, rwb, rules0, stamp, sheet):
    """v0.9 learning: (1) Settings per drawing family from how the reference fills it,
    (2) new Categories keywords where the reference consistently uses another category / Chinese
    name than the current rules, (3) the D5 code section per family and category."""
    fam = OrderedDict()
    for v in rows:
        pat = _family_pattern(v["A"])
        if pat:
            fam.setdefault(pat, []).append(v)
    # (1) Settings
    ws_s = sheet("Settings")
    head = [str(c.value or "") for c in ws_s[1]]
    for i, h in enumerate(Rules.SHEETS["Settings"], 1):          # add the v0.9 columns
        if h not in head:
            ws_s.cell(1, len(head) + 1, h).font = Font(bold=True, color="FFFFFF")
            ws_s.cell(1, len(head) + 1).fill = PatternFill("solid", fgColor="305496")
            head.append(h)
    col = {h: i for i, h in enumerate(head)}
    old = [list(r) + [None] * (len(head) - len(r)) for r in ws_s.iter_rows(min_row=2, values_only=True) if r and r[0]]
    yn = lambda b: "Y" if b else "N"
    learned = []
    for pat, lst in fam.items():
        base = next((r for r in old if str(r[0]).upper() == pat), None)
        if base is None:
            import fnmatch
            base = next((list(r) for r in old if fnmatch.fnmatch(pat.strip("*").join("AB"), str(r[0]).upper())),
                        [None] * len(head))
        row = list(base) + [None] * (len(head) - len(base))
        row[0] = pat
        sup = [v for v in lst if v["E"] == "Support"]
        if sup:
            row[col["Support QTY = 1 (Y/N)"]] = yn(sum(1 for v in sup if str(v["M"]).strip() in ("1", "1.0")) >= 0.8 * len(sup))
            # v0.9.2: one row per support = each support row carries one support number
            nums = [v for v in sup if v["T"] not in (None, "")]
            if len(nums) >= 3:
                one = sum(1 for v in nums if len(SUP_NO_CAD.findall(str(v["T"]).upper())) <= 1)
                row[col["One row per support (Y/N)"]] = yn(one >= 0.8 * len(nums))
        row[col["Fill D5 codes (Y/N)"]] = yn(any(v["X"] for v in lst))
        dl = [str(v["D"]) for v in lst if v["D"]]
        if dl:
            row[col["Pipe line no. short (Y/N)"]] = yn(sum(1 for d in dl if d.count("-") <= 2) >= 0.8 * len(dl))
        withdn = [v for v in lst if parse_dn(str(v["I"] or ""))]
        if len(withdn) >= 5:
            share = sum(1 for v in withdn if v["J"] not in (None, "")) / float(len(withdn))
            row[col["DN in J on every row (Y/N)"]] = yn(share >= 0.8)
        withl = [v for v in lst if v["L"] not in (None, "") and v["H"]]
        if len(withl) >= 5:
            norm = lambda t: re.sub(r"[^A-Z0-9]", "", latin_part(str(t)).upper().replace("0", "O"))
            share = sum(1 for v in withl if norm(v["L"]) and norm(v["L"]) in norm(v["H"])) / float(len(withl))
            row[col["Item code from description (Y/N)"]] = yn(share >= 0.6)
        if len(sup) >= 3:
            row[col["Support weight (Y/N)"]] = yn(sum(1 for v in sup if v["U"] not in (None, "")) >= 0.2 * len(sup))
        lines = {(v["A"], v["D"]) for v in lst}
        areas = [v for v in lst if v["V"] not in (None, "")]
        if len(lst) >= 10 and lines:
            row[col["Surface area once per line (Y/N)"]] = yn(len(areas) <= 1.3 * len(lines))
            pdat = [v for v in lst if v.get("PAINT") not in (None, "")]
            if pdat:
                row[col["Painting / tracing / insulation once per line (Y/N)"]] = yn(len(pdat) <= 1.3 * len(lines))
        eq = [v for v in lst if re.fullmatch(r"\s*(\d+(?:\.\d+)?)\s*[XxХх×]\s*\1\s*", str(v["I"] or ""))]
        if eq:
            row[col["J/K also for equal sizes (Y/N)"]] = yn(sum(1 for v in eq if v["J"] not in (None, "")) >= 0.5 * len(eq))
        row[col["Note"]] = "%s, %d row(s)" % (stamp, len(lst))
        learned.append(row)
    keep = [r for r in old if str(r[0]).upper() not in fam]
    _rewrite(ws_s, learned + keep)

    # (2) names per family: keyword in the description -> category + Chinese name, where the reference of
    #     this family consistently disagrees with the general Categories rules
    ws_c = sheet("Family Names")
    add = []
    for pat, lst in fam.items():
        stats, cur = {}, {}
        for v in lst:
            if not v["F"] or not v["H"] or v["E"] == "Support":
                continue
            e = {"FITTING": "Fittings", "VALVES": "Valve"}.get(str(v["E"]).upper(), v["E"])
            u = latin_part(str(v["H"])).upper()
            got = rules0.classify(str(v["H"]), None)
            for t in {t for t in re.findall(r"[A-Z][A-Z-]{3,}", u) if t not in LEARN_STOP and not t.startswith("GR")}:
                stats.setdefault(t, Counter())[(e, v["F"])] += 1
                cur.setdefault(t, Counter())["ok" if got and got[0] == e and got[1] == v["F"] else "bad"] += 1
        best = {}
        for t, st in stats.items():
            (e, f), n = st.most_common(1)[0]
            tot = sum(st.values())
            if tot >= 2 and n >= 0.8 * tot and cur[t]["bad"] >= 0.5 * tot:
                best.setdefault((e, f), []).append((t, n, tot))
        for (e, f), cand in best.items():
            for t, n, tot in sorted(cand, key=lambda x: (-x[1], x[2] - x[1], -len(x[0])))[:2]:
                add.append([pat, t, e, f, "Y", "%s: %d of %d row(s) of this family with '%s'" % (stamp, n, tot, t)])
    # v0.9.2: the Chinese name of supports per family (e.g. 管架 instead of 支架)
    for pat, lst in fam.items():
        sf = Counter(str(v["F"]).strip() for v in lst if v["E"] == "Support" and v["F"])
        if sum(sf.values()) >= 3:
            f, n = sf.most_common(1)[0]
            cur = rules0.classify("SUPPORT", "Support")
            if n >= 0.8 * sum(sf.values()) and (not cur or cur[1] != f):
                for kw in ("SUPPORT", "SUP"):
                    add.append([pat, kw, "Support", f, "Y", "%s: %d of %d support row(s) of this family named %s"
                                % (stamp, n, sum(sf.values()), f)])
    keyf = {(r[0], r[1]) for r in add}
    oldc = [list(r) + [None] * 6 for r in ws_c.iter_rows(min_row=2, values_only=True) if r and r[0]]
    _rewrite(ws_c, add + [r[:6] for r in oldc if (r[0], r[1]) not in keyf])

    # (3) D5 sections: family + category -> section of the code list
    ws_d = sheet("D5 Sections")
    rules_list = Rules.__new__(Rules)
    try:
        d5_list = OrderedDict((str(r[0]).strip(), r) for r in Rules._rows(rwb, "D5 List") if r[0])
    except Exception:
        d5_list = {}
    sec = {}
    for pat, lst in fam.items():
        for v in lst:
            if not v["X"]:
                continue
            code = str(v["X"]).strip()
            e = {"FITTING": "Fittings", "VALVES": "Valve"}.get(str(v["E"]).upper(), v["E"])
            info = d5_list.get(code)
            section = code.rsplit("-", 1)[0] if info is not None and info[3] is not None else code
            sec.setdefault((pat, e), Counter())[section] += 1
    oldd = [list(r) + [None] * 5 for r in ws_d.iter_rows(min_row=2, values_only=True) if r and r[0]]
    newd = []
    for (pat, e), cnt in sec.items():
        section, n = cnt.most_common(1)[0]
        newd.append([pat, e, section, "Y" if n >= 0.8 * sum(cnt.values()) else None,
                     "%s, %d of %d row(s)" % (stamp, n, sum(cnt.values()))])
    keyd = {(str(r[0]).upper(), str(r[1]).lower()) for r in newd}
    _rewrite(ws_d, newd + [r[:5] for r in oldd if (str(r[0]).upper(), str(r[1]).lower()) not in keyd])
    return len(learned), len(add), len(newd)


def import_d5_list(src, rules_path):
    """v0.9: copy the official D5 code list (code, unit, description) into rules.xlsx > D5 List,
    with the DN band read from the description ('DN 125-200 mm', 'Diameter 0-25mm')."""
    wb = openpyxl.load_workbook(src, data_only=True)
    ws = wb.worksheets[0]
    out = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        code = str(r[0] or "").strip()
        if not re.fullmatch(r"\d{4}(-\d\d){1,3}", code):
            continue
        desc = str(r[3] or r[2] or "").strip()
        m = re.search(r"(?:DN|Diameter|diameter)\s*(\d+)\s*-\s*(\d+)", desc) or \
            re.search(r"installation\s+(\d+)\s*-\s*(\d+)\s*mm", desc)
        lo, hi = (int(m.group(1)), int(m.group(2))) if m else (None, None)
        out.append([code, r[1], desc, lo, hi, code.rsplit("-", 1)[0]])
    rwb = openpyxl.load_workbook(rules_path)
    if "D5 List" not in rwb.sheetnames:
        w = rwb.create_sheet("D5 List")
        w.append(Rules.SHEETS["D5 List"])
        for c in w[1]:
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor="305496")
    _rewrite(rwb["D5 List"], out)
    rwb.save(rules_path)
    print("D5 List: %d codes copied from %s" % (len(out), os.path.basename(src)))
    return len(out)


def _num_or_text(v):
    try:
        f = float(v)
        return int(f) if f.is_integer() else f
    except (TypeError, ValueError):
        return v


def _rewrite(ws, rows):
    ws.delete_rows(2, ws.max_row)
    for r in rows:
        ws.append(list(r))


# ============================================================================ main
def has_measurement_sheet(path):
    wb = openpyxl.load_workbook(path, read_only=True)
    try:
        return bool(find_sheet(wb.sheetnames, SHEET_NAMES))
    finally:
        wb.close()                       # read-only mode keeps the file open (locks it on Windows)


def find_workbook(folder, out_dir):
    found = glob.glob(os.path.join(folder, "**", "*.xlsx"), recursive=True)
    for p in sorted(found, key=lambda p: (os.path.relpath(p, folder).count(os.sep), p)):   # top level first
        if os.path.basename(p).startswith("~$") or os.path.abspath(p).startswith(os.path.abspath(out_dir)):
            continue
        try:
            if has_measurement_sheet(p):
                return p
        except Exception:
            pass
    return None


def main():
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(errors="replace")
        except Exception:
            pass
    ap = argparse.ArgumentParser(description="Fill the Measurement Table from isometric PDFs")
    made = ensure_workspace()
    if made and DATA != HERE:
        print("Workspace: %s  (created: %s)" % (DATA, ", ".join(os.path.basename(m) or m for m in made)))
    ap.add_argument("folder", nargs="?", default=os.path.join(DATA, "input"))
    ap.add_argument("--workbook")
    ap.add_argument("--engine", default="auto", choices=["auto", "rapidocr", "tesseract"])
    ap.add_argument("--rules", default=os.path.join(DATA, "rules.xlsx"))
    ap.add_argument("--out", default=os.path.join(DATA, "output"))
    ap.add_argument("--force", action="store_true")
    ap.add_argument("--selftest", action="store_true")
    ap.add_argument("--import-d5", default=None, help="copy the official D5 code list (xlsx) into rules.xlsx")
    ap.add_argument("--learn", metavar="REFERENCE_XLSX", nargs="+")
    ap.add_argument("--jobs", type=int, default=max(1, min(6, (os.cpu_count() or 2) - 1)),
                    help="pages read in parallel (default: CPU cores - 1, max 6)")
    interactive = len(sys.argv) == 1
    if interactive:
        sys.argv += menu()
    a = ap.parse_args()
    try:
        run(a)
    except UserError as e:
        print("\n" + str(e))
        if not interactive:
            sys.exit(1)
    finally:
        for p in _POOLS:
            p.terminate()
        if interactive:
            try:
                input("\nPress Enter to close...")
            except EOFError:
                pass


def menu():
    print("Process Piping Measurement - AutoFill v%s" % VERSION)
    print("  1  Fill the Measurement Table from the drawings in the 'input' folder")
    print("  2  Self-test: compare with drawings already filled by hand (writes nothing)")
    print("  3  Learn from a workbook with correct answers")
    choice = input("Choose 1, 2 or 3 [1]: ").strip() or "1"
    folder = os.path.join(DATA, "input")
    if choice == "2":
        return [folder, "--selftest"]
    if choice == "3":
        p = input("Drag the reference workbook(s) into this window, then press Enter:\n").strip()
        files = [x.strip('"') for x in re.findall(r'"[^"]+"|\S+', p)]
        return ["--learn"] + files
    return [folder]


def run(a):

    if getattr(a, "import_d5", None):
        return [{"kind": "import", "rows": import_d5_list(a.import_d5, a.rules), "output": a.rules}]
    if a.learn:
        return [learn(a.learn, a.rules)]

    folder = os.path.abspath(a.folder)
    if getattr(a, "pdfs", None):                      # GUI: individual PDF files dropped
        pdfs = [os.path.abspath(p) for p in a.pdfs]
        folder = os.path.commonpath([os.path.dirname(p) for p in pdfs])
    elif os.path.isfile(folder) and folder.lower().endswith(".pdf"):
        pdfs, folder = [folder], os.path.dirname(folder)
    else:
        pdfs = sorted(glob.glob(os.path.join(folder, "**", "*.pdf"), recursive=True))
    if not pdfs:
        raise UserError("No PDF files found in %s" % folder)
    groups = group_by_workbook(folder, pdfs, a.workbook, a.out, a.selftest)
    if not groups:
        raise UserError('No workbook with a "%s" sheet found in %s' % ('" / "'.join(SHEET_NAMES), folder))

    engine = make_engine(a.engine)
    rules = Rules(a.rules)
    pool = None
    results = []
    try:
        if a.jobs > 1:
            import multiprocessing
            choice = "rapidocr" if engine.name.startswith("rapid") else "tesseract"
            pool = multiprocessing.Pool(a.jobs, initializer=_init_worker, initargs=(choice,))
            _POOLS.append(pool)
        print("AutoFill v%s | OCR engine: %s | parallel: %d" % (VERSION, engine.name, a.jobs))
        if len(groups) > 1:
            print("%d project folders found - each is filled into its own workbook." % len(groups))
        total = sum(len(g) for _, g in groups)
        emit("start", total=total, groups=len(groups), engine=engine.name)
        done = 0
        for wbp, group in groups:
            print("")
            results.append(run_group(a, wbp, group, engine, rules, pool, offset=done, total=total))
            done += len(group)
    finally:
        if pool is not None:
            pool.terminate()
            if pool in _POOLS:
                _POOLS.remove(pool)
    return results


def group_by_workbook(folder, pdfs, workbook, out_dir, selftest=False):
    """One project per folder: every PDF goes with the workbook in its own folder or the nearest
    folder above it. Answer-key workbooks (…Fine_Tuning…) are used only for the self-test."""
    if workbook:
        return [(workbook, pdfs)]
    books = []
    for p in glob.glob(os.path.join(folder, "**", "*.xlsx"), recursive=True):
        if os.path.basename(p).startswith("~$") or os.path.abspath(p).startswith(os.path.abspath(out_dir)):
            continue
        try:
            if has_measurement_sheet(p):
                books.append(os.path.abspath(p))
        except Exception:
            pass
    if not books:
        return []
    top = find_workbook(folder, out_dir)
    top = os.path.abspath(top) if top else books[0]
    groups = OrderedDict()
    for pdf in pdfs:
        d = os.path.dirname(os.path.abspath(pdf))
        own = [b for b in books if (d + os.sep).startswith(os.path.dirname(b) + os.sep)]
        # deepest folder wins; several workbooks in one folder -> the usual (alphabetical) choice
        ref = lambda b: 1 if re.search(r"fine.?tuning|supervised|spuervised|reference|SFT", os.path.basename(b), re.I) else 0
        best = max(own, key=lambda b: (os.path.dirname(b).count(os.sep), ref(b) if selftest else -ref(b),
                                      -sorted(books).index(b))) if own else top
        groups.setdefault(best, []).append(pdf)
    return list(groups.items())


def run_group(a, wbp, pdfs, engine, rules, pool, offset=0, total=None):
    """Fill (or self-test) one workbook. Returns a summary dict for the GUI."""
    total = total or len(pdfs)
    res = {"kind": "selftest" if a.selftest else "fill", "workbook": wbp, "output": None, "drawings": 0,
           "rows": 0, "checks": 0, "skipped": [], "errors": [], "score": None}
    print("Workbook : %s" % wbp)
    print("Drawings : %d PDF file(s)" % len(pdfs))

    wb = openpyxl.load_workbook(wbp)
    ws = wb[find_sheet(wb.sheetnames, SHEET_NAMES)]
    cmap = ColumnMap(ws)
    roster = roster_lookup(wb)
    ca = cmap.get("A")
    existing = {str(ws["%s%d" % (ca, r)].value).strip().upper()
                for r in range(FIRST_DATA_ROW, ws.max_row + 1) if ws["%s%d" % (ca, r)].value}
    log_rows = []

    def log(iso, sheet, level, msg):
        log_rows.append((iso, sheet, level, msg))

    debug_dir = os.path.join(a.out, "ocr_debug")
    added, score = 0, [0, 0]
    for n, pdf in enumerate(pdfs, 1):
        iso, _ = iso_from_filename(pdf)
        emit("drawing", index=offset + n, total=total, name=os.path.basename(pdf), iso=iso)
        if a.selftest and iso not in existing:
            res["skipped"].append(iso)
            continue
        if not a.selftest and iso in existing and not a.force:
            print("[%d/%d] %s  -> already in the table, skipped (use --force to add anyway)"
                  % (n, len(pdfs), iso))
            log(iso, "", "INFO", "Already in Measurement Table - skipped")
            res["skipped"].append(iso)
            continue
        print("[%d/%d] %s" % (n, len(pdfs), os.path.basename(pdf)), flush=True)
        try:
            iso, rows = process_pdf(pdf, engine, rules, roster, log, debug_dir, pool)
        except Cancelled:
            raise
        except Exception as e:
            import traceback
            traceback.print_exc()
            log(iso or os.path.basename(pdf), "", "ERROR", str(e))
            res["errors"].append("%s: %s" % (iso or os.path.basename(pdf), e))
            continue
        res["errors"] += ["%s: %s" % (x[0], x[3]) for x in log_rows if x[2] == "ERROR" and x[0] == iso]
        if not rows:
            continue
        res["drawings"] += 1
        if a.selftest:
            o, t = selftest(ws, iso, rows, cmap)
            score[0] += o
            score[1] += t
            emit("drawing_done", iso=iso, rows=len(rows), equal=o, cells=t)
            continue
        r0, r1 = write_rows(ws, rows, cmap)
        added += len(rows)
        flagged = sum(1 for r in rows for k in r["_flags"] if not k.endswith("_unverified"))
        res["checks"] += flagged
        print("    %d rows -> rows %d-%d, %d cell(s) to check" % (len(rows), r0, r1, flagged))
        emit("drawing_done", iso=iso, rows=len(rows), checks=flagged)
    res["rows"] = added

    if a.selftest:
        if score[1]:
            print("\nOVERALL: %d/%d cells equal (%.1f%%)  [engine: %s]"
                  % (score[0], score[1], 100.0 * score[0] / score[1], engine.name))
            res["score"] = (score[0], score[1])
        print("OCR text of every page saved in: %s" % debug_dir)
        return res
    if added == 0:
        print("\nNothing new to add - no output file written.")
        return res
    if "AutoFill Log" in wb.sheetnames:
        del wb["AutoFill Log"]
    lw = wb.create_sheet("AutoFill Log")
    lw.append(["Time", "ISO No.", "Sheet", "Level", "Message"])
    now = dt.datetime.now().strftime("%Y-%m-%d %H:%M")
    lw.append([now, "", "", "INFO", "AutoFill v%s, OCR engine: %s" % (VERSION, engine.name)])
    for row in log_rows:
        lw.append([now] + list(row))
    for c, w in zip("ABCDE", (17, 44, 7, 8, 100)):
        lw.column_dimensions[c].width = w
    os.makedirs(a.out, exist_ok=True)
    stem = os.path.splitext(os.path.basename(wbp))[0]
    outp = os.path.join(a.out, "%s_autofill_%s.xlsx" % (stem, dt.datetime.now().strftime("%Y%m%d_%H%M")))
    k = 2
    while os.path.exists(outp):
        outp = os.path.join(a.out, "%s_autofill_%s_%d.xlsx" % (stem, dt.datetime.now().strftime("%Y%m%d_%H%M"), k))
        k += 1
    wb.save(outp)
    res["output"] = outp
    print("\nDone. %d row(s) added. Saved to:\n  %s" % (added, outp))
    print("Yellow = please check (hover for the reason). Orange = rule not yet marked verified in rules.xlsx.")
    return res


if __name__ == "__main__":
    import multiprocessing
    multiprocessing.freeze_support()
    main()
