# -*- coding: utf-8 -*-
"""מנוע הניתוח: קורא דו"ח Word, מאתר סעיפים ותמונות, שולח ל-Claude Code (מנוי) וכותב את הניתוח למסמך."""
import base64
import copy
import io
import json
import os
import re
import shutil
import subprocess
from concurrent.futures import ThreadPoolExecutor, as_completed

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

import sys
import time
BASE = os.path.dirname(sys.executable if getattr(sys, "frozen", False) else os.path.abspath(__file__))
INSTR_DIR = os.path.join(BASE, "הוראות")
EXAMPLES_DIR = os.path.join(BASE, "דוגמאות")
MC_FALLBACK = "{http://schemas.openxmlformats.org/markup-compatibility/2006}Fallback"
BLIP = "{http://schemas.openxmlformats.org/drawingml/2006/main}blip"
R_EMBED = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed"
AI_PREFIX = "AI_"
CELL_PREFIX = "AICELL_"
STATICS_LABELS = ("Kyphosis", "Scoliosis", "Lordosis", "Coronal Imbalance", "Sagittal Imbalance", "Stiffness",
                  "Fleche Cervicale",
                  "Pelvic Rotation", "Pelvic Torsion", "Pelvic Obliquity", "instability",
                  "Vertebral Rotation", "Apical Deviation", "Trunk Length")
STATICS_CHECK_ONLY = ("Stiffness", "instability")    # שיפוט קליני של הרופא - הכלי רק בודק, לא דורס
STATICS_EXTRA = ("Vertebral Rotation", "Apical Deviation", "Trunk Length")  # שורה שהכלי מוסיף לטבלה אם חסרה
TABLE_PREFIX = "AITBL_"
STAB_TABLE = {"key": "stab_table", "title": "טבלת Stabilimeter מאוחדת"}
STATICS = {"key": "statics", "title": "טבלת מדדים סטטיים - מילוי ובדיקה"}


def fix_bidi(text):
    """סימני כיוון בלתי נראים (RLM) כדי שבשורה עברית "L2 11 מ"מ ימינה" ו-"L1 5° שמאלה" יוצגו בסדר הנכון."""
    rlm = "\u200f"
    text = re.sub(r"\b([A-Za-z]{1,3}\d{0,2})(\s+)(?=[-\u2212]?\d)", lambda m: m.group(1) + rlm + m.group(2), text)
    return text.replace("°", "°" + rlm).replace(rlm + rlm, rlm)


def _same_key(found, key):
    """True אם found הוא key עצמו או ניתוח-תמונה שלו (key#n)."""
    return bool(found) and (found == key or found.startswith(key + "#"))


def _cell_marked(p):
    return any((b.get(qn("w:name")) or "").startswith(CELL_PREFIX) for b in p.iter(qn("w:bookmarkStart")))


def load_json(name, default=None):
    path = os.path.join(BASE, name)
    if not os.path.exists(path):
        return default
    with open(path, encoding="utf-8") as f:
        return json.load(f)


SPEC = load_json("sections.json")
CONFIG = {"model": "opus", "max_parallel": 3, "timeout_sec": 420, "max_examples": 2}
CONFIG.update(load_json("config.json", {}) or {})


# ---------- קריאת המסמך ----------

def _in_fallback(el):
    return any(a.tag == MC_FALLBACK for a in el.iterancestors())


def para_text(p):
    """טקסט הפסקה, בלי טקסט של תיבות טקסט מקוננות."""
    out = []
    for t in p.iter(qn("w:t")):
        owner = next(a for a in t.iterancestors() if a.tag == qn("w:p"))
        if owner is p:
            out.append(t.text or "")
    return "".join(out).strip()


def para_images(p):
    return [b.get(R_EMBED) for b in p.iter(BLIP) if b.get(R_EMBED) and not _in_fallback(b)]


def ai_key(p):
    """מחזיר את מפתח הסעיף אם הפסקה נכתבה ע"י הכלי."""
    for b in p.iter(qn("w:bookmarkStart")):
        name = b.get(qn("w:name")) or ""
        if name.startswith(AI_PREFIX):
            return name[len(AI_PREFIX):].rsplit("_", 1)[0]
    return None


def _match(patterns, text):
    low = text.lower()
    return any(re.search(p, low) for p in patterns)


class Report:
    def __init__(self, path):
        self.path = path
        self.doc = Document(path)
        self.body = self.doc.element.body
        self.paras = [p for p in self.body.iter(qn("w:p")) if not _in_fallback(p)]
        self.texts = [para_text(p) for p in self.paras]
        self._scan()
        self._identity()

    def _scan(self):
        """מאתר כותרות סעיפים וגבולות."""
        self.heads = {}          # key -> index
        marks = []               # (index, key or None)
        all_specs = SPEC["sections"] + [SPEC["discussion"]]
        for i, t in enumerate(self.texts):
            if not t or len(t) > 110 or ai_key(self.paras[i]):
                continue
            hit = next((s["key"] for s in all_specs
                        if s["key"] not in self.heads and re.search(s["pattern"], t.lower())), None)
            if hit:
                self.heads[hit] = i
                marks.append((i, hit))
            elif _match(SPEC["boundaries"], t):
                marks.append((i, None))
        marks.sort()
        self.ranges = {}
        for n, (i, key) in enumerate(marks):
            if key:
                end = marks[n + 1][0] if n + 1 < len(marks) else len(self.paras)
                self.ranges[key] = (i, end)

    def _identity(self):
        """אוסף פרטים מזהים (שם, ת.ז, ת. לידה, טלפון, מייל) משורת הפרטים בטבלה הראשונה לצורך התממה."""
        self.id_values, self.id_rows = [], []
        tables = self.doc.tables
        if not tables:
            return
        rows = tables[0].rows
        for n, row in enumerate(rows):
            labels = {c.text.strip() for c in row.cells}
            if {"שם", "ת.ז"} <= labels and n + 1 < len(rows):
                self.id_rows = [row._tr, rows[n + 1]._tr]
                for cell in rows[n + 1].cells:
                    v = cell.text.strip()
                    if v and v not in self.id_values:
                        self.id_values.append(v)
                break

    def scrub(self, text):
        for v in self.id_values:
            text = text.replace(v, "[הוסר]")
            if re.search(r"[א-ת]", v):  # שם: מסירים גם כל מילה בנפרד
                for w in v.split():
                    if len(w) > 2:
                        text = re.sub(r"(?<![א-ת])" + re.escape(w) + r"(?![א-ת])", "[הוסר]", text)
        text = re.sub(r"[\w.+-]+@[\w-]+\.[\w.-]+", "[הוסר]", text)
        text = re.sub(r"(?<!\d)0\d{1,2}[- ]?\d{7}(?!\d)", "[הוסר]", text)
        text = re.sub(r"(?<!\d)\d{9}(?!\d)", "[הוסר]", text)
        return text

    # ---------- מידע על סעיפים ----------

    def section_images(self, key):
        if key not in self.ranges:
            return []
        s, e = self.ranges[key]
        return [(i, rid) for i in range(s + 1, e) for rid in para_images(self.paras[i])]

    def image_captions(self, key):
        """לכל תמונה בסעיף - הטקסט הקצר שמעליה (כיתוב שהרופא שם), אם יש."""
        out = []
        if key not in self.ranges:
            return out
        s, _e = self.ranges[key]
        for i, _rid in self.section_images(key):
            cap = ""
            for j in range(i - 1, s, -1):
                if para_images(self.paras[j]):
                    break
                t = self.texts[j]
                if t and not ai_key(self.paras[j]):
                    cap = t if len(t) <= 80 else ""
                    break
            out.append(self.scrub(cap))
        return out

    def has_ai_text(self, key):
        return any(_same_key(ai_key(p), key) for p in self.paras)

    def status(self):
        """רשימת הסעיפים למסך: נמצא/לא נמצא, מספר תמונות, האם כבר נותח."""
        out = []
        cells = self.statics_cells()
        out.append({"key": "statics", "title": STATICS["title"], "found": bool(cells),
                    "images": len(self.statics_images()),
                    "done": any(_cell_marked(c.paragraphs[0]._p) for _, c in cells)})
        stab = [k for k in ("stab_load", "stab_sway") if k in self.ranges]
        out.append({"key": "stab_table", "title": STAB_TABLE["title"], "found": bool(stab),
                    "images": sum(len(self.section_images(k)) for k in stab),
                    "done": any(_tbl_key(t) == "stab_table" for t in self.body.iter(qn("w:tbl")))})
        bal = [k for k in ("semitandem", "oneleg", "romberg") if k in self.ranges]
        out.append({"key": "balance_table", "title": BALANCE_TABLE["title"], "found": bool(bal),
                    "images": sum(len(self.section_images(k)) for k in bal),
                    "done": any(_tbl_key(t) == "balance_table" for t in self.body.iter(qn("w:tbl")))})
        for s in SPEC["sections"]:
            out.append({"key": s["key"], "title": s["title"], "found": s["key"] in self.ranges,
                        "images": len(self.section_images(s["key"])), "done": self.has_ai_text(s["key"])})
        d = SPEC["discussion"]
        out.append({"key": d["key"], "title": d["title"], "found": d["key"] in self.heads,
                    "images": 0, "done": self.has_ai_text(d["key"])})
        return out

    def context_text(self, include_ai, stop_at=None):
        """טקסט המסמך מותמם, בסדר המסמך, כולל טבלאות. include_ai=False משמיט ניתוחים שהכלי כתב."""
        stop_idx = stop_at if stop_at is not None else len(self.paras)
        index = {id(p): i for i, p in enumerate(self.paras)}
        lines = []

        def emit_para(p):
            i = index.get(id(p))
            if i is None or i >= stop_idx:
                return i is None
            t = self.texts[i]
            if not t or (not include_ai and ai_key(p)):
                return True
            if stop_at is None and _match(SPEC["context_stop"], t) and len(t) < 40:
                return False
            lines.append(t)
            return True

        for child in self.body.iterchildren():
            if child.tag == qn("w:p"):
                if not emit_para(child):
                    break
            elif child.tag == qn("w:tbl"):
                stop = False
                for tr in child.iter(qn("w:tr")):
                    if any(tr is x for x in self.id_rows):
                        continue
                    cells = []
                    for tc in tr.iterchildren(qn("w:tc")):
                        ps = [p for p in tc.iter(qn("w:p")) if id(p) in index and index[id(p)] < stop_idx]
                        if any(_match(SPEC["context_stop"], self.texts[index[id(p)]]) for p in ps) and stop_at is None:
                            stop = True
                        cells.append(" ".join(self.texts[index[id(p)]] for p in ps
                                              if include_ai or not ai_key(p)).strip())
                    if stop:
                        break
                    if any(cells):
                        lines.append(" | ".join(cells))
                if stop:
                    break
        return self.scrub("\n".join(lines))

    def image_payload(self, rid):
        """מחזיר (media_type, base64) של תמונה, מוקטנת אם צריך."""
        part = self.doc.part.related_parts[rid]
        data, mtype = part.blob, part.content_type
        try:
            from PIL import Image
            im = Image.open(io.BytesIO(data))
            if max(im.size) > 1568 or mtype not in ("image/png", "image/jpeg"):
                im.thumbnail((1568, 1568))
                buf = io.BytesIO()
                im.convert("RGB").save(buf, "PNG")
                data, mtype = buf.getvalue(), "image/png"
        except Exception:
            pass
        return mtype, base64.b64encode(data).decode("ascii")

    def image_crop(self, rid, box=(0.0, 0.0, 0.45, 0.42), width=1500):
        """חיתוך מוגדל של אזור בצילום (ברירת מחדל: הפינה העליונה-שמאלית, שם טבלת הפרמטרים של DIERS)."""
        try:
            from PIL import Image
            part = self.doc.part.related_parts[rid]
            im = Image.open(io.BytesIO(part.blob)).convert("RGB")
            w, h = im.size
            im = im.crop((int(box[0] * w), int(box[1] * h), int(box[2] * w), int(box[3] * h)))
            if im.width < width:
                im = im.resize((width, int(im.height * width / im.width)), Image.LANCZOS)
            buf = io.BytesIO()
            im.save(buf, "PNG")
            return "image/png", base64.b64encode(buf.getvalue()).decode("ascii")
        except Exception:
            return None

    # ---------- כתיבה למסמך ----------

    def font_rpr(self):
        """הגופן והגודל הנפוצים בטקסט הרגיל של המסמך - כדי שהטקסט של הכלי ייראה זהה."""
        if getattr(self, "_font", None) is not None:
            return self._font
        from collections import Counter
        fonts, sizes = Counter(), Counter()
        for r in self.body.iter(qn("w:r")):
            if any(a.tag == qn("w:tbl") for a in r.iterancestors()):
                continue
            txt = "".join(t.text or "" for t in r.iter(qn("w:t")))
            if len(txt.strip()) < 3:
                continue
            rpr = r.find(qn("w:rPr"))
            if rpr is None:
                continue
            f = rpr.find(qn("w:rFonts"))
            if f is not None:
                fonts[tuple(sorted((k, v) for k, v in f.attrib.items() if not k.endswith("Theme")))] += len(txt)
            z = rpr.find(qn("w:szCs")) if rpr.find(qn("w:szCs")) is not None else rpr.find(qn("w:sz"))
            if z is not None:
                sizes[z.get(qn("w:val"))] += len(txt)
        out = []
        if fonts:
            f = OxmlElement("w:rFonts")
            for k, v in fonts.most_common(1)[0][0]:
                f.set(k, v)
            out.append(f)
        if sizes:
            val = sizes.most_common(1)[0][0]
            for tag in ("w:sz", "w:szCs"):
                z = OxmlElement(tag)
                z.set(qn("w:val"), val)
                out.append(z)
        self._font = out
        return out

    def _next_bm_id(self):
        ids = [int(b.get(qn("w:id")) or 0) for b in self.body.iter(qn("w:bookmarkStart"))]
        return max(ids + [9000]) + 1

    def remove_ai(self, key):
        """מוחק את מה שהכלי כתב לסעיף (כולל ניתוחים נפרדים לכל תמונה: key#1, key#2...)."""
        for p in list(self.paras):
            if _same_key(ai_key(p), key) and p.getparent() is not None:
                p.getparent().remove(p)

    def write(self, key, text, anchor=None):
        """כותב ניתוח מתחת ל-anchor (ברירת מחדל: התמונה האחרונה בסעיף, או הכותרת). מחליף ניתוח קודם של אותו מפתח."""
        for p in list(self.paras):
            if ai_key(p) == key and p.getparent() is not None:
                p.getparent().remove(p)
        if anchor is None:
            imgs = self.section_images(key)
            anchor = self.paras[imgs[-1][0]] if imgs else self.paras[self.heads[key]]
        bm = self._next_bm_id()
        for n, (kind, line) in enumerate(reversed(parse_markup(text))):
            anchor.addnext(self._para(kind, line, bm + n, key))

    def _para(self, kind, text, bm_id, key):
        """בונה פסקה מעוצבת. kind: head (כותרת משנה) / bullet / note (לתשומת לב) / text."""
        p = OxmlElement("w:p")
        ppr = OxmlElement("w:pPr")
        if kind == "head":
            ppr.append(OxmlElement("w:keepNext"))
        ppr.append(OxmlElement("w:bidi"))
        sp = OxmlElement("w:spacing")
        sp.set(qn("w:before"), {"head": "180", "text": "100", "note": "60"}.get(kind, "0"))
        sp.set(qn("w:after"), {"head": "40", "bullet": "90"}.get(kind, "80"))
        ppr.append(sp)
        if kind in ("bullet", "note"):
            ind = OxmlElement("w:ind")
            ind.set(qn("w:right"), "360")
            ind.set(qn("w:hanging"), "220")
            ppr.append(ind)
        jc = OxmlElement("w:jc")
        jc.set(qn("w:val"), "both")
        ppr.append(jc)
        p.append(ppr)
        bs = OxmlElement("w:bookmarkStart")
        bs.set(qn("w:id"), str(bm_id))
        bs.set(qn("w:name"), f"{AI_PREFIX}{key}_{bm_id}")
        p.append(bs)

        color = CONFIG.get("note_color", "000000") if kind == "note" else CONFIG.get("ai_color", "000000")
        if kind == "head":
            segs = [(text.rstrip(":").strip() + ":", True)]
        else:
            segs = [(s, i % 2 == 1) for i, s in enumerate(re.split(r"\*\*", text)) if s]
        if kind == "note":
            label = re.match(r"^([^:*]{2,20}):\s*", text)   # למשל "הערת איכות:" - מודגש כמו שהוא
            if label:
                segs = [(label.group(1) + ": ", True)] + [
                    (s, i % 2 == 1) for i, s in enumerate(re.split(r"\*\*", text[label.end():])) if s]
            else:
                segs.insert(0, ("לתשומת לב: ", True))
        bullet = {"bullet": "•", "note": "⚠"}.get(kind)
        if bullet:
            segs.insert(0, (bullet, False))
        for s, bold in segs:
            r = OxmlElement("w:r")
            rpr = OxmlElement("w:rPr")
            if bold:
                rpr.append(OxmlElement("w:b"))
                rpr.append(OxmlElement("w:bCs"))
            for el in self.font_rpr():
                rpr.append(copy.deepcopy(el))
            c = OxmlElement("w:color")
            c.set(qn("w:val"), color)
            rpr.append(c)
            if kind == "head":
                u = OxmlElement("w:u")
                u.set(qn("w:val"), "single")
                rpr.append(u)
            rpr.append(OxmlElement("w:rtl"))
            r.append(rpr)
            t = OxmlElement("w:t")
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            t.text = fix_bidi(s)
            r.append(t)
            if s == bullet:
                r.append(OxmlElement("w:tab"))
            p.append(r)
        be = OxmlElement("w:bookmarkEnd")
        be.set(qn("w:id"), str(bm_id))
        p.append(be)
        return p

    def approve(self, keys=None):
        """אישור טקסט: הופך את הטקסט שהכלי כתב לשחור רגיל ומסיר את הסימון, כך שהכלי לא יחליף אותו עוד."""
        count = 0
        for p in self.paras:
            k = ai_key(p) or ("statics" if _cell_marked(p) else None)
            if not k or (keys and not any(_same_key(k, x) for x in keys)):
                continue
            for c in list(p.iter(qn("w:color"))):
                c.getparent().remove(c)
            for b in list(p.iter(qn("w:bookmarkStart"))) + list(p.iter(qn("w:bookmarkEnd"))):
                b.getparent().remove(b)
            count += 1
        if not keys or any(k in keys for k in TABLES):
            for t in self.body.iter(qn("w:tbl")):
                if keys and _tbl_key(t) not in keys:
                    continue
                marks = [b for b in t.iter(qn("w:bookmarkStart")) if (b.get(qn("w:name")) or "").startswith(TABLE_PREFIX)]
                for b in marks:
                    bid = b.get(qn("w:id"))
                    for e in [e for e in t.iter(qn("w:bookmarkEnd")) if e.get(qn("w:id")) == bid]:
                        e.getparent().remove(e)
                    b.getparent().remove(b)
                    count += 1
        return count

    # ---------- טבלת המדדים הסטטיים ----------

    def ensure_discussion_heading(self):
        """מוסיף כותרת 'דיון :' (בעיצוב של 'המלצות:') לפני 'המלצות:' אם אין כותרת דיון בדוח."""
        idx = next((i for i, t in enumerate(self.texts) if re.match(r"^המלצות\s*:", t)), None)
        if idx is None:
            return False
        rec = self.paras[idx]
        head = copy.deepcopy(rec)
        ts = list(head.iter(qn("w:t")))
        for n, t in enumerate(ts):
            t.text = "דיון :" if n == 0 else ""
        for b in list(head.iter(qn("w:bookmarkStart"))) + list(head.iter(qn("w:bookmarkEnd"))):
            b.getparent().remove(b)
        rec.addprevious(head)
        return True

    def statics_table(self):
        for t in self.doc.tables:
            if any("Kyphosis" in c.text for r in t.rows for c in r.cells):
                return t
        return None

    def statics_cells(self):
        """[(שם המדד, תא הערך)] - תא הערך הוא התא שמתחת לתא הכותרת."""
        t = self.statics_table()
        out = []
        if t is None:
            return out
        rows = t.rows
        for ri in range(len(rows) - 1):
            cells, nxt = rows[ri].cells, rows[ri + 1].cells
            for ci, c in enumerate(cells):
                lab = re.sub(r"\s+", " ", c.text.strip())
                low = lab.lower().replace("è", "e")
                lab = "instability" if "instability" in low else next(
                    (x for x in STATICS_LABELS if x.lower() == low), lab)
                if lab in STATICS_LABELS and ci < len(nxt) and all(lab != l for l, _ in out):
                    out.append((lab, nxt[ci]))
        return out

    def statics_images(self):
        """צילומי ה-DIERS הסטטיים: Coronal view קודם, ואחריו Sagittal ו-Angle (למדדים שאינם במבט הקורונלי)."""
        return [rid for k in ("coronal", "sagittal", "angle") for _, rid in self.section_images(k)]

    def ensure_extra_rows(self):
        """מוסיף לטבלת המדדים שורת כותרות + שורת ערכים למדדים שמופיעים ב-DIERS וחסרים בטבלה."""
        t = self.statics_table()
        if t is None or any("Vertebral Rotation" in c.text for r in t.rows for c in r.cells):
            return False
        rows = t.rows
        src = next((i for i, r in enumerate(rows) if any("Pelvic Rotation" in c.text for c in r.cells)), None)
        if src is None or src + 1 >= len(rows) or len(rows[src].cells) < 3:
            return False
        lab_tr, val_tr = copy.deepcopy(rows[src]._tr), copy.deepcopy(rows[src + 1]._tr)
        for tr, texts in ((lab_tr, STATICS_EXTRA), (val_tr, ("", "", ""))):
            tcs = list(tr.iterchildren(qn("w:tc")))
            for tc, text in zip(tcs, texts):
                ts = list(tc.iter(qn("w:t")))
                for n, el in enumerate(ts):
                    el.text = text if n == 0 else ""
                for b in list(tc.iter(qn("w:bookmarkStart"))) + list(tc.iter(qn("w:bookmarkEnd"))):
                    b.getparent().remove(b)
        rows[src + 1]._tr.addnext(lab_tr)
        lab_tr.addnext(val_tr)
        return True

    def set_cell(self, cell, text, bm_id):
        """כותב ערך לתא בטבלה, בכחול ומסומן, תוך שמירת עיצוב הגופן הקיים."""
        ps = cell.paragraphs
        p = ps[0]._p
        for extra in ps[1:]:
            extra._p.getparent().remove(extra._p)
        first = p.find(qn("w:r"))
        rpr = copy.deepcopy(first.find(qn("w:rPr"))) if first is not None and first.find(qn("w:rPr")) is not None \
            else OxmlElement("w:rPr")
        for el in list(p):
            if el.tag != qn("w:pPr"):
                p.remove(el)
        ppr = p.find(qn("w:pPr"))            # כיוון ימין-לשמאל בתא, אחרת ערכים כמו "L2 11 מ"מ" מתהפכים
        if ppr is None:
            ppr = OxmlElement("w:pPr")
            p.insert(0, ppr)
        if ppr.find(qn("w:bidi")) is None:
            bidi = OxmlElement("w:bidi")
            anchor_el = next((x for x in (ppr.find(qn("w:jc")), ppr.find(qn("w:rPr"))) if x is not None), None)
            if anchor_el is not None:
                anchor_el.addprevious(bidi)
            else:
                ppr.append(bidi)
        if rpr.find(qn("w:rtl")) is None:
            rpr.append(OxmlElement("w:rtl"))
        for c in rpr.findall(qn("w:color")):
            rpr.remove(c)
        color = OxmlElement("w:color")
        color.set(qn("w:val"), CONFIG.get("ai_color", "000000"))
        rpr.append(color)
        bs = OxmlElement("w:bookmarkStart")
        bs.set(qn("w:id"), str(bm_id))
        bs.set(qn("w:name"), f"{CELL_PREFIX}{bm_id}")
        be = OxmlElement("w:bookmarkEnd")
        be.set(qn("w:id"), str(bm_id))
        r = OxmlElement("w:r")
        r.append(rpr)
        for n, line in enumerate(text.split("\n")):
            if n:
                r.append(OxmlElement("w:br"))
            t = OxmlElement("w:t")
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            t.text = fix_bidi(line.strip())
            r.append(t)
        p.extend([bs, r, be])

    def write_after_table(self, key, table, text):
        """כותב הערות (למשל אי-התאמות) מיד אחרי טבלה."""
        self.remove_ai(key)
        bm = self._next_bm_id()
        for n, (kind, line) in enumerate(reversed(parse_markup(text))):
            table._tbl.addnext(self._para(kind, line, bm + n, key))

    def ai_markup(self, key):
        """משחזר את הטקסט שהכלי כתב לסעיף, בפורמט הסימון (לצורך ניסוח מחדש)."""
        lines = []
        for p in self.paras:
            if ai_key(p) != key:
                continue
            segs = []
            for r in p.iter(qn("w:r")):
                txt = "".join(t.text or "" for t in r.iter(qn("w:t")))
                rpr = r.find(qn("w:rPr"))
                bold = rpr is not None and rpr.find(qn("w:b")) is not None
                under = rpr is not None and rpr.find(qn("w:u")) is not None
                segs.append((txt, bold, under))
            full = "".join(s[0] for s in segs)
            if segs and segs[0][2]:
                lines.append("## " + full.rstrip(":"))
            elif full.startswith("⚠"):
                lines.append("! " + full[1:].strip())
            else:
                body = "".join(f"**{s}**" if b and s.strip() else s for s, b, _ in segs)
                lines.append("- " + body[1:].strip() if body.startswith("•") else body)
        return "\n".join(lines)

    def save(self, out_path):
        self.doc.save(out_path)


# ---------- בניית ההוראה ----------

def read_txt(folder, name):
    path = os.path.join(folder, name + ".txt")
    if os.path.exists(path):
        with open(path, encoding="utf-8-sig") as f:
            return f.read().strip()
    return ""


def examples_for(key):
    raw = read_txt(EXAMPLES_DIR, key)
    items = [x.strip() for x in re.split(r"\n-{5,}\n", raw) if x.strip()]
    return items[-int(CONFIG["max_examples"]):]


def build_prompt(report, key, note="", image_index=None):
    is_disc = key == "discussion"
    title = next(s["title"] for s in SPEC["sections"] + [SPEC["discussion"]] if s["key"] == key)
    stop = report.heads.get("discussion")  # ניתוחי הסעיפים לא רואים את הדיון; הדיון רואה הכול עד הכותרת שלו
    parts = [read_txt(INSTR_DIR, "כללי"),
             f"## הסעיף לכתיבה כעת: {title}",
             "## ההוראה לסעיף זה (מחייבת)\n" + read_txt(INSTR_DIR, key)]
    if note.strip():
        parts.append("## הערת הרופא לסעיף זה (מחייבת)\n" + note.strip())
    ex = examples_for(key)
    if ex:
        parts.append("## דוגמאות מדו\"חות קודמים של הרופא (מטופלים אחרים) - למינוח ולעומק הקליני בלבד. "
                     "אל תעתיק מהן ממצאים ואל תאמץ את המבנה שלהן: המבנה מחייב לפי 'פורמט הפלט'.\n"
                     + "\n-----\n".join(ex))
    parts.append("## נתוני המטופל הנוכחי מתוך הדו\"ח (מותמם)\n" + report.context_text(include_ai=is_disc, stop_at=stop))
    if not is_disc:
        caps = report.image_captions(key)
        n = len(caps)
        if image_index is None:
            if any(caps):
                parts.append("## כיתובי התמונות (לפי הסדר)\n" + "\n".join(
                    f"תמונה {i + 1}: {c or 'ללא כיתוב'}" for i, c in enumerate(caps)))
            parts.append(f"מצורפות {n} תמונות של הסעיף. כתוב כעת את הניתוח לסעיף \"{title}\" בלבד, לפי ההוראה.")
        else:
            cap = caps[image_index]
            parts.append(f"בסעיף \"{title}\" יש {n} תמונות, וכל אחת מקבלת ניתוח נפרד מתחתיה. מצורפת תמונה "
                         f"{image_index + 1} מתוך {n}" + (f" (כיתוב: {cap})" if cap else "") +
                         ". כתוב ניתוח של התמונה הזאת בלבד, לפי ההוראה ופורמט הפלט, עם מסקנה משלה. "
                         "אל תתאר תמונות אחרות ואל תחזור על מה שאינו בתמונה הזאת.")
    else:
        parts.append("כתוב כעת את הדיון לפי ההוראה.")
    return "\n\n".join(p for p in parts if p)


# ---------- Claude Code (מנוי) ----------

def claude_exe():
    exe = shutil.which("claude")
    if exe:
        return exe
    for cand in (os.path.join(os.environ.get("APPDATA", ""), "npm", "claude.cmd"),
                 os.path.join(os.path.expanduser("~"), ".local", "bin", "claude.exe")):
        if os.path.exists(cand):
            return cand
    raise RuntimeError("Claude Code לא מותקן במחשב. יש להריץ את 'התקנה.bat'.")


def _run(args, stdin=None, timeout=60):
    flags = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0
    return subprocess.run([claude_exe()] + args, input=stdin, capture_output=True, timeout=timeout,
                          encoding="utf-8", errors="replace", creationflags=flags,
                          cwd=os.path.expanduser("~"))


def auth_status():
    """מחזיר (מחובר?, תיאור החשבון)."""
    try:
        data = json.loads(_run(["auth", "status"]).stdout or "{}")
    except Exception as e:
        return False, str(e)
    if not data.get("loggedIn"):
        return False, "לא מחובר"
    desc = data.get("email") or data.get("orgName") or "מחובר"
    plan = data.get("subscriptionType")
    return True, f"{desc} ({plan})" if plan else desc


def open_login():
    """פותח חלון התחברות/החלפת חשבון."""
    exe = claude_exe()
    subprocess.Popen(f'start "Claude Login" cmd /k ""{exe}" auth logout & "{exe}" auth login"', shell=True)


def clean_output(text):
    """מיישר את הפלט למבנה המותר: ## כותרת, - תבליט, ! הערה, **מודגש**."""
    text = re.sub(r"^\s*[*•]\s+", "- ", text, flags=re.M)
    text = re.sub(r"^\s*\d+[.)]\s+", "- ", text, flags=re.M)
    return text.strip()


def parse_markup(text):
    """ממיר את הפלט לרשימת (סוג, טקסט): head / bullet / note / text."""
    out = []
    for raw in text.split("\n"):
        line = raw.strip()
        if not line or set(line) <= set("-_=*"):
            continue
        if not CONFIG.get("show_notes", False) and (
                line.startswith("!") or re.match(r"^[-•⚠]?\s*\**\s*(הערת איכות|לתשומת לב|סתירה)", line)):
            continue   # בלי הערות איכות / לתשומת לב במסמך
        if line.startswith("#"):
            out.append(("head", line.lstrip("#").strip().strip("*").strip()))
        elif line[:2] in ("- ", "• "):
            out.append(("bullet", line[2:].strip()))
        elif line.startswith("!"):
            out.append(("note", line.lstrip("!").strip()))
        else:
            out.append(("text", line))
    return out


def ask_claude(prompt, images=(), model=None, raw=False):
    content = [{"type": "image", "source": {"type": "base64", "media_type": m, "data": d}} for m, d in images]
    content.append({"type": "text", "text": prompt})
    msg = json.dumps({"type": "user", "message": {"role": "user", "content": content}}, ensure_ascii=False)
    args = ["-p", "--input-format", "stream-json", "--output-format", "stream-json", "--verbose",
            "--safe-mode", "--strict-mcp-config", "--tools", "", "--no-session-persistence",
            "--model", model or CONFIG["model"], "--effort", CONFIG.get("effort", "high"),
            "--system-prompt", "You are a clinical documentation assistant. Follow the instructions in the user message exactly. Reply in Hebrew with the requested text only."]
    res = _run(args, stdin=msg + "\n", timeout=int(CONFIG["timeout_sec"]))
    result = None
    for line in (res.stdout or "").splitlines():
        try:
            ev = json.loads(line)
        except ValueError:
            continue
        if ev.get("type") == "result":
            result = ev
    if not result:
        raise RuntimeError("לא התקבלה תשובה מ-Claude: " + ((res.stderr or res.stdout or "")[-400:]))
    if result.get("is_error"):
        raise RuntimeError(str(result.get("result") or result.get("subtype")))
    return result.get("result", "") if raw else clean_output(result.get("result", ""))


# ---------- הרצה ----------

WORD_EXT = (".docx", ".dotx", ".doc", ".docm", ".dotm", ".rtf")


def ensure_docx(path, log=print):
    """מחזיר נתיב ל-.docx שאפשר לנתח. תבנית (.dotx) או פורמט Word אחר מומרים לעותק .docx באותה תיקייה."""
    root, ext = os.path.splitext(path)
    ext = ext.lower()
    if ext == ".docx":
        return path
    if ext not in WORD_EXT:
        raise ValueError("זה לא קובץ Word")
    target = root + ".docx"
    if os.path.exists(target):
        target = root + " - לניתוח.docx"
    if ext == ".dotx":
        import zipfile
        with zipfile.ZipFile(path) as zin, zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as zout:
            for it in zin.infolist():
                data = zin.read(it.filename)
                if it.filename == "[Content_Types].xml":
                    data = data.replace(b"wordprocessingml.template.main+xml", b"wordprocessingml.document.main+xml")
                zout.writestr(it, data)
    else:   # .doc / .docm / .dotm / .rtf - המרה דרך Word עצמו
        ps = ("$w = New-Object -ComObject Word.Application; $w.Visible = $false; "
              "try { $d = $w.Documents.Open($env:SRC_DOC, $false, $true); $d.SaveAs2($env:DST_DOC, 16); $d.Close($false) } "
              "finally { $w.Quit() }")
        env = dict(os.environ, SRC_DOC=path, DST_DOC=target)
        flags = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0
        subprocess.run(["powershell", "-NoProfile", "-Command", ps], env=env, capture_output=True, timeout=180,
                       creationflags=flags)
        if not os.path.exists(target):
            raise RuntimeError("ההמרה ל-.docx נכשלה - פתח את הקובץ ב-Word ושמור אותו כ'מסמך Word (.docx)'")
    log(f"הקובץ '{os.path.basename(path)}' הוא {ext} - נוצר עותק לניתוח: {os.path.basename(target)}")
    return target


def default_output(path):
    root, _ = os.path.splitext(path)
    return root + " - מנותח.docx" if not root.endswith(" - מנותח") else root + ".docx"


def _json_block(text):
    m = re.search(r"\{.*\}", text, flags=re.S)
    if not m:
        raise RuntimeError("התשובה אינה בפורמט הצפוי")
    return json.loads(m.group(0))


def _norm_val(v):
    """מפתח השוואה לערך שנקרא: המספרים והצד/כיוון (R/L/A/P) בלבד."""
    if v in (None, ""):
        return None
    return tuple(re.findall(r"-?\d+(?:\.\d+)?|\b[RLAP]\b", str(v).replace("°", " ")))


def _read_statics(report, prompt, rids, model):
    """שולח את צילומי ה-DIERS + חיתוך מוגדל של טבלת הפרמטרים של כל צילום, ומחזיר את ה-JSON."""
    images = []
    for r in rids:
        images.append(report.image_payload(r))
        crop = report.image_crop(r)
        if crop:
            images.append(crop)
    return _json_block(ask_claude(prompt, images, model, raw=True))


def run_statics(report, note="", model=None, log=print, dry_run=False, fix=None):
    """ממלא את טבלת המדדים הסטטיים מצילומי ה-DIERS ובודק ערכים קיימים.
    הצילום נקרא פעמיים באופן בלתי תלוי; פועלים רק על ערך ששתי הקריאות הסכימו עליו.
    fix=True: ערך קיים שלא תואם לצילום מוחלף בערך מהצילום (ונרשם מתחת לטבלה מה שונה)."""
    fix = CONFIG.get("fix_values", True) if fix is None else fix
    rids = report.statics_images()
    if report.statics_table() is None or not rids:
        log("דילוג: מדדים סטטיים - לא נמצאה הטבלה או שאין תמונה מתחת ל-Coronal view")
        return
    if report.ensure_extra_rows():
        log("   נוספה לטבלת המדדים שורה: " + ", ".join(STATICS_EXTRA))
    cells = report.statics_cells()
    rows, fillable = [], {}
    for label, cell in cells:
        cur = re.sub(r"\s+", " ", cell.text.strip())
        ai = _cell_marked(cell.paragraphs[0]._p)
        fillable[label] = label not in STATICS_CHECK_ONLY and (ai or not re.search(r"\d", cur))
        rows.append(f'- {label} | טקסט נוכחי בתא: "{cur}" | ' +
                    ("למילוי" if fillable[label] else
                     "שיפוט קליני של הרופא - לבדיקה בלבד, לא לשנות" if label in STATICS_CHECK_ONLY else
                     "כבר מולא - לבדיקה מול הצילום"))
    prompt = "\n\n".join(x for x in [
        read_txt(INSTR_DIR, "statics"),
        ("## הערת הרופא\n" + note.strip()) if note.strip() else "",
        "## על התמונות\nלכל צילום מצורף אחריו גם חיתוך מוגדל של הפינה העליונה-שמאלית, שבה נמצאת בדרך כלל טבלת "
        "הפרמטרים (Parameter / Value / Unit). קרא את המספרים מהחיתוך המוגדל, ספרה אחרי ספרה, ובדוק שהם תואמים לשם "
        "הפרמטר באותה שורה. אל תבלבל בין Flèche Cervicale ל-Flèche Lombaire, ובין +max ל--max.",
        "## תאי הטבלה\n" + "\n".join(rows),
        "לתא שמסומן \"כבר מולא\": אם הערך לא תואם לצילום (agrees=false) - כתוב גם ב-new_text את הטקסט הנכון לתא, "
        "בניסוח של הרופא ולפי כללי ההתאמה. אם תואם - new_text הוא null.",
        'החזר JSON בלבד, בלי טקסט נוסף, במבנה: {"cells":[{"label":"...","image_value":"הערך כפי שנקרא מהתמונה, כולל צד, או null",'
        '"new_text":"הטקסט לתא או null","agrees":true/false/null,"note":"הסבר קצר בעברית אם יש אי-התאמה, אחרת null"}]}'] if x)
    if dry_run:
        log(f"[בדיקה יבשה] statics: {len(prompt)} תווים, {len(rids)} תמונות")
        return
    # שתי קריאות בלתי תלויות במקביל - הצלבה
    with ThreadPoolExecutor(max_workers=2) as pool:
        reads = list(pool.map(lambda _i: _read_statics(report, prompt, rids, model), range(2)))
    by_read = [{it.get("label"): it for it in r_.get("cells", [])} for r_ in reads]

    def consensus(label):
        """הערך ששתי קריאות לפחות הסכימו עליו; אם שתיים לא הסכימו - קריאה שלישית מכריעה."""
        items = [r_.get(label) for r_ in by_read if r_.get(label)]
        for i in range(len(items)):
            for j in range(i + 1, len(items)):
                if _norm_val(items[i].get("image_value")) == _norm_val(items[j].get("image_value")):
                    return items[i], items[j]
        return None

    pending = [l for l, _c in cells if consensus(l) is None and any(r_.get(l) for r_ in by_read)]
    if pending:   # מכריע: קריאה שלישית
        try:
            third = _read_statics(report, prompt, rids, model)
            by_read.append({it.get("label"): it for it in third.get("cells", [])})
        except Exception as e:
            log(f"   קריאה שלישית נכשלה: {e}")
    by_label = dict(cells)
    bm = report._next_bm_id() + 500
    filled, fixed, skipped = 0, 0, []
    for label, cell in cells:
        pair = consensus(label)
        if not pair:
            if any(r_.get(label) for r_ in by_read):
                skipped.append(label)
            continue
        x, y = pair
        new = (x.get("new_text") or y.get("new_text") or "").strip()
        if not new or label in STATICS_CHECK_ONLY:
            continue
        if fillable[label]:
            report.set_cell(by_label[label], new, bm)
            bm += 1
            filled += 1
        elif fix and x.get("agrees") is False and y.get("agrees") is False:
            old = re.sub(r"\s+", " ", cell.text.strip())
            report.set_cell(by_label[label], new, bm)
            bm += 1
            fixed += 1
            log(f"   {label}: \"{old}\" -> \"{new.replace(chr(10), ' / ')}\"")
    report.remove_ai("statics")       # בלי הערות במסמך
    log(f"✓ statics: מולאו {filled} תאים, תוקנו {fixed}" +
        (f"; לא ודאי (לא שונה): {', '.join(skipped)}" if skipped else ""))


STAB_ROWS = [("weight", "נשיאת משקל (%)", "50 / 50"),
             ("anterior", "עומס קדמי (%)", "40"),
             ("posterior", "עומס אחורי (%)", "60"),
             ("sway_ap", "Sway קדמי-אחורי - Sagittal (מ\"מ)", ""),
             ("sway_ml", "Sway צידי - Coronal (מ\"מ)", "")]
GREEN, RED = ("00703C", "E2F0D9"), ("C00000", "FBE4E4")


def _tbl_marked(tbl):
    return _tbl_key(tbl) is not None


def _cell_write(cell, text, bold=False, color=None, fill=None, size=18):
    tc = cell._tc
    tcpr = tc.get_or_add_tcPr()
    if fill:
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), fill)
        tcpr.append(shd)
    va = OxmlElement("w:vAlign")
    va.set(qn("w:val"), "center")
    tcpr.append(va)
    p = cell.paragraphs[0]._p
    ppr = p.get_or_add_pPr()
    ppr.append(OxmlElement("w:bidi"))
    sp = OxmlElement("w:spacing")
    sp.set(qn("w:before"), "20")
    sp.set(qn("w:after"), "20")
    ppr.append(sp)
    jc = OxmlElement("w:jc")
    jc.set(qn("w:val"), "center")
    ppr.append(jc)
    r = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr")
    if bold:
        rpr.append(OxmlElement("w:b"))
        rpr.append(OxmlElement("w:bCs"))
    if color:
        c = OxmlElement("w:color")
        c.set(qn("w:val"), color)
        rpr.append(c)
    for tag in ("w:sz", "w:szCs"):
        z = OxmlElement(tag)
        z.set(qn("w:val"), str(size))
        rpr.append(z)
    rpr.append(OxmlElement("w:rtl"))
    r.append(rpr)
    t = OxmlElement("w:t")
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = text
    r.append(t)
    p.append(r)


def _val(d):
    """ערך תא מה-JSON: מחרוזת או {"v":..,"flag":..}."""
    if d is None:
        return "-", False
    if isinstance(d, dict):
        v = d.get("v")
        return ("-" if v in (None, "") else str(v)), bool(d.get("flag"))
    return str(d), False


BLUE = ("1F3864", "DEEAF6")
GREY = (None, "F2F2F2")

BALANCE_TABLE = {"key": "balance_table", "title": "טבלת מבחני יציבות מאוחדת"}
BALANCE_ROWS = [("area", "שטח תנודה (מ\"מ²)"),
                ("path", "אורך מסלול COP (מ\"מ)"),
                ("velocity", "מהירות ממוצעת (מ\"מ/שנ')"),
                ("ml_max", "תנודה צידית מרבית (מ\"מ)"),
                ("ap_max", "תנודה קדמית-אחורית מרבית (מ\"מ)")]
BALANCE_COLS = [("Semi Tandem", BLUE, [("st_L", "שמאל"), ("st_R", "ימין")]),
                ("One Leg", BLUE, [("ol_L", "שמאל"), ("ol_R", "ימין")]),
                ("Romberg", None, [("rb_open", "עיניים פקוחות"), ("rb_closed", "עיניים עצומות")])]

TABLES = {
    "stab_table": {"sections": ("stab_load", "stab_sway"), "instr": "stab_table",
                   "caption": "**טבלה מאוחדת - Stabilimeter:** ירוק = עיניים פקוחות, אדום = עיניים עצומות; ערכים חריגים מודגשים."},
    "balance_table": {"sections": ("semitandem", "oneleg", "romberg"), "instr": "balance_table",
                      "caption": "**טבלה מאוחדת - מבחני יציבות:** השוואת ימין ושמאל; ב-Romberg ירוק = עיניים פקוחות, "
                                 "אדום = עיניים עצומות; הצד/המצב הירוד מודגש."},
}


def _tbl_key(tbl):
    for b in tbl.iter(qn("w:bookmarkStart")):
        name = b.get(qn("w:name")) or ""
        if name.startswith(TABLE_PREFIX):
            rest = name[len(TABLE_PREFIX):]
            return "stab_table" if rest.isdigit() else rest.rsplit("_", 1)[0]   # גרסה ישנה בלי שם
    return None


def _new_grid(doc, n_rows, widths):
    tbl = doc.add_table(rows=n_rows, cols=len(widths))
    tblpr = tbl._tbl.tblPr
    tblpr.append(OxmlElement("w:bidiVisual"))
    jc = OxmlElement("w:jc")
    jc.set(qn("w:val"), "center")
    tblpr.append(jc)
    borders = OxmlElement("w:tblBorders")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "4")
        el.set(qn("w:color"), "A6A6A6")
        borders.append(el)
    tblpr.append(borders)
    for row in tbl.rows:
        for w, cell in zip(widths, row.cells):
            cell.width = w * 635
    return tbl


def _place_table(report, key, tbl, anchor, data):
    """מסמן את הטבלה, מכניס אותה אחרי anchor עם כיתוב מעליה והערה (אם יש) מתחתיה."""
    t = tbl._tbl
    bm = report._next_bm_id()
    first_p = t.find(qn("w:tr")).find(qn("w:tc")).find(qn("w:p"))
    bs = OxmlElement("w:bookmarkStart")
    bs.set(qn("w:id"), str(bm))
    bs.set(qn("w:name"), f"{TABLE_PREFIX}{key}_{bm}")
    be = OxmlElement("w:bookmarkEnd")
    be.set(qn("w:id"), str(bm))
    first_p.insert(1, bs)
    first_p.append(be)
    anchor.addnext(t)
    anchor.addnext(report._para("text", TABLES[key]["caption"], bm + 1, key))
    # בלי הערות מתחת לטבלה - רק הנתונים


def build_stab_table(report, data, anchor):
    """Stabilimeter: מדד | נורמה | פקוחות (שמאל/כולל/ימין) | עצומות (שמאל/כולל/ימין)."""
    tbl = _new_grid(report.doc, 2 + len(STAB_ROWS), [2300, 900, 850, 850, 850, 850, 850, 850])
    hdr, sub = tbl.rows[0].cells, tbl.rows[1].cells
    _cell_write(hdr[0], "מדד", bold=True, fill=GREY[1])
    _cell_write(hdr[1], "נורמה", bold=True, fill=GREY[1])
    _cell_write(hdr[2], "עיניים פקוחות", bold=True, color=GREEN[0], fill=GREEN[1])
    _cell_write(hdr[5], "עיניים עצומות", bold=True, color=RED[0], fill=RED[1])
    for i, name in enumerate(("שמאל", "כולל / ליבה", "ימין")):
        _cell_write(sub[2 + i], name, bold=True, color=GREEN[0], fill=GREEN[1], size=16)
        _cell_write(sub[5 + i], name, bold=True, color=RED[0], fill=RED[1], size=16)
    _cell_write(sub[0], "", fill=GREY[1])
    _cell_write(sub[1], "", fill=GREY[1])
    rows, norms = data.get("rows", {}), data.get("norms", {})
    for n, (key, label, norm) in enumerate(STAB_ROWS):
        cells = tbl.rows[2 + n].cells
        _cell_write(cells[0], label, bold=True)
        _cell_write(cells[1], str(norms.get(key) or norm or "-"), size=16)
        for g, (cond, (color, _fill)) in enumerate((("open", GREEN), ("closed", RED))):
            vals = (rows.get(key) or {}).get(cond) or {}
            for i, side in enumerate(("L", "T", "R")):
                v, flag = _val(vals.get(side))
                _cell_write(cells[2 + 3 * g + i], v, bold=flag, color=color)
    hdr[2].merge(hdr[4])
    hdr[5].merge(hdr[7])
    hdr[0].merge(sub[0])
    hdr[1].merge(sub[1])
    _place_table(report, "stab_table", tbl, anchor, data)


def build_balance_table(report, data, anchor):
    """מבחני יציבות: מדד | Semi Tandem (שמאל/ימין) | One Leg (שמאל/ימין) | Romberg (פקוחות/עצומות)."""
    tbl = _new_grid(report.doc, 2 + len(BALANCE_ROWS), [2500, 950, 950, 950, 950, 1050, 1050])
    hdr, sub = tbl.rows[0].cells, tbl.rows[1].cells
    _cell_write(hdr[0], "מדד", bold=True, fill=GREY[1])
    _cell_write(sub[0], "", fill=GREY[1])
    col = 1
    for title, colors, subs in BALANCE_COLS:
        _cell_write(hdr[col], title, bold=True, color=(colors or GREY)[0], fill=(colors or GREY)[1])
        for i, (_k, name) in enumerate(subs):
            c = colors or (GREEN if i == 0 else RED)
            _cell_write(sub[col + i], name, bold=True, color=c[0], fill=c[1], size=16)
        col += len(subs)
    rows = data.get("rows", {})
    for n, (key, label) in enumerate(BALANCE_ROWS):
        cells = tbl.rows[2 + n].cells
        _cell_write(cells[0], label, bold=True)
        col = 1
        for _title, colors, subs in BALANCE_COLS:
            for i, (k, _name) in enumerate(subs):
                v, flag = _val((rows.get(key) or {}).get(k))
                color = colors[0] if colors else (GREEN if i == 0 else RED)[0]
                _cell_write(cells[col + i], v, bold=flag, color=color)
            col += len(subs)
    col = 1
    for _title, _c, subs in BALANCE_COLS:
        hdr[col].merge(hdr[col + len(subs) - 1])
        col += len(subs)
    hdr[0].merge(sub[0])
    _place_table(report, "balance_table", tbl, anchor, data)


BUILDERS = {"stab_table": build_stab_table, "balance_table": build_balance_table}


def run_table(report, key, note="", model=None, log=print, dry_run=False):
    """טבלה מאוחדת מתוך תמונות של כמה סעיפים. מוכנסת בסוף הסעיף האחרון בקבוצה."""
    spec = TABLES[key]
    keys = [k for k in spec["sections"] if k in report.ranges]
    if not any(report.section_images(k) for k in keys):
        log(f"דילוג: {key} - אין תמונות בסעיפים הרלוונטיים")
        return
    images, labels = [], []
    for k in keys:
        title = next(x["title"] for x in SPEC["sections"] if x["key"] == k)
        for (_, rid), cap in zip(report.section_images(k), report.image_captions(k)):
            images.append(report.image_payload(rid))
            labels.append(f"תמונה {len(images)}: מתוך הסעיף \"{title}\"" + (f" - כיתוב: {cap}" if cap else ""))
    prompt = "\n\n".join(x for x in [
        read_txt(INSTR_DIR, spec["instr"]),
        ("## הערת הרופא\n" + note.strip()) if note.strip() else "",
        "## סדר התמונות המצורפות\n" + "\n".join(labels),
        "## טקסט וטבלאות מהדו\"ח (לעיון; מקור הערכים הוא התמונות)\n" +
        report.context_text(False, report.heads.get("discussion"))[:6000],
    ] if x)
    if dry_run:
        log(f"[בדיקה יבשה] {key}: {len(prompt)} תווים, {len(images)} תמונות")
        return
    data = _json_block(ask_claude(prompt, images, model, raw=True))
    for t in [t for t in report.body.iter(qn("w:tbl")) if _tbl_key(t) == key]:
        t.getparent().remove(t)
    report.remove_ai(key)
    last = [k for k in keys if report.section_images(k)][-1]
    s, e = report.ranges[last]
    imgs = report.section_images(last)
    anchor = report.paras[imgs[-1][0]]
    for j in range(e - 1, imgs[-1][0], -1):     # הפסקה האחרונה בסעיף שעדיין במסמך ואינה בתוך טבלה
        p = report.paras[j]
        if p.getparent() is not None and not ai_key(p) and not any(a.tag == qn("w:tbl") for a in p.iterancestors()):
            anchor = p
            break
    BUILDERS[key](report, data, anchor)
    log(f"✓ {key}")


def precheck(path, keys):
    """בדיקה לפני הרצה: מחזיר רשימת אזהרות לרופא."""
    rep_ = Report(path)
    warn = []
    titles = {s["key"]: s["title"] for s in SPEC["sections"]}
    if not rep_.id_values:
        warn.append("שורת הפרטים האישיים (שם / ת.ז) ריקה או לא נמצאה בטבלה הראשונה - אם יש בדוח פרטים מזהים, ייתכן שלא יוסרו לפני השליחה!")
    no_img = [titles[k] for k in keys if k in titles and k in rep_.ranges and not rep_.section_images(k)]
    if no_img:
        warn.append("סעיפים מסומנים בלי תמונה (ידולגו): " + ", ".join(no_img))
    missing = [titles[k] for k in keys if k in titles and k not in rep_.ranges]
    if missing:
        warn.append("כותרות שלא נמצאו בדוח (ידולגו): " + ", ".join(missing))
    multi = [titles[k] for k in keys if k in titles and len(rep_.section_images(k)) > 4]
    if multi:
        warn.append("סעיפים עם יותר מ-4 תמונות - כדאי לוודא שהתמונות במקום הנכון: " + ", ".join(multi))
    ctx = rep_.context_text(False, rep_.heads.get("discussion"))
    if len(ctx) < 400:
        warn.append("כמעט ואין טקסט קליני בדוח (אנמנזה / בדיקה פיזיקלית) - הניתוח יתבסס על התמונות בלבד.")
    if "discussion" in keys:
        pending = [titles[k] for k in titles if k in rep_.ranges and rep_.section_images(k)
                   and not rep_.has_ai_text(k) and k not in keys]
        if pending:
            warn.append("הדיון ייכתב בלי הניתוחים של: " + ", ".join(pending))
    return warn


REPHRASE = {"short": "קצר את הטקסט בכשליש: השאר רק את הממצאים והמשמעויות החשובים ביותר.",
            "long": "הרחב מעט את המשמעות הקלינית והתפקודית, בלי לחרוג ממגבלת האורך של הסעיף.",
            "simple": "נסח בצורה פשוטה וברורה יותר, משפטים קצרים, פחות מונחים באנגלית."}


def rephrase(path, keys, mode="short", note="", model=None, log=print):
    """מנסח מחדש ניתוחים קיימים בלי לקרוא שוב את התמונות ובלי לשנות ממצאים. כותב לאותו קובץ."""
    started = time.time()
    report = Report(path)
    present = []
    for p in report.paras:
        k = ai_key(p)
        if k and k not in present:
            present.append(k)
    todo = [k for k in present if any(_same_key(k, base) for base in keys)
            and k.split("#")[0] not in ("statics",) and k.split("#")[0] not in TABLES]
    errors = {}

    def job(k):
        prompt = "\n\n".join(x for x in [
            read_txt(INSTR_DIR, "כללי"), "## ההוראה המקורית לסעיף\n" + read_txt(INSTR_DIR, k.split("#")[0]),
            "## הניתוח הקיים\n" + report.ai_markup(k),
            "## המשימה\nנסח מחדש את הניתוח הקיים. " + REPHRASE.get(mode, "") +
            (" דגש נוסף מהרופא: " + note.strip() if note.strip() else "") +
            "\nאסור לשנות ממצאים, ערכים, צדדים או מסקנות, ואסור להוסיף ממצאים חדשים. שמור בדיוק על אותו מבנה וסימונים "
            "(## כותרת, - תבליט, ! הערה, **מודגש**, כותרות התבליטים המודגשות). החזר את הטקסט המנוסח בלבד."] if x)
        return ask_claude(prompt, (), model)

    results = {}
    with ThreadPoolExecutor(max_workers=int(CONFIG["max_parallel"])) as pool:
        futs = {pool.submit(job, k): k for k in todo}
        for f in as_completed(futs):
            k = futs[f]
            try:
                results[k] = f.result()
                log(f"✓ נוסח מחדש: {k}")
            except Exception as e:
                errors[k] = str(e)
                log(f"✗ {k}: {e}")

    def anchor_of(k):
        ps = [p for p in report.paras if ai_key(p) == k]
        return ps[0].getprevious() if ps else None

    anchors = {k: anchor_of(k) for k in results}
    for k, text in results.items():
        report.write(k, text, anchors[k])
    report.save(path)
    log_usage("ניסוח מחדש", len(todo), model or CONFIG["model"], time.time() - started, len(errors))
    return path, errors


def log_usage(action, n_sections, model, seconds, n_errors):
    """יומן שימוש ללא פרטים מזהים (בלי שם קובץ)."""
    try:
        path = os.path.join(BASE, "יומן שימוש.csv")
        new = not os.path.exists(path)
        with open(path, "a", encoding="utf-8-sig") as f:
            if new:
                f.write("תאריך,שעה,פעולה,סעיפים,מודל,שניות,שגיאות\n")
            f.write(time.strftime("%Y-%m-%d,%H:%M") + f",{action},{n_sections},{model},{int(seconds)},{n_errors}\n")
    except Exception:
        pass


def run(path, keys, notes=None, model=None, out_path=None, log=print, dry_run=False):
    """מנתח את הסעיפים שנבחרו (keys) וכותב עותק חדש. 'discussion' רץ אחרון, אחרי שכל הניתוחים נכתבו."""
    notes = notes or {}
    started = time.time()
    out_path = out_path or default_output(path)
    report = Report(path)
    errors = {}
    if "statics" in keys:
        try:
            run_statics(report, notes.get("statics", ""), model, log, dry_run)
            report.save(out_path)
            report = Report(out_path)   # הסעיפים הבאים רואים את הטבלה המעודכנת
        except Exception as e:
            errors["statics"] = str(e)
            log(f"✗ statics: {e}")
    for tk in ("stab_table", "balance_table"):
        if tk not in keys:
            continue
        try:
            run_table(report, tk, notes.get(tk, ""), model, log, dry_run)
            report.save(out_path)
            report = Report(out_path)
        except Exception as e:
            errors[tk] = str(e)
            log(f"✗ {tk}: {e}")
    todo = []
    for k in keys:
        if k in ("discussion", "statics", "stab_table", "balance_table"):
            continue
        if k not in report.ranges:
            log(f"דילוג: הכותרת של '{k}' לא נמצאה במסמך")
        elif not report.section_images(k):
            log(f"דילוג: '{k}' - אין תמונה מתחת לכותרת, ולכן לא נותח")
        else:
            todo.append(k)

    # משימות: סעיף עם תמונה אחת = ניתוח אחד; סעיף עם כמה תמונות = ניתוח נפרד מתחת לכל תמונה
    jobs = []
    for k in todo:
        imgs = report.section_images(k)
        if len(imgs) == 1:
            jobs.append((k, k, None))
        else:
            report.remove_ai(k)
            jobs += [(k, f"{k}#{n + 1}", n) for n in range(len(imgs))]

    def job(item):
        k, sub, n = item
        imgs = report.section_images(k)
        rids = [rid for _, rid in imgs] if n is None else [imgs[n][1]]
        prompt = build_prompt(report, k, notes.get(k, ""), image_index=n)
        if dry_run:
            return f"[בדיקה יבשה] {len(prompt)} תווים, {len(rids)} תמונות"
        return ask_claude(prompt, [report.image_payload(r) for r in rids], model)

    results = {}
    with ThreadPoolExecutor(max_workers=int(CONFIG["max_parallel"])) as pool:
        futs = {pool.submit(job, it): it for it in jobs}
        for f in as_completed(futs):
            k, sub, n = futs[f]
            try:
                results[(k, sub, n)] = f.result()
                log(f"✓ {sub}")
            except Exception as e:
                errors[sub] = str(e)
                log(f"✗ {sub}: {e}")
    # כתיבה מהסוף להתחלה כדי לא לשבש מיקומים
    def pos(item):
        k, _sub, n = item
        imgs = report.section_images(k)
        return imgs[-1 if n is None else n][0]
    for item in sorted(results, key=pos, reverse=True):
        k, sub, n = item
        anchor = None if n is None else report.paras[report.section_images(k)[n][0]]
        report.write(sub, results[item], anchor)

    report.save(out_path)

    if "discussion" in keys:
        report = Report(out_path)
        if "discussion" not in report.heads and report.ensure_discussion_heading():
            report.save(out_path)
            report = Report(out_path)
            log("   נוספה כותרת 'דיון :' לפני 'המלצות:'")
        if "discussion" not in report.heads:
            log("דילוג: לא נמצאה כותרת 'דיון' ולא נמצא 'המלצות:' כדי להוסיף אותה")
        else:
            try:
                prompt = build_prompt(report, "discussion", notes.get("discussion", ""))
                text = f"[בדיקה יבשה] {len(prompt)} תווים" if dry_run else ask_claude(prompt, (), model)
                report.write("discussion", text)
                report.save(out_path)
                log("✓ discussion")
            except Exception as e:
                errors["discussion"] = str(e)
                log(f"✗ discussion: {e}")
    if not dry_run:
        log_usage("ניתוח", len(keys), model or CONFIG["model"], time.time() - started, len(errors))
    return out_path, errors


if __name__ == "__main__":
    import sys
    r = Report(sys.argv[1])
    for s in r.status():
        print(s)
    print(r.id_values)
