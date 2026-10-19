# -*- coding: utf-8 -*-
"""כלי עזר: למידת דוגמאות סגנון מדו"חות מוכנים, ויצירת עותק "ריק מניתוחים" לבדיקות.

שימוש:
  python tools.py learn "C:\\תיקייה עם דו\"חות מוכנים"
  python tools.py strip "דו\"ח מוכן.docx"      -> יוצר עותק בלי הניתוחים, להשוואה
"""
import glob
import os
import re
import sys

import engine

MIN_LEN = 40


def section_texts(report, key):
    """פסקאות הטקסט החופשי שבסעיף (בלי כותרת, טבלאות קצרות ותמונות)."""
    if key not in report.ranges:
        return []
    s, e = report.ranges[key]
    return [i for i in range(s + 1, e)
            if len(report.texts[i]) >= MIN_LEN and not engine.para_images(report.paras[i])
            and not any(a.tag == engine.qn("w:tbl") for a in report.paras[i].iterancestors())]


def learn(folder):
    os.makedirs(engine.EXAMPLES_DIR, exist_ok=True)
    keys = [s["key"] for s in engine.SPEC["sections"]] + ["discussion"]
    added = 0
    files = [f for f in glob.glob(os.path.join(folder, "**", "*.docx"), recursive=True)
             if not os.path.basename(f).startswith("~$")]
    for path in files:
        try:
            rep = engine.Report(path)
        except Exception as e:
            print("דילוג:", os.path.basename(path), e)
            continue
        for key in keys:
            text = rep.scrub("\n".join(rep.texts[i] for i in section_texts(rep, key))).strip()
            if len(text) < 80:
                continue
            target = os.path.join(engine.EXAMPLES_DIR, key + ".txt")
            old = open(target, encoding="utf-8-sig").read() if os.path.exists(target) else ""
            if text in old:
                continue
            with open(target, "w", encoding="utf-8-sig") as f:
                f.write((old.strip() + "\n-----\n" if old.strip() else "") + text + "\n")
            added += 1
    print(f"נסרקו {len(files)} דו\"חות, נוספו {added} דוגמאות לתיקיית 'דוגמאות'.")


def strip(path):
    rep = engine.Report(path)
    keys = [s["key"] for s in engine.SPEC["sections"]] + ["discussion"]
    for key in keys:
        for i in section_texts(rep, key):
            p = rep.paras[i]
            if p.getparent() is not None:
                p.getparent().remove(p)
    out = os.path.splitext(path)[0] + " - ללא ניתוחים.docx"
    rep.save(out)
    print("נשמר:", out)
    return out


if __name__ == "__main__":
    if len(sys.argv) < 3 or sys.argv[1] not in ("learn", "strip"):
        print(__doc__)
    elif sys.argv[1] == "learn":
        learn(sys.argv[2])
    else:
        strip(sys.argv[2])
