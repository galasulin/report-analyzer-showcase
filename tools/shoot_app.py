# -*- coding: utf-8 -*-
"""צילום מסך של התוכנה עם דו"ח הדגמה (חשבון ושמות קבצים בדיוניים)."""
import os
import sys
import time

SRC, DEMO, OUT = sys.argv[1], sys.argv[2], sys.argv[3]
sys.path.insert(0, SRC)
os.chdir(SRC)
import app  # noqa: E402
from PIL import ImageGrab  # noqa: E402

app.engine.auth_status = lambda: (True, "demo@clinic.example (max)")
a = app.App()
a.add_files([DEMO])
a.mark("all")
a.write("✓ statics: מולאו 9 תאים, תוקנו 0")
a.write("✓ stab_table")
a.write("✓ coronal")
a.write("✓ discussion")
a.write("   ✓ נשמר: דוח הדגמה - מנותח.docx")
for _ in range(40):
    a.update()
    time.sleep(0.05)
a.lift()
a.attributes("-topmost", True)
a.update()
time.sleep(0.6)
x, y, w, h = a.winfo_rootx(), a.winfo_rooty(), a.winfo_width(), a.winfo_height()
ImageGrab.grab(bbox=(x, y, x + w, y + h)).save(OUT)
a.destroy()
print("ok", w, h)
