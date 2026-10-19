# -*- coding: utf-8 -*-
"""ניתוח דו"ח ביומכני - חלון ההפעלה.
© כל הזכויות שמורות לגל אסולין, IT Manager."""
import os
import queue
import sys
import threading
import tkinter as tk
import webbrowser
from tkinter import filedialog, messagebox, ttk

import engine

try:   # תצוגה חדה במסכים עם הגדלה (125%/150%)
    import ctypes
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    pass

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    Base = TkinterDnD.Tk
except Exception:          # בלי גרירה - התוכנה עובדת כרגיל עם כפתור הבחירה
    DND_FILES, Base = None, tk.Tk

APP_NAME = "ניתוח דו\"ח ביומכני"
VERSION = "1.2"
AUTHOR = "גל אסולין"
AUTHOR_ROLE = "IT Manager"
AUTHOR_MAIL = ""
COPYRIGHT = f"© 2026 כל הזכויות שמורות ל{AUTHOR} · {AUTHOR_ROLE}"

MODELS = [("Claude Opus 5.5  ·  מומלץ", "claude-opus-5-5"),
          ("Claude Fable 5.1  ·  החזק ביותר, איטי יותר", "claude-fable-5-1"),
          ("Claude Sonnet 5.5  ·  מהיר וחסכוני", "claude-sonnet-5-5"),
          ("Claude Opus 5", "claude-opus-5")]
EFFORTS = [("עומק חשיבה: גבוה (מומלץ)", "high"), ("עומק חשיבה: מקסימלי", "xhigh"), ("עומק חשיבה: בינוני (מהיר)", "medium")]
REPHRASE = [("קצר יותר", "short"), ("מפורט יותר", "long"), ("פשוט וברור יותר", "simple")]
ITEMS = [engine.STATICS, engine.STAB_TABLE, engine.BALANCE_TABLE] + engine.SPEC["sections"] + [engine.SPEC["discussion"]]

# פלטת צבעים
NAVY, TEAL, BG, CARD, LINE, TEXT, MUTED = "#0F2A44", "#2EC4B6", "#F3F6FA", "#FFFFFF", "#DCE3EC", "#1E2B3A", "#6B7A8C"
OK, WARN, INFO, ERR = "#16A34A", "#D97706", "#2563EB", "#DC2626"
F = "Segoe UI"
DROP_HINT = "⇩  גרור לכל מקום בחלון קובץ Word (‎.docx / ‎.dotx / ‎.doc) - או לחץ '+ הוסף קבצים'"
LOG_FILE = os.path.join(os.path.dirname(sys.executable if getattr(sys, "frozen", False) else os.path.abspath(__file__)),
                        "שגיאות.log")


def log_error(text):
    try:
        import time
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(time.strftime("%Y-%m-%d %H:%M:%S ") + text + "\n")
    except Exception:
        pass


def resource(name):
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, name)


class App(Base):
    def __init__(self):
        super().__init__()
        self.title(f"{APP_NAME}  ·  הרופא")
        k = self.winfo_fpixels("1i") / 96.0
        w, h = int(1020 * k), int(900 * k)
        h = min(h, int(self.winfo_screenheight() * 0.92))
        self.geometry(f"{w}x{h}+{max(0, (self.winfo_screenwidth() - w) // 2)}+10")
        self.minsize(int(900 * k), int(700 * k))
        self.configure(bg=BG)
        try:
            self.iconbitmap(resource("icon.ico"))
        except Exception:
            pass
        self.files, self.status, self.rows = [], [], {}
        self.q = queue.Queue()
        self.busy = False
        self._style()
        self._build()
        self._enable_drop(self)
        if not DND_FILES:
            self.drop_hint.configure(text="גרירה לא זמינה במחשב הזה - השתמש ב'+ הוסף קבצים'")
        self.refresh_account()
        self.after(150, self._poll)

    # ---------- עיצוב ----------
    def _style(self):
        st = ttk.Style(self)
        st.theme_use("clam")
        st.configure(".", font=(F, 10), background=BG, foreground=TEXT)
        st.configure("Card.TFrame", background=CARD)
        st.configure("Card.TLabel", background=CARD, foreground=TEXT)
        st.configure("Muted.TLabel", background=CARD, foreground=MUTED, font=(F, 9))
        st.configure("CardTitle.TLabel", background=CARD, foreground=NAVY, font=(F, 11, "bold"))
        st.configure("TButton", font=(F, 10), padding=(12, 6), background="#E8EEF5", foreground=TEXT,
                     bordercolor=LINE, focusthickness=0)
        st.map("TButton", background=[("active", "#DCE6F1"), ("disabled", "#EEF1F5")])
        st.configure("Accent.TButton", font=(F, 12, "bold"), padding=(26, 10), background=TEAL, foreground="white",
                     bordercolor=TEAL)
        st.map("Accent.TButton", background=[("active", "#25A99D"), ("disabled", "#9ADDD6")])
        st.configure("Dark.TButton", padding=(10, 4), background="#1C3D5E", foreground="white", bordercolor="#1C3D5E")
        st.map("Dark.TButton", background=[("active", "#26507A")])
        st.configure("TCombobox", padding=4)
        st.configure("Teal.Horizontal.TProgressbar", troughcolor="#E3E9F0", background=TEAL, bordercolor=LINE,
                     lightcolor=TEAL, darkcolor=TEAL)
        st.configure("TCheckbutton", background=CARD)

    def card(self, parent, title, subtitle=""):
        outer = tk.Frame(parent, bg=LINE)
        inner = ttk.Frame(outer, style="Card.TFrame", padding=12)
        inner.pack(fill="both", expand=True, padx=1, pady=1)
        head = ttk.Frame(inner, style="Card.TFrame")
        head.pack(fill="x", pady=(0, 6))
        ttk.Label(head, text=title, style="CardTitle.TLabel").pack(side="right")
        if subtitle:
            ttk.Label(head, text=subtitle, style="Muted.TLabel").pack(side="right", padx=10)
        return outer, inner, head

    # ---------- בניית המסך ----------
    def _build(self):
        header = tk.Frame(self, bg=NAVY, height=int(78 * self.winfo_fpixels("1i") / 96.0))
        header.pack(fill="x")
        header.pack_propagate(False)
        tk.Frame(header, bg=TEAL, width=6).pack(side="right", fill="y")
        titles = tk.Frame(header, bg=NAVY)
        titles.pack(side="right", padx=18)
        tk.Label(titles, text=APP_NAME, bg=NAVY, fg="white", font=(F, 18, "bold")).pack(anchor="e", pady=(12, 0))
        tk.Label(titles, text="הרופא  ·  מרפאה לדוגמה  ·  ניתוח DIERS · Stabilimeter · הליכה",
                 bg=NAVY, fg="#A9BCD0", font=(F, 9)).pack(anchor="e")
        acc = tk.Frame(header, bg=NAVY)
        acc.pack(side="left", padx=16)
        self.account = tk.Label(acc, text="● בודק חשבון...", bg=NAVY, fg="#A9BCD0", font=(F, 10, "bold"))
        self.account.pack(anchor="w", pady=(14, 4))
        btns = tk.Frame(acc, bg=NAVY)
        btns.pack(anchor="w")
        ttk.Button(btns, text="החלף חשבון", style="Dark.TButton", command=self.switch_account).pack(side="left")
        ttk.Button(btns, text="רענן", style="Dark.TButton", command=self.refresh_account).pack(side="left", padx=6)

        # כותרת תחתונה
        foot = tk.Frame(self, bg=NAVY)
        foot.pack(fill="x", side="bottom")
        tk.Label(foot, text=COPYRIGHT, bg=NAVY, fg="#C9D6E3", font=(F, 9)).pack(side="right", padx=14)
        about = tk.Label(foot, text=f"אודות  ·  גרסה {VERSION}", bg=NAVY, fg=TEAL, font=(F, 9, "underline"), cursor="hand2")
        about.pack(side="left", padx=14)
        about.bind("<Button-1>", lambda e: self.about())

        body = tk.Frame(self, bg=BG)
        body.pack(fill="both", expand=True, padx=14, pady=10)


        # דו"חות
        outer, files, head = self.card(body, "דו\"חות", "גרור לכאן קבצי Word - גם כמה יחד")
        outer.pack(fill="x")
        ttk.Button(head, text="+ הוסף קבצים", command=self.choose).pack(side="left")
        ttk.Button(head, text="הסר", command=self.remove_selected).pack(side="left", padx=4)
        ttk.Button(head, text="נקה", command=self.clear_files).pack(side="left")
        self.flist = tk.Listbox(files, height=3, justify="right", activestyle="none", exportselection=False,
                                relief="flat", bg="#F7FAFD", highlightthickness=1, highlightcolor=TEAL,
                                highlightbackground=LINE, selectbackground="#D5F3EF", selectforeground=TEXT,
                                font=(F, 10))
        self.flist.pack(fill="x")
        self.flist.bind("<<ListboxSelect>>", lambda e: self.show_selected())
        self.drop_hint = ttk.Label(files, text=DROP_HINT, style="Muted.TLabel")
        self.drop_hint.pack(anchor="e", pady=(4, 0))

        # סעיפים
        outer, box, head = self.card(body, "סעיפים לניתוח", "סמן מה לנתח · אפשר להוסיף דגש לכל סעיף")
        outer.pack(fill="both", expand=True, pady=10)
        for text, mode in (("נקה", "none"), ("רק מה שלא נותח", "new"), ("כל מה שיש לו תמונה", "images"),
                           ("✦ סמן הכל + דיון", "all")):
            ttk.Button(head, text=text, command=lambda m=mode: self.mark(m)).pack(side="left", padx=2)
        grid = tk.Frame(box, bg=CARD)
        grid.pack(fill="both", expand=True)
        cols = tk.Frame(grid, bg="#EEF3F8")
        cols.pack(fill="x")
        tk.Label(cols, text="סעיף", bg="#EEF3F8", fg=MUTED, font=(F, 9, "bold"), width=46, anchor="e").pack(side="right", padx=6)
        tk.Label(cols, text="מצב", bg="#EEF3F8", fg=MUTED, font=(F, 9, "bold"), width=14, anchor="e").pack(side="right")
        tk.Label(cols, text="דגש לסעיף (לא חובה)", bg="#EEF3F8", fg=MUTED, font=(F, 9, "bold"), anchor="e").pack(side="right", fill="x", expand=True)
        for n, s in enumerate(ITEMS):
            bg = CARD if n % 2 == 0 else "#F8FAFC"
            row = tk.Frame(grid, bg=bg)
            row.pack(fill="x")
            var = tk.BooleanVar(value=False)
            cb = tk.Checkbutton(row, text=s["title"], variable=var, width=44, anchor="e", bg=bg, activebackground=bg,
                                font=(F, 10), selectcolor="white", fg=TEXT)
            cb.pack(side="right", padx=4)
            st = tk.Label(row, text="—", width=14, anchor="e", bg=bg, fg=MUTED, font=(F, 9, "bold"))
            st.pack(side="right")
            note = tk.Entry(row, justify="right", relief="flat", bg="#F3F6FA", highlightthickness=1,
                            highlightbackground=LINE, highlightcolor=TEAL, font=(F, 9))
            note.pack(side="right", fill="x", expand=True, padx=6, pady=3)
            self.rows[s["key"]] = (var, note, st, cb)

        # הגדרות והרצה
        outer, ctl, _h = self.card(body, "הרצה")
        outer.pack(fill="x")
        line1 = ttk.Frame(ctl, style="Card.TFrame")
        line1.pack(fill="x")
        self.run_btn = ttk.Button(line1, text="▶   הרץ ניתוח", style="Accent.TButton", command=self.start)
        self.run_btn.pack(side="right")
        self.model = ttk.Combobox(line1, values=[m[0] for m in MODELS], state="readonly", width=40, justify="right")
        current = engine.CONFIG.get("model", MODELS[0][1])
        self.model.current(next((i for i, m in enumerate(MODELS) if m[1] == current), 0))
        self.model.pack(side="right", padx=(16, 6))
        self.effort = ttk.Combobox(line1, values=[e[0] for e in EFFORTS], state="readonly", width=24, justify="right")
        self.effort.current(next((i for i, e in enumerate(EFFORTS) if e[1] == engine.CONFIG.get("effort")), 0))
        self.effort.pack(side="right", padx=6)
        self.fix_var = tk.BooleanVar(value=engine.CONFIG.get("fix_values", True))
        tk.Checkbutton(line1, text="תקן אוטומטית ערכים שגויים בטבלה לפי הצילום", variable=self.fix_var, bg=CARD,
                       activebackground=CARD, font=(F, 9), fg=TEXT, selectcolor="white").pack(side="right", padx=10)
        line2 = ttk.Frame(ctl, style="Card.TFrame")
        line2.pack(fill="x", pady=(10, 0))
        ttk.Button(line2, text="✔  אשר טקסט", command=self.approve).pack(side="right")
        ttk.Button(line2, text="✎  נסח מחדש מסומנים", command=self.rephrase).pack(side="right", padx=6)
        self.mode = ttk.Combobox(line2, values=[m[0] for m in REPHRASE], state="readonly", width=16, justify="right")
        self.mode.current(0)
        self.mode.pack(side="right")
        self.bar = ttk.Progressbar(line2, mode="indeterminate", style="Teal.Horizontal.TProgressbar")
        self.bar.pack(side="left", fill="x", expand=True, padx=(0, 12))

        # יומן
        logf = tk.Frame(body, bg=LINE)
        logf.pack(fill="both", pady=(10, 0))
        self.log = tk.Text(logf, height=5, state="disabled", wrap="word", relief="flat", bg="#0F1F31",
                           fg="#D6E2EE", font=(F, 9), padx=10, pady=8, insertbackground="white")
        self.log.tag_configure("r", justify="right")
        self.log.tag_configure("ok", foreground="#5EE0A0", justify="right")
        self.log.tag_configure("err", foreground="#FF8A8A", justify="right")
        self.log.pack(fill="both", padx=1, pady=1)
        self.write("גרור או הוסף דו\"חות, סמן סעיפים ולחץ 'הרץ ניתוח'.  נשלחים רק התמונות והטקסט הקליני - ללא שם, ת.ז, טלפון ומייל.")

    def about(self):
        win = tk.Toplevel(self)
        win.title("אודות")
        win.configure(bg=CARD)
        win.resizable(False, False)
        try:
            win.iconbitmap(resource("icon.ico"))
        except Exception:
            pass
        tk.Frame(win, bg=TEAL, height=6).pack(fill="x")
        tk.Label(win, text=APP_NAME, bg=CARD, fg=NAVY, font=(F, 16, "bold")).pack(padx=40, pady=(18, 0))
        tk.Label(win, text=f"גרסה {VERSION}", bg=CARD, fg=MUTED, font=(F, 10)).pack()
        tk.Label(win, text="ניתוח קליני אוטומטי של DIERS, Stabilimeter ובדיקות הליכה\n"
                           "ישירות לתוך הדו\"ח ב-Word, בעזרת Claude.",
                 bg=CARD, fg=TEXT, font=(F, 10), justify="center").pack(pady=12)
        tk.Label(win, text=f"פותח ועוצב על ידי {AUTHOR}", bg=CARD, fg=NAVY, font=(F, 11, "bold")).pack()
        tk.Label(win, text=AUTHOR_ROLE, bg=CARD, fg=MUTED, font=(F, 10)).pack()
        mail = tk.Label(win, text=AUTHOR_MAIL, bg=CARD, fg=INFO, font=(F, 10, "underline"), cursor="hand2")
        mail.pack(pady=(2, 10))
        mail.bind("<Button-1>", lambda e: webbrowser.open(f"mailto:{AUTHOR_MAIL}"))
        tk.Label(win, text=COPYRIGHT, bg=CARD, fg=MUTED, font=(F, 9)).pack(pady=(0, 6))
        tk.Label(win, text="אין להעתיק, להפיץ או לשנות ללא אישור בכתב.", bg=CARD, fg=MUTED, font=(F, 8)).pack(pady=(0, 16))
        ttk.Button(win, text="סגור", command=win.destroy).pack(pady=(0, 16))
        win.transient(self)
        win.grab_set()

    def write(self, text):
        tag = "ok" if text.strip().startswith("✓") or "נשמר" in text else "err" if ("✗" in text or "שגיאה" in text) else "r"
        self.log.configure(state="normal")
        self.log.insert("end", text + "\n", tag)
        self.log.see("end")
        self.log.configure(state="disabled")

    # ---------- חשבון ----------
    def refresh_account(self):
        def work():
            try:
                ok, desc = engine.auth_status()
            except Exception as e:
                ok, desc = False, str(e)
            self.q.put(("account", (ok, desc)))
        threading.Thread(target=work, daemon=True).start()

    def switch_account(self):
        if messagebox.askyesno("החלפת חשבון", "ייפתח חלון שחור ואחריו הדפדפן להתחברות לחשבון Claude.\n"
                                              "בסיום ההתחברות סגור את החלון השחור ולחץ 'רענן'.\n\nלהמשיך?"):
            try:
                engine.open_login()
            except Exception as e:
                messagebox.showerror("שגיאה", str(e))

    def _enable_drop(self, widget):
        if not DND_FILES:
            return
        try:
            widget.drop_target_register(DND_FILES)
            widget.dnd_bind("<<Drop>>", self.on_drop)
        except Exception:
            pass
        for child in widget.winfo_children():
            self._enable_drop(child)

    # ---------- קבצים ----------
    def choose(self):
        self.add_files(filedialog.askopenfilenames(
            title="בחר דו\"חות", filetypes=[("מסמכי Word", "*.docx *.dotx *.doc *.docm *.dotm"), ("כל הקבצים", "*.*")]))

    def on_drop(self, event):
        self.add_files(self.tk.splitlist(event.data))
        return event.action

    def add_files(self, paths):
        added = 0
        for p in paths:
            p = os.path.normpath(p)
            name = os.path.basename(p)
            if name.startswith("~$"):
                continue
            if os.path.isdir(p):
                inner = [os.path.join(p, f) for f in sorted(os.listdir(p))
                         if f.lower().endswith(engine.WORD_EXT) and not f.startswith("~$")]
                if inner:
                    self.add_files(inner)
                else:
                    self.write(f"✗ בתיקייה '{name}' אין קבצי Word")
                continue
            if not p.lower().endswith(engine.WORD_EXT):
                self.write(f"✗ לא נוסף - זה לא קובץ Word: {name}")
                continue
            try:
                p = engine.ensure_docx(p, log=self.write)
                name = os.path.basename(p)
            except Exception as e:
                self.write(f"✗ {name}: {e}")
                continue
            if p not in self.files:
                self.files.append(p)
                self.flist.insert("end", "   " + name)
                added += 1
        if added:
            self.flist.selection_clear(0, "end")
            self.flist.selection_set("end")
            self.drop_hint.configure(text=f"{len(self.files)} דו\"חות ברשימה  ·  הסימון חל על כולם")
            self.show_selected()

    def remove_selected(self):
        for i in reversed(self.flist.curselection()):
            self.flist.delete(i)
            del self.files[i]
        self.show_selected()

    def clear_files(self):
        self.flist.delete(0, "end")
        self.files = []
        self.drop_hint.configure(text=DROP_HINT)
        self.show_selected()

    def current(self):
        sel = self.flist.curselection()
        return self.files[sel[0]] if sel else (self.files[0] if self.files else None)

    def show_selected(self):
        path = self.current()
        self.status = []
        if path:
            try:
                self.status = engine.Report(path).status()
            except Exception as e:
                self.write(f"שגיאה בפתיחת {os.path.basename(path)}: {e}")
        by_key = {s["key"]: s for s in self.status}
        for key, (var, note, st, cb) in self.rows.items():
            s = by_key.get(key)
            if not s:
                st.configure(text="—", fg=MUTED)
                continue
            need_img = key != "discussion"
            s["ready"] = (s["found"] and (not need_img or s["images"] > 0)) or key == "discussion"
            if key == "discussion" and not s["found"]:
                txt, color = "תיווצר כותרת", INFO
            elif not s["found"]:
                txt, color = "לא נמצא בדו\"ח", MUTED
            elif need_img and not s["images"]:
                txt, color = "אין תמונה", WARN
            elif s["done"]:
                txt, color = "נותח ✓" + (f"  ({s['images']})" if need_img else ""), INFO
            else:
                txt, color = "מוכן" + (f"  ({s['images']} תמ')" if need_img else ""), OK
            st.configure(text=txt, fg=color)

    def mark(self, mode):
        by_key = {s["key"]: s for s in self.status}
        for key, (var, *_r) in self.rows.items():
            s = by_key.get(key, {"found": True, "ready": True, "done": False})
            if mode == "all":
                var.set(True)
            elif mode == "none" or not s.get("ready"):
                var.set(False)
            elif mode == "images":
                var.set(key != "discussion")
            else:
                var.set(not s["done"] and key != "discussion")

    def marked(self):
        return [k for k, (var, *_r) in self.rows.items() if var.get()]

    # ---------- פעולות ----------
    def _ready(self, need_keys=True):
        if self.busy:
            return None
        if not self.files:
            messagebox.showinfo("", "הוסף קודם דו\"ח אחד לפחות.")
            return None
        keys = self.marked()
        if need_keys and not keys:
            messagebox.showinfo("", "לא סומן אף סעיף.")
            return None
        return keys

    def _settings(self):
        engine.CONFIG["effort"] = EFFORTS[self.effort.current()][1]
        engine.CONFIG["fix_values"] = bool(self.fix_var.get())
        return MODELS[self.model.current()][1]

    def _background(self, work):
        self.busy = True
        self.run_btn.configure(state="disabled")
        self.bar.start(10)

        def wrapped():
            try:
                work()
            except Exception as e:
                self.q.put(("log", f"שגיאה: {e}"))
            self.q.put(("idle", None))
        threading.Thread(target=wrapped, daemon=True).start()

    def start(self):
        keys = self._ready()
        if keys is None:
            return
        files = list(self.files)
        warnings = []
        for f in files:
            try:
                for w in engine.precheck(f, keys):
                    warnings.append(f"• {os.path.basename(f)}: {w}" if len(files) > 1 else f"• {w}")
            except Exception as e:
                warnings.append(f"• {os.path.basename(f)}: לא ניתן לפתוח ({e})")
        if warnings and not messagebox.askyesno("בדיקה לפני הרצה", "\n\n".join(warnings) + "\n\nלהמשיך בכל זאת?"):
            return
        notes = {k: self.rows[k][1].get() for k in keys}
        model = self._settings()
        self.write(f"מריץ {len(keys)} סעיפים על {len(files)} דו\"חות  ·  {MODELS[self.model.current()][0].split('·')[0].strip()}")

        def work():
            for i, f in enumerate(files):
                self.q.put(("log", f"── {os.path.basename(f)}  ({i + 1}/{len(files)})"))
                out = engine.default_output(f)
                try:
                    out, errors = engine.run(f, keys, notes, model, out, log=lambda m: self.q.put(("log", "   " + m)))
                    self.q.put(("done", (f, out, errors, len(files) == 1)))
                except PermissionError:
                    self.q.put(("log", f"   ✗ לא ניתן לשמור את '{os.path.basename(out)}' - סגור אותו ב-Word ונסה שוב."))
                except Exception as e:
                    self.q.put(("log", f"   ✗ שגיאה: {e}"))
        self._background(work)

    def rephrase(self):
        keys = self._ready()
        if keys is None:
            return
        mode_name, mode = REPHRASE[self.mode.current()]
        if not messagebox.askyesno("ניסוח מחדש", f"הסעיפים המסומנים שכבר נותחו ינוסחו מחדש ({mode_name}), בלי לשנות ממצאים.\n"
                                                 "יש לסגור את הדו\"חות ב-Word.\n\nלהמשיך?"):
            return
        files = list(self.files)
        notes = " ".join(self.rows[k][1].get() for k in keys).strip()
        model = self._settings()

        def work():
            for f in files:
                self.q.put(("log", f"── {os.path.basename(f)}"))
                try:
                    out, errors = engine.rephrase(f, keys, mode, notes, model, log=lambda m: self.q.put(("log", "   " + m)))
                    self.q.put(("done", (f, out, errors, len(files) == 1)))
                except PermissionError:
                    self.q.put(("log", "   ✗ הקובץ פתוח ב-Word - סגור אותו ונסה שוב."))
        self._background(work)

    def approve(self):
        keys = self._ready(need_keys=False)
        if keys is None:
            return
        what = f"{len(keys)} הסעיפים המסומנים" if keys else "כל הסעיפים שנותחו"
        if not messagebox.askyesno("אישור טקסט", f"הטקסט של {what} יינעל, והכלי לא יחליף אותו יותר בהרצות הבאות.\n"
                                                 "יש לסגור את הדו\"חות ב-Word לפני האישור.\n\nלהמשיך?"):
            return
        for f in self.files:
            try:
                rep = engine.Report(f)
                n = rep.approve(keys or None)
                rep.save(f)
                self.write(f"✓ {os.path.basename(f)}: אושרו {n} שורות.")
            except PermissionError:
                self.write(f"✗ {os.path.basename(f)}: הקובץ פתוח ב-Word - סגור אותו ונסה שוב.")
            except Exception as e:
                self.write(f"✗ {os.path.basename(f)}: שגיאה - {e}")
        self.show_selected()

    def _poll(self):
        try:
            while True:
                kind, data = self.q.get_nowait()
                if kind == "log":
                    self.write(data)
                elif kind == "account":
                    ok, desc = data
                    self.account.configure(text=("●  " + desc) if ok else "●  לא מחובר - לחץ 'החלף חשבון'",
                                           fg="#5EE0A0" if ok else "#FF8A8A")
                elif kind == "idle":
                    self.busy = False
                    self.bar.stop()
                    self.run_btn.configure(state="normal")
                    self.show_selected()
                    self.write("✓ הסתיים.")
                elif kind == "done":
                    src, out, errors, open_it = data
                    self.write(f"   ✓ נשמר: {os.path.basename(out)}" + (f"  |  ✗ {len(errors)} סעיפים נכשלו" if errors else ""))
                    if src in self.files:
                        i = self.files.index(src)
                        self.files[i] = out
                        self.flist.delete(i)
                        self.flist.insert(i, "   " + os.path.basename(out))
                        self.flist.selection_set(i)
                    if open_it:
                        try:
                            os.startfile(out)
                        except Exception:
                            pass
        except queue.Empty:
            pass
        self.after(150, self._poll)


def main():
    import traceback
    app = App()

    def on_error(exc, val, tb):
        log_error("".join(traceback.format_exception(exc, val, tb)))
        messagebox.showerror("שגיאה", f"{val}\n\nהפרטים נשמרו בקובץ שגיאות.log בתיקיית התוכנה.")
    app.report_callback_exception = on_error
    if len(sys.argv) > 1:          # קבצים שנגררו על אייקון התוכנה או נפתחו איתה
        app.after(300, lambda: app.add_files(sys.argv[1:]))
    app.mainloop()


if __name__ == "__main__":
    try:
        main()
    except Exception:
        import traceback
        log_error(traceback.format_exc())
        raise
