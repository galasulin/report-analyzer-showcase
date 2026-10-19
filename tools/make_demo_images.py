# -*- coding: utf-8 -*-
"""יוצר צילומי DIERS / Pedoscan סינתטיים (נתונים בדיוניים לגמרי) לצורך הדגמה."""
import os
import sys

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
import numpy as np  # noqa: E402
from matplotlib.patches import Ellipse  # noqa: E402

OUT = sys.argv[1] if len(sys.argv) > 1 else "."
os.makedirs(OUT, exist_ok=True)
plt.rcParams["font.family"] = "Segoe UI"


def diers_coronal(path):
    fig = plt.figure(figsize=(14, 7.5), dpi=110)
    fig.patch.set_facecolor("white")
    ax = fig.add_axes([0.01, 0.52, 0.34, 0.46])
    ax.axis("off")
    rows = [("Flèche Cervicale", "62", "mm", ""), ("Flèche Lombaire", "38", "mm", ""),
            ("Sagittal Imbalance VP-DM", "41", "mm A", "y"), ("Coronal Imbalance VP-DM", "9", "mm R", "y"),
            ("Pelvic Obliquity", "5", "mm L", "y"), ("Pelvic Torsion DL-DR", "2", "° R", ""),
            ("Kyphotic Angle ICT-ITL (max)", "54", "°", "g"), ("Lordotic Angle ITL-ILS (max)", "39", "°", "y"),
            ("ScoliosisAngle", "8", "L", ""), ("Pelvis Rotation", "3", "° R", "y"),
            ("Apical Deviation VP-DM (+max)", "6", "mm R (~T7)", ""), ("Apical Deviation VP-DM (-max)", "4", "mm L (~L2)", "")]
    colors = {"y": "#FFD400", "g": "#0A7D2C", "": "#E6E6E6"}
    ax.text(0.0, 1.0, "Parameter", fontsize=10, va="top")
    ax.text(0.66, 1.0, "Value", fontsize=10, va="top")
    ax.text(0.76, 1.0, "Unit", fontsize=10, va="top")
    for i, (p, v, u, c) in enumerate(rows):
        y = 0.92 - i * 0.077
        ax.add_patch(plt.Rectangle((0.0, y - 0.06), 0.98, 0.07, color=colors[c], transform=ax.transAxes))
        tc = "white" if c == "g" else "black"
        ax.text(0.02, y - 0.025, p, fontsize=9.5, va="center", color=tc)
        ax.text(0.72, y - 0.025, v, fontsize=9.5, va="center", ha="right", color=tc)
        ax.text(0.76, y - 0.025, u, fontsize=9.5, va="center", color=tc)
    # עקומות projection
    z = np.linspace(0, 1, 200)
    for k, (x0, title, f) in enumerate([(0.04, "Lateral projection (mm)", lambda t: 40 * np.sin(2.2 * np.pi * t - 0.4)),
                                         (0.24, "Frontal projection (mm)", lambda t: 6 * np.sin(np.pi * t)),
                                         (0.44, "Surface rotation (°)", lambda t: 4 * np.sin(3 * np.pi * t))]):
        a = fig.add_axes([x0, 0.06, 0.16, 0.42])
        a.plot(f(z), -z * 500, color=["#1F3A93", "#1F3A93", "#D62828"][k], lw=1.8)
        a.axvline(0, color="gray", lw=0.6)
        a.set_xlim(-100 if k < 2 else -20, 100 if k < 2 else 20)
        a.grid(True, color="#DDDDDD")
        a.set_xlabel(title, fontsize=8)
        a.tick_params(labelsize=7)
    # "גב" עם מפת חום
    b = fig.add_axes([0.66, 0.05, 0.32, 0.9])
    yy, xx = np.mgrid[-1:1:300j, -1:1:200j]
    body = (xx ** 2 / (0.75 - 0.15 * yy) ** 2 + (yy + 0.1) ** 2 / 1.1 ** 2) < 1
    heat = np.sin(6 * xx) * np.cos(4 * yy) + 0.3 * np.sin(13 * yy)
    heat = np.where(body, heat, np.nan)
    b.imshow(heat, cmap="coolwarm", origin="upper")
    b.plot(100 + 6 * np.sin(np.linspace(0, np.pi, 50)), np.linspace(30, 260, 50), color="#2E2E2E", lw=6, alpha=0.6)
    b.axis("off")
    fig.savefig(path)
    plt.close(fig)


def pedoscan(path):
    """Pedoscan סטטי בדיוני: עיניים פקוחות (ירוק) מול עצומות (אדום), עם ערכים מספריים."""
    fig = plt.figure(figsize=(13, 6.4), dpi=110)
    rng = np.random.default_rng(7)
    data = {"open": {"L": (2.4, 11.8), "T": (7.6, 6.1), "R": (1.9, 7.4), "w": (46, 54), "ap": (44, 56), "col": "#0A7D2C"},
            "closed": {"L": (3.1, 15.2), "T": (9.8, 8.3), "R": (2.2, 9.9), "w": (44, 56), "ap": (48, 52), "col": "#C00000"}}
    for row, cond in enumerate(("open", "closed")):
        d = data[cond]
        label = "EYES OPEN  (selected examination)" if cond == "open" else "EYES CLOSED  (comparative examination)"
        fig.text(0.02, 0.95 - row * 0.48, label, color=d["col"], fontsize=11, weight="bold")
        tx = fig.add_axes([0.02, 0.55 - row * 0.48, 0.22, 0.36])
        tx.axis("off")
        lines = [("Weight distribution", ""), ("LEFT", f"{d['w'][0]} %"), ("RIGHT", f"{d['w'][1]} %"),
                 ("ANTERIOR", f"{d['ap'][0]} %"), ("POSTERIOR", f"{d['ap'][1]} %")]
        for i, (k, v) in enumerate(lines):
            tx.text(0.0, 0.9 - i * 0.19, k, fontsize=10, weight="bold" if i == 0 else "normal")
            tx.text(0.75, 0.9 - i * 0.19, v, fontsize=10, color=d["col"])
        for j, (side, name) in enumerate((("L", "COP left"), ("T", "COP"), ("R", "COP right"))):
            ax = fig.add_axes([0.27 + j * 0.24, 0.53 - row * 0.48, 0.2, 0.38])
            ml, ap = d[side]
            ax.add_patch(Ellipse((0, 0), 18, 26, fc="#DDEBC4", ec="black", ls="--"))
            t = np.cumsum(rng.normal(size=(300, 2)) * [ml / 9, ap / 9], axis=0)
            t -= t.mean(axis=0)
            ax.plot(t[:, 0], t[:, 1], color=d["col"], lw=1)
            ax.set_xlim(-15, 15)
            ax.set_ylim(-18, 18)
            ax.set_xticks([])
            ax.set_yticks([])
            ax.set_title(f"{name}   a/p {ap} mm   lat {ml} mm", fontsize=9)
            ax.set_aspect("equal")
    fig.text(0.5, 0.005, "PEDOSCAN STATIC  |  demo data  |  normative ellipse = green", ha="center", fontsize=8, color="gray")
    fig.savefig(path)
    plt.close(fig)


diers_coronal(os.path.join(OUT, "demo_diers_coronal.png"))
pedoscan(os.path.join(OUT, "demo_pedoscan.png"))
print("ok")
