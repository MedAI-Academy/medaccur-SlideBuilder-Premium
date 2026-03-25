"""Chart renderer — generates matplotlib/PIL charts as PNG bytes for PPTX insertion."""
import io
import math
from typing import Optional
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import numpy as np


# ══════════════════════════════════════════════════════════
# THEME COLORS (RGB tuples for matplotlib)
# ══════════════════════════════════════════════════════════
THEME_COLORS = {
    "dark": {
        "bg": "#0D2B4E", "surface": "#163060", "text": "#EAF0FF",
        "muted": "#7B9FD4", "dim": "#4A6A9A",
        "accent": "#7C6FFF", "teal": "#22D3A5", "gold": "#F5C842", "rose": "#FF5F7E",
    },
    "light": {
        "bg": "#FFFFFF", "surface": "#E2E8F0", "text": "#2D3748",
        "muted": "#718096", "dim": "#A0AEC0",
        "accent": "#3182CE", "teal": "#319795", "gold": "#D69E2E", "rose": "#E53E3E",
    },
    "gray": {
        "bg": "#34495E", "surface": "#415B73", "text": "#ECF0F1",
        "muted": "#BDC3C7", "dim": "#7F8C8D",
        "accent": "#E74C3C", "teal": "#1ABC9C", "gold": "#F39C12", "rose": "#E74C3C",
    },
    "pharma": {
        "bg": "#F0F9FF", "surface": "#DBEAFE", "text": "#1E3A5F",
        "muted": "#4A6FA5", "dim": "#93B4D4",
        "accent": "#2563EB", "teal": "#0D9488", "gold": "#D97706", "rose": "#DC2626",
    },
    "premium": {
        "bg": "#111111", "surface": "#1C1C1C", "text": "#E8E8E8",
        "muted": "#999999", "dim": "#555555",
        "accent": "#C9A84C", "teal": "#C9A84C", "gold": "#C9A84C", "rose": "#B33030",
    },
}


def _get_colors(theme: str) -> dict:
    return THEME_COLORS.get(theme, THEME_COLORS["dark"])


# ══════════════════════════════════════════════════════════
# KAPLAN-MEIER CURVE (approximation)
# ══════════════════════════════════════════════════════════
def render_km_curve(
    drug_name: str,
    drug_mpfs_months: float,
    control_mpfs_months: float,
    n_total: str = "N/A",
    theme: str = "dark",
    dpi: int = 200,
    width_inches: float = 7.0,
    height_inches: float = 4.0,
) -> bytes:
    """Render an approximated KM curve as PNG bytes."""
    c = _get_colors(theme)
    fig, ax = plt.subplots(figsize=(width_inches, height_inches), dpi=dpi)
    fig.patch.set_facecolor(c["bg"])
    ax.set_facecolor(c["surface"])

    max_t = max(drug_mpfs_months, control_mpfs_months) * 1.5
    t = np.linspace(0, max_t, 200)

    # Exponential decay: S(t) = exp(-ln2 * t / median)
    drug_surv = np.exp(-0.693 * t / drug_mpfs_months) * 100
    ctrl_surv = np.exp(-0.693 * t / control_mpfs_months) * 100

    # Step-function style
    ax.step(t, drug_surv, where="post", color=c["teal"], linewidth=2.5, label=f"{drug_name}: {drug_mpfs_months} mo")
    ax.step(t, ctrl_surv, where="post", color=c["rose"], linewidth=2.5, label=f"Control: {control_mpfs_months} mo")

    # Median lines
    ax.axhline(y=50, color=c["dim"], linewidth=0.8, linestyle="--", alpha=0.5)
    ax.axvline(x=drug_mpfs_months, color=c["teal"], linewidth=1, linestyle=":", alpha=0.6)
    ax.axvline(x=control_mpfs_months, color=c["rose"], linewidth=1, linestyle=":", alpha=0.6)

    ax.set_xlabel("Months", fontsize=9, color=c["muted"], fontfamily="sans-serif")
    ax.set_ylabel("PFS Probability (%)", fontsize=9, color=c["muted"], fontfamily="sans-serif")
    ax.set_xlim(0, max_t)
    ax.set_ylim(0, 105)
    ax.tick_params(colors=c["dim"], labelsize=8)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["bottom"].set_color(c["dim"])
    ax.spines["left"].set_color(c["dim"])
    ax.grid(axis="y", color=c["dim"], alpha=0.2, linewidth=0.5)

    legend = ax.legend(
        loc="upper right", fontsize=8, frameon=True,
        facecolor=c["surface"], edgecolor=c["dim"], labelcolor=c["text"],
    )

    # Patients at risk annotation
    ax.text(
        0.02, -0.12, f"Patients at Risk: N={n_total}",
        transform=ax.transAxes, fontsize=7, color=c["dim"], fontfamily="monospace",
    )

    buf = io.BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight", facecolor=fig.get_facecolor(), edgecolor="none")
    plt.close(fig)
    buf.seek(0)
    return buf.read()


# ══════════════════════════════════════════════════════════
# HORIZONTAL BAR CHART (mPFS comparison)
# ══════════════════════════════════════════════════════════
def render_mpfs_bar_chart(
    rows: list[dict],
    focus_drug: str = "",
    theme: str = "dark",
    dpi: int = 200,
    width_inches: float = 9.0,
    height_inches: float = 5.5,
) -> bytes:
    """Render horizontal mPFS comparison bar chart as PNG bytes."""
    c = _get_colors(theme)

    # Parse mPFS values
    items = []
    for r in rows:
        name = str(r.get("drug_trial", ""))[:35]
        mpfs_raw = str(r.get("mpfs", ""))
        is_focus = bool(r.get("is_focus", False))
        # Extract numeric value
        val = 0.0
        import re
        m = re.search(r"([\d.]+)\s*mo", mpfs_raw)
        if m:
            val = float(m.group(1))
        elif not re.search(r"(?i)nr|not\s*reached", mpfs_raw):
            m2 = re.search(r"~?([\d.]+)", mpfs_raw)
            if m2:
                val = float(m2.group(1))
        if val > 0:
            items.append({"name": name, "val": val, "is_focus": is_focus or (focus_drug and focus_drug.lower() in name.lower())})

    if len(items) < 2:
        return b""

    items.sort(key=lambda x: x["val"])
    items = items[-12:]  # Top 12

    fig, ax = plt.subplots(figsize=(width_inches, height_inches), dpi=dpi)
    fig.patch.set_facecolor(c["bg"])
    ax.set_facecolor(c["bg"])

    names = [it["name"] for it in items]
    vals = [it["val"] for it in items]
    colors = [c["teal"] if it["is_focus"] else c["accent"] for it in items]
    alphas = [1.0 if it["is_focus"] else 0.75 for it in items]

    bars = ax.barh(names, vals, color=colors, height=0.65, edgecolor="none")
    for bar, alpha in zip(bars, alphas):
        bar.set_alpha(alpha)

    # Value labels
    for bar, val in zip(bars, vals):
        ax.text(
            bar.get_width() + 0.3, bar.get_y() + bar.get_height() / 2,
            f"{val:.1f} mo", va="center", fontsize=8, color=c["text"], fontweight="bold",
        )

    ax.set_xlabel("Median PFS (months)", fontsize=9, color=c["muted"])
    ax.tick_params(axis="y", colors=c["text"], labelsize=8)
    ax.tick_params(axis="x", colors=c["dim"], labelsize=8)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["bottom"].set_color(c["dim"])
    ax.spines["left"].set_color(c["dim"])
    ax.grid(axis="x", color=c["dim"], alpha=0.15, linewidth=0.5)
    ax.set_xlim(0, max(vals) * 1.2)

    # Legend
    from matplotlib.patches import Patch
    legend_elements = [
        Patch(facecolor=c["teal"], label="★ Focus Drug"),
        Patch(facecolor=c["accent"], alpha=0.75, label="Competitor"),
    ]
    ax.legend(handles=legend_elements, loc="lower right", fontsize=8,
              frameon=True, facecolor=c["surface"], edgecolor=c["dim"], labelcolor=c["text"])

    buf = io.BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight", facecolor=fig.get_facecolor(), edgecolor="none")
    plt.close(fig)
    buf.seek(0)
    return buf.read()


# ══════════════════════════════════════════════════════════
# SWOT QUADRANT (visual)
# ══════════════════════════════════════════════════════════
def render_swot_chart(
    strengths: list[str],
    weaknesses: list[str],
    opportunities: list[str],
    threats: list[str],
    theme: str = "dark",
    dpi: int = 200,
    width_inches: float = 10.0,
    height_inches: float = 6.0,
) -> bytes:
    """Render SWOT as a visual quadrant chart."""
    c = _get_colors(theme)
    fig, axes = plt.subplots(2, 2, figsize=(width_inches, height_inches), dpi=dpi)
    fig.patch.set_facecolor(c["bg"])
    fig.subplots_adjust(hspace=0.15, wspace=0.1)

    quads = [
        (axes[0, 0], "Strengths", strengths, c["teal"]),
        (axes[0, 1], "Weaknesses", weaknesses, c["rose"]),
        (axes[1, 0], "Opportunities", opportunities, c["accent"]),
        (axes[1, 1], "Threats", threats, c["gold"]),
    ]

    for ax, title, items, color in quads:
        ax.set_facecolor(c["surface"])
        ax.set_xlim(0, 1)
        ax.set_ylim(0, 1)
        ax.axis("off")

        # Title bar
        ax.add_patch(plt.Rectangle((0, 0.88), 1, 0.12, facecolor=color, transform=ax.transAxes, clip_on=False))
        ax.text(0.5, 0.94, title.upper(), transform=ax.transAxes, ha="center", va="center",
                fontsize=12, fontweight="bold", color=c["bg"], fontfamily="sans-serif")

        # Bullets
        for i, item in enumerate(items[:5]):
            text = str(item) if isinstance(item, str) else str(item.get("bullet", item.get("text", item.get("title", str(item)))))
            y = 0.80 - i * 0.16
            if y < 0.05:
                break
            ax.text(0.05, y, f"• {text[:80]}", transform=ax.transAxes, fontsize=8,
                    color=c["text"], fontfamily="sans-serif", va="top", wrap=True)

    buf = io.BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight", facecolor=fig.get_facecolor(), edgecolor="none")
    plt.close(fig)
    buf.seek(0)
    return buf.read()
