"""Chart renderer — generates matplotlib/PIL charts as PNG bytes for PPTX insertion.

Charts:
  - KM curves (stepped, endpoint-aware: PFS or OS)
  - mPFS horizontal bar chart
  - ORR grouped bar chart (NEW)
  - SWOT quadrant (fixed text overflow)
"""
import io
import math
import re
from typing import Optional
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import matplotlib.patches as mpatches
import numpy as np


# ══════════════════════════════════════════════════════════
# THEME COLORS
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
    "normal": {
        "bg": "#0D2B4E", "surface": "#163060", "text": "#EAF0FF",
        "muted": "#7B9FD4", "dim": "#4A6A9A",
        "accent": "#7C6FFF", "teal": "#22D3A5", "gold": "#F5C842", "rose": "#FF5F7E",
    },
    "gold": {
        "bg": "#111111", "surface": "#1C1C1C", "text": "#E8E8E8",
        "muted": "#999999", "dim": "#555555",
        "accent": "#C9A84C", "teal": "#C9A84C", "gold": "#D4AF37", "rose": "#B33030",
    },
    "aquarell": {
        "bg": "#F0F4F8", "surface": "#E2E8F0", "text": "#2D3748",
        "muted": "#718096", "dim": "#A0AEC0",
        "accent": "#3182CE", "teal": "#319795", "gold": "#D69E2E", "rose": "#E53E3E",
    },
}


def _get_colors(theme: str) -> dict:
    return THEME_COLORS.get(theme, THEME_COLORS["dark"])


def _generate_km_steps(median_months, n_patients=80, seed=42):
    """Generate realistic stepped KM curve from median survival using Weibull."""
    rng = np.random.RandomState(seed)
    scale = median_months / (np.log(2) ** (1.0 / 1.2))
    event_times = rng.weibull(1.2, n_patients) * scale
    censor_times = rng.uniform(0, median_months * 2.5, n_patients)
    observed = event_times < censor_times
    times = np.minimum(event_times, censor_times)

    order = np.argsort(times)
    times = times[order]
    observed = observed[order]

    n_at_risk = n_patients
    surv = 1.0
    t_steps = [0.0]
    s_steps = [100.0]

    for t, is_event in zip(times, observed):
        if is_event and n_at_risk > 0:
            surv *= (1.0 - 1.0 / n_at_risk)
            t_steps.append(float(t))
            s_steps.append(surv * 100.0)
        n_at_risk -= 1

    return np.array(t_steps), np.array(s_steps)


# ══════════════════════════════════════════════════════════
# KAPLAN-MEIER CURVE (realistic stepped)
# ══════════════════════════════════════════════════════════
def render_km_curve(
    drug_name: str,
    drug_median: float,
    control_median: float,
    endpoint: str = "PFS",
    n_total: str = "N/A",
    hr: str = "",
    p_value: str = "",
    theme: str = "dark",
    dpi: int = 200,
    width_inches: float = 7.0,
    height_inches: float = 4.0,
    # Legacy parameter names (backwards compatible)
    drug_mpfs_months: float = 0,
    control_mpfs_months: float = 0,
) -> bytes:
    """Render a realistic stepped KM curve as PNG bytes."""
    # Handle legacy parameter names
    if drug_mpfs_months > 0 and drug_median == 0:
        drug_median = drug_mpfs_months
    if control_mpfs_months > 0 and control_median == 0:
        control_median = control_mpfs_months

    if drug_median <= 0 or control_median <= 0:
        return b""

    c = _get_colors(theme)
    fig, ax = plt.subplots(figsize=(width_inches, height_inches), dpi=dpi)
    fig.patch.set_facecolor(c["bg"])
    ax.set_facecolor(c["surface"])

    # Different seeds per arm AND per endpoint so PFS and OS look different
    seed_offset = 0 if endpoint == "PFS" else 500
    t_drug, s_drug = _generate_km_steps(drug_median, n_patients=90,
                                         seed=int(drug_median * 100) + seed_offset)
    t_ctrl, s_ctrl = _generate_km_steps(control_median, n_patients=90,
                                         seed=int(control_median * 100) + 7 + seed_offset)

    ax.step(t_drug, s_drug, where="post", color=c["teal"], linewidth=2.2,
            label=f"{drug_name}: {drug_median} mo")
    ax.step(t_ctrl, s_ctrl, where="post", color=c["rose"], linewidth=2.2,
            label=f"Control: {control_median} mo")

    # Shaded area between curves
    max_t = max(t_drug[-1], t_ctrl[-1])
    common_t = np.linspace(0, max_t, 300)
    d_interp = np.interp(common_t, t_drug, s_drug)
    c_interp = np.interp(common_t, t_ctrl, s_ctrl)
    ax.fill_between(common_t, d_interp, c_interp, alpha=0.07, color=c["teal"])

    # Median lines
    ax.axhline(y=50, color=c["dim"], linewidth=0.8, linestyle="--", alpha=0.4)
    ax.axvline(x=drug_median, color=c["teal"], linewidth=1, linestyle=":", alpha=0.5)
    ax.axvline(x=control_median, color=c["rose"], linewidth=1, linestyle=":", alpha=0.5)

    y_label = f"{'Overall Survival' if endpoint == 'OS' else 'PFS'} Probability (%)"
    ax.set_xlabel("Months", fontsize=9, color=c["muted"])
    ax.set_ylabel(y_label, fontsize=9, color=c["muted"])
    ax.set_xlim(0, max_t * 1.05)
    ax.set_ylim(0, 105)
    ax.tick_params(colors=c["dim"], labelsize=8)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["bottom"].set_color(c["dim"])
    ax.spines["left"].set_color(c["dim"])
    ax.grid(axis="y", color=c["dim"], alpha=0.15, linewidth=0.5)

    ax.legend(loc="upper right", fontsize=8, frameon=True,
              facecolor=c["surface"], edgecolor=c["dim"], labelcolor=c["text"])

    # Annotation
    parts = []
    if hr:
        parts.append(f"HR {hr}")
    if p_value:
        parts.append(f"p={p_value}")
    parts.append(f"N={n_total}")
    ax.text(0.02, -0.12, "  |  ".join(parts),
            transform=ax.transAxes, fontsize=7, color=c["dim"], fontfamily="monospace")

    buf = io.BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight", facecolor=fig.get_facecolor(), edgecolor="none")
    plt.close(fig)
    buf.seek(0)
    return buf.read()


# ══════════════════════════════════════════════════════════
# HORIZONTAL BAR CHART (mPFS comparison)
# ══════════════════════════════════════════════════════════
def render_mpfs_bar_chart(
    rows: list,
    focus_drug: str = "",
    theme: str = "dark",
    dpi: int = 200,
    width_inches: float = 9.0,
    height_inches: float = 5.5,
) -> bytes:
    """Render horizontal mPFS comparison bar chart."""
    c = _get_colors(theme)

    items = []
    for r in rows:
        name = str(r.get("drug_trial", ""))[:40]
        mpfs_raw = str(r.get("mpfs", ""))
        is_focus = bool(r.get("is_focus", False))
        val = 0.0
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
    items = items[-14:]

    fig, ax = plt.subplots(figsize=(width_inches, height_inches), dpi=dpi)
    fig.patch.set_facecolor(c["bg"])
    ax.set_facecolor(c["bg"])

    names = [it["name"] for it in items]
    vals = [it["val"] for it in items]
    colors = [c["teal"] if it["is_focus"] else c["accent"] for it in items]

    bars = ax.barh(names, vals, color=colors, height=0.6, edgecolor="none")
    for bar, it in zip(bars, items):
        bar.set_alpha(1.0 if it["is_focus"] else 0.7)

    for bar, val in zip(bars, vals):
        ax.text(bar.get_width() + 0.3, bar.get_y() + bar.get_height() / 2,
                f"{val:.1f} mo", va="center", fontsize=8, color=c["text"], fontweight="bold")

    ax.set_xlabel("Median PFS (months)", fontsize=9, color=c["muted"])
    ax.tick_params(axis="y", colors=c["text"], labelsize=8)
    ax.tick_params(axis="x", colors=c["dim"], labelsize=8)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["bottom"].set_color(c["dim"])
    ax.spines["left"].set_color(c["dim"])
    ax.grid(axis="x", color=c["dim"], alpha=0.15, linewidth=0.5)
    ax.set_xlim(0, max(vals) * 1.2)

    legend_elements = [
        mpatches.Patch(facecolor=c["teal"], label="★ Focus Drug"),
        mpatches.Patch(facecolor=c["accent"], alpha=0.7, label="Competitor"),
    ]
    ax.legend(handles=legend_elements, loc="lower right", fontsize=8,
              frameon=True, facecolor=c["surface"], edgecolor=c["dim"], labelcolor=c["text"])

    buf = io.BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight", facecolor=fig.get_facecolor(), edgecolor="none")
    plt.close(fig)
    buf.seek(0)
    return buf.read()


# ══════════════════════════════════════════════════════════
# ORR GROUPED BAR CHART (NEW)
# ══════════════════════════════════════════════════════════
def render_orr_bar_chart(
    rows: list,
    focus_drug: str = "",
    theme: str = "dark",
    dpi: int = 200,
    width_inches: float = 9.0,
    height_inches: float = 5.5,
) -> bytes:
    """Render vertical ORR comparison bar chart."""
    c = _get_colors(theme)

    items = []
    for r in rows:
        name = str(r.get("drug_trial", ""))[:28]
        orr_raw = str(r.get("orr", ""))
        is_focus = bool(r.get("is_focus", False))
        vals = re.findall(r"([\d.]+)\s*%?", orr_raw)
        if not vals:
            continue
        drug_orr = float(vals[0])
        ctrl_orr = float(vals[1]) if len(vals) > 1 else None
        if drug_orr > 0:
            items.append({
                "name": name, "drug_orr": drug_orr, "ctrl_orr": ctrl_orr,
                "is_focus": is_focus or (focus_drug and focus_drug.lower() in name.lower()),
            })

    if len(items) < 2:
        return b""

    items.sort(key=lambda x: x["drug_orr"], reverse=True)
    items = items[:12]

    fig, ax = plt.subplots(figsize=(width_inches, height_inches), dpi=dpi)
    fig.patch.set_facecolor(c["bg"])
    ax.set_facecolor(c["bg"])

    x = np.arange(len(items))
    bar_w = 0.35
    has_ctrl = any(it["ctrl_orr"] is not None for it in items)
    drug_colors = [c["teal"] if it["is_focus"] else c["accent"] for it in items]
    drug_vals = [it["drug_orr"] for it in items]

    if has_ctrl:
        bars1 = ax.bar(x - bar_w/2, drug_vals, bar_w, color=drug_colors, edgecolor="none", alpha=0.9)
        ctrl_vals = [it["ctrl_orr"] or 0 for it in items]
        bars2 = ax.bar(x + bar_w/2, ctrl_vals, bar_w, color=c["dim"], edgecolor="none", alpha=0.5)
        for bar, val in zip(bars2, ctrl_vals):
            if val > 0:
                ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 1,
                        f"{val:.0f}%", ha="center", va="bottom", fontsize=7, color=c["dim"])
    else:
        bars1 = ax.bar(x, drug_vals, bar_w * 1.5, color=drug_colors, edgecolor="none", alpha=0.9)

    for bar, val, it in zip(bars1, drug_vals, items):
        color = c["teal"] if it["is_focus"] else c["text"]
        ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 1,
                f"{val:.0f}%", ha="center", va="bottom", fontsize=8, color=color, fontweight="bold")

    ax.set_xticks(x)
    ax.set_xticklabels([it["name"] for it in items], rotation=35, ha="right", fontsize=7)
    ax.set_ylabel("Overall Response Rate (%)", fontsize=9, color=c["muted"])
    ax.set_ylim(0, min(max(drug_vals) * 1.25, 110))
    ax.tick_params(axis="x", colors=c["text"], labelsize=7)
    ax.tick_params(axis="y", colors=c["dim"], labelsize=8)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["bottom"].set_color(c["dim"])
    ax.spines["left"].set_color(c["dim"])
    ax.grid(axis="y", color=c["dim"], alpha=0.15, linewidth=0.5)

    legend_elements = [
        mpatches.Patch(facecolor=c["teal"], label="★ Focus Drug"),
        mpatches.Patch(facecolor=c["accent"], alpha=0.7, label="Competitor"),
    ]
    if has_ctrl:
        legend_elements.append(mpatches.Patch(facecolor=c["dim"], alpha=0.5, label="Control Arm"))
    ax.legend(handles=legend_elements, loc="upper right", fontsize=8,
              frameon=True, facecolor=c["surface"], edgecolor=c["dim"], labelcolor=c["text"])

    fig.tight_layout()
    buf = io.BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight", facecolor=fig.get_facecolor(), edgecolor="none")
    plt.close(fig)
    buf.seek(0)
    return buf.read()


# ══════════════════════════════════════════════════════════
# SWOT QUADRANT (fixed text overflow)
# ══════════════════════════════════════════════════════════
def render_swot_chart(
    strengths: list,
    weaknesses: list,
    opportunities: list,
    threats: list,
    theme: str = "dark",
    dpi: int = 200,
    width_inches: float = 10.0,
    height_inches: float = 6.0,
) -> bytes:
    """Render SWOT as a visual quadrant chart with proper text wrapping."""
    c = _get_colors(theme)
    fig, axes = plt.subplots(2, 2, figsize=(width_inches, height_inches), dpi=dpi)
    fig.patch.set_facecolor(c["bg"])
    fig.subplots_adjust(hspace=0.12, wspace=0.08)

    quads = [
        (axes[0, 0], "STRENGTHS", strengths, c["teal"]),
        (axes[0, 1], "WEAKNESSES", weaknesses, c["rose"]),
        (axes[1, 0], "OPPORTUNITIES", opportunities, c["accent"]),
        (axes[1, 1], "THREATS", threats, c["gold"]),
    ]

    def _clean(item):
        if isinstance(item, str):
            return item.strip()
        if isinstance(item, dict):
            return str(item.get("bullet", item.get("text", item.get("title", str(item))))).strip()
        return str(item).strip()

    def _truncate(text, max_chars=52):
        if len(text) <= max_chars:
            return text
        return text[:max_chars - 1] + "…"

    for ax, title, items, color in quads:
        ax.set_facecolor(c["surface"])
        ax.set_xlim(0, 1)
        ax.set_ylim(0, 1)
        ax.axis("off")

        # Title bar
        title_rect = mpatches.FancyBboxPatch(
            (0.02, 0.87), 0.96, 0.11,
            boxstyle="round,pad=0.01", facecolor=color, edgecolor="none",
            transform=ax.transAxes, clip_on=False
        )
        ax.add_patch(title_rect)
        ax.text(0.5, 0.925, title, transform=ax.transAxes, ha="center", va="center",
                fontsize=11, fontweight="bold", color="white", fontfamily="sans-serif")

        # Bullets
        max_items = min(len(items), 5)
        spacing = 0.15 if max_items <= 4 else 0.125

        for i in range(max_items):
            text = _truncate(_clean(items[i]))
            y = 0.78 - i * spacing
            if y < 0.02:
                break
            ax.text(0.06, y, f"• {text}", transform=ax.transAxes, fontsize=7.5,
                    color=c["text"], fontfamily="sans-serif", va="top", clip_on=True)

        # Border
        for spine in ax.spines.values():
            spine.set_visible(True)
            spine.set_color(c["dim"])
            spine.set_linewidth(0.5)

    buf = io.BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight", facecolor=fig.get_facecolor(), edgecolor="none")
    plt.close(fig)
    buf.seek(0)
    return buf.read()
