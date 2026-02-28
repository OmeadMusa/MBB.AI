#!/usr/bin/env python3
"""
Chart generation utilities for MBB deliverables.

Uses matplotlib for static chart images (embeddable in PPTX).
Uses openpyxl charts for native Excel charts.

All functions use the MBB color palette by default and export at 150 DPI.
"""

from __future__ import annotations

import os

# ── MBB Color Palette ────────────────────────────────────────────────────────

MBB_COLORS = [
    "#003A70",  # Dark Navy (primary)
    "#4472C4",  # Steel Blue (secondary)
    "#00B0F0",  # Teal (accent)
    "#00B050",  # Green (positive)
    "#FF0000",  # Red (negative)
    "#808080",  # Gray (neutral)
    "#FFC000",  # Amber
    "#7030A0",  # Purple
]

MBB_POSITIVE = "#00B050"
MBB_NEGATIVE = "#FF0000"
MBB_NEUTRAL = "#4472C4"

DEFAULT_DPI = 150
DEFAULT_FONT = "Calibri"


def _check_matplotlib():
    """Check if matplotlib is available."""
    try:
        import matplotlib
        matplotlib.use("Agg")  # Non-interactive backend
        import matplotlib.pyplot as plt
        return plt
    except ImportError:
        print("matplotlib not installed. Install with: pip install matplotlib")
        print("Chart generation requires matplotlib.")
        return None


# ── Matplotlib Chart Functions ───────────────────────────────────────────────

def create_bar_chart(data: list[float], labels: list[str], title: str,
                     filepath: str, colors: list[str] = None,
                     horizontal: bool = True, source: str = "",
                     value_labels: bool = True) -> str | None:
    """
    Create a bar chart (horizontal by default — MBB standard for comparisons).

    Args:
        data: Values for each bar
        labels: Labels for each bar
        title: Chart title (should be an assertion)
        filepath: Output path for image
        colors: Optional custom colors (defaults to MBB palette)
        horizontal: If True, horizontal bars (default)
        source: Source citation text
        value_labels: If True, show values on bars

    Returns:
        filepath if successful, None if matplotlib unavailable
    """
    plt = _check_matplotlib()
    if plt is None:
        return None

    colors = colors or MBB_COLORS[:len(data)]
    if len(colors) < len(data):
        colors = colors * (len(data) // len(colors) + 1)

    fig, ax = plt.subplots(figsize=(10, max(4, len(data) * 0.6)))

    if horizontal:
        bars = ax.barh(labels, data, color=colors[:len(data)])
        if value_labels:
            for bar, val in zip(bars, data):
                ax.text(bar.get_width() + max(data) * 0.01, bar.get_y() + bar.get_height() / 2,
                        f"{val:,.0f}" if isinstance(val, (int, float)) else str(val),
                        va="center", fontsize=9, fontfamily=DEFAULT_FONT)
        ax.invert_yaxis()
    else:
        bars = ax.bar(labels, data, color=colors[:len(data)])
        if value_labels:
            for bar, val in zip(bars, data):
                ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + max(data) * 0.01,
                        f"{val:,.0f}" if isinstance(val, (int, float)) else str(val),
                        ha="center", fontsize=9, fontfamily=DEFAULT_FONT)

    ax.set_title(title, fontsize=13, fontweight="bold", fontfamily=DEFAULT_FONT,
                 color="#003A70", pad=15)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)

    if source:
        fig.text(0.1, 0.02, f"Source: {source}", fontsize=7, fontstyle="italic",
                 color="#808080", fontfamily=DEFAULT_FONT)

    plt.tight_layout()
    os.makedirs(os.path.dirname(filepath) if os.path.dirname(filepath) else ".", exist_ok=True)
    fig.savefig(filepath, dpi=DEFAULT_DPI, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    return filepath


def create_waterfall_chart(values: list[float], labels: list[str], title: str,
                           filepath: str, source: str = "",
                           total_label: str = "Total") -> str | None:
    """
    Create a waterfall / bridge chart (standard for profit bridges, revenue walks).

    Args:
        values: Incremental values (positive or negative). Last value is treated as total.
        labels: Labels for each bar
        title: Chart title (assertion)
        filepath: Output path
        source: Source citation
        total_label: Label for the total bar (default "Total")

    Returns:
        filepath if successful, None if matplotlib unavailable
    """
    plt = _check_matplotlib()
    if plt is None:
        return None

    import numpy as np

    n = len(values)
    cumulative = [0] * n
    cumulative[0] = values[0]
    for i in range(1, n - 1):
        cumulative[i] = cumulative[i - 1] + values[i]
    cumulative[-1] = values[-1]  # Total is absolute, not incremental

    # Calculate bottoms for stacking
    bottoms = [0] * n
    bottoms[0] = 0
    for i in range(1, n - 1):
        bottoms[i] = cumulative[i - 1] if values[i] >= 0 else cumulative[i]
    bottoms[-1] = 0

    # Colors: positive = green, negative = red, total = navy
    colors = []
    for i, v in enumerate(values):
        if i == 0 or i == n - 1:
            colors.append(MBB_COLORS[0])  # Navy for start/total
        elif v >= 0:
            colors.append(MBB_POSITIVE)
        else:
            colors.append(MBB_NEGATIVE)

    fig, ax = plt.subplots(figsize=(10, 5))
    bar_values = [abs(v) if i not in (0, n - 1) else v for i, v in enumerate(values)]
    bars = ax.bar(labels, bar_values, bottom=bottoms, color=colors, width=0.6)

    # Value labels
    for i, (bar, val) in enumerate(zip(bars, values)):
        y_pos = bottoms[i] + bar.get_height() + max(abs(v) for v in values) * 0.02
        prefix = "+" if val > 0 and i not in (0, n - 1) else ""
        ax.text(bar.get_x() + bar.get_width() / 2, y_pos,
                f"{prefix}{val:,.0f}", ha="center", fontsize=9,
                fontfamily=DEFAULT_FONT, fontweight="bold")

    # Connector lines between bars
    for i in range(n - 2):
        ax.plot([i + 0.3, i + 0.7], [cumulative[i], cumulative[i]],
                color="#808080", linewidth=0.8, linestyle="--")

    ax.set_title(title, fontsize=13, fontweight="bold", fontfamily=DEFAULT_FONT,
                 color="#003A70", pad=15)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.axhline(y=0, color="#333333", linewidth=0.5)

    if source:
        fig.text(0.1, 0.02, f"Source: {source}", fontsize=7, fontstyle="italic",
                 color="#808080", fontfamily=DEFAULT_FONT)

    plt.tight_layout()
    os.makedirs(os.path.dirname(filepath) if os.path.dirname(filepath) else ".", exist_ok=True)
    fig.savefig(filepath, dpi=DEFAULT_DPI, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    return filepath


def create_scatter_matrix(x_data: list[float], y_data: list[float],
                          labels: list[str], title: str, filepath: str,
                          x_label: str = "", y_label: str = "",
                          bubble_sizes: list[float] = None,
                          quadrant_labels: list[str] = None,
                          source: str = "") -> str | None:
    """
    Create a scatter / bubble chart (BCG matrix style, positioning charts).

    Args:
        x_data: X-axis values
        y_data: Y-axis values
        labels: Labels for each point
        title: Chart title
        filepath: Output path
        x_label: X-axis label
        y_label: Y-axis label
        bubble_sizes: Optional sizes for bubble chart
        quadrant_labels: Optional labels for 4 quadrants [top-left, top-right, bottom-left, bottom-right]
        source: Source citation

    Returns:
        filepath if successful, None if matplotlib unavailable
    """
    plt = _check_matplotlib()
    if plt is None:
        return None

    fig, ax = plt.subplots(figsize=(10, 8))

    sizes = bubble_sizes or [100] * len(x_data)
    scatter = ax.scatter(x_data, y_data, s=sizes, c=MBB_COLORS[:len(x_data)],
                         alpha=0.7, edgecolors="white", linewidth=1.5)

    # Point labels
    for i, label in enumerate(labels):
        ax.annotate(label, (x_data[i], y_data[i]),
                    textcoords="offset points", xytext=(8, 8),
                    fontsize=9, fontfamily=DEFAULT_FONT)

    # Quadrant lines and labels
    if quadrant_labels:
        x_mid = (max(x_data) + min(x_data)) / 2
        y_mid = (max(y_data) + min(y_data)) / 2
        ax.axvline(x=x_mid, color="#D9D9D9", linewidth=1, linestyle="--")
        ax.axhline(y=y_mid, color="#D9D9D9", linewidth=1, linestyle="--")

        x_range = max(x_data) - min(x_data)
        y_range = max(y_data) - min(y_data)
        positions = [
            (min(x_data) + x_range * 0.15, max(y_data) - y_range * 0.05),  # top-left
            (max(x_data) - x_range * 0.15, max(y_data) - y_range * 0.05),  # top-right
            (min(x_data) + x_range * 0.15, min(y_data) + y_range * 0.05),  # bottom-left
            (max(x_data) - x_range * 0.15, min(y_data) + y_range * 0.05),  # bottom-right
        ]
        for ql, (qx, qy) in zip(quadrant_labels, positions):
            ax.text(qx, qy, ql, fontsize=11, fontfamily=DEFAULT_FONT,
                    fontweight="bold", color="#808080", ha="center",
                    fontstyle="italic", alpha=0.6)

    ax.set_xlabel(x_label, fontsize=11, fontfamily=DEFAULT_FONT)
    ax.set_ylabel(y_label, fontsize=11, fontfamily=DEFAULT_FONT)
    ax.set_title(title, fontsize=13, fontweight="bold", fontfamily=DEFAULT_FONT,
                 color="#003A70", pad=15)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)

    if source:
        fig.text(0.1, 0.02, f"Source: {source}", fontsize=7, fontstyle="italic",
                 color="#808080", fontfamily=DEFAULT_FONT)

    plt.tight_layout()
    os.makedirs(os.path.dirname(filepath) if os.path.dirname(filepath) else ".", exist_ok=True)
    fig.savefig(filepath, dpi=DEFAULT_DPI, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    return filepath


def create_line_chart(x_data: list, y_series: dict[str, list[float]], title: str,
                      filepath: str, x_label: str = "", y_label: str = "",
                      source: str = "") -> str | None:
    """
    Create a line chart for time series data.

    Args:
        x_data: X-axis values (e.g., years, quarters)
        y_series: Dict of {series_name: [values]}
        title: Chart title
        filepath: Output path
        x_label: X-axis label
        y_label: Y-axis label
        source: Source citation

    Returns:
        filepath if successful, None if matplotlib unavailable
    """
    plt = _check_matplotlib()
    if plt is None:
        return None

    fig, ax = plt.subplots(figsize=(10, 5))

    for i, (name, values) in enumerate(y_series.items()):
        color = MBB_COLORS[i % len(MBB_COLORS)]
        ax.plot(x_data, values, marker="o", color=color, linewidth=2,
                markersize=6, label=name)
        # Label last point
        ax.text(x_data[-1], values[-1], f" {values[-1]:,.0f}",
                fontsize=9, fontfamily=DEFAULT_FONT, color=color,
                va="center")

    ax.set_xlabel(x_label, fontsize=11, fontfamily=DEFAULT_FONT)
    ax.set_ylabel(y_label, fontsize=11, fontfamily=DEFAULT_FONT)
    ax.set_title(title, fontsize=13, fontweight="bold", fontfamily=DEFAULT_FONT,
                 color="#003A70", pad=15)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.legend(fontsize=10, frameon=False)

    if source:
        fig.text(0.1, 0.02, f"Source: {source}", fontsize=7, fontstyle="italic",
                 color="#808080", fontfamily=DEFAULT_FONT)

    plt.tight_layout()
    os.makedirs(os.path.dirname(filepath) if os.path.dirname(filepath) else ".", exist_ok=True)
    fig.savefig(filepath, dpi=DEFAULT_DPI, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    return filepath


# ── Native Excel Chart Helpers (openpyxl) ────────────────────────────────────

def add_excel_bar_chart(ws, data_range: str, title: str, position: str = "E2",
                        category_range: str = None, horizontal: bool = False):
    """
    Add a native bar chart to an Excel worksheet.

    Args:
        ws: openpyxl worksheet
        data_range: Cell range for data (e.g., "B1:D10")
        title: Chart title
        position: Cell where chart top-left is anchored
        category_range: Cell range for category labels
        horizontal: If True, horizontal bar chart
    """
    from openpyxl.chart import BarChart, Reference

    chart = BarChart()
    chart.type = "bar" if horizontal else "col"
    chart.title = title
    chart.style = 10
    chart.y_axis.title = ""
    chart.x_axis.title = ""

    data = Reference(ws, range_string=data_range)
    chart.add_data(data, titles_from_data=True)

    if category_range:
        cats = Reference(ws, range_string=category_range)
        chart.set_categories(cats)

    chart.shape = 4
    chart.width = 18
    chart.height = 10
    ws.add_chart(chart, position)


def add_excel_line_chart(ws, data_range: str, title: str, position: str = "E2",
                         category_range: str = None):
    """
    Add a native line chart to an Excel worksheet.

    Args:
        ws: openpyxl worksheet
        data_range: Cell range for data
        title: Chart title
        position: Cell where chart top-left is anchored
        category_range: Cell range for category labels
    """
    from openpyxl.chart import LineChart, Reference

    chart = LineChart()
    chart.title = title
    chart.style = 10

    data = Reference(ws, range_string=data_range)
    chart.add_data(data, titles_from_data=True)

    if category_range:
        cats = Reference(ws, range_string=category_range)
        chart.set_categories(cats)

    chart.width = 18
    chart.height = 10
    ws.add_chart(chart, position)
