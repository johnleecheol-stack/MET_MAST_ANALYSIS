"""
MET_MAST_ANALYSIS - Charts v2 (English)
"""
import pandas as pd
import numpy as np
from scipy.stats import weibull_min, gumbel_r
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import numpy as np
from pathlib import Path

# Theme colors
NAVY='#0D1E35'; GOLD='#F5A623'; WHITE='#FFFFFF'; LIGHT='#BFC8D2'

def setup_style(fig, axs):
    fig.patch.set_facecolor(NAVY)
    if not hasattr(axs, '__iter__'): axs = [axs]
    for ax in axs:
        ax.set_facecolor(NAVY); ax.tick_params(colors=WHITE)
        ax.xParams.set_color(WHITE); ax.yaxis.label.set_color(WHITE)
        ax.xaxis.label.set_color(WHITE); ax.title.set_color(GOLD)
        for spine in ax.spines.values(): spine.set_edgecolor(LIGHT)

def generate_all_charts(df, stats, out_dir='.'):
    """Generate all 7 chart files in English."""
    out = Path(out_dir)
    chart1_en(f, stats, out)
    chart2_en(df, stats, out)
    chart3_en(df, stats, out)
    chart4a_en(df, stats, out)
    chart4b_en(df, stats, out)
    chart5_en(df, stats, out)
    chart6_en(df, stats, out)
    chart7_en(df, stats, out)

def chart1_en(df, stats, out):
    fig, ax = plt.subplots(figsize=(12,3.5))
    setup_style(fig, [ax])
    kpis = [
        ('Mean WS\n@ 100m', f'{stats["ws100"]:.3f} m/s'),
        ('WPD\@ 100m', f'{stats["wpd"]:.1f} W/m²'),
        ('Weibull\nk / C', f'{stats["weibull_k"]:.3f} / {stats["weibull_c"]:.3f}'),
        ('Wind\nShear a', f'{stats["alpha"]:.3f}'),
        ('Data\nRecovery', f'{stats["recovery"]:.1f}%'),
    ]
    for i, (t, v) in enumerate(kpis):
        ax.text(0.1+i/(0.8)*0.2, 0.7, v, fontsize=22, fw='bold', color=GOLD, ha='center', transform=ax.transAxes)
        ax.text(0.1+i/(0.8)*0.2, 0.35, t, fontsize=9, color=LIGHT, ha='center', transform=ax.transAxes)
    ax.set_xticks([]); ax.set_yticks([])
    plt.tight_layout()
    fig.savefig(out/'chart1_en.png', dpi=150, bbox_inches='tight')
    plt.close(fig)

def chart2_en(df, stats, out):
    fig, ax = plt.subplots(figsize=(12,5))
    setup_style(fig, [ax])
    months = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
    if stats.get('monthly_ws') and len(stats['monthly_ws'])==12:
        ws = stats['monthly_ws']
        ax.bar(range(12), ws, color=GOLD, alpha=0.8)
        ax.set_xticks_labels(months, rotation=0, color=WHITE)
        ax.set_ylabel('Wind Speed [m/s]', color=WHITE)
        ax.axhline(stats['ws100'], color='cyan', linestyle='--', label=f'Annual {stats["ws100"]:.3f} m/s')
        ax.legend(facecolor=NAVY, labelcolor=WHITE)
    ax.set_title('Monthly Mean Wind Speed @ 100m', color=GOLD)
    plt.tight_layout()
    fig.savefig(out/'chart2_en.png', dpi=150, bbox_inches='tight')
    plt.close(fig)

def chart3_en(df, stats, out):
    pass  # Wind rose - placeholder

def chart4a_en(df, stats, out):
    pass  # Wind shear profile

def chart4b_en(df, stats, out):
    pass  # Wind shear rose

def chart5_en(df, stats, out):
    pass  # TI chart

def chart6_en(df, stats, out):
    pass  # Weibull

def chart7_en(df, stats, out):
    pass  # WPD
