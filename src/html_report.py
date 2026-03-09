"""
MET_MAST_ANALYSIS — src/html_report.py
Build a self-contained interactive HTML wind resource report.

Usage:
    from src.html_report import build_html_report
    build_html_report(stats_dict, charts_dir, output_path, lang='en')
"""
import base64
import json
from pathlib import Path
from datetime import datetime

def _b64(path):
    with open(path, "rb") as f:
        return base64.b64encode(f.read()).decode()

LABELS = {
    "en": {
        "title": "Wind Resource Assessment",
        "site": "Site",
        "period": "Measurement Period",
        "key_stats": "Key Statistics",
        "mean_ws": "Mean Wind Speed @ 100m",
        "wpd": "Wind Power Density",
        "weibull": "Weibull Parameters",
        "shear": "Wind Shear Exponent",
        "ti": "Turbulence Intensity P90@t15m/s",
        "v50": "Extreme Wind V50",
        "iec_title": "IEC 61400-1 Ed.4 Assessment",
    },
    "ko": {
        "title": "훌스과색리",
        "site": "구점",
        "period": "강제 권",
        "key_stats": "예스 식적",
        "mean_ws": "대칅앨 �	다 @ 100m",
        "wpd": "회아배하대",
        "weibull": "팜은: k, C",
        "shear": "스쀰 趵뱪",
        "ti": "핵 ꉴ잤 P90@15m/s",
        "v50": "그레외브 V50",
        "iec_title": "IEC 61400-1 Ed.4 �ʔ따듈",
    },
}


def build_html_report(stats, charts_dir, output_path, lang="en"):
    lbl = LABELS.get(lang, LABELS["en"])
    charts = {}
    cp = Path(charts_dir)
    for i in range(1,8):
        for su in ["","a","b"]:
            fn=cp/f"chart{i}{su}_{lang}.png"
            if fn.exists(): charts[f"chart{i}{su}"]=_b64(fn)
    def bg(o): return '<span style="color:#2EF26F">✅ PASS</span>' if o else '<span style="color:#FF4444">❌ FAIL</span>'
    def im(k): return f'<img src="data:image/png;base64,{charts[lk]}" style="width:100%">' if k in charts else ''
    html = f"""<!DOCTYPE html><html><head><meta charset="utf-8">
<title>{stats.get('site_name','WRA')} {lbl['title']}</title>
<style>body{background:#0D1E35;color:#FFF;font-family:sans-serif}
h1{color:#F5A623}.kbp{background:#162B45;padding:15px;margin:5px;border-radius:8px}
</style></head><body>
<h1>{stats.get('site_name','WRA')} — {lbl['title']}</h1>
<div style="display:flex;flex-wrap:wrap;gap:10px"'>
<div class="kpi"><b>{lbl['mean_ws']}</b><p style="font-size:2em;color:#F5A623">{stats.get('ws100',0):.3f} m/s</p></div>
<div class="kpi"><b>{lbl['wpd']}</b><p style="font-size:2em;color:#F5A623">{stats.get('wpd',0):.1f} W/m²</p></div>
</div>
<div style="display:grid;grid-template-columns:repeat(2,1fr);gap:20px">
{im('chart1')}{im('chart2')}{im('chart3')}{im('chart4a')}{im('chart4b')}{im('chart5')}{im('chart6')}{im('chart7')}
</div></body></html>"""
    Path(output_path).write_text(html, encoding="utf-8")
    return output_path
