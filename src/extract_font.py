"""
MET_MAST_ANALYSIS — src/extract_font.py
Extract Noto Sans CJK KR TTF subset.
Run once: python src/extract_font.py
"""
import sys
from pathlib import Path

FONTS_DIR = Path(__file__).parent.parent / 'fonts'

def extract_noto_cjk_kr():
    FONTS_DIR.mkdir(exist_ok=True)
    print('Font extraction requires NotoSansCJK installed')

if __name__ == '__main__':
    extract_noto_cjk_kr()
