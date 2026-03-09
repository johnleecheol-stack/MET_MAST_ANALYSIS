#!/usr/bin/env python3
"""
MET_MAST_ANALYSIS — run.py
CLI entry point.
"""
import argparse
import subprocess
import sys
from pathlib import Path

def main():
    p = argparse.ArgumentParser(description="MET_MAST_ANALYSIS WRA Pipeline")
    p.add_argument("--data", required=True)
    p.add_argument("--config", default="config/project_config.yaml")
    p.add_argument("--out", default="output")
    p.add_argument("--lang", default="both", choices=["en", "ko", "both"])
    p.add_argument("--skip-pptx", action="store_true")
    p.add_argument("--skip-html", action="store_true")
    args = p.parse_args()
    print(f"Running WRA pipeline for {args.data}")

if __name__ == "__main__":
    main()
