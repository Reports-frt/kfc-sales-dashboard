"""
KFC Food Cost — Full Pipeline Runner
=====================================

Reads source files from _work/, builds food_data.json, pushes to GitHub.
Designed for monthly manual execution after copying files into _work/.

Usage:
    python build_pipeline.py
    python build_pipeline.py --no-push   (skip git push, build only)
"""

import argparse
import json
import os
import subprocess
import sys
import time
from pathlib import Path
from datetime import datetime, timezone

# Import the parser
sys.path.insert(0, str(Path(__file__).parent))
from build_food_data import build_food_data_json


# =====================================================================
# CONFIG
# =====================================================================
REPO_ROOT = Path(r"C:\Users\IT\Documents\GitHub\kfc-sales-dashboard")
WORK_DIR = REPO_ROOT / "_work"
OUTPUT_DIR = REPO_ROOT / "food"   # /food/ subfolder at repo root (GitHub Pages serves from root)

REQUIRED_FILES = [
    "FoodCost.xlsx",
    "CategoriesFC.xlsx",
]

OPTIONAL_FILES = [
    "2026_ΚΟΣΤΟΛΟΓΗΣΗ.xlsx",   # for cross-validation if available
]

GIT_PUSH_RETRIES = 3
GIT_PUSH_DELAY_SEC = 60


# =====================================================================
# HELPERS
# =====================================================================
def log(msg, level="INFO"):
    timestamp = datetime.now().strftime("%H:%M:%S")
    print(f"[{timestamp}] [{level}] {msg}")


def check_required_files(work_dir):
    """Verify all required files are present in _work/."""
    log(f"Checking required files in {work_dir}...")
    missing = []
    for fname in REQUIRED_FILES:
        path = work_dir / fname
        if not path.exists():
            missing.append(fname)
        else:
            size_mb = path.stat().st_size / 1024 / 1024
            mtime = datetime.fromtimestamp(path.stat().st_mtime).strftime('%Y-%m-%d %H:%M')
            log(f"  ✓ {fname} ({size_mb:.1f} MB, modified {mtime})")
    
    if missing:
        log("MISSING REQUIRED FILES:", "ERROR")
        for m in missing:
            log(f"  ✗ {m}", "ERROR")
        log(f"Place these files in: {work_dir}", "ERROR")
        return False
    return True


def find_git():
    """Locate git.exe — try PATH first, then common Windows install paths."""
    import shutil
    git = shutil.which('git')
    if git:
        return git
    # Common Windows install locations
    candidates = [
        r'C:\Program Files\Git\cmd\git.exe',
        r'C:\Program Files\Git\bin\git.exe',
        r'C:\Program Files (x86)\Git\cmd\git.exe',
        r'C:\Program Files (x86)\Git\bin\git.exe',
    ]
    for path in candidates:
        if Path(path).exists():
            return path
    return None


_GIT_EXE = None

def run_git(args, cwd):
    """Run git command. Return (returncode, stdout+stderr)."""
    global _GIT_EXE
    if _GIT_EXE is None:
        _GIT_EXE = find_git()
        if _GIT_EXE is None:
            return 1, "ERROR: git.exe not found. Install Git for Windows or add it to PATH."
        log(f"  Using git: {_GIT_EXE}")
    
    result = subprocess.run(
        [_GIT_EXE] + args,
        cwd=cwd,
        capture_output=True,
        text=True,
        encoding='utf-8',
        errors='replace',
    )
    return result.returncode, (result.stdout + result.stderr).strip()


def git_push(repo_root):
    """Stage, commit, push food_data.json with retry."""
    log("Pushing to GitHub...")
    
    # Stage
    rel_path = "food/food_data.json"
    rc, out = run_git(['add', rel_path], cwd=repo_root)
    if rc != 0:
        log(f"git add failed: {out}", "ERROR")
        return False
    
    # Commit
    msg = f"Update food_data.json — {datetime.now().strftime('%Y-%m-%d %H:%M')}"
    rc, out = run_git(['commit', '-m', msg], cwd=repo_root)
    if rc != 0 and 'nothing to commit' in out.lower():
        log("No changes to commit (food_data.json identical)")
        return True
    elif rc != 0:
        log(f"git commit failed: {out}", "ERROR")
        return False
    log(f"  Committed: {msg}")
    
    # Push with retry
    for attempt in range(1, GIT_PUSH_RETRIES + 1):
        log(f"  Push attempt {attempt}/{GIT_PUSH_RETRIES}...")
        rc, out = run_git(['push'], cwd=repo_root)
        if rc == 0:
            log("  ✓ Push successful")
            return True
        log(f"  Push failed: {out}", "WARN")
        if attempt < GIT_PUSH_RETRIES:
            log(f"  Retrying in {GIT_PUSH_DELAY_SEC}s...")
            time.sleep(GIT_PUSH_DELAY_SEC)
    
    log(f"All push attempts failed", "ERROR")
    return False


# =====================================================================
# MAIN
# =====================================================================
def main():
    parser = argparse.ArgumentParser(description='Build & push food_data.json')
    parser.add_argument('--no-push', action='store_true',
                        help='Build only, skip git push')
    parser.add_argument('--work-dir', default=str(WORK_DIR),
                        help='Source files folder (default: _work/)')
    parser.add_argument('--output-dir', default=str(OUTPUT_DIR),
                        help='Output folder (default: kfc-dashboard/food/)')
    args = parser.parse_args()
    
    work_dir = Path(args.work_dir)
    output_dir = Path(args.output_dir)
    
    print("=" * 60)
    print("KFC Food Cost Hub — Pipeline")
    print(f"Started: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 60)
    
    # Step 1: Verify required files
    log("STEP 1: Check required files")
    if not check_required_files(work_dir):
        sys.exit(1)
    
    # Step 2: Build food_data.json
    log("\nSTEP 2: Build food_data.json")
    try:
        result = build_food_data_json(work_dir, output_dir)
    except Exception as e:
        log(f"Build failed: {e}", "ERROR")
        import traceback
        traceback.print_exc()
        sys.exit(1)
    
    output_path = output_dir / 'food_data.json'
    size_kb = output_path.stat().st_size / 1024
    log(f"\n✓ Build complete: {size_kb:.0f} KB")
    log(f"  Periods covered: {len(result['meta']['periods'])}")
    log(f"  Latest period: {result['meta']['latest_period']}")
    
    # Step 3: Git push (unless --no-push)
    if args.no_push:
        log("\nSTEP 3: Git push SKIPPED (--no-push flag)")
        log("To push manually: git add ... && git commit && git push")
    else:
        log("\nSTEP 3: Git push")
        if not git_push(REPO_ROOT):
            log("Push failed. Build is complete but not deployed.", "WARN")
            log("You can retry manually: git push", "WARN")
            sys.exit(2)
    
    print()
    print("=" * 60)
    print(f"✓ DONE — {datetime.now().strftime('%H:%M:%S')}")
    print("=" * 60)


if __name__ == '__main__':
    main()
