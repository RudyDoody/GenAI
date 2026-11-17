"""Rebuild every deliverable from src/facts.py.

    python src/build_all.py

Needs: Python 3.10+ (pandas, numpy, scikit-learn, matplotlib, openpyxl, python-docx), Node 18+ with
`npm install` run in src/ (pptxgenjs, react-icons, react, react-dom, sharp), and Playwright for the creatives.
Optional: set XLSX_RECALC to a LibreOffice-based recalc script to cache formula values (Excel recalculates on open anyway),
and PPTX_SKILL to a folder containing scripts/apply_theme.js to write theme colours into the decks.
"""
import os
import subprocess
import sys
from pathlib import Path

SRC = Path(__file__).resolve().parent
ROOT = SRC.parent
PY = [sys.executable]

STEPS = ["gen_data.py", "analysis.py", "export_facts.py", "build_team.py", "build_phase1.py", "creatives.py", "build_p1.py",
         "build_p2.py", "build_p3.py", "build_p4.py", "build_phase3.py", "build_roi.py", "build_governance.py",
         "build_change_roadmap.py", "build_phase5.py"]
DECKS = ["build_p5_sales_deck.js", "build_board_deck.js"]


def run(cmd, env=None):
    print(">", " ".join(map(str, cmd)))
    subprocess.run(cmd, cwd=SRC, check=True, env=env)


if __name__ == "__main__":
    for s in STEPS:
        run(PY + [s])
    env = {**os.environ, "NODE_PATH": str(SRC / "node_modules")}
    if env.get("PPTX_SKILL"):
        for js in DECKS:
            run(["node", js], env=env)
    else:
        print("PPTX_SKILL not set: decks not rebuilt (committed .pptx files are current).")
    rc = os.environ.get("XLSX_RECALC")
    if rc:
        for x in ROOT.glob("0*/**/*.xlsx"):
            run(PY + [rc, str(x), "60"])
    print("Done.")
