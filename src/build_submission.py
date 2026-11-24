"""Submission pack: every Word document and deck exported to PDF in the brief's order, plus one combined PDF.
Needs LibreOffice (soffice) and poppler (pdfunite). Spreadsheets and data files stay in their phase folders."""
import shutil
import subprocess
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "06_Submission"

ORDER = [
    ("00_Team_and_Project_Management/Team_Charter_and_Workplan.docx", "01_Team_Charter_and_Workplan"),
    ("01_Diagnose/Opportunity_Scan.docx", "02_Phase1_Opportunity_Scan"),
    ("02_Build/P1_Marketing/Marketing_Pack.docx", "03_P1_Marketing_Pack"),
    ("02_Build/P2_Customer_Support_Agent/Agent_Build_Pack.docx", "04_P2_Support_Agent_Build_Pack"),
    ("02_Build/P3_Finance/Peer_Annual_Report_Summary.docx", "05_P3_Finance_Peer_Summary_and_Budget"),
    ("02_Build/P4_Operations/Operations_Pack.docx", "06_P4_Operations_Forecast_and_Decision"),
    ("02_Build/P5_Sales/B2B_Proposal_Deck_Nordlicht.pptx", "07_P5_B2B_Proposal_Deck"),
    ("02_Build/P5_Sales/Sales_Prompt_Log.docx", "08_P5_Sales_Prompt_Log"),
    ("03_Stress_Test/Shock_E_Response_Competitor_Move.docx", "09_Phase3_Shock_E_Response"),
    ("04_Decide/Governance_Pack.docx", "10_Phase4_Governance_Pack"),
    ("04_Decide/Change_Plan_ADKAR.docx", "11_Phase4_Change_Plan_ADKAR"),
    ("04_Decide/Roadmap_18_Month.docx", "12_Phase4_Roadmap_18_Month"),
    ("05_Defend/Board_Pitch_Deck.pptx", "13_Phase5_Board_Pitch_Deck"),
    ("05_Defend/Panel_QA_Demo_and_Run_of_Show.docx", "14_Phase5_Panel_QA_and_Demo"),
    ("05_Defend/Individual_Viva_Prep.docx", "15_Phase5_Individual_Viva_Prep"),
]


def soffice():
    for c in ("soffice", "libreoffice"):
        if shutil.which(c):
            return c
    raise SystemExit("LibreOffice not found")


def build():
    OUT.mkdir(exist_ok=True)
    pdfs = []
    with tempfile.TemporaryDirectory() as tmp:
        for src, name in ORDER:
            s = ROOT / src
            subprocess.run([soffice(), "--headless", "--convert-to", "pdf", "--outdir", tmp, str(s)],
                           check=True, capture_output=True, env={"HOME": tmp, "PATH": "/usr/bin:/bin"})
            dst = OUT / f"{name}.pdf"
            shutil.move(str(Path(tmp) / (s.stem + ".pdf")), dst)
            pdfs.append(dst)
            print(dst.name)
    combined = OUT / "00_WeaveWorks_Group4_Complete_Submission.pdf"
    subprocess.run(["pdfunite", *map(str, pdfs), str(combined)], check=True)
    print(combined.name)
    return pdfs, combined


if __name__ == "__main__":
    build()
