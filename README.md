# AI Transformation Consultancy Challenge: Group 4

**Client:** WeaveWorks Exports, a fictional textile/knitwear exporter in Tiruppur with ₹120 crore revenue.
Its built-in tensions: EU buyers demand compliance, and the workforce has low digital skills.
**Brief:** ₹50 lakh and 18 months to adopt Gen AI.
**Shock card E:** the main rival launches AI personalisation and cuts prices 10%.
**Team:** 4 members. The CEO role (roadmap and pitch) is shared.

## Our recommendation in one paragraph

Approve **₹47.61 lakh** for **10 AI use cases** in three stage-gated horizons (Jan 2027 – Jun 2028). That figure includes
₹11.0 lakh of hidden costs (training, human oversight, security/legal) and ₹4.5 lakh contingency, and leaves
₹2.39 lakh in reserve until Stage-gate 1. The base-case 18-month programme ROI is **+125%**
(pessimistic −12%, optimistic +315%), and the programme breaks even if we deliver **45%** of base-case benefits.

Our answer to the rival's 10% cut is **value, not price**. Matching it for every buyer would cost ~₹12 crore a year,
which is more than our ~₹10.8 crore EBITDA (assumed). Instead we offer value-engineered quotes to at-risk accounts and
our own personalisation portal with guardrails ("WeaveWorks Studio"). Compliance and traceability are the moat.

## Team roles

Names are set once in `src/facts.py` (`TEAM`). To swap roles, swap the names there and rebuild (see the last section).

| Member | Role | Owns |
|---|---|---|
| Rudransh Bajpai | CMO + Sales lead | P1 Marketing, P5 B2B proposal deck, Shock E response |
| Het Laheri | CCO + Chief AI & Risk Officer | P2 Support Agent, Governance pack, ADKAR plan |
| Shrey Doda | CFO | P3 Peer annual-report summary + budget, ROI model |
| Anushka Baniya | COO | P4 Forecast + DAX + decision matrix, roadmap integration |

## Where everything is (brief → file)

| Phase | Deliverable | File(s) |
|---|---|---|
| — | Team charter, RACI, 4-week plan, stand-up log, AI-use register, submission checklist | `00_Team_and_Project_Management/Team_Charter_and_Workplan.docx` |
| 1 Diagnose | Value chain, 17 use cases, value × feasibility 2x2, 5 picked | `01_Diagnose/Opportunity_Scan.docx` |
| 2 Build: P1 | Synthetic dataset (300 buyers), 3 segments (k-means), 2 personas, campaign, 3 creatives, 3-email sequence, ads for LinkedIn/Google/Meta | `02_Build/P1_Marketing/` |
| 2 Build: P2 | No-code Agent: instructions, 9 knowledge files, escalation rules, policy limits, 25-test suite; **backup live demo** ([Weave Buyer Support](https://claude.ai/artifact/CNcLbDjCa2Smmh5xujpYaD), built from the same files by `src/build_agent_page.py`) | `02_Build/P2_Customer_Support_Agent/` |
| 2 Build: P3 | AI summary of **S.P. Apparels Ltd** FY2024-25 annual report + verification table; 3-scenario budget | `02_Build/P3_Finance/` |
| 2 Build: P4 | Demand forecast (Excel formulas + Power BI star schema + DAX variance measure); supply-disruption decision matrix | `02_Build/P4_Operations/` |
| 2 Build: P5 | B2B proposal deck for the largest segment | `02_Build/P5_Sales/` |
| 3 Stress-test | Shock E response; roadmap and budget changes | `03_Stress_Test/Shock_E_Response_Competitor_Move.docx` |
| 4 Decide | ROI model (every use case, 3 scenarios, break-even driver, payback) | `04_Decide/ROI_Model.xlsx` |
| 4 Decide | AI acceptable-use policy, risk register, approval workflow, DPDP Act + sector rules | `04_Decide/Governance_Pack.docx`, `04_Decide/Risk_Register.xlsx` |
| 4 Decide | ADKAR change plan (merchandising team) | `04_Decide/Change_Plan_ADKAR.docx` |
| 4 Decide | 18-month roadmap, 3 horizons, ₹50 lakh incl. hidden costs | `04_Decide/Roadmap_18_Month.docx` |
| 5 Defend | Board pitch (13 slides + appendix, speaker notes with timings) | `05_Defend/Board_Pitch_Deck.pptx` |
| 5 Defend | Run-of-show, live demo script, panel Q&A (sceptical CFO, regulator, guest), investment-committee plan | `05_Defend/Panel_QA_Demo_and_Run_of_Show.docx` |
| 5 Defend | Investment-committee scorecard (allocates ₹50 lakh play money) | `05_Defend/Investment_Committee_Scorecard.xlsx` |
| 5 Defend | Individual viva prep: 60-second pitch, 5 numbers, rejected AI output, likely questions | `05_Defend/Individual_Viva_Prep.docx` |

Every prototype document has a **"Prompt log"** section (prompt → raw AI output → human-edited version → why) and a
highlighted **"Rejected AI output"** box for the viva.

## What the team still has to do (cannot be done from this repository)

1. ~~Build the Agent~~ **Done.** Custom GPT "Weave - WeaveWorks Buyer Support" built in ChatGPT; all 25 tests run: 24 pass (96%), all 7 critical pass, go-live gate **GO** (`Agent_Test_Suite.xlsx`). T06 failed on a knowledge-base gap, now fixed in `knowledge_base/03_Sampling_Policy.md`: **re-upload that one file to the GPT and re-run T06**, then review and countersign the test sheet (the Tester column currently says Claude).
2. ~~Canva creatives~~ **Done.** The three creatives are editable Canva designs in the folder [WeaveWorks Exports - Group 4 - P1 Marketing Creatives](https://www.canva.com/folder/FAHbuaqQHOs) (imported from `P1_Marketing/creatives/canva_import/*.pdf`, built by `src/build_canva_creatives.js`, then import defects fixed in Canva). Rudransh should open each once so he can explain how they were made in the viva.
3. **Verify the peer figures.** Open S.P. Apparels' official Annual Report 2024-25 PDF and fill the "Page ref" column in
   `Finance_Budget_and_Peer.xlsx > Verification`. Our figures were checked against filings, investor materials and
   earnings-call summaries, but not yet against the AR PDF itself.
4. **Build the Power BI page** from `P4_Operations/powerbi_data/*.csv` using `DAX_Measures.md`. A .pbix cannot be generated here.
5. **Record backup videos** of both live demos the day before the pitch.
6. **Prompt logs:** the raw outputs were generated with Claude when this pack was drafted. If your instructor wants logs
   from your own tools, re-run the prompts (written out in full) and keep the edit rationale.
7. Fill in the weekly stand-up log (Team Charter §4).

## Ground rules: how we comply

- **No real personal or customer data.** All buyer, order and knowledge-base data is synthetic (`src/gen_data.py`, `src/build_p2_kb.py`).
  The prospect "Nordlicht Basics GmbH", the personas and the rival "KnitSphere Apparels" are fictional.
- **Every statistic is cited or labelled "assumed".** Everything about WeaveWorks is assumed; real-world facts (DPDP Act,
  EU regulations, S.P. Apparels financials) carry sources in each document.
- **Every AI output has a human owner.** The owner is named on every document cover and in the RACI.

## Rebuilding everything

All numbers come from one file, `src/facts.py`. Change an assumption there and rebuild:

```bash
pip install pandas numpy scikit-learn matplotlib openpyxl python-docx
cd src && npm install && cd ..
python src/build_all.py
```

The scripts are `gen_data.py` (synthetic data) → `analysis.py` (segmentation, forecast, charts) → one `build_*.py` per deliverable
→ `build_*_deck.js` for the two PowerPoint decks. The Excel forecast and ROI models use live formulas, and their results
reconcile exactly with the Python versions.
