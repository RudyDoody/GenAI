"""Phase 4 - Governance pack: AI acceptable-use policy, use-case risk register (docx + xlsx), approval workflow,
DPDP Act obligations and sector-specific rules."""
from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import PatternFill

import facts
from docx_lib import Doc, FIG, ROOT
from xlsx_lib import INPUT_FILL, cell, header, note, title, widths

OUT = ROOT / "04_Decide"

RISKS = [
    # id, use case, risk, L, I, controls, resL, resI, owner, tier
    ("R01", "UC04 Agent", "Agent promises an out-of-policy refund, credit or delivery date to a buyer", 3, 4,
     "Hard limits in instructions; 7 critical tests must pass; weekly transcript sampling; escalation to humans; zero money authority", 1, 4, "CCO", 3),
    ("R02", "UC04 Agent", "Agent states wrong order status (hallucination or stale data)", 3, 3,
     "Answers only from daily ERP export with 'last updated' date; 'not found' rule; test T03/T05", 1, 3, "Merchandising Mgr", 2),
    ("R03", "UC04/UC17", "Prompt injection or data leakage (other buyers' data, instructions)", 2, 4,
     "Per-buyer data scoping; injection tests T10/T16/T23; no internal files in knowledge base", 1, 4, "CARO", 3),
    ("R04", "All", "Employee pastes buyer or personal data into a free/unapproved AI tool", 4, 4,
     "AUP + data classes; approved tools with training opt-out; browser block list; AI literacy for all; 1-hour incident reporting", 2, 4, "CARO", 0),
    ("R05", "UC03 Marketing", "Emails to EU contacts without valid consent (GDPR) or misleading claims", 2, 3,
     "Opt-in only; unsubscribe; claims checklist (no 'guaranteed compliant'); CMO sign-off", 1, 3, "CMO", 2),
    ("R06", "UC03/UC02/UC17", "AI-generated imagery or designs infringe IP; buyer designs used to train shared models", 2, 4,
     "Tools with commercial-use terms; no buyer designs in tools that train; written consent; AI imagery labelled", 1, 4, "Design Head", 3),
    ("R07", "UC09 Compliance", "Assistant drafts an incorrect due-diligence or DPP statement sent to a buyer", 3, 5,
     "Source-linked answers only; Compliance Manager signs every outgoing answer; quarterly sample audit", 1, 5, "Compliance Mgr", 3),
    ("R08", "UC15 Costing", "Costing assistant under-prices (stale yarn price, wrong efficiency)", 3, 4,
     "Daily yarn-price feed; margin floor check; merchandiser approval; CFO approval below floor", 1, 4, "CFO", 2),
    ("R09", "UC06 Forecast", "Over-reliance on forecast leads to wrong capacity booking", 2, 3,
     "Forecast is advisory; monthly S&OP review; MAPE tracked; buyer confirmations override", 1, 3, "COO", 1),
    ("R10", "UC13 HR", "Screening AI discriminates (gender, age, region, caste or disability proxies)", 3, 4,
     "No auto-reject; names/photos/addresses removed; planted-pair bias test before use; HR + CARO approval; candidate notice", 1, 4, "HR Head", 3),
    ("R11", "UC11 Voice SOP", "Voice assistant gives an unsafe machine instruction to an operator", 2, 5,
     "Closed knowledge base of approved SOPs only; safety officer signs content; 'stop and call supervisor' default", 1, 5, "Safety Officer", 3),
    ("R12", "UC17 Portal", "Buyer staff personal data processed abroad without safeguards", 2, 3,
     "Data-processing agreement; EU hosting option; privacy notice; DPDP/GDPR rights process", 1, 3, "CARO", 2),
    ("R13", "All", "Vendor lock-in, price rise or outage of a SaaS AI tool", 3, 2,
     "Exportable data; two-vendor shortlist; contract price caps; manual fallback procedure", 2, 2, "CFO", 1),
    ("R14", "All", "Staff resistance or fear of job loss slows adoption", 3, 3,
     "ADKAR plan; no AI-driven job losses for 18 months; champions; time-saved reinvested", 2, 3, "CEO (shared)", 1),
    ("R15", "All", "Model drift or vendor model change degrades output unnoticed", 3, 3,
     "Quarterly re-run of test suites; KPI dashboards; change log for vendor updates", 1, 3, "CARO", 1),
    ("R16", "All", "AI tool account compromise / cyber incident", 2, 4,
     "SSO + MFA; least privilege; logs kept 180 days; CERT-In 6-hour reporting playbook", 1, 4, "IT Head", 2),
    ("R17", "UC03/UC17", "'AI-washing': overstated AI or sustainability claims in marketing", 2, 3,
     "Claims checklist; evidence file per claim; legal review of campaign", 1, 3, "CMO", 2),
]

DATA_CLASSES = [
    ("Public", "Website text, published catalogues, public reports", "Any tool, including free tools", "None"),
    ("Internal", "SOPs, process notes, anonymised production data", "Approved company AI tools only", "Remove names"),
    ("Confidential", "Buyer designs/tech-packs, prices, costings, contracts, order data", "Approved tools with no-training contract; never free tools", "Need-to-know; buyer consent for design use"),
    ("Restricted", "Personal data (employees, candidates, buyer staff), Aadhaar, bank, health, union membership", "Only systems approved for that purpose by the AI Board; never general chat tools", "DPDP/GDPR basis; minimisation; logging"),
]

DPDP = [
    ("Lawful basis: consent with notice, or a 'legitimate use' (s.4-7)", "Buyer contacts: consent via opt-in or contract performance; employees: legitimate use for employment (s.7) with notice",
     "UC03, UC04, UC13, UC17"),
    ("Notice in clear language (s.5)", "Privacy & AI notice in the Agent and portal; candidate notice for UC13 in English and Tamil", "UC04, UC13, UC17"),
    ("Purpose limitation & minimisation (s.6, 8)", "Data classes table; AI tools get only fields needed; no Restricted data in general tools", "All"),
    ("Accuracy when data drives decisions (s.8(3))", "Human review of any AI output that affects a person (hiring, claims)", "UC13, UC04"),
    ("Reasonable security safeguards (s.8(5))", "SSO, MFA, encryption, logging; vendor security review", "All"),
    ("Breach intimation to the Data Protection Board and each affected person (s.8(6); DPDP Rules 2025)", "Incident playbook: contain, assess, notify; aligns with CERT-In 6-hour rule for cyber incidents", "All"),
    ("Erasure when purpose is served (s.8(7))", "Chat logs 90 days; candidate data 6 months unless hired", "UC04, UC13"),
    ("Processors under contract (s.8(2))", "Data-processing agreements with every AI vendor; no-training clause", "All"),
    ("Grievance redressal and contact person (s.8(9)-(10))", "privacy@ contact; CARO named as data-protection contact", "All"),
    ("Children's data (s.9)", "Not processed: B2B only. Kidswear marketing targets buyers, not children", "n/a"),
    ("Cross-border transfer (s.16)", "Permitted unless a country is restricted by notification; keep vendor-location register", "UC04, UC17"),
    ("Penalties (Schedule): up to Rs 250 crore for security-safeguard failures", "Board-level risk; included in R04/R16", "-"),
]

SECTOR = [
    ("EU GDPR (Reg. 2016/679)", "Applies to EU buyer staff's personal data when we target or monitor them (marketing, chatbot)", "Consent, unsubscribe, rights process, EU-standard DPAs"),
    ("EU AI Act (Reg. 2024/1689)", "Art. 50: tell people they are talking to AI and label AI-generated content (from Aug 2026); Art. 4 AI literacy as good practice; employment AI is 'high-risk' in the EU, our benchmark for UC13",
     "AI disclosure line in Agent; image labels; staff AI literacy; UC13 Tier 3"),
    ("CSDDD (Dir. 2024/1760, as amended 2025)", "Large EU buyers cascade due-diligence questions to suppliers", "UC09 answers signed by Compliance Manager; no 'compliant' claims"),
    ("ESPR / Digital Product Passport (Reg. 2024/1781)", "Textile DPP data requirements expected via delegated act", "Accurate traceability data; UC09 data model"),
    ("EU Forced Labour Regulation (Reg. 2024/3015)", "Products made with forced labour banned from Dec 2027; buyers need supply-chain evidence", "Spinner and job-work traceability; worker-voice records"),
    ("REACH Annex XVII; ZDHC MRSL", "Chemical restrictions (e.g. azo dyes)", "AI never alters lab test data; test reports source-linked"),
    ("GOTS / OEKO-TEX / amfori BSCI", "Certification and audit integrity", "AI must not draft or alter certificates or transaction certificates"),
    ("CERT-In Directions 2022 (IT Act s.70B)", "Report specified cyber incidents within 6 hours; keep logs 180 days", "Incident playbook; log retention"),
    ("Labour codes (Code on Wages 2019; Industrial Relations Code 2020)", "No discrimination in recruitment/wages; notice before changes to conditions of service from new technology",
     "UC13 bias audit; consult worker representatives before UC11 rollout"),
]


def risk_xlsx():
    wb = Workbook()
    ws = wb.active
    ws.title = "Risk_Register"
    title(ws, "WeaveWorks AI use-case risk register", "Likelihood and impact 1-5 (assumed team scores). Score = L x I. Tier 0 = applies to all use cases. "
          "Review: Tier 3 monthly, Tier 2 quarterly, Tier 1 half-yearly.")
    hdr = ["ID", "Use case", "Risk", "L", "I", "Inherent score", "Controls", "Residual L", "Residual I", "Residual score", "Owner", "Approval tier", "Status", "Next review"]
    header(ws, 4, hdr, height=36)
    for i, r in enumerate(RISKS):
        row = 5 + i
        vals = [r[0], r[1], r[2], r[3], r[4], f"=D{row}*E{row}", r[5], r[6], r[7], f"=H{row}*I{row}", r[8], r[9] if r[9] else "All (policy)", "Open", ""]
        for j, v in enumerate(vals):
            c = cell(ws, row, j + 1, v, wrap=True, color="000000" if j not in (3, 4, 7, 8) else None,
                     fill=INPUT_FILL if j in (12, 13) else None, fmt="0" if j in (3, 4, 5, 7, 8, 9) else None)
    last = 4 + len(RISKS)
    red, amber, green = (PatternFill("solid", fgColor=c) for c in ("F8C9C9", "FDE7B8", "CFEBDD"))
    for col in ("F", "J"):
        rng = f"{col}5:{col}{last}"
        ws.conditional_formatting.add(rng, CellIsRule(operator="greaterThanOrEqual", formula=["12"], fill=red))
        ws.conditional_formatting.add(rng, CellIsRule(operator="between", formula=["6", "11"], fill=amber))
        ws.conditional_formatting.add(rng, CellIsRule(operator="lessThan", formula=["6"], fill=green))
    widths(ws, [5, 14, 40, 4, 4, 9, 52, 7, 7, 9, 14, 9, 9, 11])
    ws.freeze_panes = "D5"
    s = wb.create_sheet("Heat_Summary")
    title(s, "Risk counts by band")
    header(s, 3, ["Band", "Inherent", "Residual"])
    for i, (band, lo, hi) in enumerate([("High (12-25)", 12, 25), ("Medium (6-11)", 6, 11), ("Low (1-5)", 1, 5)]):
        r = 4 + i
        cell(s, r, 1, band, color="000000")
        cell(s, r, 2, f'=COUNTIFS(Risk_Register!$F$5:$F${last},">={lo}",Risk_Register!$F$5:$F${last},"<={hi}")', fmt="0")
        cell(s, r, 3, f'=COUNTIFS(Risk_Register!$J$5:$J${last},">={lo}",Risk_Register!$J$5:$J${last},"<={hi}")', fmt="0")
    widths(s, [18, 10, 10])
    p = OUT / "Risk_Register.xlsx"
    wb.save(p)
    return p


def doc():
    d = Doc("AI Governance Pack", "Phase 4 - acceptable-use policy, use-case risk register, approval workflow, DPDP Act and sector rules",
            "Member 2 - Chief AI & Risk Officer (CARO)", "04-Governance")
    d.callout(["**Three rules everyone remembers:** (1) A named person owns every AI output that leaves the company. "
               "(2) No buyer or personal data in unapproved AI tools, ever. (3) AI never makes a promise about money, dates, people or compliance.",
               "**Contents:** 1 Acceptable-use policy - 2 Data classes - 3 Approval workflow - 4 Risk register - 5 DPDP Act obligations - "
               "6 Sector and market rules - 7 Incident response - 8 Roles."],
              kind="info", title="Summary")

    d.h1("1. AI acceptable-use policy (v1.0)")
    d.h3("1.1 Scope and principles")
    d.p("Applies to every employee, contractor and job-work partner using AI tools for WeaveWorks work. Principles: **human accountability, "
        "privacy by design, transparency, fairness, security and proportionality** (controls scale with risk).")
    d.h3("1.2 Approved tools")
    d.table(["Tier", "Tools", "Allowed data"], [
        ("Approved (enterprise)", "Company-licensed AI assistants and builders with a data-processing agreement, SSO and a no-training setting; "
                                  "the no-code Agent platform; Power BI", "Public, Internal, Confidential"),
        ("Limited", "Free/consumer AI tools", "Public data only"),
        ("Prohibited", "Unvetted browser extensions, AI note-takers in buyer calls without consent, any tool without clear data terms", "None"),
    ], widths=[3.4, 9.6, 4.0], size=8.5, first_col_bold=True)
    d.h3("1.3 You must")
    d.bullets([
        "Check AI output before use; you own it. Facts and figures need a source or the label 'assumed'.",
        "Tell people when they are dealing with AI (chatbot disclosure; label AI-generated images sent to buyers).",
        "Get sign-off before AI-drafted prices, delivery dates, claims decisions or compliance statements go to a buyer.",
        "Complete AI-literacy training (2 hours, English or Tamil) before using an approved tool.",
        "Report any AI incident or data slip to the CARO within 1 hour. You will not be punished for honest reporting.",
    ])
    d.h3("1.4 You must not")
    d.bullets([
        "Paste Confidential or Restricted data (buyer designs, prices, contracts, anyone's personal data) into Limited or Prohibited tools.",
        "Use AI to make or automate decisions about hiring, pay, discipline or termination without human review and HR approval.",
        "Use AI to create, alter or summarise certificates, test reports or audit evidence for submission.",
        "Monitor workers with AI (e.g. camera productivity scoring) without an approved, notified purpose.",
        "Create deepfakes, impersonate people or generate claims you cannot evidence ('guaranteed compliant', 'best in India').",
    ])
    d.h1("2. Data classes")
    d.table(["Class", "Examples", "AI tools allowed", "Extra rule"], DATA_CLASSES, widths=[2.4, 6.0, 5.0, 3.6], size=8.5, first_col_bold=True)

    d.h1("3. Approval workflow")
    d.image(FIG / "approval_workflow.png", caption="Figure 1. Six-step approval workflow for any new AI use case or material change.")
    d.table(["Tier", "When", "Approvers", "Before go-live"], [
        ("1 - Low", "Internal productivity, no personal or confidential data (e.g. forecasting, SOP drafting)", "Department head + CARO (notified)", "Basic test; owner named"),
        ("2 - Medium", "Confidential business data or internal decisions with money impact (costing, marketing)", "Department head + CARO + IT/security", "Test plan; data-processing agreement; KPI baseline"),
        ("3 - High", "Customer-facing AI, personal data, decisions about people, compliance evidence (Agent, Studio, UC09, UC11, UC13)",
         "AI Board: CEO, CFO, CARO, function head (+ HR for people decisions)", "Full test suite; DPIA-style privacy check; bias test where relevant; human-in-loop design; rollback plan"),
    ], widths=[2.2, 6.4, 4.4, 4.0], size=8.5, first_col_bold=True)
    d.p("The AI Board meets monthly (30 minutes) and at each stage-gate. Every approval is logged in the risk register with an owner and review date.")

    d.h1("4. Use-case risk register (summary)")
    d.p("Full register with live scoring: `Risk_Register.xlsx` (17 risks). Highest inherent risks and how we bring them down:")
    top = sorted(RISKS, key=lambda r: -r[3] * r[4])[:8]
    d.table(["ID", "Use case", "Risk", "Inherent", "Key controls", "Residual", "Owner"],
            [(r[0], r[1], r[2], r[3] * r[4], r[5], r[6] * r[7], r[8]) for r in top],
            widths=[1.0, 2.0, 4.2, 1.4, 5.6, 1.4, 1.4], size=7.8)

    d.h1("5. India's DPDP Act 2023: obligations and how we meet them")
    d.table(["Obligation", "What WeaveWorks does", "Use cases"], DPDP, widths=[5.4, 8.6, 3.0], size=8)
    d.callout("The DPDP Rules 2025 phase in obligations over about 18 months from notification (Nov 2025), so most duties fall due inside our "
              "roadmap window. We plan as if all obligations apply from M1. Check the current status of each rule with counsel (budgeted under 'security, legal & audit').",
              kind="warn", title="Timing")

    d.h1("6. Sector and market rules")
    d.table(["Rule", "Why it matters to WeaveWorks", "Our control"], SECTOR, widths=[4.2, 7.2, 5.6], size=8)

    d.h1("7. AI incident response (one page)")
    d.table(["Step", "Within", "Action", "Lead"], [
        ("Detect & report", "1 hour", "Any staff member reports to CARO (WhatsApp hotline + email). Agent transcripts flagged by monitoring.", "Anyone"),
        ("Contain", "4 hours", "Pause the tool or feature; revoke shared links; ask platforms to remove leaked content.", "CARO + IT"),
        ("Assess", "24 hours", "What data, whose, how many people; is it a personal-data breach or a CERT-In reportable incident (6 hours)?", "CARO"),
        ("Notify", "As required", "Data Protection Board and affected persons (DPDP); CERT-In; buyers whose data is involved (contract terms).", "CEO + CARO"),
        ("Communicate", "Same day", "Holding statement; one spokesperson; no blame on individuals in public.", "CMO"),
        ("Learn", "2 weeks", "Root cause, fix, re-test, update risk register and training.", "CARO"),
    ], widths=[2.6, 1.8, 10.2, 2.4], size=8.5, first_col_bold=True)

    d.h1("8. Roles")
    d.table(["Role", "Responsibility"], [
        ("CEO (shared)", "Chairs AI Board; owns AI strategy and risk appetite"),
        ("Chief AI & Risk Officer (Member 2)", "Policy owner; triage and tiering; risk register; incident lead; data-protection contact"),
        ("CFO (Member 3)", "Budget, ROI tracking, stage-gate financial review, vendor contracts"),
        ("Function heads", "Own use cases, KPIs and human review in their teams"),
        ("AI champions (1 per department)", "First-line help, feedback, adoption"),
    ], widths=[5.2, 11.8], size=8.5, first_col_bold=True)
    d.h1("Prompt log: prompt, raw output, edited version")
    d.prompt_log(GOV_LOG)
    d.callout("Rejected: the AI's claim that the DPDP Act 'does not apply to B2B companies'. It does: buyer staff and our own employees are "
              "data principals. Accepting it would have left the Agent, marketing and HR streams without a lawful basis.", kind="ai",
              title="Rejected AI output (for the viva)")
    d.sources(["DPDP", "DPDP_RULES", "CERTIN", "GDPR", "EU_AI_ACT", "CSDDD", "ESPR", "EU_FLR", "REACH"],
              extra=["Code on Wages, 2019; Industrial Relations Code, 2020 (Government of India).",
                     "This pack is a student consultancy exercise, not legal advice."])
    return d.save(OUT / "Governance_Pack.docx")


GOV_LOG = [
    dict(title="Draft the acceptable-use policy",
         prompt="Draft a 1-page AI acceptable-use policy for a 1,150-employee Indian textile exporter whose workforce has low digital skills. Include "
                "approved tools, prohibited uses, data rules, human review and incident reporting. Plain English.",
         raw="Comprehensive but 4 pages; banned 'all use of ChatGPT and similar tools'; incident reporting 'within 72 hours'; no mention of buyer IP.",
         edited="Tiered tools (approved / limited / prohibited) instead of a ban; 1-hour internal reporting; buyer designs as Confidential; Tamil training.",
         why="A blanket ban drives shadow use (the Shock B scenario). 72 hours is too slow for CERT-In's 6-hour clock. Buyer IP is our biggest confidentiality exposure."),
    dict(title="Map DPDP obligations",
         prompt="List the obligations of a Data Fiduciary under India's DPDP Act 2023 that apply to an exporter using AI for customer support, "
                "marketing and HR screening. Cite sections.",
         raw="'...The DPDP Act primarily governs B2C data; B2B companies have limited obligations...' Section numbers partly wrong (breach notice cited as s.12).",
         edited="Obligation table with sections checked against the Act's text (s.4-9, 16, Schedule); B2B contact and employee data explicitly in scope.",
         why="Wrong scope and wrong citations would fail in front of the 'regulator' on the panel. Every section number was checked by Member 2 against the bare Act."),
    dict(title="Generate the risk register",
         prompt="Create a risk register for these 11 AI use cases [list]: risk, likelihood, impact, controls, owner.",
         raw="20 generic risks ('AI may be biased', 'data may leak'), all scored 3x3, no residual scoring, owner 'IT' for everything.",
         edited="17 specific risks tied to use cases and our shock scenarios; inherent and residual scores; business owners; approval tier.",
         why="Generic, uniformly scored risks cannot drive prioritisation. Owners must be the people who run the process, not IT."),
]

if __name__ == "__main__":
    print(risk_xlsx())
    print(doc())
