"""Phase 3 - Stress-test: Shock E response (rival AI personalisation + 10% price cut).  Also P5 prompt log."""
import facts
from docx_lib import Doc, ROOT

CO = facts.CO
REV = CO["revenue_cr"]
EBITDA = REV * CO["ebitda_margin"]
# revenue at risk by segment (share of segment revenue at risk - ASSUMED)
AT_RISK = {"A": 0.05, "B": 0.25, "C": 0.20}


def at_risk():
    rows, tot = [], 0
    for k, share in AT_RISK.items():
        seg_rev = REV * facts.SEGMENT_REV_SHARE[k]
        r = seg_rev * share
        tot += r
        rows.append((f"{k}. {facts.SEGMENT_NAMES[k]}", f"{seg_rev:.1f}", f"{share:.0%}", f"{r:.1f}"))
    return rows, tot


def pre_shock_total():
    uc = 0
    for r in facts.roi_rows():
        if r["pre_shock_golive"] is None:
            continue
        uc += r["one_time"] + r["run_pm"] * facts.months_live(r["pre_shock_golive"])
    bs = facts.budget_summary()
    return uc + bs["hidden"] + bs["contingency"]


def roadmap_changes():
    out = []
    for r in facts.roi_rows():
        pre, post = r["pre_shock_golive"], r["golive"]
        if not r["included"]:
            out.append((r["id"], r["name"], f"M{pre}", "Deferred beyond M18", f"-{r['cost']:.1f}", "Frees budget for UC17; needs labelled defect images first"))
        elif pre is None:
            out.append((r["id"], r["name"], "Not in plan", f"M{post}", f"+{r['cost']:.1f}", "Direct answer to rival's AI personalisation"))
        elif pre != post:
            pre_cost = r["one_time"] + r["run_pm"] * facts.months_live(pre)
            out.append((r["id"], r["name"], f"M{pre}", f"M{post}", f"+{r['cost'] - pre_cost:.1f} (extra run months)",
                        "Pulled forward to protect margin / speed"))
    return out


def build_shock():
    rows, tot = at_risk()
    d = Doc("Shock Response: Competitor Move", "Phase 3 - Stress-test. Card E: the main rival launches AI personalisation and cuts prices 10%",
            "Whole team; lead: Member 1 (CMO) with Member 3 (CFO)", "03-Stress-Test")
    d.callout([f"**Shock:** {facts.RIVAL} (fictional Tiruppur rival) launches an AI buyer portal that generates design variants in each buyer's "
               "brand style and personalised offers, and cuts FOB prices by 10%.",
               f"**Our answer:** do NOT match the price cut across the board. A 10% cut on Rs {REV:.0f} crore is Rs {REV*0.1:.0f} crore a year, more than "
               f"our entire EBITDA (about Rs {EBITDA:.1f} crore, assumed). Instead: (1) value-engineered offers for the at-risk value retailers, "
               "(2) accelerate our own personalisation with guardrails ('WeaveWorks Studio', UC17), (3) make compliance and traceability the reason to stay.",
               "**Roadmap impact:** UC17 added; costing assistant and virtual sampling pulled forward; computer-vision QC deferred. Still within Rs 50 lakh."],
              kind="info", title="Summary")

    d.h1("1. What happened and what it threatens")
    d.table(["Segment", "Revenue (Rs cr)", "Share at risk (assumed)", "Revenue at risk (Rs cr)"], rows + [("Total", f"{REV:.1f}", "", f"{tot:.1f}")],
            widths=[6.5, 3.0, 3.8, 3.7], size=9)
    d.bullets([
        "**Segment B (value retailers)** is most exposed: highest price sensitivity (8.8/10 in our data) and lowest sustainability requirement.",
        "**Segment C (agile labels)** is drawn to personalisation and speed, which is exactly what the rival's portal offers.",
        "**Segment A (sustainable brands)** is least exposed: switching suppliers means re-doing due diligence, and they value traceability.",
    ])
    d.h1("2. Options we considered")
    d.table(["Option", "Upside", "Downside", "Verdict"], [
        ("1. Match the 10% cut for all buyers", "Keeps volume short term",
         f"Costs about Rs {REV*0.1:.0f} crore a year, more than total EBITDA; trains buyers to expect cuts", "Reject"),
        ("2. Do nothing; wait for the rival to stumble", "No cost now", f"Up to Rs {tot:.0f} crore revenue at risk; capacity under-used", "Reject"),
        ("3. Exit Segment B, focus on A", "Higher margin mix", "37% of revenue; factories need volume to absorb fixed cost", "Reject"),
        ("4. Targeted value-match + accelerated personalisation with guardrails + compliance as the moat", "Protects margin and the at-risk accounts; differentiates",
         "Needs fast delivery of UC15/UC02/UC17; some B volume may still leave", "RECOMMENDED"),
    ], widths=[5.0, 3.6, 5.8, 2.6], size=8.5, first_col_bold=True)

    d.h1("3. The response in three moves")
    d.h3("Move 1 - Value, not price (Segment B, weeks 0-6)")
    d.bullets([
        "Identify the at-risk B accounts with the segmentation model plus order-intake signals; CFO sets a floor margin per account.",
        "Offer **value-engineered** alternatives instead of a discount: fabric GSM optimisation, trim and packaging rationalisation, "
        "larger colour-runs. Target 4-6% lower FOB at the same margin (assumed range, validated style by style).",
        "Use the costing assistant (UC15, pulled forward to M4) to turn these around in 24 hours, with assumptions shown to the buyer.",
        "Show total cost per piece, not the quote (P5 deck slide 6): hidden costs of re-sampling, compliance admin and late delivery.",
    ])
    d.h3("Move 2 - Personalisation with guardrails (Segments C and A, months 2-10)")
    d.bullets([
        "**UC17 'WeaveWorks Studio'**: buyer portal with AI design variants in the buyer's style, instant indicative quote and the "
        "traceability sheet. Built on UC02 (3D sampling, pulled forward to M7) and UC15 (costing).",
        "Guardrails the rival may not have: buyer designs never used to train shared models without written consent; every price "
        "is indicative until a merchandiser confirms; AI-generated imagery labelled; Tier-3 approval before launch (Governance pack).",
        "Speed vs risk: closed beta with 5 friendly Segment C buyers at M8, general launch at M9 only after the test suite passes.",
    ])
    d.h3("Move 3 - Compliance as the moat (Segment A, ongoing)")
    d.bullets([
        "Compliance document assistant (UC09) stays at M7: per-style traceability sheets and faster due-diligence questionnaire answers.",
        "Marketing campaign 'Traceable by Design' leads with what a cheaper quote does not include.",
    ])

    d.h1("4. What changes in the roadmap and budget")
    d.table(["ID", "Use case", "Before shock", "After shock", "Budget change (Rs lakh)", "Why"], roadmap_changes(),
            widths=[1.2, 4.6, 2.0, 2.6, 2.4, 4.2], size=8.5)
    bs = facts.budget_summary()
    d.p(f"Net effect: total committed **Rs {bs['total']:.2f} lakh** (Rs {pre_shock_total():.2f} lakh before the shock); "
        f"Rs {bs['headroom']:.2f} lakh unallocated reserve remains for Stage-gate 1. Base-case 18-month benefit of UC17: "
        f"Rs {[r for r in facts.roi_rows() if r['id']=='UC17'][0]['benefit'][1]:.1f} lakh (break-even if it retains about "
        f"Rs {[r for r in facts.roi_rows() if r['id']=='UC17'][0]['breakeven_driver']:.1f} crore of revenue a year). See ROI_Model.xlsx.")

    d.h1("5. 30-60-90 day plan")
    d.table(["When", "Actions", "Owner", "Signal we watch"], [
        ("Days 0-30", "Key-account calls to all 9 Segment B accounts; value-engineering offers on top 20 styles; win/loss tracking starts; "
                      "C3 creative + total-cost slide live", "CMO + CFO", "Weekly RFQ win rate; lost-deal reasons"),
        ("Days 31-60", "Costing assistant live (M4); Studio design sprint; governance Tier-3 review of Studio; buyer advisory group (5 buyers)",
         "COO + CCO/CARO", "Quote turnaround < 24 h; B-account order intake vs forecast"),
        ("Days 61-90", "3D sampling pilot; Studio closed beta plan; board update on retained revenue", "Whole team", "Retained at-risk revenue vs Rs 1.5 crore base case"),
    ], widths=[2.2, 8.2, 2.8, 3.8], size=8.5, first_col_bold=True)

    d.h1("6. Communication")
    d.h3("To buyers (key-account script, 90 seconds)")
    d.p("\"You may have seen lower quotes in the market. We will not ask you to pay for things you do not need, so here are three "
        "value-engineered options for your top styles, each costed in 24 hours. We also want to show you what is not in a cheaper quote: "
        "traceability data your auditors accept, fewer sample rounds and on-time delivery we track every week. If price alone decides, "
        "we will tell you honestly which styles we cannot win.\"", italic=True)
    d.h3("To employees")
    d.p("\"A competitor is using AI and lower prices. Our answer is to use AI to remove waste and make your work faster, not to remove jobs. "
        "Merchandisers will spend less time chasing status and more time with buyers. Training starts this month.\" (Linked to the ADKAR plan.)", italic=True)
    d.h1("7. What we will not do, even under pressure")
    d.bullets([
        "Launch customer-facing AI without the Tier-3 approval and the test suite (a rushed bot is how out-of-policy promises happen).",
        "Use buyers' designs or data to train shared models without written consent.",
        "Let any AI tool send a price, delivery date or compliance statement to a buyer without human sign-off.",
        "Cut training, oversight or security budget to fund the response.",
    ])
    d.sources(["CSDDD", "ESPR"], extra=["Revenue-at-risk shares, EBITDA margin and value-engineering range are ASSUMED team estimates."])
    return d.save(ROOT / "03_Stress_Test" / "Shock_E_Response_Competitor_Move.docx")


P5_LOG = [
    dict(title="Outline the B2B proposal deck",
         prompt="Outline a 9-slide B2B proposal deck from an Indian knitwear exporter to a sustainable mid-market EU fashion brand (persona pasted). "
                "Lead with the buyer's priorities, then our offer, commercial terms, pilot and next steps.",
         raw="Slide 1 'About WeaveWorks - 20 years of excellence'; slide 2 'Our state-of-the-art factories'; ... slide 8 'Why we are the best in India'.",
         edited="Reordered to buyer-first: what we heard -> who we are -> three commitments -> process -> total cost -> pilot -> comparison -> next steps.",
         why="The raw outline was supplier-centric and full of unprovable superlatives. A sourcing head buys solutions to her problems, not our history."),
    dict(title="Total-cost comparison against the price-cutting rival",
         prompt="Our rival cut prices 10%. Build an illustrative per-piece total-cost comparison (FOB + hidden costs) showing why our offer is "
                "better value. FOB EUR 4.60 vs 4.14.",
         raw="Hidden costs for the rival: sampling EUR 0.45, compliance EUR 0.60, delays EUR 0.50 -> 'WeaveWorks is 20% cheaper overall'.",
         edited="Hidden costs set to modest, defensible assumptions (0.20 / 0.28 / 0.18 vs 0.06 / 0.05 / 0.04), giving EUR 4.75 vs 4.80, "
                "labelled illustrative, with an offer to rebuild using the buyer's data.",
         why="The raw numbers were inflated to win the argument; a sourcing professional would see through them and distrust everything else. "
             "A small, honest advantage plus a pilot is more credible."),
    dict(title="Competitor comparison table",
         prompt="Create a comparison table: WeaveWorks vs KnitSphere Apparels vs a large integrated mill, on traceability, sampling speed, costing speed, MOQ, certificates.",
         raw="Named KnitSphere with 'No' on GOTS and 'poor quality record'.",
         edited="Columns are 'Typical low-price offer' and 'Large integrated mill'; no named competitor; 'Partly' where it varies.",
         why="We have no evidence about the rival's certificates or quality; naming and disparaging a competitor is unfair and a legal risk."),
]


def build_p5_log():
    d = Doc("Sales Prototype: Prompt Log & Speaker Guide", "Prototype P5 - B2B proposal deck for the largest segment (Sustainable mid-market brands)",
            "Member 1 - CMO / Sales lead", "P5-Sales")
    d.p("Deck: B2B_Proposal_Deck_Nordlicht.pptx (9 slides, speaker notes included). The prospect 'Nordlicht Basics GmbH' and its sourcing head "
        "are fictional, built from the Segment A persona. Segment A is the largest segment: 48% of revenue (assumed).")
    d.table(["Slide", "Message", "Evidence used"], [
        ("1", "Traceable by Design - proposal", "Campaign concept (P1)"),
        ("2", "What we heard: 4 priorities", "Persona Katrin A. (P1)"),
        ("3", "WeaveWorks at a glance", "Company facts (assumed)"),
        ("4", "Three measurable commitments", "UC02, UC06, UC09, UC15, Agent (P2)"),
        ("5", "Tech-pack to warehouse; AI vs people", "Governance pack: human sign-off rule"),
        ("6", "Total cost, not the quote", "Shock E response; illustrative numbers"),
        ("7", "90-day pilot with exit clause", "Indicative terms (assumed)"),
        ("8", "Why WeaveWorks (no named competitor)", "Team assessment"),
        ("9", "Next steps", "-"),
    ], widths=[1.4, 7.6, 8.0], size=8.5)
    d.h1("Prompt log: prompt, raw output, edited version")
    d.prompt_log(P5_LOG)
    d.callout("Rejected: the AI's total-cost numbers that made WeaveWorks '20% cheaper overall'. They were invented to win the slide, and one "
              "challenge from the buyer would have destroyed our credibility.", kind="ai", title="Rejected AI output (for the viva)")
    d.sources([], extra=["All prices, costs and terms in the deck are illustrative assumptions."])
    return d.save(ROOT / "02_Build" / "P5_Sales" / "Sales_Prompt_Log.docx")


if __name__ == "__main__":
    print(build_shock())
    print(build_p5_log())
