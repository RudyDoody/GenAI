"""Phase 4 - ADKAR change plan + 18-month roadmap (three horizons, within Rs 50 lakh, hidden costs included)."""
import json

import facts
from docx_lib import Doc, FIG, ROOT

R = json.loads((ROOT / "data" / "results.json").read_text())
HZ = {"H1": (1, 6), "H2": (7, 12), "H3": (13, 18)}
HIDDEN_SPLIT = [(0.5, 0.3, 0.2), (1 / 3, 1 / 3, 1 / 3), (0.4, 0.3, 0.3)]  # training, oversight, security - assumed phasing


def spend_by_horizon():
    """One-time cost lands in the month before go-live (build); run cost from go-live month."""
    out = {h: 0.0 for h in HZ}
    per_uc = []
    for r in facts.roi_rows():
        if not r["included"]:
            continue
        row = {h: 0.0 for h in HZ}
        build_m = max(1, r["golive"] - 1)
        for h, (a, b) in HZ.items():
            if a <= build_m <= b:
                row[h] += r["one_time"]
            months = max(0, min(b, 18) - max(a, r["golive"]) + 1)
            row[h] += r["run_pm"] * months
        per_uc.append((r["id"], r["name"], row))
        for h in HZ:
            out[h] += row[h]
    hidden = []
    for (name, v), split in zip(facts.HIDDEN_COSTS, HIDDEN_SPLIT):
        row = {h: v * s for h, s in zip(HZ, split)}
        hidden.append((name, row))
        for h in HZ:
            out[h] += row[h]
    return per_uc, hidden, out


def roadmap():
    per_uc, hidden, tot = spend_by_horizon()
    bs = facts.budget_summary()
    d = Doc("18-Month AI Roadmap", "Phase 4 - three horizons, Rs 50 lakh including hidden costs (training, oversight, contingency); reflects Shock E",
            "Shared CEO role; integrated by Member 4 (COO) with Member 3 (CFO)", "04-Roadmap")
    n_funded = sum(1 for r in facts.roi_rows() if r["included"])
    d.callout([f"**Plan:** {n_funded} funded AI use cases in three horizons, January 2027 to June 2028. Total committed **Rs {bs['total']:.2f} lakh** = "
               f"use cases Rs {bs['uc_cost']:.2f} lakh + hidden costs Rs {bs['hidden']:.2f} lakh + contingency Rs {bs['contingency']:.2f} lakh; "
               f"**Rs {bs['headroom']:.2f} lakh** unallocated reserve released only at Stage-gate 1.",
               "**Shock E reflected:** UC17 WeaveWorks Studio added (H2); costing assistant pulled into H1; virtual sampling pulled forward; "
               "computer-vision QC deferred; HR stream (UC13) in H3 as the brief requires."],
              kind="info", title="Summary")
    d.image(FIG / "roadmap_gantt.png", caption="Figure 1. Roadmap Gantt (M1 = January 2027).")
    d.h1("1. Three horizons")
    d.table(["Horizon", "Theme", "What goes live", "Exit criteria (stage-gate)"], [
        ("H1  M1-6\nJan-Jun 2027", "Prove & protect", "Governance + AUP; AI literacy for all; Buyer support Agent (M3); segmentation & campaigns (M3); "
         "RFQ/proposal generator (M4); demand forecast in Power BI (M4); costing assistant (M4)",
         "Gate 1 (M6): Agent resolves >= 35% of queries with 0 policy breaches; quote turnaround < 24 h; forecast MAPE < 12% on real data; "
         "spend within plan. Release the reserve."),
        ("H2  M7-12\nJul-Dec 2027", "Scale & differentiate", "Compliance document assistant (M7); virtual 3D sampling (M7); WeaveWorks Studio portal (M9, after beta); "
         "Tamil voice SOP assistant (M9)",
         "Gate 2 (M12): sample lead time <= 12 days; Studio retaining at-risk revenue; scale or stop each use case on ROI evidence."),
        ("H3  M13-18\nJan-Jun 2028", "Embed & extend", "HR stream: bias-checked JDs & screening (M13); export-docs/LC checker business case; "
         "computer-vision QC data collection (go/no-go at M18); annual AI audit; FY29 roadmap",
         "M18 board review: programme ROI vs base case; audit findings closed; next-phase budget request."),
    ], widths=[2.4, 2.4, 6.6, 5.6], size=8.5, first_col_bold=True)

    d.h1("2. Budget by horizon (Rs lakh)")
    rows = [(i, n, f"{r['H1']:.2f}", f"{r['H2']:.2f}", f"{r['H3']:.2f}", f"{sum(r.values()):.2f}") for i, n, r in per_uc]
    rows += [("Hidden", n.split(" (")[0], f"{r['H1']:.2f}", f"{r['H2']:.2f}", f"{r['H3']:.2f}", f"{sum(r.values()):.2f}") for n, r in hidden]
    rows.append(("", "Contingency (held centrally; released by CFO)", "-", "-", "-", f"{facts.CONTINGENCY_LAKH:.2f}"))
    rows.append(("", "TOTAL COMMITTED", f"{tot['H1']:.2f}", f"{tot['H2']:.2f}", f"{tot['H3']:.2f}", f"{bs['total']:.2f}"))
    rows.append(("", "Unallocated reserve (Stage-gate 1)", "", "", "", f"{bs['headroom']:.2f}"))
    rows.append(("", "BUDGET", "", "", "", f"{facts.BUDGET_LAKH:.2f}"))
    d.table(["ID", "Line", "H1", "H2", "H3", "18 months"], rows, widths=[1.4, 9.4, 1.5, 1.5, 1.5, 1.7], size=8, align_right=(2, 3, 4, 5))
    hid = bs["hidden"] + bs["contingency"]
    d.p(f"Hidden costs (training, oversight, security/legal) plus contingency are **Rs {hid:.1f} lakh, {hid / bs['total']:.0%} of the committed budget**. "
        "We deliberately did not cut them to fund more use cases: with a low-digital-skills workforce, adoption is the binding constraint, not software.")

    d.h1("3. Value we expect (base case, assumed drivers)")
    rr = [r for r in facts.roi_rows() if r["included"]]
    ben = [sum(r["benefit"][k] for r in rr) for k in range(3)]
    d.table(["", "Pessimistic", "Base", "Optimistic"], [
        ("18-month benefit (Rs lakh)", f"{ben[0]:.1f}", f"{ben[1]:.1f}", f"{ben[2]:.1f}"),
        ("Programme cost incl. hidden + contingency", f"{bs['total']:.1f}", f"{bs['total']:.1f}", f"{bs['total']:.1f}"),
        ("Programme ROI", f"{(ben[0]-bs['total'])/bs['total']:.0%}", f"{(ben[1]-bs['total'])/bs['total']:.0%}", f"{(ben[2]-bs['total'])/bs['total']:.0%}"),
    ], widths=[7.0, 3.3, 3.3, 3.4], size=9, first_col_bold=True)
    d.p(f"Break-even requires delivering **{bs['total'] / ben[1]:.0%} of base-case benefits** in 18 months. The pessimistic case loses money inside "
        "the window, which is why stage-gates can stop weak use cases early. Detail: ROI_Model.xlsx.")

    d.h1("4. KPIs the board will see each quarter")
    d.table(["KPI", "Baseline (assumed)", "M6", "M12", "M18"], [
        ("Buyer queries resolved by the Agent", "0%", "35%", "45%", "50%"),
        ("Quote turnaround", "3-5 days", "< 24 h", "< 24 h", "< 12 h"),
        ("Sample lead time (working days)", "21", "18", "12", "10"),
        ("RFQ win rate", f"{facts.CO['rfq_win_rate']:.0%}", "19%", "20%", "21%"),
        ("Forecast MAPE (real data)", f"{facts.CO['forecast_mape_now']:.0%}", "15%", "12%", "10%"),
        ("Staff completing AI literacy", "0%", "80%", "95%", "100%"),
        ("AI incidents with policy breach", "n/a", "0", "0", "0"),
    ], widths=[6.6, 3.0, 2.4, 2.4, 2.6], size=8.5)

    d.h1("5. Dependencies and the 'other stream'")
    d.bullets([
        "ERP order data export (daily CSV) is needed for the Agent (UC04) and forecast (UC06): IT, M1-M2.",
        "3D sampling (UC02) and costing (UC15) are prerequisites for the Studio portal (UC17).",
        "Clean product and supplier master data is a prerequisite for the compliance assistant (UC09).",
        "**HR stream** (the stream we did not prototype): bias-checked JDs and screening of operator/supervisor candidates (UC13) at H3, "
        "after the governance pack is proven. Tier-3 approval, planted-pair bias test before first use, no automatic rejection.",
    ])
    d.sources([], extra=["Phasing of hidden costs (training 50/30/20, oversight even, security 40/30/30) and all driver values are ASSUMED."])
    return d.save(ROOT / "04_Decide" / "Roadmap_18_Month.docx")


ADKAR = [
    ("Awareness", "Why change?", "Merchandisers hear 'AI' as 'replacement'; a rival is winning on speed and price",
     ["CEO town-hall (Tamil + English): the rival's move and what it means for orders and jobs",
      "Share the facts: 1,600 queries/month, 42% repetitive, 21-day sampling vs 10 the market expects",
      "Written commitment: no AI-driven job losses for 18 months; time saved goes to buyer development"],
     "Pulse survey: 'I understand why we are changing' >= 80% agree", "M1"),
    ("Desire", "What's in it for me?", "Fear of looking incompetent with new tools; extra work during transition",
     ["Each merchandiser picks one daily task to hand to AI first (their choice)",
      "AI champions from within the team (2 of 28); peer demos, not vendor demos",
      "Recognition: monthly 'hours given back' shout-out; KPI on buyer time, not on AI usage"],
     "Volunteers for pilot >= 10 of 28", "M1-M2"),
    ("Knowledge", "How do I do it?", "Low digital skills on the floor; English-first tools",
     ["2-hour AI literacy (Tamil/English): what AI is, the 3 rules, data classes",
      "Role-based labs: Agent escalation inbox, costing assistant, Power BI forecast page",
      "One-page job aids with screenshots; WhatsApp help group run by champions"],
     "Training completion >= 95%; 10-question quiz pass >= 80%", "M2-M4"),
    ("Ability", "Can I do it at work?", "Time pressure in peak season; fear of mistakes with buyers",
     ["Go-live outside peak months; buddy system for the first 4 weeks",
      "Human sign-off built into every flow, so no merchandiser is exposed alone",
      "Weekly 15-min clinic; champions sit with low-confidence users"],
     "Agent escalations handled within target >= 90%; error rate trend", "M3-M6"),
    ("Reinforcement", "Will it stick?", "Old habits return under pressure; leaders fall back to email chains",
     ["Stage-gate reviews include adoption metrics; leaders model use in their own meetings",
      "Update job descriptions: 'buyer development' time in each merchandiser's role",
      "Celebrate wins in town-halls; fix complaints visibly within 2 weeks"],
     "Usage retained at M9 and M12; voluntary attrition of merchandisers not above baseline", "M6-M18"),
]


def change_plan():
    d = Doc("Change Plan (ADKAR)", "Phase 4 - change plan for the employee group most affected: the merchandising team",
            "Member 2 (CARO) with Member 4 (COO); CEO role shared", "04-Change-Plan")
    d.callout([f"**Most affected group:** the {facts.CO['merchandisers']} merchandisers. Five of our AI use cases change their daily work: "
               "the support Agent takes routine buyer queries, the RFQ generator and costing assistant change quoting, the forecast changes "
               "capacity booking, and Studio changes sampling.",
               "**Secondary group:** about 120 production supervisors (Tamil voice SOP assistant, H2) - same ADKAR structure, delivered on the floor, "
               "in Tamil, with worker-representative consultation before rollout."],
              kind="info", title="Summary")
    d.h1("1. Stakeholder map")
    d.table(["Group", "Impact", "Likely reaction", "What they need"], [
        ("Merchandisers (28)", "High: tasks change daily", "Anxious, some curious", "Job security, simple tools, credit for buyer wins"),
        ("Merchandising Manager", "High: new KPIs and approvals", "Supportive if workload drops", "Dashboards, authority over escalations"),
        ("Production supervisors (~120)", "Medium (H2)", "Sceptical; low digital confidence", "Tamil voice, no extra typing, safety assurance"),
        ("Worker representatives", "Indirect", "Wary of surveillance and job cuts", "Consultation, written no-surveillance and no-job-loss commitments"),
        ("Sales & founders", "Medium", "Want speed against the rival", "Fast wins without reputational risk"),
        ("Buyers (external)", "Medium", "Expect faster answers; wary of bots", "Human escalation, accurate data, AI disclosure"),
    ], widths=[3.6, 3.6, 3.8, 6.0], size=8.5, first_col_bold=True)
    d.h1("2. ADKAR plan for the merchandising team")
    for stage, q, barrier, acts, measure, when in ADKAR:
        d.h3(f"{stage}  |  {q}  |  {when}")
        d.table(["", ""], [("Barrier today", barrier), ("Actions", "\n".join("- " + a for a in acts)), ("How we measure", measure)],
                widths=[3.2, 13.8], show_header=False, size=8.5, first_col_bold=True)
    d.h1("3. Resistance we expect, and our response")
    d.table(["Resistance", "Response"], [
        ("'AI will take my job'", "Written 18-month commitment; show how saved time is used (buyer development, new accounts); no headcount targets linked to AI."),
        ("'The bot will embarrass me with a buyer'", "Hard limits; humans own every commitment; merchandiser sees every escalation first."),
        ("'I'm not good with computers'", "Tamil sessions, champions, job aids, voice where possible; no one is assessed on AI skill in H1."),
        ("'Management will watch us through AI'", "Policy bans AI productivity surveillance without approved, notified purpose; worker reps briefed."),
        ("'We tried software before and it failed'", "Start with 1 task each; quick visible wins; fix complaints in 2 weeks; stage-gate can stop tools that don't help."),
    ], widths=[5.0, 12.0], size=8.5, first_col_bold=True)
    d.h1("4. Communication calendar")
    d.table(["When", "What", "Channel", "Who"], [
        ("M1 week 1", "Why we are changing; job commitment; Shock E context", "Town-hall (Tamil/English) + letter", "CEO (shared) + COO"),
        ("M1 week 2", "AI rules: 3 golden rules, data classes", "Team huddles + poster", "CARO + champions"),
        ("Monthly", "Wins, hours given back, what we fixed", "WhatsApp broadcast + notice board", "Champions"),
        ("Each go-live", "What changes on Monday; who to call", "15-min stand-up + job aid", "Function head"),
        ("Stage-gates", "Results, what we stop and what we scale", "Town-hall", "CEO + CFO"),
    ], widths=[2.6, 6.8, 4.4, 3.2], size=8.5)
    d.h1("5. Budget link")
    d.p(f"Change management is funded from the **Rs {facts.HIDDEN_COSTS[0][1]:.1f} lakh training & change** line (sessions, champions' time, "
        "Tamil job aids, translation) and is ring-fenced: the pessimistic-budget cut list never touches it.")
    d.h1("Prompt log: prompt, raw output, edited version")
    d.prompt_log([dict(
        title="Draft the ADKAR plan",
        prompt="Write an ADKAR change plan for 28 merchandisers at an Indian knitwear exporter adopting a support chatbot, AI costing and forecasting. "
               "Workforce has low digital skills. Include actions and measures per stage.",
        raw="'Awareness: send an email from the CEO explaining AI benefits. Desire: offer a bonus to employees who use AI the most. "
            "Knowledge: e-learning modules in English. Ability: ... Reinforcement: track AI usage per employee and include it in appraisals.'",
        edited="Town-hall in Tamil/English with a job-security commitment; no usage-based bonus or appraisal metric; hands-on labs; champions; "
               "reinforcement through buyer-time KPIs.",
        why="An English email and English e-learning ignore the low-digital-skills tension. Rewarding or appraising 'AI usage' drives "
            "box-ticking and feels like surveillance, which feeds resistance and the strike risk.")])
    d.callout("Rejected: tracking AI usage per employee in appraisals. It looks data-driven, but it measures activity, not value, and would "
              "be read as surveillance by a workforce already worried about jobs.", kind="ai", title="Rejected AI output (for the viva)")
    d.sources(["ADKAR"], extra=["Group sizes (28 merchandisers, ~120 supervisors) are ASSUMED."])
    return d.save(ROOT / "04_Decide" / "Change_Plan_ADKAR.docx")


if __name__ == "__main__":
    print(roadmap())
    print(change_plan())
    _, _, tot = spend_by_horizon()
    print({k: round(v, 2) for k, v in tot.items()}, round(sum(tot.values()), 2))
