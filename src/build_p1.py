"""P1 - Marketing pack: dataset, 3 segments, 2 personas, campaign, 3 Canva creatives,
3-email automated sequence, ad copy for 3 platforms, prompt log."""
import json
import shutil

import facts
from docx_lib import Doc, FIG, ROOT

R = json.loads((ROOT / "data" / "results.json").read_text())
OUT = ROOT / "02_Build" / "P1_Marketing"
CRE = OUT / "creatives"

# ------------------------------------------------------------------ ad copy
LINKEDIN = {
    "Intro text (<=150 chars shown before 'see more')":
        "EU due-diligence rules are coming. Is your knitwear supplier ready with farm-to-finish data? See what traceable looks like.",
    "Headline (<=70)": "Knit in Tiruppur. Traceable to the fibre.",
    "Description (<=100)": "GOTS & OEKO-TEX knitwear with product data your compliance team can use.",
    "CTA button": "Download",
}
GOOGLE_H = ["Traceable Organic Knitwear", "GOTS Certified T-Shirt Maker", "Knitwear Exporter, Tiruppur",
            "Samples in 10 Days", "DPP-Ready Product Data"]
GOOGLE_D = ["Farm-to-finish traceability for EU brands. GOTS & OEKO-TEX certified. Get a sample pack.",
            "Private-label tees, polos, kidswear & loungewear. Firm costing in 24 hrs. Talk to us."]
META = {
    "Primary text (<=125 shown)": "From tech-pack to approved sample in 10 days. AI 3D previews first, so the sample you get is the one you approve.",
    "Headline (<=40)": "Samples in 10 days, not 21",
    "Description (<=30)": "Small MOQs for new labels",
    "CTA button": "Learn more",
}


def lim(label):
    import re
    m = re.search(r"<=(\d+)", label)
    return int(m.group(1)) if m else None


def check_ads():
    rows = []
    for k, v in LINKEDIN.items():
        rows.append(("LinkedIn (Segment A)", k, v, len(v), lim(k)))
    for i, h in enumerate(GOOGLE_H, 1):
        rows.append(("Google Search RSA", f"Headline {i} (<=30)", h, len(h), 30))
    for i, dd in enumerate(GOOGLE_D, 1):
        rows.append(("Google Search RSA", f"Description {i} (<=90)", dd, len(dd), 90))
    for k, v in META.items():
        rows.append(("Meta / Instagram (Segment C)", k, v, len(v), lim(k)))
    for r in rows:
        assert r[4] is None or r[3] <= r[4], f"Ad copy over limit: {r}"
    return rows


EMAILS = [
    dict(n=1, day="Day 0 (trigger)", trig="Contact downloads the 'EU Knitwear Compliance Checklist' or is scanned at a trade fair; opted in",
         subj="Your EU knitwear compliance checklist", prev="12 questions to ask any supplier before 2027",
         body=["Hi {{first_name}},",
               "Thanks for downloading the checklist. It covers the 12 questions EU sourcing teams now ask: fibre origin, "
               "certificates, chemical safety, worker welfare and the product data a Digital Product Passport will need.",
               "We built it from the questionnaires our own buyers send us every season. If you are reviewing suppliers, "
               "it should save your team a few hours.",
               "Over the next week I'll share how we answer those questions for a real garment.",
               "Regards, {{sender}}, WeaveWorks Exports, Tiruppur"],
         cta="Open the checklist (PDF)"),
    dict(n=2, day="Day 4", trig="Sent if email 1 delivered; skip if contact already booked a call",
         subj="How we traced one organic polo, field to finish",
         prev="Five steps, one data sheet, no spreadsheets",
         body=["Hi {{first_name}},",
               "Here is one garment's journey: certified organic cotton bale, spinner, knitting, ZDHC-aligned dye house, "
               "stitching line, shipment. Each step has a certificate or a transaction record behind it.",
               "We now give every style a one-page traceability sheet in a format you can drop into your own "
               "due-diligence file. [Illustrative example - figures are sample data.]",
               "Would a sheet like this help your team?"],
         cta="See the sample traceability sheet"),
    dict(n=3, day="Day 10", trig="Sent to all; if contact clicked emails 1 or 2, also alert the account manager to call within 48 h",
         subj="A 10-day sample, on us",
         prev="Send a tech-pack; we'll send a 3D preview, then the sample",
         body=["Hi {{first_name}},",
               "If you have a style for the next season, send us the tech-pack. Within 10 working days you get an AI-assisted "
               "3D preview to approve, then a physical sample and its traceability sheet. The first sample is free for new brands "
               "(one style, standard fabrics).",
               "No obligation. If the fit or the data isn't right, tell us why; we learn either way.",
               "Reply to this email or book 20 minutes with our sourcing team."],
         cta="Book a 20-minute call"),
]

PERSONAS = [
    dict(name="Katrin A. (fictional)", seg="A - Sustainable mid-market brands",
         role="Head of Sourcing, 38, womenswear & basics brand, Hamburg; 140 staff; ~40 styles/season",
         goals=["Prove supply-chain due diligence to her board and retail partners",
                "Organic/recycled share above 70% of the range",
                "Fewer, longer supplier relationships"],
         pains=["Suppliers send certificates as scattered PDFs; data gaps found at audit time",
                "Sample rounds slip, compressing the calendar",
                "Audit fatigue: the same questionnaire answered five ways"],
         triggers="New EU due-diligence and DPP requirements; a supplier failing an audit",
         channels="LinkedIn, Munich Fabric Start, industry newsletters, peer referrals",
         message="'Traceable to the fibre': one clean data sheet per style, certificates in one place.",
         quote="I don't need the cheapest tee. I need the one I can defend in an audit."),
    dict(name="Marta K. (fictional)", seg="B - Volume value retailers",
         role="Senior Buyer, kids & basics, 44, discount retail chain, Poznan; 300 stores; ~1.2 mn pcs/yr in her category",
         goals=["Hit a fixed price point for each season", "On-time delivery for promotions",
                "Pass OEKO-TEX and amfori BSCI checks without drama"],
         pains=["Quotes change after yarn-price moves", "Late shipments mean empty promo shelves",
                "Re-inspecting quality because suppliers' reports aren't trusted"],
         triggers="Yarn price spikes; a competitor supplier offering 10% lower prices",
         channels="Sourcing agents, email, trade fairs; low social media use",
         message="'Price you can plan on': firm costing in 24 h, assumptions shown, on-time tracking.",
         quote="A cheaper quote that moves later is not cheaper."),
]


def build():
    seg = R["segmentation"]
    segs = {s["seg"]: s for s in seg["segments"]}
    shutil.copy(ROOT / "data" / "buyers_synthetic.csv", OUT / "buyers_synthetic.csv")
    shutil.copy(ROOT / "data" / "buyers_segmented.csv", OUT / "buyers_segmented.csv")

    d = Doc("Marketing Pack", "Prototype P1 - segmentation, personas, campaign, creatives, email automation, ad copy",
            "Member 1 - Chief Marketing Officer", "P1-Marketing")
    d.callout(["Who we sell to: three segments found in a synthetic 300-account buyer database. Campaign 'Traceable by Design' "
               "leads with Segment A (48% of revenue, assumed) and adapts proof points for B and C.",
               "Deliverables: dataset (CSV) - 3 segments - 2 personas - campaign concept - 3 creatives (Canva build spec + mock-ups) - "
               "3-email automated sequence - ad copy for LinkedIn, Google Search and Meta/Instagram - prompt log with one rejected output."],
              kind="info", title="Summary")

    d.h1("1. Synthetic dataset")
    d.p("No real customer data was used (ground rule). We generated **300 B2B buyer accounts** (38 active customers, "
        "40 lapsed, 222 prospects) with 18 fields: country, buyer type, annual volume, order quantity per style, styles per "
        "season, sustainability requirement score, price sensitivity, expected sample turnaround, digital engagement, target "
        "FOB price, payment terms, certifications required, preferred channel, relationship status, tenure, WeaveWorks revenue "
        "and email opt-in. Files: `buyers_synthetic.csv` (input) and `buyers_segmented.csv` (with segment labels). "
        "Generator: `src/gen_data.py` (seeded, reproducible).")
    d.h1("2. Segmentation: three segments")
    d.p("Method: 8 numeric features standardised (log for volume and order size), **k-means with k = 3**, chosen by the elbow "
        "check and because three segments are easy for a 14-person sales and customer-service team to act on. "
        f"Note: because the data are synthetic, the clusters match the generator's hidden groups "
        f"({seg['agreement_with_latent']:.0%} agreement); real CRM data will be messier.")
    d.image(FIG / "segments_profile.png", width_cm=15.5,
            caption="Figure 1. Segment profiles: mean of each feature in standard deviations from the all-buyer average.")
    rows = []
    for k in "ABC":
        s = segs[k]
        rows.append((f"{k}. {s['name']}", s["accounts"], s["active"], f"{s['revenue_lakh']/100:.1f} cr ({s['rev_share']:.0%})",
                     f"{s['est_annual_volume_k_pcs']:.0f}k", f"{s['avg_order_qty_per_style']:,.0f}",
                     f"{s['sustainability_requirement_score']:.0f}", f"{s['price_sensitivity_1to10']:.1f}",
                     f"{s['expected_sample_turnaround_days']:.0f}", f"EUR {s['target_fob_eur']:.2f}", s["top_channel"]))
    d.table(["Segment", "Accts", "Active", "WW revenue", "Vol/yr", "Qty/style", "Sustain.", "Price sens.", "Sample days",
             "Target FOB", "Top channel"], rows, widths=[3.6, 1.1, 1.1, 2.1, 1.3, 1.4, 1.3, 1.3, 1.3, 1.5, 1.6], size=7.5)
    d.table(["Segment", "What they want", "How we win", "Role in campaign"], [
        ("A. Sustainable mid-market brands", "Traceability, certificates, stable long-term partner",
         "Farm-to-finish data sheet per style; compliance pack; DPP-ready data", "PRIMARY: lead message, sales deck (P5)"),
        ("B. Volume value retailers", "Price point, on-time delivery, low admin",
         "Firm 24-h costing (UC15), on-time tracking, value-engineering", "Defend: most exposed to rival's 10% price cut"),
        ("C. Agile DTC & emerging labels", "Speed, small MOQs, many styles, digital service",
         "10-day sampling via 3D previews; small-batch MOQs; self-serve portal (UC17)", "Grow: social + search, lower ticket"),
    ], widths=[3.8, 4.2, 5.2, 3.8], size=8.5, first_col_bold=True)

    d.h1("3. Two personas")
    for ps in PERSONAS:
        d.h3(f"{ps['name']}  |  Segment {ps['seg']}")
        d.table(["", ""], [("Profile", ps["role"]), ("Goals", "\n".join("- " + g for g in ps["goals"])),
                           ("Pains", "\n".join("- " + g for g in ps["pains"])), ("Buying triggers", ps["triggers"]),
                           ("Channels", ps["channels"]), ("Message that lands", ps["message"]),
                           ("In her words (illustrative)", f"\"{ps['quote']}\"")],
                widths=[3.6, 13.4], show_header=False, size=8.5, first_col_bold=True)
    d.p("Personas are composites built from the synthetic segment profiles. Names are fictional; no real people.",
        size=8.5, italic=True)

    d.h1("4. Campaign concept: 'Traceable by Design'")
    d.table(["Element", "Plan"], [
        ("Big idea", "EU buyers are being asked to prove where their clothes come from. WeaveWorks makes that proof part of the product: "
                     "every style ships with a farm-to-finish data sheet and certificates in one place."),
        ("Tagline", "Knit in Tiruppur. Traceable to the fibre."),
        ("Objective (Q1-Q2 2027)", "40 qualified sample requests and 12 RFQs from new EU brands; re-activate 8 lapsed accounts (targets assumed)."),
        ("Audience", "Primary Segment A (Katrin); adapted proof points for B (Marta: price certainty) and C (speed)."),
        ("Lead magnet", "'EU Knitwear Compliance Checklist' (12 questions) - gated PDF on the website."),
        ("Channels", "LinkedIn sponsored content and InMail (A), Google Search (all), Instagram/Meta (C), 3-email automation, "
                     "trade-fair follow-up (Munich Fabric Start), account-manager calls."),
        ("Shock E link", "Rival KnitSphere now offers AI personalisation and 10% lower prices. We do not compete on price alone: "
                         "we show traceability, 10-day sampling and firm costing, the things a cheaper quote does not include."),
        ("Budget (inside UC03)", "Rs 1.4 lakh set-up + Rs 0.08 lakh/month tools; paid media from the existing marketing budget (assumed)."),
        ("KPIs", "Cost per sample request; sample-to-RFQ rate; RFQ win rate; email click rate (target 4%, assumed); unsubscribe < 0.5%."),
        ("Compliance", "Email only to opted-in contacts (GDPR for EU recipients, DPDP for Indian data processing); unsubscribe in every email; "
                       "no unverifiable claims such as 'guaranteed compliant'."),
    ], widths=[3.6, 13.4], size=8.5, first_col_bold=True)

    d.h1("5. Three creatives (Canva)")
    d.p("The PNGs below are exact-size **mock-ups** rendered from HTML so the layout is fixed. The CMO rebuilds each in Canva "
        "(templates: 'LinkedIn Post' 1200x627 and 'Instagram Post' 1080x1080) using the spec, then exports PNG. "
        "Brand: Indigo #1F2A5A, Teal #0F7C78, Marigold #E9A23B, Cotton #F6F4EF; font Inter (Canva) or Arial.")
    specs = [
        ("C1_LinkedIn_Traceable_by_Design.png", "C1 - LinkedIn 1200x627 - Segment A",
         "Headline 'Knit in Tiruppur. Traceable to the fibre.' Sub: farm-to-finish data. Badges: GOTS / OEKO-TEX / DPP-ready data. CTA: Get the traceability sample pack."),
        ("C2_Instagram_10_Day_Sampling.png", "C2 - Instagram 1080x1080 - Segment C",
         "Hero number '10 days', from tech-pack to approved sample; 21 -> 10 days comparison; AI-assisted 3D previews line."),
        ("C3_LinkedIn_Value_Programmes.png", "C3 - LinkedIn 1200x627 - Segment B",
         "'Price you can plan on. Quality you never chase.' Three proof points; quote card '24h from tech-pack to firm price'."),
    ]
    for fn, title, spec in specs:
        d.h3(title)
        d.image(CRE / fn, width_cm=11.5 if "Instagram" not in fn else 8.5)
        d.p(spec, size=9)
    d.callout("Claims check: '10 days' and '24 h' are TARGETS that depend on UC02 and UC15 going live (roadmap M7 and M4). "
              "Do not run C2 or C3 publicly until the service level is proven for 8 consecutive weeks.", kind="warn",
              title="Before publishing")

    d.h1("6. Three-email automated sequence")
    d.p("Tool: any marketing-automation platform with a free tier (e.g. Brevo, Mailchimp or HubSpot). Audience: opted-in "
        "contacts tagged Segment A or C. Exit rule: contact books a call or replies. Personalisation tokens: {{first_name}}, {{sender}}.")
    for e in EMAILS:
        d.h3(f"Email {e['n']}  |  {e['day']}")
        d.table(["", ""], [("Trigger / rule", e["trig"]), ("Subject", e["subj"]), ("Preview text", e["prev"]),
                           ("Body", "\n".join(e["body"])), ("Button", e["cta"]),
                           ("Footer", "WeaveWorks Exports, Tiruppur, India - You receive this because you opted in. Unsubscribe | Privacy notice")],
                widths=[3.0, 14.0], show_header=False, size=8.5, first_col_bold=True)

    d.h1("7. Ad copy for three platforms (character limits checked)")
    rows = [(p, f, t, f"{n}/{l}" if l else str(n)) for p, f, t, n, l in check_ads()]
    d.table(["Platform", "Field", "Copy", "Chars"], rows, widths=[3.0, 3.4, 9.2, 1.4], size=8)
    d.p("Character counts are checked automatically by src/build_p1.py; the build fails if any line is over its limit.",
        size=8.5, italic=True)

    d.h1("8. Prompt log: prompt, raw output, edited version")
    d.prompt_log(MKT_LOG)
    d.callout("Persona raw output (Log 3) invented a statistic: '73% of German buyers will switch suppliers by 2026 because of "
              "CSDDD'. It had no source, so we deleted it. Email raw output (Log 4) promised 'guaranteed CSDDD compliance', a legal "
              "claim we cannot make. We replaced it with 'data your compliance team can use'.",
              kind="ai", title="Rejected AI output (for the viva)")
    d.sources(["CSDDD", "ESPR", "GDPR", "DPDP"], extra=[
        "LinkedIn, Google Ads and Meta character limits: platform help pages (checked Nov 2026; limits change, re-check before launch)."])
    return d.save(OUT / "Marketing_Pack.docx")


MKT_LOG = [
    dict(title="Design the synthetic buyer dataset",
         prompt="You are a B2B marketing analyst for an Indian knitwear exporter selling to EU brands. Propose a schema for a synthetic "
                "buyer database of 300 accounts that would let us segment buyers. For each field give type, realistic range and why it matters. "
                "No real company names.",
         raw="Proposed 14 fields incl. 'company_name' (e.g. 'H&M-like retailer'), annual revenue, number of stores, 'decision maker age', "
             "Instagram followers, sustainability score, price sensitivity, payment terms...",
         edited="Kept 11 fields, added 'expected sample turnaround', 'certifications required', 'email opt-in', 'relationship status'. "
                "Replaced company names with IDs (WW-B0001...).",
         why="Real-brand look-alikes and a decision-maker's age are personal/irrelevant data. Added fields that link to our operational "
             "promises (sampling speed) and to GDPR consent."),
    dict(title="Interpret the k-means clusters",
         prompt="Here are the mean standardised values of 8 features for 3 clusters [table pasted]. Name each cluster in 3-5 words, describe it "
                "for a sales team and suggest one offer per cluster. Flag anything the numbers do not support.",
         raw="Cluster 0 'Eco-Premium Brands' ... Cluster 1 'Price Warriors' ... Cluster 2 'Gen-Z Fast Fashion' - target with TikTok influencers.",
         edited="Renamed to neutral, behaviour-based names (Sustainable mid-market brands / Volume value retailers / Agile DTC & emerging labels). "
                "Dropped the TikTok idea.",
         why="'Price Warriors' is disrespectful in a customer-facing deck. Nothing in the data shows age or TikTok usage, so that "
             "recommendation was not supported."),
    dict(title="Write two B2B personas",
         prompt="Using segments A and B [profiles pasted], write two B2B buyer personas: role, goals, pains, triggers, channels, message. "
                "Max 120 words each. Do not invent statistics.",
         raw="Katrin, 38, Head of Sourcing... 'Research shows 73% of German buyers will switch suppliers by 2026 because of CSDDD.' "
             "Marta, 44, ... 'As a mother of two she values kidswear safety...'",
         edited="Removed the 73% statistic and the family detail; added buying triggers tied to our shock (rival's 10% price cut).",
         why="The statistic had no source (ground rule: cite or label assumed). Family status is a stereotype and irrelevant to a B2B purchase."),
    dict(title="Draft the 3-email nurture sequence",
         prompt="Write a 3-email automated sequence for EU sourcing managers who downloaded our compliance checklist. Day 0, 4, 10. Subject, "
                "preview, body under 120 words, one CTA each. Tone: practical, no hype.",
         raw="Email 3 subject: 'Become CSDDD-compliant in 10 days - guaranteed'. Body: '...our AI makes your supply chain 100% compliant...'",
         edited="Subject 'A 10-day sample, on us'. Body promises a 3D preview + sample + traceability sheet; no compliance guarantee.",
         why="A supplier cannot make a buyer 'compliant'; the claim could be misleading advertising and creates legal exposure. "
             "We promise only what we control."),
    dict(title="Ad copy for LinkedIn, Google, Meta",
         prompt="Write ad copy for (1) LinkedIn sponsored content, (2) a Google responsive search ad with 5 headlines and 2 descriptions, "
                "(3) an Instagram feed ad, for the 'Traceable by Design' campaign. Respect each platform's character limits.",
         raw="Google headline: 'Sustainable Organic Knitwear Manufacturer India' (47 chars) ... Instagram: 'Best quality, cheapest prices in India!'",
         edited="All Google headlines cut to 30 characters or fewer and checked automatically; 'best' and 'cheapest' removed.",
         why="The raw output broke Google's 30-character headline limit. Superlatives cannot be substantiated and contradict our "
             "'value, not price' response to the rival."),
]

if __name__ == "__main__":
    print(build())
