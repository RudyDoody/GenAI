"""P1 - three campaign creatives as HTML mock-ups, rendered to PNG with Playwright.
The CMO rebuilds them in Canva from the spec in the Marketing Pack (sizes, hex colours, fonts, copy)."""
import os
import subprocess
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "02_Build" / "P1_Marketing" / "creatives"
OUT.mkdir(parents=True, exist_ok=True)

INDIGO, TEAL, MARIGOLD, COTTON, INK = "#1F2A5A", "#0F7C78", "#E9A23B", "#F6F4EF", "#151515"

BASE = f"""<style>
*{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:'Inter',Arial,sans-serif;}}
.knit{{background-image:repeating-linear-gradient(45deg,rgba(255,255,255,.07) 0 6px,transparent 6px 18px),
repeating-linear-gradient(-45deg,rgba(255,255,255,.07) 0 6px,transparent 6px 18px);}}
.tag{{display:inline-block;font-size:20px;letter-spacing:.12em;text-transform:uppercase;font-weight:700}}
</style>"""

C1 = BASE + f"""<div class="knit" style="width:1200px;height:627px;background-color:{INDIGO};color:#fff;padding:64px 72px;position:relative">
<div class="tag" style="color:{MARIGOLD}">WeaveWorks Exports - Tiruppur</div>
<h1 style="font-size:76px;line-height:1.02;font-weight:800;margin-top:28px;max-width:820px">Knit in Tiruppur.<br>Traceable to the fibre.</h1>
<p style="font-size:28px;margin-top:28px;max-width:720px;color:#dfe3f1">Organic-cotton knitwear with farm-to-finish data your EU compliance team can actually use.</p>
<div style="position:absolute;left:72px;bottom:56px;display:flex;gap:14px">
<span style="background:{MARIGOLD};color:{INK};font-weight:700;font-size:24px;padding:14px 26px;border-radius:40px">Get the traceability sample pack</span></div>
<div style="position:absolute;right:72px;bottom:56px;display:flex;gap:10px">
{''.join(f'<span style="border:2px solid rgba(255,255,255,.55);border-radius:40px;padding:10px 18px;font-size:19px">{b}</span>' for b in ['GOTS','OEKO-TEX','DPP-ready data'])}</div>
<div style="position:absolute;right:72px;top:64px;width:170px;height:170px;border-radius:50%;border:10px solid {TEAL};"></div>
<div style="position:absolute;right:112px;top:104px;width:90px;height:90px;border-radius:50%;background:{MARIGOLD}"></div>
</div>"""

C2 = BASE + f"""<div style="width:1080px;height:1080px;background:{COTTON};color:{INK};padding:80px;position:relative;overflow:hidden">
<div class="tag" style="color:{TEAL}">For fast-moving labels</div>
<div style="font-size:250px;font-weight:800;line-height:.9;margin-top:40px;color:{INDIGO}">10<span style="font-size:90px"> days</span></div>
<div style="font-size:60px;font-weight:800;margin-top:10px;line-height:1.05">from tech-pack to<br>approved sample.</div>
<div style="display:flex;align-items:center;gap:22px;margin-top:56px;font-size:30px">
<span style="text-decoration:line-through;color:#8a8984">21 days</span><span style="font-size:34px">&#8594;</span>
<span style="background:{TEAL};color:#fff;padding:10px 22px;border-radius:12px;font-weight:700">10 days</span></div>
<p style="font-size:28px;margin-top:40px;max-width:760px;color:#3b3b3b">AI-assisted 3D previews first, so the physical sample you receive is the one you approve.</p>
<div style="position:absolute;bottom:70px;left:80px;font-size:26px;font-weight:700;color:{INDIGO}">weaveworks.example/studio</div>
<div style="position:absolute;right:-120px;bottom:-120px;width:460px;height:460px;border-radius:50%;background:{MARIGOLD};opacity:.9"></div>
<div style="position:absolute;right:-30px;bottom:-30px;width:280px;height:280px;border-radius:50%;background:{INDIGO}"></div>
</div>"""

C3 = BASE + f"""<div class="knit" style="width:1200px;height:627px;background-color:{TEAL};color:#fff;padding:60px 72px;position:relative">
<div class="tag" style="color:#fff;opacity:.85">Value programmes - kids & basics</div>
<h1 style="font-size:54px;line-height:1.08;font-weight:800;margin-top:22px;max-width:720px">Price you can plan on.<br>Quality you never chase.</h1>
<ul style="list-style:none;margin-top:30px;font-size:23px;line-height:1.65;max-width:700px">
<li>&#10003;&nbsp; Firm costing in 24 hours, assumptions shown</li>
<li>&#10003;&nbsp; OEKO-TEX 100 certified, amfori BSCI audited</li>
<li>&#10003;&nbsp; On-time shipment tracked every week</li></ul>
<div style="position:absolute;right:72px;top:60px;bottom:60px;width:300px;background:#fff;color:{INK};border-radius:24px;padding:36px 30px">
<div style="font-size:20px;color:{TEAL};font-weight:700;letter-spacing:.08em">YOUR QUOTE</div>
<div style="font-size:22px;margin-top:22px;color:#52514e">Kids crew-neck tee<br>180 gsm, 100% cotton</div>
<div style="font-size:64px;font-weight:800;margin-top:22px;color:{INDIGO}">24h</div>
<div style="font-size:20px;color:#52514e">from tech-pack to firm price</div>
<div style="margin-top:40px;background:{MARIGOLD};color:{INK};font-weight:700;font-size:22px;padding:14px;border-radius:40px;text-align:center">Request a quote</div></div>
</div>"""

CREATIVES = [
    ("C1_LinkedIn_Traceable_by_Design.png", C1, 1200, 627),
    ("C2_Instagram_10_Day_Sampling.png", C2, 1080, 1080),
    ("C3_LinkedIn_Value_Programmes.png", C3, 1200, 627),
]


def render():
    js = ["const { chromium } = require('playwright');", "(async () => {", "const b = await chromium.launch();"]
    for fn, html, w, h in CREATIVES:
        hp = OUT / (fn.replace(".png", ".html"))
        hp.write_text("<!doctype html><meta charset='utf-8'>" + html)
        js += [f"{{ const p = await b.newPage({{viewport:{{width:{w},height:{h}}}, deviceScaleFactor:1}});",
               f"await p.goto('file://{hp}'); await p.screenshot({{path:'{OUT / fn}'}}); await p.close(); }}"]
    js += ["await b.close(); })();"]
    script = ROOT / "src" / "_render_creatives.js"
    script.write_text("\n".join(js))
    subprocess.run(["node", str(script)], check=True, env={**os.environ, "NODE_PATH": "/opt/node22/lib/node_modules"})
    for fn, *_ in CREATIVES:
        (OUT / fn.replace(".png", ".html")).unlink()
    script.unlink()
    return [OUT / c[0] for c in CREATIVES]


if __name__ == "__main__":
    print(render())
