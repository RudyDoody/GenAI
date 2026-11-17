"""Dump facts.py to data/facts.json so the JavaScript deck builders use the same numbers."""
import json
from pathlib import Path
import facts

ROOT = Path(__file__).resolve().parents[1]
out = dict(CO=facts.CO, TEAM=facts.TEAM, SEGMENT_NAMES=facts.SEGMENT_NAMES, SEGMENT_REV_SHARE=facts.SEGMENT_REV_SHARE,
           BUDGET_LAKH=facts.BUDGET_LAKH, HIDDEN_COSTS=facts.HIDDEN_COSTS, CONTINGENCY_LAKH=facts.CONTINGENCY_LAKH,
           USE_CASES=facts.USE_CASES, ROI=facts.roi_rows(), BUDGET=facts.budget_summary(), SHOCK=facts.SHOCK,
           RIVAL=facts.RIVAL, TODAY=facts.TODAY)
(ROOT / "data" / "facts.json").write_text(json.dumps(out, indent=1, default=float))
print("ok")
