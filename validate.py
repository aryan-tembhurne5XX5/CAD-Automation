import json
from pathlib import Path
from collections import defaultdict

# =====================================================
# CONFIG
# =====================================================
GROUPED_HOLES_JSON = Path(r"E:\Phase 1\extractions\grouped_holes.json")
RULES_JSON         = Path(r"E:\Phase 1\extractions\rules.json")
BOM_JSON           = Path(r"E:\Phase 1\extractions\bom.json")  # optional
OUTPUT_JSON        = Path(r"E:\Phase 1\extractions\validation.json")

# =====================================================
# LOAD DATA
# =====================================================
with open(GROUPED_HOLES_JSON, "r", encoding="utf-8") as f:
    grouped_holes = json.load(f)

with open(RULES_JSON, "r", encoding="utf-8") as f:
    rules = json.load(f)

bom = {}
if BOM_JSON.exists():
    with open(BOM_JSON, "r", encoding="utf-8") as f:
        bom = json.load(f)

# =====================================================
# BUILD RULE LOOKUP
# =====================================================
rule_confidence = {}

for r in rules:
    if r["constraint_type"] == "Insert" and r["mandatory"]:
        key = (r["source_part_type"], r["target_part_type"])
        rule_confidence[key] = r["confidence"]

# =====================================================
# PHASE-4 VALIDATION / COMPLETION
# =====================================================
results = []

for entry in grouped_holes:
    plate = entry["plate"]
    fastener = entry["fastener_type"]
    expected = entry["hole_count"]
    present = expected

    # If BOM exists, validate against BOM
    if fastener in bom:
        present = bom[fastener]

    missing = max(0, expected - present)

    confidence = entry["confidence"]
    status = "OK" if missing == 0 else "INCOMPLETE"

    results.append({
        "plate": plate,
        "fastener_type": fastener,
        "expected_count": expected,
        "present_count": present,
        "missing": missing,
        "confidence": confidence,
        "status": status
    })

# =====================================================
# SAVE OUTPUT
# =====================================================
with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
    json.dump(results, f, indent=4)

print("✅ Phase-4 complete")
print(f"   Output → {OUTPUT_JSON}")