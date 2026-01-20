import json
from collections import defaultdict
from pathlib import Path

# =====================================================
# CONFIG
# =====================================================
INPUT_JSON = Path("assembly_extraction.json")
OUT_NORMALIZED = Path("normalized_constraints.json")
OUT_RULES = Path("rules.json")

# =====================================================
# LOAD DATA
# =====================================================
with open(INPUT_JSON, "r", encoding="utf-8") as f:
    data = json.load(f)

occurrences = data["occurrences"]
constraints = data["constraints"]

# =====================================================
# PART CLASSIFICATION (DETERMINISTIC)
# =====================================================
part_type = {}

for occ in occurrences:
    desc = (occ.get("description") or "").upper()
    hole_count = occ.get("hole_count", 0)

    if "RIVET" in desc or "NUT" in desc or "SCREW" in desc:
        part_type[occ["name"]] = "Fastener"
    elif hole_count > 0:
        part_type[occ["name"]] = "Plate"
    else:
        part_type[occ["name"]] = "Structural"

# =====================================================
# CONSTRAINT NORMALIZATION
# =====================================================
def constraint_signature(c):
    return (
        c["constraint_type"],
        tuple(sorted([c["occurrence_1"], c["occurrence_2"]])),
        tuple(sorted([c["entity_1_type"], c["entity_2_type"]]))
    )

normalized = {}
for c in constraints:
    sig = constraint_signature(c)
    if sig not in normalized:
        normalized[sig] = c

normalized_constraints = list(normalized.values())

# =====================================================
# SAVE NORMALIZED CONSTRAINTS
# =====================================================
with open(OUT_NORMALIZED, "w", encoding="utf-8") as f:
    json.dump(normalized_constraints, f, indent=4)

# =====================================================
# RULE MINING
# =====================================================
rule_counter = defaultdict(int)
rule_examples = {}

for c in normalized_constraints:
    src = c["occurrence_1"]
    tgt = c["occurrence_2"]

    src_type = part_type.get(src, "Unknown")
    tgt_type = part_type.get(tgt, "Unknown")

    rule_key = (
        c["constraint_type"],
        tuple(sorted([c["entity_1_type"], c["entity_2_type"]])),
        src_type,
        tgt_type
    )

    rule_counter[rule_key] += 1
    rule_examples.setdefault(rule_key, c)

# =====================================================
# BUILD RULES
# =====================================================
rules = []
max_occurrence = max(rule_counter.values()) if rule_counter else 1

for i, (key, count) in enumerate(rule_counter.items(), start=1):
    constraint_type, entity_pair, src_type, tgt_type = key

    confidence = round(count / max_occurrence, 3)

    rule = {
        "rule_id": f"RULE_{i:03d}",
        "constraint_type": constraint_type,
        "entity_pair": list(entity_pair),
        "source_part_type": src_type,
        "target_part_type": tgt_type,
        "occurrences_seen": count,
        "confidence": confidence,
        "mandatory": confidence >= 0.9
    }

    rules.append(rule)

# =====================================================
# SAVE RULES
# =====================================================
with open(OUT_RULES, "w", encoding="utf-8") as f:
    json.dump(rules, f, indent=4)

print("✅ Phase-2 complete")
print(f"   → {OUT_NORMALIZED}")
print(f"   → {OUT_RULES}")