import json
from collections import defaultdict
from pathlib import Path

# =====================================================
# CONFIG
# =====================================================
INPUT_JSON  = Path(r"E:\Phase 1\extractions\assembly_extraction.json")
OUTPUT_JSON = Path(r"E:\Phase 1\extractions\inferred_holes.json")

# =====================================================
# LOAD DATA
# =====================================================
with open(INPUT_JSON, "r", encoding="utf-8") as f:
    data = json.load(f)

occurrences = {o["name"]: o for o in data["occurrences"]}
constraints = data["constraints"]

# =====================================================
# FASTENER IDENTIFICATION
# =====================================================
def is_fastener(occ):
    desc = (occ.get("description") or "").upper()
    return any(k in desc for k in ["RIVET", "BOLT", "SCREW", "PIN", "NUT"])

fasteners = {
    name for name, occ in occurrences.items()
    if is_fastener(occ)
}

# =====================================================
# BUILD INSERT GRAPH
# =====================================================
insert_graph = defaultdict(set)

for c in constraints:
    if c["constraint_type"] != "Insert":
        continue

    o1 = c.get("occurrence_1")
    o2 = c.get("occurrence_2")

    if not o1 or not o2:
        continue

    insert_graph[o1].add(o2)
    insert_graph[o2].add(o1)

# =====================================================
# INFER HOLES (SINGLE OR MULTI PLATE)
# =====================================================
holes = []

for fastener in fasteners:
    connected = insert_graph.get(fastener, set())

    if not connected:
        continue

    plates = [
        o for o in connected
        if o in occurrences and not is_fastener(occurrences[o])
    ]

    if len(plates) >= 1:   # 🔥 FIXED LOGIC
        holes.append({
            "fastener": fastener,
            "hole_stack": plates,
            "hole_type": "single_plate" if len(plates) == 1 else "multi_plate",
            "confidence": round(0.75 + 0.05 * len(plates), 2)
        })

# =====================================================
# SAVE OUTPUT
# =====================================================
with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
    json.dump(holes, f, indent=4)

print("✅ Phase-3 complete")
print(f"   Inferred holes: {len(holes)}")
print(f"   Output → {OUTPUT_JSON}")