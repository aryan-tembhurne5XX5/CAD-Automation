import json
from collections import defaultdict
from pathlib import Path

# =====================================================
# CONFIG
# =====================================================
INPUT_JSON  = Path(r"E:\Phase 1\extractions\assembly_extraction.json")
OUTPUT_JSON = Path(r"E:\Phase 1\extractions\inferred_holes.json")

# =====================================================
# LOAD PHASE-1 DATA
# =====================================================
with open(INPUT_JSON, "r", encoding="utf-8") as f:
    data = json.load(f)

occurrences = {o["name"]: o for o in data["occurrences"]}
constraints = data["constraints"]

# =====================================================
# FASTENER IDENTIFICATION (SAME LOGIC AS PHASE-2)
# =====================================================
def is_fastener(occ):
    desc = (occ.get("description") or "").upper()
    return any(k in desc for k in ["RIVET", "BOLT", "SCREW", "PIN", "NUT"])

fasteners = {
    name for name, occ in occurrences.items()
    if is_fastener(occ)
}

# =====================================================
# BUILD INSERT-CONSTRAINT GRAPH
# =====================================================
graph = defaultdict(set)

for c in constraints:
    if c["constraint_type"] != "Insert":
        continue

    o1 = c.get("occurrence_1")
    o2 = c.get("occurrence_2")

    if not o1 or not o2:
        continue

    graph[o1].add(o2)
    graph[o2].add(o1)

# =====================================================
# INFER HOLE STACKS
# =====================================================
hole_results = []

for fastener in fasteners:
    connected = graph.get(fastener, set())

    # A hole stack must pass through >= 2 parts
    plates = [
        o for o in connected
        if o in occurrences and not is_fastener(occurrences[o])
    ]

    if len(plates) >= 2:
        hole_results.append({
            "fastener": fastener,
            "hole_stack": plates,
            "confidence": round(0.8 + 0.05 * len(plates), 2)
        })

# =====================================================
# SAVE OUTPUT
# =====================================================
with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
    json.dump(hole_results, f, indent=4)

print(f"✅ Phase-3 complete")
print(f"   Inferred hole stacks: {len(hole_results)}")
print(f"   Output → {OUTPUT_JSON}")