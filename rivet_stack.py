import json
from pathlib import Path
from collections import defaultdict

# =====================================================
# CONFIG
# =====================================================
ASM_JSON   = Path(r"E:\Phase 1\extractions\assembly_extraction.json")
AXIS_JSON  = Path(r"E:\Phase 1\extractions\geometry_fastener_axes.json")
OUT_JSON   = Path(r"E:\Phase 1\extractions\rivet_stacks.json")

# =====================================================
# LOAD DATA
# =====================================================
assembly = json.loads(ASM_JSON.read_text(encoding="utf-8"))
axes_raw = json.loads(AXIS_JSON.read_text(encoding="utf-8"))

# =====================================================
# INDEX FASTENER AXES
# =====================================================
axes = {a["occurrence"]: a for a in axes_raw}

# =====================================================
# CLASSIFY PARTS
# =====================================================
occ_by_name = {o["name"]: o for o in assembly["occurrences"]}

def is_fastener(o):
    return "RIVET" in (o.get("description") or "").upper()

def is_plate(o):
    return o["document_type"] == "Part" and not is_fastener(o)

fasteners = {o["name"] for o in assembly["occurrences"] if is_fastener(o)}

# =====================================================
# BUILD FASTENER → PLATE MAP FROM INSERT CONSTRAINTS
# =====================================================
stack_map = defaultdict(set)

for c in assembly["constraints"]:
    if c["constraint_type"] != "Insert":
        continue

    a = c["occurrence_1"]
    b = c["occurrence_2"]

    if a in fasteners and b in occ_by_name and is_plate(occ_by_name[b]):
        stack_map[a].add(b)

    elif b in fasteners and a in occ_by_name and is_plate(occ_by_name[a]):
        stack_map[b].add(a)

# =====================================================
# BUILD STACK OUTPUT
# =====================================================
stacks = []

for fastener, plates in stack_map.items():
    if fastener not in axes:
        continue  # geometry missing → skip

    stacks.append({
        "fastener": fastener,
        "plates": sorted(plates),
        "stack_size": len(plates),
        "stack_type": "blind_rivet",
        "confidence": 0.95 if len(plates) >= 1 else 0.7
    })

# =====================================================
# SAVE
# =====================================================
OUT_JSON.write_text(json.dumps(stacks, indent=4), encoding="utf-8")

print("✅ Phase-5 rivet stack inference complete")
print(f"   → {OUT_JSON}")
print(f"   → stacks inferred: {len(stacks)}")