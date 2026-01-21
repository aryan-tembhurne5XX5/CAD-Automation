import json
import math
from pathlib import Path

# =====================================================
# CONFIG
# =====================================================
ASM_JSON = Path(r"E:\Phase 1\extractions\assembly_extraction.json")
AXIS_JSON = Path(r"E:\Phase 1\extractions\geometry_fastener_axes.json")
OUT_JSON  = Path(r"E:\Phase 1\extractions\blind_rivet_stacks.json")

# Geometry thresholds (mm)
MAX_RIVET_GRIP = 20.0
HOLE_RADIUS_THRESHOLD = 3.0

# =====================================================
# VECTOR HELPERS
# =====================================================
def dot(a, b):
    return sum(x*y for x, y in zip(a, b))

def sub(a, b):
    return [x-y for x, y in zip(a, b)]

def norm(v):
    return math.sqrt(dot(v, v))

def normalize(v):
    n = norm(v)
    return [x/n for x in v] if n else v

def distance(a, b):
    return norm(sub(a, b))

# =====================================================
# LOAD DATA
# =====================================================
with open(ASM_JSON, "r", encoding="utf-8") as f:
    asm = json.load(f)

with open(AXIS_JSON, "r", encoding="utf-8") as f:
    axes = json.load(f)

occ_map = {
    o["name"]: o
    for o in asm["occurrences"]
}

# =====================================================
# PLATE FILTER
# =====================================================
def is_plate(occ):
    desc = (occ.get("description") or "").upper()
    return not any(k in desc for k in ["RIVET", "NUT", "SCREW", "FASTENER"])

# =====================================================
# PHASE-5 STACK INFERENCE
# =====================================================
stacks = []

for fastener in axes:
    f_name = fastener["occurrence"]
    axis_o = fastener["origin"]
    axis_d = normalize(fastener["direction"])

    plates = []

    for occ in asm["occurrences"]:
        if not is_plate(occ):
            continue

        occ_name = occ["name"]
        if occ_name == f_name:
            continue

        # Use occurrence transform origin if present
        if "origin" not in occ:
            continue

        p = occ["origin"]
        v = sub(p, axis_o)
        t = dot(v, axis_d)

        if t < 0 or t > MAX_RIVET_GRIP:
            continue

        radial = norm(sub(v, [t*x for x in axis_d]))
        if radial > HOLE_RADIUS_THRESHOLD:
            continue

        plates.append((occ_name, t))

    if len(plates) < 2:
        continue

    plates.sort(key=lambda x: x[1])

    stacks.append({
        "fastener": f_name,
        "plates": [p[0] for p in plates],
        "stack_size": len(plates),
        "stack_type": "blind_rivet",
        "confidence": round(min(0.95, 0.7 + 0.05 * len(plates)), 2)
    })

# =====================================================
# SAVE
# =====================================================
with open(OUT_JSON, "w", encoding="utf-8") as f:
    json.dump(stacks, f, indent=4)

print("✅ Phase-5 blind rivet stacks inferred correctly")
print(f"   → {OUT_JSON}")