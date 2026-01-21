import json
import math
from pathlib import Path

# =====================================================
# CONFIG
# =====================================================
ASSEMBLY_JSON = Path(r"E:\Phase 1\extractions\assembly_extraction.json")
AXIS_JSON     = Path(r"E:\Phase 1\extractions\geometry_fastener_axes.json")
OUT_JSON      = Path(r"E:\Phase 1\extractions\blind_rivet_stacks.json")

# ---- Tunable parameters ----
AXIS_DISTANCE_TOL = 8.0     # mm (axis proximity)
MAX_STACK_LENGTH  = 12.0    # mm (blind rivet allowance)
MIN_STACK_PLATES  = 1

# =====================================================
# VECTOR HELPERS
# =====================================================
def dot(a, b):
    return a[0]*b[0] + a[1]*b[1] + a[2]*b[2]

def sub(a, b):
    return [a[0]-b[0], a[1]-b[1], a[2]-b[2]]

def norm(v):
    return math.sqrt(dot(v, v))

def normalize(v):
    n = norm(v)
    return [v[0]/n, v[1]/n, v[2]/n] if n else [0,0,0]

def dist_point_to_axis(p, o, d):
    v = sub(p, o)
    proj = dot(v, d)
    closest = [o[i] + proj*d[i] for i in range(3)]
    return norm(sub(p, closest)), proj

# =====================================================
# LOAD DATA
# =====================================================
with open(ASSEMBLY_JSON, "r", encoding="utf-8") as f:
    asm = json.load(f)

with open(AXIS_JSON, "r", encoding="utf-8") as f:
    axes = json.load(f)

occurrences = asm["occurrences"]

# -----------------------------------------------------
# Classify parts
# -----------------------------------------------------
plates = {}
fasteners = {}

for occ in occurrences:
    name = occ["name"]
    desc = (occ.get("description") or "").upper()

    if "RIVET" in desc:
        fasteners[name] = occ
    else:
        plates[name] = occ

# =====================================================
# PHASE-5 STACK INFERENCE
# =====================================================
results = []

for fast_name, axis_data in axes.items():

    origin = axis_data["origin"]
    direction = normalize(axis_data["direction"])

    candidates = []

    for plate_name, plate in plates.items():

        plate_origin = plate.get("origin")
        if not plate_origin:
            continue

        dist, proj = dist_point_to_axis(
            plate_origin, origin, direction
        )

        if dist > AXIS_DISTANCE_TOL:
            continue

        if 0 < proj < MAX_STACK_LENGTH:
            candidates.append((proj, plate_name))

    if len(candidates) < MIN_STACK_PLATES:
        continue

    candidates.sort(key=lambda x: x[0])

    stack_plates = [p for _, p in candidates]

    confidence = min(0.95, 0.6 + 0.15 * len(stack_plates))

    results.append({
        "fastener": fast_name,
        "plates": stack_plates,
        "stack_size": len(stack_plates),
        "stack_type": "blind_rivet",
        "confidence": round(confidence, 3)
    })

# =====================================================
# SAVE OUTPUT
# =====================================================
with open(OUT_JSON, "w", encoding="utf-8") as f:
    json.dump(results, f, indent=4)

print("✅ Phase-5 complete")
print(f"   → {OUT_JSON}")
print(f"   → stacks inferred: {len(results)}")