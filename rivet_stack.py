import json
import math
from pathlib import Path

# =====================================================
# CONFIG
# =====================================================
ASSEMBLY_JSON = Path(r"E:\Phase 1\extractions\assembly_extraction.json")
AXIS_JSON     = Path(r"E:\Phase 1\extractions\geometry_fastener_axes.json")
OUTPUT_JSON   = Path(r"E:\Phase 1\extractions\inferred_rivet_stacks.json")

AXIS_TOLERANCE = 3.0   # mm (safe for sheet metal)

# =====================================================
# VECTOR UTILS
# =====================================================
def dot(a, b):
    return sum(x*y for x, y in zip(a, b))

def norm(v):
    return math.sqrt(dot(v, v))

def sub(a, b):
    return [x - y for x, y in zip(a, b)]

def scale(v, s):
    return [x * s for x in v]

# =====================================================
# LOAD DATA
# =====================================================
assembly = json.loads(ASSEMBLY_JSON.read_text())
axes     = json.loads(AXIS_JSON.read_text())

occurrences = assembly["occurrences"]

# Plate candidates (non-fasteners)
plates = {
    occ["name"]: occ
    for occ in occurrences
    if occ["document_type"] == "Part"
    and "RIVET" not in (occ.get("description") or "").upper()
}

# Plate centroids ≈ transform origins
plate_positions = {
    name: [
        occ.get("transform", {}).get("translation", [0,0,0])[0]
        if "transform" in occ else 0,
        occ.get("transform", {}).get("translation", [0,0,0])[1]
        if "transform" in occ else 0,
        occ.get("transform", {}).get("translation", [0,0,0])[2]
        if "transform" in occ else 0,
    ]
    for name, occ in plates.items()
}

# =====================================================
# STACK INFERENCE
# =====================================================
results = []

for rivet in axes:
    O = rivet["origin"]
    D = rivet["direction"]
    D_norm = norm(D)
    if D_norm == 0:
        continue
    D = scale(D, 1.0 / D_norm)

    pierced = []

    for plate_name, P in plate_positions.items():
        V = sub(P, O)
        t = dot(V, D)
        closest = scale(D, t)
        d = norm(sub(V, closest))

        if d < AXIS_TOLERANCE:
            pierced.append((plate_name, t))

    if not pierced:
        continue

    pierced.sort(key=lambda x: x[1])

    results.append({
        "fastener": rivet["occurrence"],
        "plates": [p[0] for p in pierced],
        "stack_size": len(pierced),
        "stack_type": "blind_rivet" if len(pierced) >= 2 else "single_plate",
        "confidence": round(min(0.9 + 0.05 * len(pierced), 0.99), 2)
    })

# =====================================================
# SAVE
# =====================================================
OUTPUT_JSON.write_text(json.dumps(results, indent=4))
print(f"✅ Phase-5 complete → {OUTPUT_JSON}")
print(f"   Inferred stacks: {len(results)}")