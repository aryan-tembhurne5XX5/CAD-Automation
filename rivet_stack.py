import json
import math
from pathlib import Path

# =====================================================
# CONFIG
# =====================================================
ASSEMBLY_JSON = Path(r"E:\Phase 1\extractions\assembly_extraction.json")
AXIS_JSON     = Path(r"E:\Phase 1\extractions\geometry_fastener_axes.json")
OUT_JSON      = Path(r"E:\Phase 1\extractions\rivet_stacks.json")

HOLE_DIAMETER_TOL = 0.3   # mm
AXIS_DIST_TOL     = 2.0   # mm
MIN_STACK_SIZE    = 1

# =====================================================
# HELPERS
# =====================================================
def dist_point_to_axis(p, o, d):
    px, py, pz = p
    ox, oy, oz = o
    dx, dy, dz = d

    vx, vy, vz = px - ox, py - oy, pz - oz
    cx = vy*dz - vz*dy
    cy = vz*dx - vx*dz
    cz = vx*dy - vy*dx

    num = math.sqrt(cx*cx + cy*cy + cz*cz)
    den = math.sqrt(dx*dx + dy*dy + dz*dz)
    return num / den if den else float("inf")

def is_rivet(desc):
    return "RIVET" in (desc or "").upper()

# =====================================================
# LOAD DATA
# =====================================================
assembly = json.loads(ASSEMBLY_JSON.read_text(encoding="utf-8"))
axes_raw = json.loads(AXIS_JSON.read_text(encoding="utf-8"))

# =====================================================
# INDEX AXES (FIX)
# =====================================================
axes = {
    a["occurrence"]: a
    for a in axes_raw
}

# =====================================================
# CLASSIFY PARTS
# =====================================================
plates = []
fasteners = []

for occ in assembly["occurrences"]:
    if occ["document_type"] != "Part":
        continue

    if is_rivet(occ.get("description")):
        fasteners.append(occ)
    elif occ.get("hole_count", 0) > 0:
        plates.append(occ)

# =====================================================
# RIVET STACK INFERENCE
# =====================================================
stacks = []

for fast in fasteners:
    fname = fast["name"]
    axis = axes.get(fname)
    if not axis:
        continue

    origin = axis["origin"]
    direction = axis["direction"]

    matched_plates = []

    for plate in plates:
        for hole in plate.get("holes", []):
            # diameter compatibility
            if abs(hole["diameter"] - hole["diameter"]) > HOLE_DIAMETER_TOL:
                continue

            dist = dist_point_to_axis(origin, origin, direction)
            if dist <= AXIS_DIST_TOL:
                matched_plates.append(plate["name"])
                break

    if len(matched_plates) >= MIN_STACK_SIZE:
        stacks.append({
            "fastener": fname,
            "plates": matched_plates,
            "stack_size": len(matched_plates),
            "stack_type": "blind_rivet",
            "confidence": round(0.85 + 0.05 * min(len(matched_plates), 3), 2)
        })

# =====================================================
# SAVE OUTPUT
# =====================================================
OUT_JSON.write_text(json.dumps(stacks, indent=4), encoding="utf-8")

print("✅ Phase-5 rivet stack inference complete")
print(f"   → {OUT_JSON}")
print(f"   → stacks found: {len(stacks)}")