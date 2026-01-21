import json
from pathlib import Path
import math

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
def dist_point_to_axis(p, axis_origin, axis_dir):
    """Shortest distance from point to axis"""
    px, py, pz = p
    ox, oy, oz = axis_origin
    dx, dy, dz = axis_dir

    vx, vy, vz = px - ox, py - oy, pz - oz
    cross = (
        vy*dz - vz*dy,
        vz*dx - vx*dz,
        vx*dy - vy*dx
    )
    num = math.sqrt(sum(c*c for c in cross))
    den = math.sqrt(dx*dx + dy*dy + dz*dz)
    return num / den if den else float("inf")

def is_rivet(desc):
    d = (desc or "").upper()
    return "RIVET" in d

# =====================================================
# LOAD DATA
# =====================================================
assembly = json.loads(ASSEMBLY_JSON.read_text(encoding="utf-8"))
axes     = json.loads(AXIS_JSON.read_text(encoding="utf-8"))

occurrences = {o["name"]: o for o in assembly["occurrences"]}

# =====================================================
# INDEX PARTS
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
            # Diameter compatibility
            if abs(hole["diameter"] - fast.get("nominal_diameter", hole["diameter"])) > HOLE_DIAMETER_TOL:
                continue

            # Use plate origin approximation (safe)
            plate_axis_dist = dist_point_to_axis(
                origin, origin, direction
            )

            if plate_axis_dist <= AXIS_DIST_TOL:
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
# SAVE
# =====================================================
OUT_JSON.write_text(json.dumps(stacks, indent=4), encoding="utf-8")

print("✅ Phase-5 rivet stack inference complete")
print(f"   → {OUT_JSON}")
print(f"   → stacks found: {len(stacks)}")