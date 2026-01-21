import json
import math
from pathlib import Path

# =====================================================
# CONFIG
# =====================================================
GEOM_JSON = Path(r"E:\Phase 1\extractions\geometry_hooks.json")
ASM_JSON  = Path(r"E:\Phase 1\extractions\assembly_extraction.json")
OUT_JSON  = Path(r"E:\Phase 1\extractions\blind_rivet_stacks.json")

PROXIMITY_TOL_MM = 3.0

# =====================================================
# HELPERS
# =====================================================
def dist(a, b):
    return math.sqrt(sum((a[i] - b[i]) ** 2 for i in range(3)))

# =====================================================
# LOAD DATA
# =====================================================
geometry = json.loads(GEOM_JSON.read_text())
assembly = json.loads(ASM_JSON.read_text())

occ_geom = geometry["occurrences"]
occ_info = {o["name"]: o for o in assembly["occurrences"]}
constraints = assembly["constraints"]

# =====================================================
# CLASSIFY PARTS
# =====================================================
fasteners = set()
plates = set()

for name, o in occ_info.items():
    desc = (o.get("description") or "").upper()
    if "RIVET" in desc or "NUT" in desc:
        fasteners.add(name)
    else:
        plates.add(name)

# =====================================================
# INSERT CONSTRAINT MAP
# =====================================================
insert_map = {}

for c in constraints:
    if c["constraint_type"] != "Insert":
        continue
    insert_map.setdefault(c["occurrence_1"], set()).add(c["occurrence_2"])
    insert_map.setdefault(c["occurrence_2"], set()).add(c["occurrence_1"])

# =====================================================
# STACK INFERENCE
# =====================================================
stacks = []

for f in fasteners:
    if f not in occ_geom:
        continue

    f_origin = occ_geom[f]["origin"]
    connected = insert_map.get(f, set())

    stack = []
    for p in connected:
        if p not in plates or p not in occ_geom:
            continue
        d = dist(f_origin, occ_geom[p]["origin"])
        if d <= PROXIMITY_TOL_MM:
            stack.append((p, d))

    if not stack:
        continue

    stack.sort(key=lambda x: x[1])
    plate_stack = [p for p, _ in stack]

    stack_type = (
        "single_plate" if len(plate_stack) == 1
        else "blind_rivet"
    )

    stacks.append({
        "fastener": f,
        "plates": plate_stack,
        "stack_size": len(plate_stack),
        "stack_type": stack_type,
        "confidence": round(min(1.0, 0.7 + 0.1 * len(plate_stack)), 2)
    })

# =====================================================
# SAVE
# =====================================================
OUT_JSON.write_text(json.dumps(stacks, indent=4))
print("✅ Phase-5 blind rivet stacks inferred")
print(f"→ {OUT_JSON}")