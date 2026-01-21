import json
import math
from pathlib import Path

# =====================================================
# CONFIG
# =====================================================
ASM_JSON   = Path(r"E:\Phase 1\extractions\assembly_extraction.json")
GEOM_JSON  = Path(r"E:\Phase 1\extractions\geometry_hooks.json")
OUT_JSON   = Path(r"E:\Phase 1\extractions\blind_rivet_stacks.json")

# Geometry tolerances (mm)
AXIS_DISTANCE_TOL = 3.0      # radial distance to axis
MAX_STACK_LENGTH  = 20.0     # max plate stack depth

# =====================================================
# VECTOR MATH
# =====================================================
def dot(a, b):
    return sum(x*y for x, y in zip(a, b))

def sub(a, b):
    return [x - y for x, y in zip(a, b)]

def norm(v):
    return math.sqrt(dot(v, v))

def normalize(v):
    n = norm(v)
    return [x / n for x in v] if n > 0 else [0, 0, 0]

# =====================================================
# LOAD INPUTS
# =====================================================
assembly = json.loads(ASM_JSON.read_text())
geometry = json.loads(GEOM_JSON.read_text())["occurrences"]

occurrences = assembly["occurrences"]

# =====================================================
# CLASSIFY PARTS
# =====================================================
fasteners = {}
plates = {}

for occ in occurrences:
    name = occ["name"]
    desc = (occ.get("description") or "").upper()

    if "RIVET" in desc or "BLIND" in desc:
        fasteners[name] = occ
    else:
        plates[name] = occ

# =====================================================
# PHASE-5 INFERENCE
# =====================================================
results = []

for fname in fasteners:
    if fname not in geometry:
        continue

    f_origin = geometry[fname]["origin"]
    f_axis   = normalize(geometry[fname]["z_axis"])

    stack = []

    for pname in plates:
        if pname not in geometry:
            continue

        p_origin = geometry[pname]["origin"]
        v = sub(p_origin, f_origin)

        # Projection along axis
        t = dot(v, f_axis)

        # Reject behind rivet or too deep
        if t < 0 or t > MAX_STACK_LENGTH:
            continue

        # Radial distance to axis
        radial = sub(v, [t * d for d in f_axis])
        dist = norm(radial)

        if dist <= AXIS_DISTANCE_TOL:
            stack.append((t, pname))

    if not stack:
        continue

    # Sort plates along axis
    stack.sort(key=lambda x: x[0])

    results.append({
        "fastener": fname,
        "plates": [p for _, p in stack],
        "stack_size": len(stack),
        "stack_type": "blind_rivet",
        "confidence": 0.95
    })

# =====================================================
# SAVE OUTPUT
# =====================================================
OUT_JSON.parent.mkdir(parents=True, exist_ok=True)
OUT_JSON.write_text(json.dumps(results, indent=4))

print("✅ Phase-5 blind rivet inference complete")
print(f"→ {OUT_JSON}")
print(f"→ {len(results)} rivets inferred")