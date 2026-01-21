import json
import math
from pathlib import Path

# =====================================================
# CONFIG
# =====================================================
ASM_JSON   = Path(r"E:\Phase 1\extractions\assembly_extraction.json")
HOOKS_JSON = Path(r"E:\Phase 1\extractions\geometry_hooks.json")
AXIS_JSON  = Path(r"E:\Phase 1\extractions\geometry_fastener_axes.json")
OUT_JSON   = Path(r"E:\Phase 1\extractions\blind_rivet_stacks.json")

MAX_RIVET_GRIP = 20.0      # mm
HOLE_RADIUS    = 3.0       # mm

# =====================================================
# VECTOR MATH
# =====================================================
def dot(a, b): return sum(x*y for x, y in zip(a, b))
def sub(a, b): return [x-y for x, y in zip(a, b)]
def norm(v): return math.sqrt(dot(v, v))
def normalize(v):
    n = norm(v)
    return [x/n for x in v] if n else v

# =====================================================
# LOAD FILES
# =====================================================
asm   = json.loads(ASM_JSON.read_text())
hooks = json.loads(HOOKS_JSON.read_text())
axes  = json.loads(AXIS_JSON.read_text())

occ_info = {o["name"]: o for o in asm["occurrences"]}
occ_pos  = {
    name: data["translation"]
    for name, data in hooks["occurrence_transforms"].items()
}

# =====================================================
# HELPERS
# =====================================================
def is_plate(occ):
    desc = (occ.get("description") or "").upper()
    return not any(k in desc for k in ("RIVET", "NUT", "SCREW", "FASTENER"))

# =====================================================
# PHASE-5: STACK INFERENCE
# =====================================================
results = []

for fast in axes:
    f_name = fast["occurrence"]
    axis_o = fast["origin"]
    axis_d = normalize(fast["direction"])

    stack = []

    for occ_name, occ in occ_info.items():
        if occ_name == f_name:
            continue
        if not is_plate(occ):
            continue
        if occ_name not in occ_pos:
            continue

        p = occ_pos[occ_name]
        v = sub(p, axis_o)
        t = dot(v, axis_d)

        if t < 0 or t > MAX_RIVET_GRIP:
            continue

        radial = norm(sub(v, [t*x for x in axis_d]))
        if radial > HOLE_RADIUS:
            continue

        stack.append((occ_name, t))

    if len(stack) < 2:
        continue

    stack.sort(key=lambda x: x[1])

    results.append({
        "fastener": f_name,
        "plates": [s[0] for s in stack],
        "stack_size": len(stack),
        "stack_type": "blind_rivet",
        "confidence": 0.95
    })

# =====================================================
# SAVE
# =====================================================
OUT_JSON.write_text(json.dumps(results, indent=4))
print("✅ Phase-5 complete → blind rivet stacks inferred")
print(f"→ {OUT_JSON}")